"""PDF Reader (PyMuPDF 기반).

PDF에는 "문단"이라는 개념이 없다. 글리프가 좌표에 찍혀 있을 뿐이다. 그래서
읽기는 항상 **재구성**이며, 아래 순서로 원래 구조를 복원한다.

1. `page.find_tables()` 로 표 영역을 먼저 확정하고 그 사각형을 기억한다.
2. `page.get_text("dict")` 로 텍스트 블록/라인/스팬을 얻는다.
3. 표 영역 안에 들어가는 텍스트 블록은 건너뛴다(중복 방지).
4. 남은 블록을 y좌표 순으로 정렬해 문단으로 만든다.
5. 본문 글자 크기의 중앙값을 구해, 그보다 뚜렷이 큰 문단을 제목으로 승격한다.

스캔 PDF(이미지만 있고 텍스트 레이어가 없는 문서)는 텍스트가 0자로 나온다.
이 경우 경고를 남긴다 - OCR은 이 프로그램의 범위 밖이다.
"""

from __future__ import annotations

import statistics
from pathlib import Path
from typing import Any, Iterable, Optional

from ..errors import CorruptFileError, EncryptedFileError, ReadError
from ..ir import (
    Align,
    Block,
    Cell,
    DocKind,
    DocMeta,
    Document,
    Image,
    PageBreak,
    Paragraph,
    Run,
    Table,
)

#: 이 비율 이상 크면 제목으로 승격한다.
HEADING_RATIO_H1 = 1.55
HEADING_RATIO_H2 = 1.28
HEADING_RATIO_H3 = 1.12

#: 표 영역과 텍스트 블록이 이만큼 겹치면 표에 속한 것으로 본다.
TABLE_OVERLAP = 0.55


class PdfReader:
    def __init__(
        self,
        path: Path | str,
        *,
        password: Optional[str] = None,
        extract_tables: bool = True,
        extract_images: bool = True,
        page_range: Optional[tuple[int, int]] = None,
    ) -> None:
        self.path = Path(path)
        self.password = password
        self.extract_tables = extract_tables
        self.extract_images = extract_images
        self.page_range = page_range
        self.warnings: list[str] = []

    def read(self) -> Document:
        try:
            import pymupdf as fitz
        except ImportError:
            try:
                import fitz  # 구버전 이름
            except ImportError as e:  # pragma: no cover
                raise ReadError("PyMuPDF가 필요합니다. pip install pymupdf") from e

        try:
            pdf = fitz.open(str(self.path))
        except Exception as e:
            raise CorruptFileError(
                "PDF를 열 수 없습니다.", detail=f"{self.path}: {e}"
            ) from e

        if pdf.needs_pass:
            if not self.password or not pdf.authenticate(self.password):
                pdf.close()
                raise EncryptedFileError(str(self.path))

        doc = Document(kind=DocKind.DOCUMENT, meta=self._meta(pdf))
        first = pdf[0] if pdf.page_count else None
        if first is not None:
            doc.page_width_pt = first.rect.width
            doc.page_height_pt = first.rect.height

        lo, hi = self.page_range or (0, pdf.page_count)
        lo = max(0, lo)
        hi = min(pdf.page_count, hi)

        total_chars = 0
        for pno in range(lo, hi):
            if pno > lo:
                doc.blocks.append(PageBreak())
            page = pdf[pno]
            blocks, n = self._page_blocks(page, fitz)
            total_chars += n
            doc.blocks.extend(blocks)

        if total_chars == 0 and pdf.page_count:
            self.warnings.append(
                "텍스트 레이어가 없는 PDF입니다(스캔 문서로 보입니다). "
                "글자를 추출하려면 OCR이 필요하며 이 프로그램은 지원하지 않습니다."
            )

        self._promote_headings(doc)
        doc.compact()
        pdf.close()
        if not doc.meta.title:
            doc.meta.title = _first_text(doc)
        return doc

    # -- 메타 -------------------------------------------------------------

    @staticmethod
    def _meta(pdf) -> DocMeta:
        m = pdf.metadata or {}
        return DocMeta(
            title=(m.get("title") or "").strip() or None,
            author=(m.get("author") or "").strip() or None,
            subject=(m.get("subject") or "").strip() or None,
            keywords=(m.get("keywords") or "").strip() or None,
            created=(m.get("creationDate") or "").strip() or None,
            modified=(m.get("modDate") or "").strip() or None,
            source_format="pdf",
            extra={"pages": pdf.page_count, "producer": m.get("producer") or ""},
        )

    # -- 페이지 -----------------------------------------------------------

    def _page_blocks(self, page, fitz) -> tuple[list[Block], int]:
        out: list[Block] = []
        table_rects: list[Any] = []

        if self.extract_tables:
            for tbl, rect in self._find_tables(page):
                out.append((_rect_tuple(rect)[1], tbl))
                table_rects.append(_rect_tuple(rect))

        text_items: list[tuple[float, Block]] = []
        n_chars = 0
        try:
            data = page.get_text("dict", sort=True)
        except Exception as e:
            self.warnings.append(f"{page.number + 1}쪽 텍스트 추출 실패: {e}")
            data = {"blocks": []}

        for blk in data.get("blocks", []):
            if blk.get("type") != 0:  # 0=텍스트, 1=이미지
                continue
            bbox = blk.get("bbox")
            if bbox and self._inside_table(bbox, table_rects):
                continue
            para = self._block_to_paragraph(blk)
            if para is not None and para.runs:
                n_chars += len(para.text)
                text_items.append((bbox[1] if bbox else 0.0, para))

        if self.extract_images:
            for y, img in self._page_images(page, fitz):
                text_items.append((y, img))

        merged = [(y, b) for y, b in out] + text_items
        merged.sort(key=lambda t: t[0])
        return [b for _, b in merged], n_chars

    def _find_tables(self, page) -> Iterable[tuple[Table, Any]]:
        try:
            finder = page.find_tables()
        except Exception as e:
            self.warnings.append(f"{page.number + 1}쪽 표 인식 실패: {e}")
            return
        for t in finder.tables:
            try:
                grid = t.extract()
            except Exception:
                continue
            if not grid or all(not any(c for c in row) for row in grid):
                continue
            tbl = Table()
            for row in grid:
                tbl.rows.append([Cell.of((c or "").strip()) for c in row])
            tbl.header_row = _header_from_finder(t)
            tbl.normalize()
            yield tbl, t.bbox if hasattr(t, "bbox") else page.rect

    @staticmethod
    def _inside_table(bbox: Iterable[float], rects: list[Any]) -> bool:
        x0, y0, x1, y1 = bbox
        area = max(1e-6, (x1 - x0) * (y1 - y0))
        for r in rects:
            rx0, ry0, rx1, ry1 = _rect_tuple(r)
            ix = max(0.0, min(x1, rx1) - max(x0, rx0))
            iy = max(0.0, min(y1, ry1) - max(y0, ry0))
            if ix * iy / area >= TABLE_OVERLAP:
                return True
        return False

    def _block_to_paragraph(self, blk: dict) -> Optional[Paragraph]:
        """PyMuPDF의 텍스트 블록 하나를 문단 하나로 만든다.

        블록 안의 여러 줄은 대개 같은 문단이므로 공백으로 잇는다. 다만 줄 끝이
        하이픈이면 단어가 잘린 것으로 보고 하이픈을 제거하고 붙인다.
        """
        para = Paragraph()
        lines = blk.get("lines", [])
        for li, line in enumerate(lines):
            spans = line.get("spans", [])
            for span in spans:
                text = span.get("text", "")
                if not text:
                    continue
                para.runs.append(_span_to_run(span))
            if li < len(lines) - 1 and para.runs:
                last = para.runs[-1]
                if last.text.endswith("-"):
                    last.text = last.text[:-1]
                elif not last.text.endswith((" ", " ")):
                    # CJK끼리는 공백을 넣지 않는다. 넣으면 어색하게 벌어진다.
                    nxt = _first_char_of_next_line(lines, li + 1)
                    if not (_is_cjk(last.text[-1:]) and _is_cjk(nxt)):
                        last.text += " "
        if not para.runs:
            return None
        para.merge_runs()
        para.align = _guess_align(blk, para)
        return para

    def _page_images(self, page, fitz) -> Iterable[tuple[float, Image]]:
        try:
            infos = page.get_image_info(xrefs=True)
        except Exception:
            infos = []
        seen: set[int] = set()
        for info in infos:
            xref = info.get("xref", 0)
            if not xref or xref in seen:
                continue
            seen.add(xref)
            try:
                d = page.parent.extract_image(xref)
            except Exception:
                continue
            bbox = info.get("bbox") or (0, 0, 0, 0)
            yield (
                bbox[1],
                Image(
                    data=d["image"],
                    fmt=d.get("ext", "png"),
                    width_pt=abs(bbox[2] - bbox[0]) or None,
                    height_pt=abs(bbox[3] - bbox[1]) or None,
                    name=f"page{page.number + 1}_img{xref}.{d.get('ext', 'png')}",
                ),
            )

    # -- 제목 추론 --------------------------------------------------------

    def _promote_headings(self, doc: Document) -> None:
        """본문 크기 대비 큰 문단을 제목으로 올린다.

        PDF에는 스타일 정보가 없으므로 글자 크기가 유일한 단서다. 중앙값을
        본문 크기로 보고 비율로 판정한다.
        """
        sizes: list[float] = []
        for b in doc.blocks:
            if isinstance(b, Paragraph):
                for r in b.runs:
                    if r.size_pt and r.text.strip():
                        sizes.extend([r.size_pt] * len(r.text.strip()))
        if not sizes:
            return
        body = statistics.median(sizes)
        if body <= 0:
            return

        for b in doc.blocks:
            if not isinstance(b, Paragraph) or b.is_blank():
                continue
            cand = [r.size_pt for r in b.runs if r.size_pt and r.text.strip()]
            if not cand:
                continue
            size = max(cand)
            ratio = size / body
            text = b.text.strip()
            # 제목은 대체로 짧다. 긴 문단은 크기가 커도 본문으로 둔다.
            if len(text) > 120:
                continue
            if ratio >= HEADING_RATIO_H1:
                b.heading = 1
            elif ratio >= HEADING_RATIO_H2:
                b.heading = 2
            elif ratio >= HEADING_RATIO_H3 and all(
                r.bold for r in b.runs if r.text.strip()
            ):
                b.heading = 3


# --------------------------------------------------------------------------
# 보조
# --------------------------------------------------------------------------


def _rect_tuple(r: Any) -> tuple[float, float, float, float]:
    """PyMuPDF의 사각형 표현을 (x0, y0, x1, y1) 튜플로 통일한다.

    `find_tables()` 가 돌려주는 bbox는 버전에 따라 Rect일 수도 평범한
    튜플일 수도 있다. 한쪽만 가정하면 AttributeError로 터진다.
    """
    if hasattr(r, "x0"):
        return (float(r.x0), float(r.y0), float(r.x1), float(r.y1))
    t = tuple(float(v) for v in r)
    if len(t) >= 4:
        return (t[0], t[1], t[2], t[3])
    return (0.0, 0.0, 0.0, 0.0)


_CJK_RANGES = (
    (0x1100, 0x11FF),  # 한글 자모
    (0x2E80, 0x9FFF),  # CJK 부수 ~ 통합한자
    (0xA960, 0xA97F),
    (0xAC00, 0xD7AF),  # 한글 음절
    (0xF900, 0xFAFF),
    (0xFF00, 0xFFEF),  # 전각
)


def _is_cjk(ch: str) -> bool:
    if not ch:
        return False
    o = ord(ch[0])
    return any(lo <= o <= hi for lo, hi in _CJK_RANGES)


def _first_char_of_next_line(lines: list, idx: int) -> str:
    try:
        for span in lines[idx].get("spans", []):
            t = span.get("text", "")
            if t:
                return t[0]
    except (IndexError, KeyError):
        pass
    return ""


def _span_to_run(span: dict) -> Run:
    flags = int(span.get("flags", 0))
    color = span.get("color")
    rgb = None
    if isinstance(color, int) and color >= 0:
        rgb = ((color >> 16) & 0xFF, (color >> 8) & 0xFF, color & 0xFF)
        if rgb == (0, 0, 0):
            rgb = None
    font = span.get("font") or None
    return Run(
        text=span.get("text", ""),
        # PyMuPDF flags: 1=위첨자 2=이탤릭(합성) 4=serif 8=고정폭 16=굵게 반영
        bold=bool(flags & 2**4) or _name_says_bold(font),
        italic=bool(flags & 2**1) or _name_says_italic(font),
        superscript=bool(flags & 2**0),
        size_pt=float(span.get("size", 0)) or None,
        font=font,
        color=rgb,
    )


def _name_says_bold(font: Optional[str]) -> bool:
    return bool(font) and any(
        k in font.lower() for k in ("bold", "black", "heavy", "semibold")
    )


def _name_says_italic(font: Optional[str]) -> bool:
    return bool(font) and any(k in font.lower() for k in ("italic", "oblique"))


def _guess_align(blk: dict, para: Paragraph) -> Align:
    """블록의 좌우 여백 비대칭으로 정렬을 추정한다. 판단이 서지 않으면 LEFT."""
    lines = blk.get("lines", [])
    if len(lines) < 1:
        return Align.LEFT
    xs0 = [ln["bbox"][0] for ln in lines if "bbox" in ln]
    xs1 = [ln["bbox"][2] for ln in lines if "bbox" in ln]
    if not xs0:
        return Align.LEFT
    # 여러 줄이고 왼쪽이 들쭉날쭉한데 오른쪽이 가지런하면 오른쪽 정렬로 본다.
    if len(xs0) >= 2:
        spread_l = max(xs0) - min(xs0)
        spread_r = max(xs1) - min(xs1)
        if spread_l > 12 and spread_r < 3:
            return Align.RIGHT
        if spread_l > 8 and spread_r > 8:
            return Align.CENTER
    return Align.LEFT


def _header_from_finder(t) -> bool:
    try:
        h = t.header
        return bool(h and getattr(h, "external", False) is False and h.names)
    except Exception:
        return False


def _first_text(doc: Document) -> Optional[str]:
    for b in doc.blocks:
        if isinstance(b, Paragraph) and not b.is_blank():
            return b.text.strip()[:120]
        if isinstance(b, Table):
            for row in b.rows:
                for c in row:
                    if c.text.strip():
                        return c.text.strip()[:120]
    return None


def read(path: Path | str, **kw: Any) -> Document:
    r = PdfReader(path, **kw)
    return _attach(r.read(), r)


def _attach(doc, reader):
    """Reader가 모은 경고를 Document에 실어 파이프라인까지 전달한다."""
    if getattr(reader, "warnings", None):
        doc.meta.extra.setdefault("warnings", []).extend(reader.warnings)
    return doc
