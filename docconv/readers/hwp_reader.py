"""HWP 5.0(구형 바이너리) Reader.

한/글 5.0 문서는 OLE 복합 문서(CFB) 안에 zlib(raw deflate)로 압축된 바이너리
레코드 트리를 담는다. 공개된 "한글 문서 파일 형식 5.0" 명세를 따라 구현했다.

스트림 구성::

    FileHeader                 256바이트 평문. 시그니처/버전/플래그
    DocInfo                    글꼴, 글자모양, 문단모양, 스타일 테이블
    BodyText/Section0..N       본문
    BinData/BIN####.png        임베디드 이미지
    PrvText                    UTF-16LE 평문 미리보기(본문 파싱 실패 시 폴백)
    \\x05HwpSummaryInformation 메타데이터

레코드 헤더(32비트 리틀엔디언)::

    tag_id = v & 0x3FF          (10비트)
    level  = (v >> 10) & 0x3FF  (10비트)
    size   = (v >> 20) & 0xFFF  (12비트)
    size == 0xFFF 이면 이어지는 UINT32가 실제 크기

쓰기는 지원하지 않는다. 유효한 HWP를 생성하려면 사실상 한컴오피스가 필요하다.
대신 HWPX로 내보내면 한/글에서 열어 .hwp로 저장할 수 있다.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Optional

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

# -- 레코드 태그 ------------------------------------------------------------
HWPTAG_BEGIN = 0x010

# DocInfo
HWPTAG_DOCUMENT_PROPERTIES = HWPTAG_BEGIN
HWPTAG_ID_MAPPINGS = HWPTAG_BEGIN + 1
HWPTAG_BIN_DATA = HWPTAG_BEGIN + 2
HWPTAG_FACE_NAME = HWPTAG_BEGIN + 3
HWPTAG_BORDER_FILL = HWPTAG_BEGIN + 4
HWPTAG_CHAR_SHAPE = HWPTAG_BEGIN + 5
HWPTAG_TAB_DEF = HWPTAG_BEGIN + 6
HWPTAG_NUMBERING = HWPTAG_BEGIN + 7
HWPTAG_BULLET = HWPTAG_BEGIN + 8
HWPTAG_PARA_SHAPE = HWPTAG_BEGIN + 9
HWPTAG_STYLE = HWPTAG_BEGIN + 10

# BodyText
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51
HWPTAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52
HWPTAG_PARA_LINE_SEG = HWPTAG_BEGIN + 53
HWPTAG_PARA_RANGE_TAG = HWPTAG_BEGIN + 54
HWPTAG_CTRL_HEADER = HWPTAG_BEGIN + 55
HWPTAG_LIST_HEADER = HWPTAG_BEGIN + 56
HWPTAG_PAGE_DEF = HWPTAG_BEGIN + 57
HWPTAG_FOOTNOTE_SHAPE = HWPTAG_BEGIN + 58
HWPTAG_PAGE_BORDER_FILL = HWPTAG_BEGIN + 59
HWPTAG_SHAPE_COMPONENT = HWPTAG_BEGIN + 60
HWPTAG_TABLE = HWPTAG_BEGIN + 61
HWPTAG_SHAPE_COMPONENT_PICTURE = HWPTAG_BEGIN + 67

# -- PARA_TEXT 제어 문자 분류 ----------------------------------------------
#: 1 WCHAR만 차지하는 제어 문자
CHAR_CONTROLS = frozenset({0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31})
#: 8 WCHAR(16바이트)를 차지하는 제어 문자 (인라인 + 확장)
LONG_CONTROLS = frozenset(
    {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23}
)

SIGNATURE = b"HWP Document File"


# --------------------------------------------------------------------------
# 저수준: 레코드 스트림
# --------------------------------------------------------------------------


@dataclass
class Record:
    tag: int
    level: int
    payload: bytes

    @property
    def size(self) -> int:
        return len(self.payload)


def iter_records(data: bytes) -> Iterator[Record]:
    """레코드 스트림을 순회한다. 손상 시 조용히 중단한다."""
    pos, n = 0, len(data)
    while pos + 4 <= n:
        (v,) = struct.unpack_from("<I", data, pos)
        pos += 4
        tag = v & 0x3FF
        level = (v >> 10) & 0x3FF
        size = (v >> 20) & 0xFFF
        if size == 0xFFF:
            if pos + 4 > n:
                return
            (size,) = struct.unpack_from("<I", data, pos)
            pos += 4
        if size < 0 or pos + size > n:
            return
        yield Record(tag, level, data[pos : pos + size])
        pos += size


class Cursor:
    """바이트 payload를 순차적으로 읽는 커서. 범위를 벗어나면 0/빈값을 준다."""

    __slots__ = ("d", "p")

    def __init__(self, data: bytes) -> None:
        self.d = data
        self.p = 0

    @property
    def remaining(self) -> int:
        return len(self.d) - self.p

    def _take(self, n: int) -> Optional[bytes]:
        if self.p + n > len(self.d):
            self.p = len(self.d)
            return None
        b = self.d[self.p : self.p + n]
        self.p += n
        return b

    def u8(self) -> int:
        b = self._take(1)
        return b[0] if b else 0

    def i8(self) -> int:
        b = self._take(1)
        return struct.unpack("<b", b)[0] if b else 0

    def u16(self) -> int:
        b = self._take(2)
        return struct.unpack("<H", b)[0] if b else 0

    def i16(self) -> int:
        b = self._take(2)
        return struct.unpack("<h", b)[0] if b else 0

    def u32(self) -> int:
        b = self._take(4)
        return struct.unpack("<I", b)[0] if b else 0

    def i32(self) -> int:
        b = self._take(4)
        return struct.unpack("<i", b)[0] if b else 0

    def skip(self, n: int) -> None:
        self.p = min(len(self.d), self.p + n)

    def wstr(self) -> str:
        """UINT16 길이 + UTF-16LE 문자열."""
        ln = self.u16()
        b = self._take(ln * 2)
        return b.decode("utf-16-le", errors="replace") if b else ""


# --------------------------------------------------------------------------
# DocInfo: 서식 테이블
# --------------------------------------------------------------------------


@dataclass
class HwpCharShape:
    size_pt: float = 10.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    superscript: bool = False
    subscript: bool = False
    color: Optional[tuple[int, int, int]] = None
    font: Optional[str] = None


@dataclass
class HwpParaShape:
    align: Align = Align.JUSTIFY
    indent_pt: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    line_spacing: Optional[float] = None


@dataclass
class DocInfo:
    fonts: list[str] = field(default_factory=list)
    char_shapes: list[HwpCharShape] = field(default_factory=list)
    para_shapes: list[HwpParaShape] = field(default_factory=list)
    styles: list[tuple[str, int, int]] = field(default_factory=list)
    #: BIN_DATA 인덱스(1-base) -> 확장자
    bin_ext: dict[int, str] = field(default_factory=dict)

    def char(self, idx: int) -> HwpCharShape:
        if 0 <= idx < len(self.char_shapes):
            return self.char_shapes[idx]
        return HwpCharShape()

    def para(self, idx: int) -> HwpParaShape:
        if 0 <= idx < len(self.para_shapes):
            return self.para_shapes[idx]
        return HwpParaShape()

    def style_name(self, idx: int) -> Optional[str]:
        if 0 <= idx < len(self.styles):
            return self.styles[idx][0]
        return None


_ALIGN_BY_CODE = {
    0: Align.JUSTIFY,
    1: Align.LEFT,
    2: Align.RIGHT,
    3: Align.CENTER,
    4: Align.DISTRIBUTE,
    5: Align.DISTRIBUTE,
}


def _color(v: int) -> Optional[tuple[int, int, int]]:
    if v < 0:
        return None
    rgb = (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)
    return None if rgb == (0, 0, 0) else rgb


def parse_docinfo(data: bytes) -> DocInfo:
    info = DocInfo()
    bin_idx = 0
    for rec in iter_records(data):
        if rec.tag == HWPTAG_FACE_NAME:
            c = Cursor(rec.payload)
            prop = c.u8()
            info.fonts.append(c.wstr())
            # 이후 대체글꼴/패널티/기본글꼴 정보는 건너뛴다.
        elif rec.tag == HWPTAG_CHAR_SHAPE:
            info.char_shapes.append(_parse_char_shape(rec.payload, info))
        elif rec.tag == HWPTAG_PARA_SHAPE:
            info.para_shapes.append(_parse_para_shape(rec.payload))
        elif rec.tag == HWPTAG_STYLE:
            c = Cursor(rec.payload)
            name = c.wstr()
            c.wstr()  # 영문 이름
            c.u8()  # 속성
            c.u8()  # 다음 스타일
            c.u16()  # 언어
            para_id = c.u16()
            char_id = c.u16()
            info.styles.append((name, para_id, char_id))
        elif rec.tag == HWPTAG_BIN_DATA:
            bin_idx += 1
            c = Cursor(rec.payload)
            prop = c.u16()
            ty = prop & 0x000F
            if ty == 0:  # LINK
                c.wstr()
                c.wstr()
            else:  # EMBEDDING / STORAGE
                c.u16()  # binary data ID
                ext = c.wstr() if ty == 1 else ""
                info.bin_ext[bin_idx] = (ext or "png").lower()
    return info


def _parse_char_shape(payload: bytes, info: DocInfo) -> HwpCharShape:
    c = Cursor(payload)
    face_ids = [c.u16() for _ in range(7)]
    c.skip(7)  # ratio
    c.skip(7)  # char spacing
    c.skip(7)  # rel size
    c.skip(7)  # char offset
    base = c.i32()  # 기준 크기 (1/100 pt)
    prop = c.u32()

    cs = HwpCharShape()
    cs.size_pt = base / 100.0 if base > 0 else 10.0
    cs.italic = bool(prop & 0x01)
    cs.bold = bool(prop & 0x02)
    cs.underline = ((prop >> 2) & 0x03) != 0
    cs.superscript = bool(prop & (1 << 15))
    cs.subscript = bool(prop & (1 << 16))
    cs.strike = ((prop >> 18) & 0x07) != 0

    c.i8()  # shadow gap x
    c.i8()  # shadow gap y
    cs.color = _color(c.u32())
    if face_ids and 0 <= face_ids[0] < len(info.fonts):
        cs.font = info.fonts[face_ids[0]]
    return cs


def _parse_para_shape(payload: bytes) -> HwpParaShape:
    c = Cursor(payload)
    prop1 = c.u32()
    left = c.i32()
    right = c.i32()
    indent = c.i32()
    prev_sp = c.i32()
    next_sp = c.i32()

    ps = HwpParaShape()
    ps.align = _ALIGN_BY_CODE.get((prop1 >> 2) & 0x07, Align.JUSTIFY)
    # HWPUNIT(1/7200인치) -> pt
    ps.indent_pt = max(0.0, left / 100.0)
    ps.space_before_pt = max(0.0, prev_sp / 100.0)
    ps.space_after_pt = max(0.0, next_sp / 100.0)

    line_spacing = c.i32()
    if line_spacing > 0 and (prop1 & 0x03) == 0:  # 0 = 글자에 따라(%)
        ps.line_spacing = line_spacing / 100.0
    return ps


# --------------------------------------------------------------------------
# 본문
# --------------------------------------------------------------------------


class HwpReader:
    def __init__(self, path: Path | str) -> None:
        self.path = Path(path)
        self.warnings: list[str] = []
        self._info = DocInfo()
        self._bin: dict[str, bytes] = {}

    def read(self) -> Document:
        try:
            import olefile
        except ImportError as e:  # pragma: no cover
            raise ReadError("olefile이 필요합니다. pip install olefile") from e

        if not olefile.isOleFile(str(self.path)):
            raise CorruptFileError(
                "HWP 파일이 아닙니다(OLE 복합 문서가 아님).", detail=str(self.path)
            )

        with olefile.OleFileIO(str(self.path)) as ole:
            compressed = self._check_header(ole)
            self._load_bindata(ole, compressed)

            # DocInfo는 서식 테이블일 뿐이다. 여기서 실패해도 본문은 읽어야
            # 하므로 예외를 삼키고 기본 서식으로 진행한다. 본문(BodyText)의
            # 실패와 달리 치명적이지 않다.
            if ole.exists("DocInfo"):
                try:
                    raw = self._stream(ole, "DocInfo", compressed)
                    if raw:
                        self._info = parse_docinfo(raw)
                except Exception as e:
                    self.warnings.append(
                        f"서식 정보를 읽지 못해 기본 서식으로 진행합니다 ({e})."
                    )

            sections = sorted(
                (
                    "/".join(p)
                    for p in ole.listdir()
                    if len(p) == 2 and p[0] == "BodyText" and p[1].startswith("Section")
                ),
                key=lambda s: int(s.rsplit("Section", 1)[-1] or 0),
            )
            if not sections:
                return self._fallback_preview(ole)

            doc = Document(kind=DocKind.DOCUMENT, meta=self._meta(ole))
            for i, name in enumerate(sections):
                if i > 0:
                    doc.blocks.append(PageBreak())
                raw = self._stream(ole, name, compressed)
                doc.blocks.extend(self._parse_section(raw))

            doc.compact()
            if doc.is_empty():
                self.warnings.append(
                    "본문 레코드에서 텍스트를 얻지 못해 미리보기 텍스트로 대체했습니다."
                )
                fb = self._fallback_preview(ole)
                if not fb.is_empty():
                    fb.meta = doc.meta
                    return fb
            if not doc.meta.title:
                doc.meta.title = _first_text(doc)
            return doc

    # -- 헤더/스트림 ------------------------------------------------------

    def _check_header(self, ole) -> bool:
        if not ole.exists("FileHeader"):
            raise CorruptFileError("FileHeader가 없습니다. HWP 문서가 아닙니다.")
        head = ole.openstream("FileHeader").read(256)
        if not head.startswith(SIGNATURE):
            raise CorruptFileError("HWP 시그니처가 올바르지 않습니다.")
        if len(head) < 40:
            raise CorruptFileError("FileHeader가 손상되었습니다.")
        ver = struct.unpack_from("<BBBB", head, 32)
        (prop,) = struct.unpack_from("<I", head, 36)
        compressed = bool(prop & 0x01)
        if prop & 0x02:
            raise EncryptedFileError(str(self.path))
        if prop & 0x04:
            raise ReadError(
                "배포용(복사 방지) 문서는 읽을 수 없습니다. "
                "한/글에서 배포용 설정을 해제한 뒤 다시 시도하세요.",
                detail=str(self.path),
            )
        self.warnings.append("") if False else None
        self._version = f"{ver[3]}.{ver[2]}.{ver[1]}.{ver[0]}"
        return compressed

    @staticmethod
    def _stream(ole, name: str, compressed: bool) -> bytes:
        raw = ole.openstream(name).read()
        if not compressed:
            return raw
        try:
            # HWP는 zlib 헤더 없는 raw deflate를 쓴다.
            return zlib.decompress(raw, -15)
        except zlib.error:
            try:
                return zlib.decompress(raw)
            except zlib.error as e:
                raise CorruptFileError(
                    f"'{name}' 스트림 압축 해제에 실패했습니다.", detail=str(e)
                ) from e

    def _load_bindata(self, ole, compressed: bool) -> None:
        for p in ole.listdir():
            if len(p) == 2 and p[0] == "BinData":
                try:
                    self._bin[p[1]] = self._stream(ole, "/".join(p), compressed)
                except Exception:
                    continue

    def _meta(self, ole) -> DocMeta:
        meta = DocMeta(
            source_format="hwp", extra={"hwp_version": getattr(self, "_version", "")}
        )
        try:
            md = ole.get_metadata()
            meta.title = (
                (md.title or b"").decode("utf-8", "ignore") or None
                if isinstance(md.title, bytes)
                else md.title
            )
            meta.author = (
                (md.author or b"").decode("utf-8", "ignore") or None
                if isinstance(md.author, bytes)
                else md.author
            )
            if md.create_time:
                meta.created = md.create_time.isoformat()
            if md.last_saved_time:
                meta.modified = md.last_saved_time.isoformat()
        except Exception:
            pass
        return meta

    # -- 섹션 파싱 --------------------------------------------------------

    def _parse_section(self, data: bytes) -> list[Block]:
        """레코드 스트림을 블록 목록으로.

        표는 CTRL_HEADER(ctrl_id='tbl ') 다음에 TABLE 레코드와 셀별
        LIST_HEADER가 이어지는 구조다. level로 소속을 판별한다.
        """
        blocks: list[Block] = []
        recs = list(iter_records(data))
        i = 0
        n = len(recs)

        while i < n:
            r = recs[i]
            if r.tag == HWPTAG_PARA_HEADER:
                para, consumed, extra = self._read_paragraph(recs, i)
                if para is not None:
                    blocks.append(para)
                blocks.extend(extra)
                i += consumed
                continue
            i += 1
        return blocks

    def _read_paragraph(
        self, recs: list[Record], i: int
    ) -> tuple[Optional[Paragraph], int, list[Block]]:
        """PARA_HEADER 하나와 그에 딸린 레코드들을 소비해 문단을 만든다."""
        head = recs[i]
        base_level = head.level
        c = Cursor(head.payload)
        n_chars = c.u32() & 0x7FFFFFFF
        c.u32()  # control mask
        para_shape_id = c.u16()
        style_id = c.u8()

        text = ""
        char_pos: list[tuple[int, int]] = []
        extra: list[Block] = []

        j = i + 1
        while j < len(recs):
            r = recs[j]
            if r.tag == HWPTAG_PARA_HEADER and r.level <= base_level:
                break
            if r.level <= base_level and r.tag not in (
                HWPTAG_PARA_TEXT,
                HWPTAG_PARA_CHAR_SHAPE,
                HWPTAG_PARA_LINE_SEG,
                HWPTAG_PARA_RANGE_TAG,
                HWPTAG_CTRL_HEADER,
            ):
                break

            if r.tag == HWPTAG_PARA_TEXT:
                text = _decode_para_text(r.payload)
            elif r.tag == HWPTAG_PARA_CHAR_SHAPE:
                char_pos = _decode_char_shape_map(r.payload)
            elif r.tag == HWPTAG_CTRL_HEADER:
                ctrl_id = _ctrl_id(r.payload)
                if ctrl_id == "tbl ":
                    tbl, used = self._read_table(recs, j, r.level)
                    if tbl is not None:
                        extra.append(tbl)
                    j += used
                    continue
                if ctrl_id == "gso ":
                    img, used = self._read_picture(recs, j, r.level)
                    if img is not None:
                        extra.append(img)
                    j += used
                    continue
            j += 1

        consumed = max(1, j - i)
        if not text.strip() and not extra:
            return (Paragraph(), consumed, []) if text else (None, consumed, extra)

        ps = self._info.para(para_shape_id)
        para = Paragraph(
            align=ps.align,
            indent_pt=ps.indent_pt,
            space_before_pt=ps.space_before_pt,
            space_after_pt=ps.space_after_pt,
            line_spacing=ps.line_spacing,
            style_name=self._info.style_name(style_id),
        )
        para.heading = _heading_from_name(para.style_name)
        para.runs = list(self._make_runs(text, char_pos))
        para.merge_runs()
        return (para if para.runs else None), consumed, extra

    def _make_runs(self, text: str, char_pos: list[tuple[int, int]]) -> Iterator[Run]:
        """문자 위치별 글자모양 매핑을 Run 목록으로 변환한다."""
        if not text:
            return
        if not char_pos:
            yield Run(text=text)
            return
        bounds = char_pos + [(len(text), -1)]
        for k in range(len(char_pos)):
            start = bounds[k][0]
            end = bounds[k + 1][0]
            if start >= len(text):
                break
            chunk = text[start : min(end, len(text))]
            if not chunk:
                continue
            cs = self._info.char(bounds[k][1])
            yield Run(
                text=chunk,
                bold=cs.bold,
                italic=cs.italic,
                underline=cs.underline,
                strike=cs.strike,
                superscript=cs.superscript,
                subscript=cs.subscript,
                size_pt=cs.size_pt,
                font=cs.font,
                color=cs.color,
            )

    # -- 표 ---------------------------------------------------------------

    def _read_table(
        self, recs: list[Record], i: int, ctrl_level: int
    ) -> tuple[Optional[Table], int]:
        """CTRL_HEADER('tbl ') 이후의 TABLE + 셀 LIST_HEADER들을 읽는다."""
        j = i + 1
        n_rows = n_cols = 0
        row_sizes: list[int] = []

        # TABLE 레코드 찾기
        while j < len(recs) and recs[j].level > ctrl_level:
            if recs[j].tag == HWPTAG_TABLE:
                n_rows, n_cols, row_sizes = _parse_table_rec(recs[j].payload)
                j += 1
                break
            j += 1

        if n_rows <= 0 or n_cols <= 0:
            return None, max(1, j - i)

        grid: list[list[Optional[Cell]]] = [[None] * n_cols for _ in range(n_rows)]
        seq = 0  # LIST_HEADER 순서. 좌표 파싱 실패 시 폴백으로 쓴다.

        while j < len(recs) and recs[j].level > ctrl_level:
            r = recs[j]
            if r.tag != HWPTAG_LIST_HEADER:
                j += 1
                continue

            col, row, cspan, rspan = _parse_cell_header(r.payload)
            cell_level = r.level

            # 셀 안의 문단들을 읽는다.
            blocks: list[Block] = []
            k = j + 1
            while k < len(recs) and recs[k].level > cell_level:
                if recs[k].tag == HWPTAG_PARA_HEADER:
                    p, used, ex = self._read_paragraph(recs, k)
                    if p is not None:
                        blocks.append(p)
                    blocks.extend(ex)
                    k += used
                    continue
                k += 1

            if (
                not (0 <= row < n_rows and 0 <= col < n_cols)
                or grid[row][col] is not None
            ):
                row, col = divmod(seq, n_cols) if n_cols else (0, 0)
            if 0 <= row < n_rows and 0 <= col < n_cols:
                cell = Cell(
                    blocks=blocks, col_span=max(1, cspan), row_span=max(1, rspan)
                )
                grid[row][col] = cell
                for dr in range(cell.row_span):
                    for dc in range(cell.col_span):
                        if dr or dc:
                            rr, cc = row + dr, col + dc
                            if (
                                0 <= rr < n_rows
                                and 0 <= cc < n_cols
                                and grid[rr][cc] is None
                            ):
                                grid[rr][cc] = Cell(merged_placeholder=True)
            seq += 1
            j = k

        tbl = Table(
            rows=[[c if c is not None else Cell() for c in row] for row in grid]
        )
        return tbl, max(1, j - i)

    # -- 그림 -------------------------------------------------------------

    def _read_picture(
        self, recs: list[Record], i: int, ctrl_level: int
    ) -> tuple[Optional[Image], int]:
        j = i + 1
        while j < len(recs) and recs[j].level > ctrl_level:
            if recs[j].tag == HWPTAG_SHAPE_COMPONENT_PICTURE:
                bin_id = _parse_picture_rec(recs[j].payload)
                if bin_id:
                    img = self._image_by_id(bin_id)
                    if img is not None:
                        while j < len(recs) and recs[j].level > ctrl_level:
                            j += 1
                        return img, max(1, j - i)
                break
            j += 1
        while j < len(recs) and recs[j].level > ctrl_level:
            j += 1
        return None, max(1, j - i)

    def _image_by_id(self, bin_id: int) -> Optional[Image]:
        ext = self._info.bin_ext.get(bin_id, "")
        for name, data in self._bin.items():
            stem = name.upper()
            if f"{bin_id:04X}" in stem or f"BIN{bin_id:04d}" in stem:
                real_ext = ext or Path(name).suffix.lstrip(".") or _sniff_image(data)
                return Image(
                    data=data,
                    fmt="jpeg" if real_ext in ("jpg", "jpeg") else (real_ext or "png"),
                    name=name,
                )
        return None

    # -- 폴백 -------------------------------------------------------------

    def _fallback_preview(self, ole) -> Document:
        """본문을 못 읽을 때 PrvText(평문 미리보기)라도 살린다."""
        doc = Document(kind=DocKind.DOCUMENT, meta=self._meta(ole))
        if not ole.exists("PrvText"):
            raise CorruptFileError(
                "본문(BodyText)을 찾을 수 없습니다.", detail=str(self.path)
            )
        raw = ole.openstream("PrvText").read()
        text = raw.decode("utf-16-le", errors="replace")
        self.warnings.append(
            "본문 대신 미리보기 텍스트를 사용했습니다. 내용이 잘렸을 수 있습니다."
        )
        for line in text.replace("\r\n", "\n").split("\n"):
            doc.blocks.append(Paragraph.of(line.strip("\x00").rstrip()))
        doc.compact()
        return doc


# --------------------------------------------------------------------------
# 레코드 페이로드 파서
# --------------------------------------------------------------------------


def _decode_para_text(payload: bytes) -> str:
    """PARA_TEXT를 문자열로. 제어 문자의 가변 길이를 정확히 건너뛴다.

    이 규칙을 틀리면 텍스트에 쓰레기 글자가 섞이거나 뒷부분이 통째로 밀린다.
    """
    out: list[str] = []
    n = len(payload) // 2
    k = 0
    while k < n:
        (ch,) = struct.unpack_from("<H", payload, k * 2)
        if ch in LONG_CONTROLS:
            if ch == 9:
                out.append("\t")
            elif ch == 11:
                pass  # 표/그리기 개체. 별도 블록으로 처리한다.
            k += 8
            continue
        if ch in CHAR_CONTROLS:
            if ch == 10:
                out.append("\n")
            elif ch == 13:
                out.append("\n")
            elif ch == 24:
                out.append("-")
            elif ch in (30, 31):
                out.append(" ")
            k += 1
            continue
        out.append(chr(ch))
        k += 1
    return "".join(out)


def _decode_char_shape_map(payload: bytes) -> list[tuple[int, int]]:
    """(문자위치, char_shape_id) 배열."""
    out: list[tuple[int, int]] = []
    for off in range(0, len(payload) - 7, 8):
        pos, cid = struct.unpack_from("<II", payload, off)
        out.append((pos, cid))
    out.sort(key=lambda t: t[0])
    return out


def _ctrl_id(payload: bytes) -> str:
    """CTRL_HEADER의 4바이트 ID. 리틀엔디언이라 뒤집어 읽는다."""
    if len(payload) < 4:
        return ""
    return payload[:4][::-1].decode("ascii", errors="replace")


def _parse_table_rec(payload: bytes) -> tuple[int, int, list[int]]:
    c = Cursor(payload)
    c.u32()  # property
    n_rows = c.u16()
    n_cols = c.u16()
    c.u16()  # cell spacing
    c.skip(8)  # margins (4 x HWPUNIT16)
    sizes = [c.u16() for _ in range(min(n_rows, 4096))]
    return n_rows, n_cols, sizes


def _parse_cell_header(payload: bytes) -> tuple[int, int, int, int]:
    """표 셀 LIST_HEADER에서 (col, row, colspan, rowspan)."""
    c = Cursor(payload)
    c.i16()  # 문단 수
    c.u32()  # 속성
    col = c.u16()
    row = c.u16()
    cspan = c.u16()
    rspan = c.u16()
    return col, row, max(1, cspan), max(1, rspan)


def _parse_picture_rec(payload: bytes) -> int:
    """SHAPE_COMPONENT_PICTURE에서 BinData ID를 뽑는다.

    앞부분은 테두리/자르기/그림자 정보라 길이가 버전에 따라 다르다. 실무적으로
    안정적인 방법은 구조 오프셋(0x24 부근)을 시도하고, 실패하면 뒤에서부터
    유효 범위의 UINT16을 찾는 것이다.
    """
    if len(payload) < 60:
        return 0
    for off in (0x24, 0x26, 0x28):
        if off + 2 <= len(payload):
            (v,) = struct.unpack_from("<H", payload, off)
            if 0 < v < 4096:
                return v
    return 0


def _sniff_image(data: bytes) -> str:
    if data.startswith(b"\x89PNG"):
        return "png"
    if data.startswith(b"\xff\xd8"):
        return "jpeg"
    if data.startswith(b"GIF8"):
        return "gif"
    if data.startswith(b"BM"):
        return "bmp"
    return "png"


_HEADING_PREFIXES = ("개요 ", "heading ", "제목 ")


def _heading_from_name(name: Optional[str]) -> int:
    if not name:
        return 0
    low = name.strip().lower()
    if low in ("title", "제목"):
        return 1
    for pre in _HEADING_PREFIXES:
        if low.startswith(pre):
            try:
                return max(1, min(9, int(low[len(pre) :].strip())))
            except ValueError:
                return 0
    return 0


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


def read(path: Path | str) -> Document:
    r = HwpReader(path)
    return _attach(r.read(), r)


def _attach(doc, reader):
    """Reader가 모은 경고를 Document에 실어 파이프라인까지 전달한다."""
    if getattr(reader, "warnings", None):
        doc.meta.extra.setdefault("warnings", []).extend(reader.warnings)
    return doc
