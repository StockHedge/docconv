"""PDF Writer (PyMuPDF Story 기반).

한글이 깨지지 않게 하는 것이 이 모듈의 최우선 과제다. PDF 표준 14폰트에는
한글 글리프가 없어서, 시스템 글꼴 파일을 읽어 문서에 **임베딩**해야 한다.
그렇지 않으면 모든 한글이 네모(□)로 나온다.

구현 방식
---------
IR을 HTML로 직렬화한 뒤 PyMuPDF의 `Story` 로 조판한다. 직접 좌표를 계산하는
것보다 나은 이유는 셋이다.

* 줄바꿈/페이지 넘김을 Story가 알아서 처리한다.
* 표를 `<table>` 로 그리면 열 너비와 셀 병합을 그대로 넘길 수 있다.
* 글꼴은 `@font-face` 로 한 번만 등록하면 문서 전체에 적용된다.

Story가 실패하는 경우(아주 오래된 PyMuPDF 등)를 위해 저수준 TextWriter
폴백을 둔다. 폴백은 서식 없이 텍스트만 배치한다.
"""

from __future__ import annotations

import html as _html
import io
from pathlib import Path
from typing import Iterable, Optional

from ..errors import WriteError
from ..ir import (
    Align,
    Block,
    Border,
    BorderStyle,
    Cell,
    Document,
    Image,
    ListKind,
    PageBreak,
    Paragraph,
    Run,
    Table,
)
from ..options import ConvertOptions
from ..util import fonts, units

#: Story에 등록할 글꼴의 내부 이름
_FONT_FAMILY = "docconv-kr"
_FONT_FILE = "docconv-kr.ttf"
_FONT_FILE_BOLD = "docconv-kr-bold.ttf"


def write_pdf(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    fitz = _import_fitz()
    font = _resolve_font(opts)

    try:
        _write_with_story(fitz, doc, Path(path), opts, font)
    except WriteError:
        raise
    except Exception as e:
        # Story 실패 시 단순 텍스트라도 남긴다.
        try:
            _write_fallback(fitz, doc, Path(path), opts, font)
        except Exception:
            raise WriteError(
                "PDF 생성에 실패했습니다.", detail=f"{type(e).__name__}: {e}"
            ) from e


def _import_fitz():
    try:
        import pymupdf as fitz

        return fitz
    except ImportError:
        try:
            import fitz

            return fitz
        except ImportError as e:  # pragma: no cover
            raise WriteError("PyMuPDF가 필요합니다. pip install pymupdf") from e


def _resolve_font(opts: ConvertOptions) -> Optional[fonts.FontPair]:
    if opts.pdf_font_path:
        p = Path(opts.pdf_font_path)
        if p.is_file():
            return fonts.FontPair(name=p.stem, regular=p, bold=p)
    return fonts.find_korean_font()


# --------------------------------------------------------------------------
# Story 경로
# --------------------------------------------------------------------------


def _write_with_story(
    fitz,
    doc: Document,
    path: Path,
    opts: ConvertOptions,
    font: Optional[fonts.FontPair],
) -> None:
    archive = fitz.Archive()
    css_parts: list[str] = []

    if font is not None and not font.is_collection:
        try:
            archive.add(font.regular.read_bytes(), _FONT_FILE)
            css_parts.append(
                f"@font-face {{ font-family: {_FONT_FAMILY}; "
                f"src: url('{_FONT_FILE}'); }}"
            )
            if font.bold != font.regular and font.bold.is_file():
                archive.add(font.bold.read_bytes(), _FONT_FILE_BOLD)
                css_parts.append(
                    f"@font-face {{ font-family: {_FONT_FAMILY}; font-weight: bold; "
                    f"src: url('{_FONT_FILE_BOLD}'); }}"
                )
        except Exception:
            font = None
    elif font is not None and font.is_collection:
        # TTC는 Story가 다루지 못한다. 저수준 경로에서 인덱스로 처리한다.
        font = None

    images = _collect_images(doc, archive) if opts.include_images else {}
    font_css = list(css_parts)  # @font-face 선언. 두 번째 조판에서도 그대로 쓴다.
    css = "\n".join(font_css + [_base_css(opts, use_kfont=font is not None)])

    data = _typeset(fitz, _to_html(doc, opts, images), css, archive, doc)

    # 표가 쪽을 넘기면 Story가 첫 쪽의 셀 배경을 이후 모든 쪽에 **유령처럼
    # 복사**한다(같은 x·폭·색, 높이만 6pt로 뭉갠 채). 그 띠가 본문 글줄 위에
    # 얹혀 취소선처럼 보인다. 어떤 HTML/CSS로도 피하지 못해서, 조판 결과를
    # 살펴보고 유령이 생겼을 때만 셀 배경을 빼고 다시 조판한다.
    #
    # 한 쪽짜리 문서나 표가 쪽을 넘기지 않는 문서는 배경을 그대로 살린다.
    if _has_ghost_fills(fitz, data):
        # CSS의 `th { background-color }` 도 함께 꺼야 한다. 여기 남겨 두면
        # 머리글 칸에 배경이 그대로 붙어 유령 띠가 계속 나온다.
        css_nobg = "\n".join(
            font_css + [_base_css(opts, use_kfont=font is not None, cell_bg=False)]
        )
        data = _typeset(
            fitz, _to_html(doc, opts, images, cell_bg=False), css_nobg, archive, doc
        )

    _finalize(fitz, data, path, doc)


def _typeset(fitz, html: str, css: str, archive, doc: Document) -> bytes:
    """HTML을 조판해 PDF 바이트를 만든다.

    DocumentWriter를 파일 경로에 직접 걸면 Windows에서 핸들이 남아 후처리가
    조용히 실패한다. 그래서 메모리 버퍼에 담아 돌려준다.
    """
    story = fitz.Story(html=html, user_css=css, archive=archive)
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)

    mediabox = fitz.Rect(0, 0, doc.page_width_pt, doc.page_height_pt)
    left, top, right, bottom = doc.margin_pt
    where = mediabox + (left, top, -right, -bottom)

    more = 1
    guard = 0
    while more and guard < 10_000:
        guard += 1
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    return buf.getvalue()


#: 유령 배경으로 볼 높이 범위(pt).
#
# 위쪽 한계: 이보다 두꺼우면 정상 셀 배경이다(글자 한 줄을 담는다).
# 아래쪽 한계: 이보다 얇으면 **표 테두리 선**이다. 이 하한을 빠뜨리면 테두리를
# 유령으로 오판해, 배경이 멀쩡한 문서까지 매번 다시 조판하며 배경을 버린다.
_GHOST_MIN_HEIGHT = 2.0
_GHOST_MAX_HEIGHT = 8.0


def _has_ghost_fills(fitz, data: bytes) -> bool:
    """둘째 쪽부터 나타나는 비정상적으로 납작한 색 채움을 찾는다.

    정상 셀 배경은 셀 높이만큼(수십 pt) 칠해진다. 유령은 높이가 6pt 안팎으로
    고정된 채 본문 글줄을 가로지른다.
    """
    try:
        pdf = fitz.open("pdf", data)
    except Exception:
        return False
    try:
        if pdf.page_count < 2:
            return False
        for pno in range(1, pdf.page_count):
            for dr in pdf[pno].get_drawings():
                fill = dr.get("fill")
                if not fill:
                    continue
                rgb = tuple(round(v, 2) for v in fill)
                if rgb in ((1.0, 1.0, 1.0), (0.0, 0.0, 0.0)):
                    continue
                rect = dr["rect"]
                if (
                    _GHOST_MIN_HEIGHT <= rect.height <= _GHOST_MAX_HEIGHT
                    and rect.width > 20
                ):
                    return True
        return False
    except Exception:
        return False
    finally:
        pdf.close()


def _base_css(opts: ConvertOptions, *, use_kfont: bool, cell_bg: bool = True) -> str:
    fam = f"{_FONT_FAMILY}, sans-serif" if use_kfont else "sans-serif"
    size = opts.pdf_base_size_pt
    # 셀 배경을 끄는 재조판에서는 머리글 배경도 함께 꺼야 한다. 여기 남겨 두면
    # <th> 에 CSS 배경이 그대로 붙어 유령 띠가 계속 나온다.
    th_bg = " background-color: #f0f0f0;" if cell_bg else ""
    return f"""
    * {{ font-family: {fam}; }}
    /* body 기본 여백을 0으로 두지 않으면 표가 좌측으로 10pt 밀려
       오른쪽이 본문 폭을 넘어 잘린다(실측). */
    body {{ font-size: {size}pt; line-height: 1.5; margin: 0; padding: 0; }}
    p {{ margin: 0 0 {size * 0.35:.1f}pt 0; }}
    h1 {{ font-size: {size * 1.9:.1f}pt; font-weight: bold; margin: {size}pt 0 {size * 0.6:.1f}pt 0; }}
    h2 {{ font-size: {size * 1.55:.1f}pt; font-weight: bold; margin: {size * 0.9:.1f}pt 0 {size * 0.5:.1f}pt 0; }}
    h3 {{ font-size: {size * 1.3:.1f}pt; font-weight: bold; margin: {size * 0.8:.1f}pt 0 {size * 0.4:.1f}pt 0; }}
    h4, h5, h6 {{ font-size: {size * 1.12:.1f}pt; font-weight: bold; margin: {size * 0.7:.1f}pt 0 {size * 0.35:.1f}pt 0; }}
    table {{ border-collapse: collapse; margin: {size * 0.5:.1f}pt 0; }}
    td, th {{ padding: 3pt 4pt; vertical-align: top; text-align: left; }}
    th {{ font-weight: bold;{th_bg} }}
    ul, ol {{ margin: 0 0 {size * 0.35:.1f}pt 0; padding-left: {size * 1.6:.1f}pt; }}
    img {{ max-width: 100%; }}
    """


def _collect_images(doc: Document, archive) -> dict[int, str]:
    """이미지를 Archive에 넣고 (id -> 파일명) 매핑을 만든다."""
    out: dict[int, str] = {}
    idx = 0
    for b in _iter_all_blocks(doc.blocks):
        if not isinstance(b, Image) or not b.data:
            continue
        ext = "png" if b.fmt in ("emf", "wmf", "svg") else b.fmt
        if b.fmt in ("emf", "wmf", "svg"):
            continue  # Story가 렌더링하지 못한다.
        name = f"img{idx}.{ext}"
        try:
            archive.add(b.data, name)
        except Exception:
            continue
        out[id(b)] = name
        idx += 1
    return out


def _iter_all_blocks(blocks: Iterable[Block]) -> Iterable[Block]:
    for b in blocks:
        yield b
        if isinstance(b, Table):
            for row in b.rows:
                for c in row:
                    yield from _iter_all_blocks(c.blocks)


# --------------------------------------------------------------------------
# IR -> HTML
# --------------------------------------------------------------------------


_ALIGN_CSS = {
    Align.LEFT: "left",
    Align.CENTER: "center",
    Align.RIGHT: "right",
    Align.JUSTIFY: "justify",
    Align.DISTRIBUTE: "justify",
}


def _to_html(
    doc: Document,
    opts: ConvertOptions,
    images: dict[int, str],
    *,
    cell_bg: bool = True,
) -> str:
    # 표 열 너비를 pt 절대값으로 지정하려면 실제로 쓸 수 있는 본문 폭이 필요하다.
    left, _, right, _ = doc.margin_pt
    usable = max(72.0, doc.page_width_pt - left - right)
    parts: list[str] = ["<html><body>"]
    list_open: Optional[str] = None

    for b in doc.blocks:
        if isinstance(b, Paragraph) and b.list_kind is not ListKind.NONE:
            tag = "ul" if b.list_kind is ListKind.BULLET else "ol"
            if list_open != tag:
                if list_open:
                    parts.append(f"</{list_open}>")
                parts.append(f"<{tag}>")
                list_open = tag
            parts.append(f"<li>{_runs_html(b.runs, opts)}</li>")
            continue
        if list_open:
            parts.append(f"</{list_open}>")
            list_open = None
        parts.append(_block_html(b, opts, images, usable, cell_bg=cell_bg))

    if list_open:
        parts.append(f"</{list_open}>")
    parts.append("</body></html>")
    return "".join(parts)


def _block_html(
    b: Block,
    opts: ConvertOptions,
    images: dict[int, str],
    usable_pt: float = 453.0,
    *,
    cell_bg: bool = True,
) -> str:
    if isinstance(b, Paragraph):
        inner = _runs_html(b.runs, opts)
        if not inner.strip():
            return "<p>&#160;</p>"
        style = []
        if b.align is not Align.LEFT:
            style.append(f"text-align:{_ALIGN_CSS.get(b.align, 'left')}")
        if b.indent_pt:
            style.append(f"margin-left:{b.indent_pt:.1f}pt")
        attr = f' style="{";".join(style)}"' if style else ""
        if b.heading:
            lvl = min(6, max(1, b.heading))
            return f"<h{lvl}{attr}>{inner}</h{lvl}>"
        return f"<p{attr}>{inner}</p>"

    if isinstance(b, Table):
        return _table_html(b, opts, images, usable_pt, cell_bg=cell_bg)

    if isinstance(b, Image):
        name = images.get(id(b))
        if not name:
            return f"<p>[그림: {_esc(b.name or 'image')}]</p>"
        w = f' width="{b.width_pt:.0f}"' if b.width_pt else ""
        return f'<p><img src="{name}"{w}/></p>'

    if isinstance(b, PageBreak):
        # Story는 CSS page-break를 부분 지원한다. 안 먹으면 빈 문단으로 남는다.
        return '<div style="page-break-after: always;"></div>'

    return ""


def _table_html(
    t: Table,
    opts: ConvertOptions,
    images: dict[int, str],
    usable_pt: float = 453.0,
    *,
    cell_bg: bool = True,
) -> str:
    """표를 HTML로. 열 너비는 **pt 절대값**으로 지정한다.

    MuPDF Story는 ``<td width="83%">`` 같은 백분율을 존중하지 않는다. 실측에서
    83%를 지정한 열이 실제로는 35%만 차지했다(그 결과 제목이 좁은 칸에 밀려
    두 줄로 깨졌다). 반면 pt 절대값은 정확히 반영된다. 그래서 원본의 열 너비
    비율을 실제 본문 폭에 맞춰 pt로 환산해 넘긴다.

    셀마다 배경과 테두리를 따로 그린다. 한/글 문서는 칸마다 테두리 유무가
    다른 경우가 흔해서(제목 칸만 상자, 본문은 좌우선만), 일괄 격자로 그리면
    원본과 인상이 크게 달라진다.
    """
    t.normalize()
    if not t.rows:
        return ""

    widths_pt = _column_widths_pt(t, usable_pt)
    default_border = "0.5pt solid #9a9a9a"

    out: list[str] = ["<table>"]
    for ri, row in enumerate(t.rows):
        out.append("<tr>")
        for ci, c in enumerate(row):
            if c.merged_placeholder:
                continue
            tag = "th" if (t.header_row and ri == 0) else "td"
            attrs = ""
            if c.col_span > 1:
                attrs += f' colspan="{c.col_span}"'
            if c.row_span > 1:
                attrs += f' rowspan="{c.row_span}"'

            style: list[str] = []
            if widths_pt and ci < len(widths_pt):
                span_w = sum(widths_pt[ci : ci + max(1, c.col_span)])
                if span_w > 0:
                    style.append(f"width:{span_w:.1f}pt")

            if opts.keep_formatting:
                if c.background and cell_bg:
                    style.append(f"background-color:{units.rgb_to_hex(c.background)}")
                if c.borders is not None:
                    # 좌/상/우/하 순서로 저장돼 있다.
                    for side, b in zip(("left", "top", "right", "bottom"), c.borders):
                        style.append(f"border-{side}:{b.css()}")
                else:
                    style.append(f"border:{default_border}")
            else:
                style.append(f"border:{default_border}")

            if style:
                attrs += f' style="{";".join(style)}"'
            inner = "".join(
                _block_html(bb, opts, images, usable_pt, cell_bg=cell_bg)
                for bb in c.blocks
                if isinstance(bb, (Paragraph, Table, Image))
            )
            out.append(f"<{tag}{attrs}>{inner or '&#160;'}</{tag}>")
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


#: 열 하나가 지정 너비 외에 추가로 먹는 폭(pt).
#
# padding 좌우 8pt + 테두리 몫이다. 실측으로 정했다 - 8.0으로 두면 열 수와
# 무관하게 1~3pt씩 본문 폭을 넘어 표 오른쪽이 잘린다. 9.5면 2/3/5열 모두
# 여유를 두고 들어간다.
_CELL_PAD_LR = 9.5


def _column_widths_pt(t: Table, usable_pt: float) -> Optional[list[float]]:
    """원본 비율을 실제 본문 폭에 맞춰 pt로 환산한다.

    Story는 셀 너비에 padding을 더해 배치하므로, 그만큼을 미리 빼지 않으면
    표 전체가 본문 폭을 넘어 오른쪽이 잘린다.
    """
    n = t.n_cols
    if n <= 0:
        return None
    ratios = t.col_widths_pt or _auto_col_widths(t)
    if not ratios:
        return None
    ratios = list(ratios[:n]) + [0.0] * max(0, n - len(ratios))
    total = sum(r for r in ratios if r > 0)
    if total <= 0:
        return None

    content = max(40.0, usable_pt - _CELL_PAD_LR * n)
    # 비율이 없는 열(0)은 최소 폭만 준다.
    out: list[float] = []
    for r in ratios:
        out.append(content * (r / total) if r > 0 else 12.0)
    return out


#: 열 너비 추정에서 한 열이 가질 수 있는 최대 선호 폭(문자 단위).
_MAX_PREF_CHARS = 44.0


def _auto_col_widths(t: Table) -> Optional[list[float]]:
    """원본에 열 너비 정보가 없을 때 내용으로 추정한다.

    이게 없으면 Story가 열 너비를 자유롭게 정하는데, **대부분의 행이 비어 있고
    한 행에만 긴 값이 있는 열**이 극단적으로 좁아져 그 값이 잘린다. 실측에서
    긴 영문 낱말의 마지막 글자가 잘려 나갔다.

    두 값을 함께 본다.

    * 최장 단어 폭 - 단어 중간에서는 줄바꿈이 안 되므로 이보다 좁으면 잘린다.
      이 값이 최소 보장선이다.
    * 선호 폭 - 셀 내용 길이의 대푯값. 상한을 두어 긴 문단이 있는 열이 표를
      독점하지 않게 한다.
    """
    n = t.n_cols
    if n <= 1:
        return None

    min_w = [1.0] * n
    pref_w = [1.0] * n
    for row in t.rows:
        for ci, c in enumerate(row[:n]):
            if c.merged_placeholder or c.col_span > 1:
                continue  # 병합 셀은 한 열의 너비를 대표하지 못한다.
            text = c.text
            if not text.strip():
                continue
            longest_word = max(
                (_visual_len(w) for w in text.replace("\n", " ").split(" ")),
                default=0.0,
            )
            min_w[ci] = max(min_w[ci], min(longest_word, _MAX_PREF_CHARS))
            pref_w[ci] = max(pref_w[ci], min(_visual_len(text), _MAX_PREF_CHARS))

    widths = [max(m, p) for m, p in zip(min_w, pref_w)]
    if sum(widths) <= 0:
        return None
    return widths


def _visual_len(s: str) -> float:
    """CJK는 라틴의 약 2배 폭을 차지한다는 근사."""
    return sum(2.0 if ord(ch) > 0x1100 else 1.0 for ch in s)


def _runs_html(runs: list[Run], opts: ConvertOptions) -> str:
    parts: list[str] = []
    for r in runs:
        text = (
            _esc(r.text)
            .replace("\n", "<br/>")
            .replace("\t", "&#160;&#160;&#160;&#160;")
        )
        if not text:
            continue
        if not opts.keep_formatting:
            parts.append(text)
            continue
        style: list[str] = []
        if r.size_pt:
            style.append(f"font-size:{r.size_pt:.1f}pt")
        if r.color:
            style.append(f"color:{units.rgb_to_hex(r.color)}")
        if r.strike:
            style.append("text-decoration:line-through")
        elif r.underline:
            style.append("text-decoration:underline")
        if style:
            text = f'<span style="{";".join(style)}">{text}</span>'
        if r.bold:
            text = f"<b>{text}</b>"
        if r.italic:
            text = f"<i>{text}</i>"
        parts.append(text)
    return "".join(parts)


def _esc(s: str) -> str:
    return _html.escape(s, quote=False)


# --------------------------------------------------------------------------
# 폴백: 저수준 텍스트 배치
# --------------------------------------------------------------------------


def _write_fallback(
    fitz,
    doc: Document,
    path: Path,
    opts: ConvertOptions,
    font: Optional[fonts.FontPair],
) -> None:
    """Story가 안 될 때 최소한 텍스트는 남긴다. 서식은 포기한다."""
    pdf = fitz.open()
    fontname = "kfont"
    fontfile = None
    fontindex = 0
    if font is not None:
        fontfile = str(font.regular)
        fontindex = font.regular_index

    size = opts.pdf_base_size_pt
    left, top, right, bottom = doc.margin_pt
    width, height = doc.page_width_pt, doc.page_height_pt
    lines = _plain_lines(doc, opts)

    page = None
    y = height  # 첫 루프에서 새 페이지를 만들도록
    for line in lines:
        if page is None or y > height - bottom:
            page = pdf.new_page(width=width, height=height)
            if fontfile:
                page.insert_font(fontname=fontname, fontfile=fontfile, fontbuffer=None)
            y = top + size
        try:
            page.insert_text(
                (left, y),
                line,
                fontname=fontname if fontfile else "helv",
                fontfile=fontfile,
                fontsize=size,
            )
        except Exception:
            page.insert_text((left, y), line, fontsize=size)
        y += size * 1.6
    if page is None:
        pdf.new_page(width=width, height=height)
    pdf.save(str(path))
    pdf.close()


def _plain_lines(doc: Document, opts: ConvertOptions) -> list[str]:
    out: list[str] = []
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            out.extend(_wrap(b.text, 90))
        elif isinstance(b, Table):
            for row in b.to_grid():
                out.extend(_wrap(opts.table_text_sep.join(row), 90))
            out.append("")
        elif isinstance(b, Image):
            out.append(f"[그림: {b.name or 'image'}]")
    return out or [""]


def _wrap(text: str, width: int) -> list[str]:
    """CJK를 고려한 단순 줄바꿈. 폴백 전용이라 정밀도보다 안정성을 택한다."""
    if not text:
        return [""]
    out: list[str] = []
    cur = ""
    cur_w = 0
    for ch in text:
        w = 2 if ord(ch) > 0x1100 else 1
        if cur_w + w > width:
            out.append(cur)
            cur, cur_w = "", 0
        cur += ch
        cur_w += w
    if cur:
        out.append(cur)
    return out


def _finalize(fitz, data: bytes, path: Path, doc: Document) -> None:
    """조판 결과 바이트를 받아 서브셋/메타데이터를 적용하고 저장한다.

    **글꼴 서브셋이 이 함수의 존재 이유다.** 한글 글꼴은 글리프가 2만 자를
    넘어 파일 자체가 10MB를 우습게 넘는데, PDF는 임베딩 시 글꼴 파일을 통째로
    넣는다. 실측에서 맑은 고딕 Regular(13.4MB) + Bold(12.6MB)가 그대로 들어가
    26MB짜리 PDF가 나왔다. 문서에 실제 쓰인 글자만 남기면 236KB로 줄어든다.
    garbage collect만으로는 14.8MB에 그쳐, 서브셋이 유일한 해법이다.
    """
    pdf = fitz.open("pdf", data)
    try:
        # 순서가 중요하다. 쪽 번호를 **서브셋보다 먼저** 그려야 한다.
        # 서브셋 뒤에 텍스트를 넣으면 그 글꼴이 통째로 다시 임베딩되어
        # 236KB짜리 결과가 7.6MB로 되돌아간다(실측).
        _stamp_page_numbers(pdf, doc)

        try:
            pdf.subset_fonts(verbose=False)
        except TypeError:
            pdf.subset_fonts()
        except Exception:
            pass  # 서브셋 실패해도 문서 자체는 정상이다. 크기만 커진다.

        pdf.set_metadata(
            {
                "title": doc.meta.title or "",
                "author": doc.meta.author or "",
                "subject": doc.meta.subject or "",
                "keywords": doc.meta.keywords or "",
                "producer": "docconv",
                "creator": "docconv",
            }
        )
        pdf.save(str(path), garbage=4, deflate=True, clean=True)
    finally:
        pdf.close()


def _stamp_page_numbers(pdf, doc: Document) -> None:
    """쪽 번호를 각 페이지에 직접 그린다.

    Story는 머리말/꼬리말을 지원하지 않으므로 조판이 끝난 뒤 얹는다. 한/글
    문서는 쪽 번호를 흔히 쓰기 때문에, 빠지면 결과물이 눈에 띄게 달라 보인다.
    """
    spec = doc.page_number
    if spec is None:
        return

    fp = fonts.find_korean_font()
    fontfile = str(fp.regular) if fp is not None and not fp.is_collection else None
    size = 9.0
    left, top, right, bottom = doc.margin_pt
    pos = (spec.position or "BOTTOM_CENTER").upper()

    for i, page in enumerate(pdf):
        num = spec.start + i
        label = (
            f"{spec.side_char} {num} {spec.side_char}".strip()
            if spec.side_char
            else str(num)
        )
        w, h = page.rect.width, page.rect.height
        # 여백 안쪽에 배치한다. 본문과 겹치지 않도록 여백의 절반 지점에 둔다.
        y = h - bottom * 0.45 if pos.startswith("BOTTOM") else top * 0.6
        text_w = size * 0.55 * len(label)
        if pos.endswith("LEFT"):
            x = left
        elif pos.endswith("RIGHT"):
            x = w - right - text_w
        else:
            x = (w - text_w) / 2.0
        try:
            page.insert_text(
                (x, y), label, fontname="pgnum", fontfile=fontfile, fontsize=size
            )
        except Exception:
            try:
                page.insert_text((x, y), label, fontsize=size)
            except Exception:
                return
