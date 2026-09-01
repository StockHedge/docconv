"""DOCX Writer.

한글 관련 함정
--------------
python-docx의 ``run.font.name = "맑은 고딕"`` 은 OOXML의 ``w:rFonts`` 중
``w:ascii`` 와 ``w:hAnsi`` 만 설정한다. 그런데 Word는 문자 계열별로 글꼴을
따로 고른다 - 한글/한자/가나는 ``w:eastAsia`` 속성을 본다. 그래서 이 속성을
직접 써 주지 않으면 한글만 엉뚱한 기본 글꼴(대개 굴림)로 렌더링된다.
`_set_font()` 가 세 속성을 함께 설정한다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from ..errors import WriteError
from ..ir import (
    Align,
    Block,
    Cell,
    DocKind,
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

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag: str) -> str:
    return "{%s}%s" % (W_NS, tag)


def write_docx(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    try:
        import docx
        from docx.enum.table import WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
        from docx.oxml.ns import qn
        from docx.shared import Pt, RGBColor
    except ImportError as e:  # pragma: no cover
        raise WriteError("python-docx가 필요합니다. pip install python-docx") from e

    d = docx.Document()
    _setup_page(d, doc)
    _setup_default_font(d, qn)

    align_map = {
        Align.LEFT: WD_ALIGN_PARAGRAPH.LEFT,
        Align.CENTER: WD_ALIGN_PARAGRAPH.CENTER,
        Align.RIGHT: WD_ALIGN_PARAGRAPH.RIGHT,
        Align.JUSTIFY: WD_ALIGN_PARAGRAPH.JUSTIFY,
        Align.DISTRIBUTE: WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    }

    ctx = _Ctx(d, opts, align_map, qn, Pt, RGBColor, WD_BREAK)

    for b in doc.blocks:
        _write_block(ctx, b, doc)

    _set_core_props(d, doc)
    try:
        d.save(str(path))
    except OSError as e:
        raise WriteError(f"파일을 저장할 수 없습니다: {path}", detail=str(e)) from e


class _Ctx:
    """Writer가 들고 다니는 것들. 인자 폭발을 막는다."""

    def __init__(self, d, opts, align_map, qn, Pt, RGBColor, WD_BREAK) -> None:
        self.d = d
        self.opts = opts
        self.align_map = align_map
        self.qn = qn
        self.Pt = Pt
        self.RGBColor = RGBColor
        self.WD_BREAK = WD_BREAK
        fp = fonts.find_korean_font()
        self.default_font = fp.name if fp else None


# --------------------------------------------------------------------------
# 블록
# --------------------------------------------------------------------------


def _write_block(ctx: _Ctx, b: Block, doc: Document, container=None) -> None:
    target = container if container is not None else ctx.d
    if isinstance(b, Paragraph):
        _write_paragraph(ctx, b, target)
    elif isinstance(b, Table):
        _write_table(ctx, b, target)
    elif isinstance(b, Image):
        if ctx.opts.include_images:
            _write_image(ctx, b, target)
    elif isinstance(b, PageBreak):
        if ctx.opts.keep_page_breaks and container is None:
            ctx.d.add_paragraph().add_run().add_break(ctx.WD_BREAK.PAGE)


def _write_paragraph(ctx: _Ctx, p: Paragraph, target) -> None:
    style = None
    if p.heading and 1 <= p.heading <= 9:
        style = f"Heading {p.heading}"
    elif p.list_kind is ListKind.BULLET:
        style = (
            "List Bullet"
            if p.list_level == 0
            else f"List Bullet {min(3, p.list_level + 1)}"
        )
    elif p.list_kind is ListKind.NUMBER:
        style = (
            "List Number"
            if p.list_level == 0
            else f"List Number {min(3, p.list_level + 1)}"
        )

    try:
        wp = target.add_paragraph(style=style) if style else target.add_paragraph()
    except KeyError:
        # 템플릿에 해당 스타일이 없으면 기본 문단으로 떨어뜨린다.
        wp = target.add_paragraph()

    wp.alignment = ctx.align_map.get(p.align)
    pf = wp.paragraph_format
    if p.indent_pt:
        pf.left_indent = ctx.Pt(p.indent_pt)
    if p.space_before_pt:
        pf.space_before = ctx.Pt(p.space_before_pt)
    if p.space_after_pt:
        pf.space_after = ctx.Pt(p.space_after_pt)
    if p.line_spacing:
        pf.line_spacing = p.line_spacing

    for r in p.runs:
        _write_run(ctx, wp, r, is_heading=bool(p.heading))


def _write_run(ctx: _Ctx, wp, r: Run, *, is_heading: bool = False) -> None:
    # 문단 안의 줄바꿈은 별도 run + break로 표현한다.
    parts = r.text.split("\n")
    for i, piece in enumerate(parts):
        wr = wp.add_run(piece)
        if i < len(parts) - 1:
            wr.add_break()
        if not ctx.opts.keep_formatting:
            continue
        f = wr.font
        if r.bold:
            f.bold = True
        if r.italic:
            f.italic = True
        if r.underline:
            f.underline = True
        if r.strike:
            f.strike = True
        if r.superscript:
            f.superscript = True
        if r.subscript:
            f.subscript = True
        if r.size_pt and not is_heading:
            f.size = ctx.Pt(max(1.0, min(409.0, r.size_pt)))
        if r.color:
            try:
                f.color.rgb = ctx.RGBColor(*r.color)
            except Exception:
                pass
        name = r.font or ctx.default_font
        if name:
            _set_font(wr, name, ctx.qn)


def _set_font(wr, name: str, qn) -> None:
    """ascii / hAnsi / eastAsia 세 계열에 같은 글꼴을 지정한다.

    eastAsia를 빼먹으면 한글만 다른 글꼴로 렌더링된다 - 이 모듈에서 가장
    자주 틀리는 부분이다.
    """
    wr.font.name = name
    rPr = wr._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _write_table(ctx: _Ctx, t: Table, target) -> None:
    t.normalize()
    if not t.rows or not t.n_cols:
        return

    n_rows, n_cols = t.n_rows, t.n_cols
    try:
        wt = target.add_table(rows=n_rows, cols=n_cols)
    except Exception:
        wt = ctx.d.add_table(rows=n_rows, cols=n_cols)
    try:
        wt.style = "Table Grid"
    except KeyError:
        pass

    # 내용을 먼저 전부 채우고, 병합은 그 다음에 한 번에 한다.
    #
    # 병합과 내용 쓰기를 한 루프에서 섞으면 안 된다. python-docx의
    # `Table.cell(r, c)` 는 tc 요소를 훑어 격자를 매번 재계산하는데, 중간에
    # 병합이 일어나면 그 이후 행의 좌표 매핑이 어긋난다. 실측에서 3행까지만
    # 채워지고 나머지 14행이 빈 셀로 남았다(11,889자 -> 86자).
    for ri, row in enumerate(t.rows):
        for ci, c in enumerate(row):
            if ci >= n_cols:
                continue
            try:
                wc = wt.cell(ri, ci)
            except (IndexError, ValueError):
                continue
            _clear_cell(wc)
            if c.merged_placeholder:
                # 병합될 자리다. 빈 문단 하나만 둔다(셀은 문단이 최소 1개 필요).
                wc.add_paragraph()
                continue
            wrote = False
            for b in c.blocks:
                if isinstance(b, Paragraph):
                    _write_paragraph(ctx, b, wc)
                    wrote = True
                elif isinstance(b, Table):
                    _write_table(ctx, b, wc)
                    wrote = True
                elif isinstance(b, Image) and ctx.opts.include_images:
                    _write_image(ctx, b, wc)
                    wrote = True
            if not wrote:
                wc.add_paragraph()
            if c.background and ctx.opts.keep_formatting:
                _shade_cell(wc, c.background, ctx.qn)

    _merge_spans(wt, t, n_rows, n_cols)

    if t.col_widths_pt and ctx.opts.keep_formatting:
        _apply_col_widths(wt, t.col_widths_pt, ctx.Pt, n_cols)

    if t.header_row:
        _mark_header_row(wt, ctx.qn)


def _merge_spans(wt, t: Table, n_rows: int, n_cols: int) -> None:
    """내용을 다 채운 뒤 병합을 적용한다.

    뒤쪽 좌표부터 처리하면, 앞쪽 셀을 병합해 격자가 줄어들어도 아직 처리하지
    않은 좌표가 영향을 덜 받는다. 병합 시 합쳐지는 빈 문단은 제거한다.
    """
    for ri in range(n_rows - 1, -1, -1):
        row = t.rows[ri] if ri < len(t.rows) else []
        for ci in range(min(n_cols, len(row)) - 1, -1, -1):
            c = row[ci]
            if c.merged_placeholder or (c.col_span <= 1 and c.row_span <= 1):
                continue
            r2 = min(n_rows - 1, ri + c.row_span - 1)
            c2 = min(n_cols - 1, ci + c.col_span - 1)
            if (r2, c2) == (ri, ci):
                continue
            try:
                merged = wt.cell(ri, ci).merge(wt.cell(r2, c2))
            except (IndexError, ValueError, AttributeError):
                continue
            _drop_trailing_empty_paragraphs(merged)


def _drop_trailing_empty_paragraphs(wc) -> None:
    """병합으로 딸려 들어온 빈 문단을 정리한다. 최소 하나는 남긴다."""
    paras = list(wc.paragraphs)
    for p in reversed(paras[1:]):
        if p.text.strip():
            break
        p._element.getparent().remove(p._element)


def _clear_cell(wc) -> None:
    for p in list(wc.paragraphs):
        p._element.getparent().remove(p._element)


def _shade_cell(wc, rgb, qn) -> None:
    hexv = units.rgb_to_hex(rgb, "") or "FFFFFF"
    tcPr = wc._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {})
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexv)
    tcPr.append(shd)


def _apply_col_widths(wt, widths: list[float], Pt, n_cols: int) -> None:
    try:
        wt.autofit = False
    except Exception:
        pass
    for ci in range(min(n_cols, len(widths))):
        w = widths[ci]
        if not w or w <= 0:
            continue
        for row in wt.rows:
            try:
                row.cells[ci].width = Pt(w)
            except (IndexError, Exception):
                break


def _mark_header_row(wt, qn) -> None:
    """머리글 행이 페이지마다 반복되도록 tblHeader를 설정한다."""
    try:
        tr = wt.rows[0]._tr
    except IndexError:
        return
    trPr = tr.get_or_add_trPr()
    el = trPr.makeelement(qn("w:tblHeader"), {})
    el.set(qn("w:val"), "true")
    trPr.append(el)
    for c in wt.rows[0].cells:
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold = True


def _write_image(ctx: _Ctx, img: Image, target) -> None:
    import io

    if not img.data:
        return
    try:
        stream = io.BytesIO(img.data)
        width = ctx.Pt(img.width_pt) if img.width_pt else None
        p = target.add_paragraph()
        run = p.add_run()
        if width:
            run.add_picture(stream, width=width)
        else:
            run.add_picture(stream)
    except Exception:
        # EMF/WMF 등 python-docx가 못 다루는 형식은 자리 표시만 남긴다.
        target.add_paragraph(f"[그림: {img.name or 'image'}]")


# --------------------------------------------------------------------------
# 문서 설정
# --------------------------------------------------------------------------


def _setup_page(d, doc: Document) -> None:
    from docx.shared import Pt

    try:
        s = d.sections[0]
    except IndexError:
        return
    s.page_width = Pt(doc.page_width_pt)
    s.page_height = Pt(doc.page_height_pt)
    left, top, right, bottom = doc.margin_pt
    s.left_margin = Pt(left)
    s.top_margin = Pt(top)
    s.right_margin = Pt(right)
    s.bottom_margin = Pt(bottom)


def _setup_default_font(d, qn) -> None:
    """Normal 스타일에 한글 글꼴을 심어 문서 전체 기본값으로 만든다."""
    fp = fonts.find_korean_font()
    if fp is None:
        return
    try:
        style = d.styles["Normal"]
    except KeyError:
        return
    style.font.name = fp.name
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), fp.name)


def _set_core_props(d, doc: Document) -> None:
    cp = d.core_properties
    if doc.meta.title:
        cp.title = doc.meta.title[:255]
    if doc.meta.author:
        cp.author = doc.meta.author[:255]
    if doc.meta.subject:
        cp.subject = doc.meta.subject[:255]
    if doc.meta.keywords:
        cp.keywords = doc.meta.keywords[:255]
