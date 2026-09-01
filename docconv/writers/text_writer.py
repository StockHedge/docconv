"""텍스트 계열 Writer: TXT / Markdown / HTML / JSON.

TXT와 HTML은 인코딩을 UTF-8로 고정한다. HTML은 `<meta charset>` 을 반드시
넣는다 - 없으면 브라우저가 로캘 기본 인코딩으로 추측해 한글이 깨진다.
"""

from __future__ import annotations

import html as _html
import json
from pathlib import Path
from typing import Any, Iterable

from ..errors import WriteError
from ..ir import (
    Align,
    Block,
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
from ..util import units


# --------------------------------------------------------------------------
# TXT
# --------------------------------------------------------------------------


def write_txt(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    lines: list[str] = []
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            prefix = ""
            if b.list_kind is ListKind.BULLET:
                prefix = "  " * b.list_level + "- "
            elif b.list_kind is ListKind.NUMBER:
                prefix = "  " * b.list_level + "1. "
            lines.append(prefix + b.text)
        elif isinstance(b, Table):
            if b.name:
                lines.append(f"[{b.name}]")
            lines.extend(_table_text(b, opts.table_text_sep))
            lines.append("")
        elif isinstance(b, Image):
            lines.append(f"[그림: {b.name or 'image'}]")
        elif isinstance(b, PageBreak) and opts.keep_page_breaks:
            lines.append("\f")
    _write(path, "\n".join(lines) + "\n", "utf-8")


def _table_text(t: Table, sep: str) -> list[str]:
    return [sep.join(c.replace("\n", " ") for c in row) for row in t.to_grid()]


# --------------------------------------------------------------------------
# Markdown
# --------------------------------------------------------------------------


def write_md(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    out: list[str] = []
    if doc.meta.title:
        out.append(f"# {doc.meta.title}\n")

    for b in doc.blocks:
        if isinstance(b, Paragraph):
            out.append(_md_paragraph(b, opts))
        elif isinstance(b, Table):
            if b.name:
                out.append(f"\n**{b.name}**\n")
            out.append(_md_table(b))
        elif isinstance(b, Image):
            out.append(f"![{b.alt_text or b.name or 'image'}]({b.name or 'image'})\n")
        elif isinstance(b, PageBreak) and opts.keep_page_breaks:
            out.append("\n---\n")
    _write(path, "\n".join(out).replace("\n\n\n", "\n\n") + "\n", "utf-8")


def _md_paragraph(p: Paragraph, opts: ConvertOptions) -> str:
    text = "".join(_md_run(r, opts) for r in p.runs).strip()
    if not text:
        return ""
    if p.heading:
        return f"\n{'#' * min(6, p.heading)} {text}\n"
    if p.list_kind is ListKind.BULLET:
        return f"{'  ' * p.list_level}- {text}"
    if p.list_kind is ListKind.NUMBER:
        return f"{'  ' * p.list_level}1. {text}"
    return text + "\n"


def _md_run(r: Run, opts: ConvertOptions) -> str:
    t = r.text.replace("\n", "  \n")
    if not opts.keep_formatting or not t.strip():
        return t
    # 마크다운 특수문자 최소 이스케이프
    for ch in ("\\", "`", "*", "_"):
        t = t.replace(ch, "\\" + ch)
    if r.bold and r.italic:
        t = f"***{t}***"
    elif r.bold:
        t = f"**{t}**"
    elif r.italic:
        t = f"*{t}*"
    if r.strike:
        t = f"~~{t}~~"
    if r.href:
        t = f"[{t}]({r.href})"
    return t


def _md_table(t: Table) -> str:
    grid = t.to_grid()
    if not grid:
        return ""
    grid = [[c.replace("\n", " ").replace("|", "\\|") for c in row] for row in grid]
    n = len(grid[0])
    lines = ["| " + " | ".join(grid[0]) + " |", "|" + "|".join([" --- "] * n) + "|"]
    for row in grid[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines) + "\n"


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------


def write_html(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    import base64

    title = _html.escape(doc.meta.title or Path(path).stem)
    body: list[str] = []
    list_open: str | None = None

    for b in doc.blocks:
        if isinstance(b, Paragraph) and b.list_kind is not ListKind.NONE:
            tag = "ul" if b.list_kind is ListKind.BULLET else "ol"
            if list_open != tag:
                if list_open:
                    body.append(f"</{list_open}>")
                body.append(f"<{tag}>")
                list_open = tag
            body.append(f"<li>{_html_runs(b.runs, opts)}</li>")
            continue
        if list_open:
            body.append(f"</{list_open}>")
            list_open = None
        body.append(_html_block(b, opts, base64))
    if list_open:
        body.append(f"</{list_open}>")

    page = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{
    font-family: "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif;
    line-height: 1.7; max-width: 52rem; margin: 0 auto; padding: 2rem 1.25rem;
  }}
  table {{ border-collapse: collapse; width: 100%; margin: 1rem 0; }}
  th, td {{ border: 1px solid #bbb; padding: 0.45rem 0.6rem; vertical-align: top; text-align: left; }}
  th {{ background: rgba(128,128,128,0.14); font-weight: 600; }}
  img {{ max-width: 100%; height: auto; }}
  hr {{ border: none; border-top: 1px dashed #999; margin: 2rem 0; }}
  .page-break {{ break-after: page; }}
</style>
</head>
<body>
{chr(10).join(body)}
</body>
</html>
"""
    _write(path, page, "utf-8")


def _html_block(b: Block, opts: ConvertOptions, base64) -> str:
    if isinstance(b, Paragraph):
        inner = _html_runs(b.runs, opts)
        if not inner.strip():
            return "<p>&nbsp;</p>"
        style = ""
        if b.align is not Align.LEFT:
            style = f' style="text-align:{b.align.value}"'
        if b.heading:
            lvl = min(6, b.heading)
            return f"<h{lvl}{style}>{inner}</h{lvl}>"
        return f"<p{style}>{inner}</p>"

    if isinstance(b, Table):
        return _html_table(b, opts)

    if isinstance(b, Image):
        if not opts.include_images or not b.data:
            return f"<p>[그림: {_html.escape(b.name or 'image')}]</p>"
        mime = f"image/{'jpeg' if b.fmt in ('jpg', 'jpeg') else b.fmt}"
        data = base64.b64encode(b.data).decode("ascii")
        alt = _html.escape(b.alt_text or b.name or "image")
        return f'<p><img src="data:{mime};base64,{data}" alt="{alt}"></p>'

    if isinstance(b, PageBreak):
        return '<hr class="page-break">'
    return ""


def _html_table(t: Table, opts: ConvertOptions) -> str:
    t.normalize()
    out = ["<table>"]
    if t.name:
        out.append(f"<caption>{_html.escape(t.name)}</caption>")
    for ri, row in enumerate(t.rows):
        out.append("<tr>")
        for c in row:
            if c.merged_placeholder:
                continue
            tag = "th" if (t.header_row and ri == 0) else "td"
            attrs = ""
            if c.col_span > 1:
                attrs += f' colspan="{c.col_span}"'
            if c.row_span > 1:
                attrs += f' rowspan="{c.row_span}"'
            if c.background and opts.keep_formatting:
                attrs += f' style="background:{units.rgb_to_hex(c.background)}"'
            inner = "".join(
                _html_block(bb, opts, __import__("base64")) for bb in c.blocks
            )
            out.append(f"<{tag}{attrs}>{inner or '&nbsp;'}</{tag}>")
        out.append("</tr>")
    out.append("</table>")
    return "".join(out)


def _html_runs(runs: list[Run], opts: ConvertOptions) -> str:
    parts: list[str] = []
    for r in runs:
        t = _html.escape(r.text).replace("\n", "<br>").replace("\t", "&emsp;")
        if not t:
            continue
        if opts.keep_formatting:
            style: list[str] = []
            if r.size_pt:
                style.append(f"font-size:{r.size_pt:.1f}pt")
            if r.color:
                style.append(f"color:{units.rgb_to_hex(r.color)}")
            if r.highlight:
                style.append(f"background:{units.rgb_to_hex(r.highlight)}")
            if style:
                t = f'<span style="{";".join(style)}">{t}</span>'
            if r.bold:
                t = f"<strong>{t}</strong>"
            if r.italic:
                t = f"<em>{t}</em>"
            if r.underline:
                t = f"<u>{t}</u>"
            if r.strike:
                t = f"<s>{t}</s>"
            if r.superscript:
                t = f"<sup>{t}</sup>"
            elif r.subscript:
                t = f"<sub>{t}</sub>"
        if r.href:
            t = f'<a href="{_html.escape(r.href, quote=True)}">{t}</a>'
        parts.append(t)
    return "".join(parts)


# --------------------------------------------------------------------------
# JSON (IR 덤프)
# --------------------------------------------------------------------------


def write_json(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    """IR을 그대로 JSON으로. 디버깅과 외부 파이프라인 연계용."""

    def run(r: Run) -> dict[str, Any]:
        d: dict[str, Any] = {"text": r.text}
        for k in ("bold", "italic", "underline", "strike", "superscript", "subscript"):
            if getattr(r, k):
                d[k] = True
        for k in ("size_pt", "font", "href"):
            v = getattr(r, k)
            if v:
                d[k] = v
        if r.color:
            d["color"] = units.rgb_to_hex(r.color)
        return d

    def cell(c: Cell) -> dict[str, Any]:
        d: dict[str, Any] = {"blocks": [block(b) for b in c.blocks]}
        if c.col_span > 1:
            d["col_span"] = c.col_span
        if c.row_span > 1:
            d["row_span"] = c.row_span
        if c.merged_placeholder:
            d["merged"] = True
        if c.raw_value is not None and not isinstance(
            c.raw_value, (str, int, float, bool)
        ):
            d["raw_value"] = str(c.raw_value)
        elif c.raw_value is not None:
            d["raw_value"] = c.raw_value
        return d

    def block(b: Block) -> dict[str, Any]:
        if isinstance(b, Paragraph):
            d: dict[str, Any] = {"type": "paragraph", "runs": [run(r) for r in b.runs]}
            if b.heading:
                d["heading"] = b.heading
            if b.align is not Align.LEFT:
                d["align"] = b.align.value
            if b.list_kind is not ListKind.NONE:
                d["list"] = b.list_kind.value
                d["list_level"] = b.list_level
            return d
        if isinstance(b, Table):
            return {
                "type": "table",
                "name": b.name,
                "header_row": b.header_row,
                "rows": [[cell(c) for c in row] for row in b.rows],
            }
        if isinstance(b, Image):
            return {
                "type": "image",
                "format": b.fmt,
                "name": b.name,
                "bytes": len(b.data),
                "width_pt": b.width_pt,
                "height_pt": b.height_pt,
            }
        if isinstance(b, PageBreak):
            return {"type": "page_break"}
        return {"type": "unknown"}

    payload = {
        "meta": {
            k: v
            for k, v in {
                "title": doc.meta.title,
                "author": doc.meta.author,
                "subject": doc.meta.subject,
                "keywords": doc.meta.keywords,
                "created": doc.meta.created,
                "modified": doc.meta.modified,
                "source_format": doc.meta.source_format,
                "extra": doc.meta.extra or None,
            }.items()
            if v
        },
        "kind": doc.kind.value,
        "page": {
            "width_pt": doc.page_width_pt,
            "height_pt": doc.page_height_pt,
            "margin_pt": list(doc.margin_pt),
        },
        "stats": doc.stats(),
        "blocks": [block(b) for b in doc.blocks],
    }
    _write(path, json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")


# --------------------------------------------------------------------------


def _write(path: Path | str, text: str, encoding: str) -> None:
    try:
        Path(path).write_text(text, encoding=encoding, newline="")
    except OSError as e:
        raise WriteError(f"파일을 저장할 수 없습니다: {path}", detail=str(e)) from e
