"""텍스트 계열 Reader: TXT / Markdown / HTML / RTF / ODT.

핵심 8종은 아니지만 IR 파이프라인 덕분에 구현 비용이 낮고, 실무에서
"일단 텍스트로 뽑아 보고 싶다" 같은 요구를 잘 받아준다.
"""

from __future__ import annotations

import html as _html
import re
from pathlib import Path
from typing import Any, Iterator, Optional

from ..errors import CorruptFileError, ReadError
from ..ir import (
    Align,
    Block,
    Cell,
    DocKind,
    DocMeta,
    Document,
    ListKind,
    Paragraph,
    Run,
    Table,
)
from .sheet_reader import detect_encoding


def _read_text(path: Path) -> tuple[str, str]:
    raw = path.read_bytes()
    enc = detect_encoding(raw)
    return raw.decode(enc, errors="replace"), enc


# --------------------------------------------------------------------------
# TXT
# --------------------------------------------------------------------------


def read_txt(path: Path | str) -> Document:
    p = Path(path)
    text, enc = _read_text(p)
    doc = Document(
        meta=DocMeta(title=p.stem, source_format="txt", extra={"encoding": enc})
    )
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        doc.blocks.append(Paragraph.of(line))
    doc.compact()
    return doc


# --------------------------------------------------------------------------
# Markdown
# --------------------------------------------------------------------------

_MD_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
_MD_BULLET = re.compile(r"^(\s*)[-*+]\s+(.*)$")
_MD_NUMBER = re.compile(r"^(\s*)\d+[.)]\s+(.*)$")
_MD_TABLE_SEP = re.compile(r"^\s*\|?[\s:|-]+\|[\s:|-]*$")
_MD_BOLD = re.compile(r"\*\*(.+?)\*\*|__(.+?)__")
_MD_ITALIC = re.compile(r"(?<!\*)\*([^*]+)\*(?!\*)|(?<!_)_([^_]+)_(?!_)")
_MD_CODE = re.compile(r"`([^`]+)`")
_MD_LINK = re.compile(r"\[([^\]]*)\]\(([^)]+)\)")


def read_md(path: Path | str) -> Document:
    p = Path(path)
    text, enc = _read_text(p)
    lines = text.replace("\r\n", "\n").split("\n")
    doc = Document(
        meta=DocMeta(title=p.stem, source_format="md", extra={"encoding": enc})
    )

    i = 0
    in_code = False
    while i < len(lines):
        line = lines[i]

        if line.strip().startswith("```"):
            in_code = not in_code
            i += 1
            continue
        if in_code:
            doc.blocks.append(Paragraph(runs=[Run(text=line, font="Consolas")]))
            i += 1
            continue

        # 표: 헤더 줄 + 구분 줄 + 본문
        if "|" in line and i + 1 < len(lines) and _MD_TABLE_SEP.match(lines[i + 1]):
            tbl, consumed = _md_table(lines, i)
            doc.blocks.append(tbl)
            i += consumed
            continue

        m = _MD_HEADING.match(line)
        if m:
            para = _md_inline(m.group(2))
            para.heading = len(m.group(1))
            doc.blocks.append(para)
            i += 1
            continue

        m = _MD_BULLET.match(line) or _MD_NUMBER.match(line)
        if m:
            para = _md_inline(m.group(2))
            para.list_kind = (
                ListKind.BULLET if _MD_BULLET.match(line) else ListKind.NUMBER
            )
            para.list_level = len(m.group(1)) // 2
            doc.blocks.append(para)
            i += 1
            continue

        doc.blocks.append(_md_inline(line))
        i += 1

    doc.compact()
    if doc.blocks:
        for b in doc.blocks:
            if isinstance(b, Paragraph) and b.heading == 1:
                doc.meta.title = b.text.strip()
                break
    return doc


def _md_table(lines: list[str], start: int) -> tuple[Table, int]:
    def split_row(s: str) -> list[str]:
        s = s.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|"):
            s = s[:-1]
        return [c.strip() for c in s.split("|")]

    tbl = Table(header_row=True)
    tbl.rows.append([Cell(blocks=[_md_inline(c)]) for c in split_row(lines[start])])
    i = start + 2
    while i < len(lines) and "|" in lines[i] and lines[i].strip():
        tbl.rows.append([Cell(blocks=[_md_inline(c)]) for c in split_row(lines[i])])
        i += 1
    tbl.normalize()
    return tbl, i - start


def _md_inline(text: str) -> Paragraph:
    """굵게/기울임/코드/링크를 Run으로 분해한다."""
    para = Paragraph()
    # 링크를 먼저 텍스트로 치환하되 href를 기억한다.
    segments: list[tuple[str, dict[str, Any]]] = []
    pos = 0
    for m in _MD_LINK.finditer(text):
        if m.start() > pos:
            segments.append((text[pos : m.start()], {}))
        segments.append((m.group(1), {"href": m.group(2)}))
        pos = m.end()
    if pos < len(text):
        segments.append((text[pos:], {}))
    if not segments:
        segments = [(text, {})]

    for seg, attrs in segments:
        for chunk, style in _split_emphasis(seg):
            if not chunk:
                continue
            para.runs.append(Run(text=chunk, **style, **attrs))
    if not para.runs:
        para.runs.append(Run(text=""))
    para.merge_runs()
    return para


def _split_emphasis(s: str) -> Iterator[tuple[str, dict[str, Any]]]:
    """** __ * _ ` 를 처리한다. 중첩은 지원하지 않는다(실무상 충분)."""
    pattern = re.compile(
        r"\*\*(?P<b>.+?)\*\*|__(?P<b2>.+?)__|"
        r"(?<!\*)\*(?P<i>[^*]+)\*(?!\*)|(?<!_)_(?P<i2>[^_]+)_(?!_)|"
        r"`(?P<c>[^`]+)`"
    )
    pos = 0
    for m in pattern.finditer(s):
        if m.start() > pos:
            yield s[pos : m.start()], {}
        if m.group("b") is not None:
            yield m.group("b"), {"bold": True}
        elif m.group("b2") is not None:
            yield m.group("b2"), {"bold": True}
        elif m.group("i") is not None:
            yield m.group("i"), {"italic": True}
        elif m.group("i2") is not None:
            yield m.group("i2"), {"italic": True}
        elif m.group("c") is not None:
            yield m.group("c"), {"font": "Consolas"}
        pos = m.end()
    if pos < len(s):
        yield s[pos:], {}


# --------------------------------------------------------------------------
# HTML
# --------------------------------------------------------------------------

_BLOCK_TAGS = {
    "p",
    "div",
    "section",
    "article",
    "li",
    "blockquote",
    "pre",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
}
_SKIP_TAGS = {"script", "style", "head", "meta", "link", "noscript", "svg"}


def read_html(path: Path | str) -> Document:
    p = Path(path)
    text, enc = _read_text(p)
    try:
        from bs4 import BeautifulSoup
    except ImportError as e:
        raise ReadError("HTML을 읽으려면 beautifulsoup4가 필요합니다.") from e

    soup = BeautifulSoup(text, "lxml" if _has_lxml() else "html.parser")
    for t in soup(list(_SKIP_TAGS)):
        t.decompose()

    doc = Document(
        meta=DocMeta(
            title=(
                soup.title.string.strip()
                if soup.title and soup.title.string
                else p.stem
            ),
            source_format="html",
            extra={"encoding": enc},
        )
    )
    body = soup.body or soup
    for block in _html_blocks(body):
        doc.blocks.append(block)
    doc.compact()
    return doc


def _has_lxml() -> bool:
    try:
        import lxml  # noqa: F401

        return True
    except ImportError:
        return False


def _html_blocks(node) -> Iterator[Block]:
    from bs4 import NavigableString, Tag

    for child in node.children:
        if isinstance(child, NavigableString):
            s = str(child).strip()
            if s:
                yield Paragraph.of(_html.unescape(s))
            continue
        if not isinstance(child, Tag):
            continue
        name = child.name.lower()
        if name in _SKIP_TAGS:
            continue
        if name == "table":
            yield _html_table(child)
            continue
        if name in _BLOCK_TAGS:
            para = _html_paragraph(child)
            if para.runs:
                if name.startswith("h") and len(name) == 2 and name[1].isdigit():
                    para.heading = int(name[1])
                if name == "li":
                    para.list_kind = ListKind.BULLET
                yield para
            # 블록 안에 중첩 블록/표가 있으면 이어서 처리
            for sub in child.find_all(["table"], recursive=False):
                yield _html_table(sub)
            continue
        yield from _html_blocks(child)


def _html_paragraph(tag) -> Paragraph:
    from bs4 import NavigableString, Tag

    para = Paragraph()

    def walk(node, style: dict[str, Any]) -> None:
        for ch in node.children:
            if isinstance(ch, NavigableString):
                s = _html.unescape(str(ch))
                if s.strip():
                    para.runs.append(Run(text=re.sub(r"\s+", " ", s), **style))
                continue
            if not isinstance(ch, Tag):
                continue
            nm = ch.name.lower()
            if nm in ("table", "ul", "ol"):
                continue
            st = dict(style)
            if nm in ("b", "strong"):
                st["bold"] = True
            elif nm in ("i", "em"):
                st["italic"] = True
            elif nm == "u":
                st["underline"] = True
            elif nm in ("s", "strike", "del"):
                st["strike"] = True
            elif nm == "a":
                href = ch.get("href")
                if href:
                    st["href"] = href
            elif nm == "br":
                para.runs.append(Run(text="\n"))
                continue
            elif nm in ("code", "kbd", "samp"):
                st["font"] = "Consolas"
            walk(ch, st)

    walk(tag, {})
    para.merge_runs()
    return para


def _html_table(tag) -> Table:
    tbl = Table()
    for tr in tag.find_all("tr"):
        row: list[Cell] = []
        for td in tr.find_all(["td", "th"], recursive=False) or tr.find_all(
            ["td", "th"]
        ):
            c = Cell(blocks=[_html_paragraph(td)])
            try:
                c.col_span = max(1, int(td.get("colspan", 1)))
                c.row_span = max(1, int(td.get("rowspan", 1)))
            except (TypeError, ValueError):
                pass
            row.append(c)
        if row:
            tbl.rows.append(row)
    if tbl.rows and tag.find("th"):
        tbl.header_row = True
    tbl.normalize()
    return tbl


# --------------------------------------------------------------------------
# RTF
# --------------------------------------------------------------------------


def read_rtf(path: Path | str) -> Document:
    p = Path(path)
    try:
        from striprtf.striprtf import rtf_to_text
    except ImportError as e:
        raise ReadError("RTF를 읽으려면 striprtf가 필요합니다.") from e

    raw = p.read_bytes()
    try:
        text = rtf_to_text(raw.decode("cp949", errors="replace"), errors="ignore")
    except Exception:
        text = rtf_to_text(raw.decode("latin-1", errors="replace"), errors="ignore")

    doc = Document(meta=DocMeta(title=p.stem, source_format="rtf"))
    for line in text.replace("\r\n", "\n").split("\n"):
        doc.blocks.append(Paragraph.of(line.rstrip()))
    doc.compact()
    return doc


# --------------------------------------------------------------------------
# ODT
# --------------------------------------------------------------------------

_ODF_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
_ODF_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_ODF_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"


def read_odt(path: Path | str) -> Document:
    from ..util.safety import SafeZip

    p = Path(path)
    with SafeZip(p) as z:
        if not z.has("content.xml"):
            raise CorruptFileError("ODT에 content.xml이 없습니다.", detail=str(p))
        root = z.read_xml("content.xml")

    doc = Document(meta=DocMeta(title=p.stem, source_format="odt"))
    body = root.find(".//{%s}text" % _ODF_OFFICE)
    if body is None:
        return doc

    for el in body:
        tag = el.tag.split("}")[-1] if isinstance(el.tag, str) else ""
        if tag in ("p", "h"):
            para = Paragraph.of("".join(el.itertext()))
            if tag == "h":
                lvl = el.get("{%s}outline-level" % _ODF_TEXT)
                para.heading = int(lvl) if lvl and lvl.isdigit() else 1
            doc.blocks.append(para)
        elif tag == "table":
            tbl = Table(name=el.get("{%s}name" % _ODF_TABLE))
            for tr in el.iter("{%s}table-row" % _ODF_TABLE):
                row = [
                    Cell.of("".join(tc.itertext()))
                    for tc in tr.iter("{%s}table-cell" % _ODF_TABLE)
                ]
                if row:
                    tbl.rows.append(row)
            tbl.normalize()
            if tbl.rows:
                doc.blocks.append(tbl)
        elif tag == "list":
            for li in el.iter("{%s}list-item" % _ODF_TEXT):
                para = Paragraph.of("".join(li.itertext()))
                para.list_kind = ListKind.BULLET
                doc.blocks.append(para)
    doc.compact()
    return doc
