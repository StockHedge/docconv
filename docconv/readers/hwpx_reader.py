"""HWPX(OWPML) Reader.

HWPX는 ZIP 컨테이너 + XML로, 사실상 한글판 OOXML이다. 한컴오피스 설치 없이
순수 파이썬으로 완전히 읽을 수 있다.

    version.xml
    mimetype
    META-INF/container.xml, manifest.xml
    Contents/header.xml       스타일 정의(글꼴/글자모양/문단모양/스타일)
    Contents/section0.xml     본문 (구역마다 파일 하나)
    Contents/content.hpf      OPF 형식 매니페스트 + 메타데이터
    BinData/image1.png        임베디드 바이너리

구현하면서 실제 파일에서 확인한 함정
------------------------------------
* ``paraPr`` 의 여백/줄간격이 ``hp:switch > hp:case | hp:default`` 분기 안에
  들어 있다. 평범하게 ``paraPr/margin`` 으로 찾으면 전부 놓친다.
* 표에서 병합으로 가려진 칸은 ``<tc>`` 자체가 없다. 행마다 셀 수가 다르므로
  ``cellAddr`` + ``cellSpan`` 으로 그리드를 역산해야 한다.
* 글자 크기 ``height`` 는 1/100 pt 단위다(1200 = 12pt). 길이 계열 속성의
  HWPUNIT(1/7200인치)과 단위가 다르니 혼동하지 말 것.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator, Optional

from lxml import etree

from ..errors import CorruptFileError
from ..ir import (
    Align,
    Block,
    Border,
    BorderStyle,
    Borders,
    Cell,
    DocKind,
    DocMeta,
    Document,
    Image,
    ListKind,
    PageBreak,
    PageNumberSpec,
    Paragraph,
    Run,
    Table,
)
from ..util import units
from ..util.safety import SafeZip

# 네임스페이스 -------------------------------------------------------------
NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
}
HP = NS["hp"]
HH = NS["hh"]
HC = NS["hc"]
HS = NS["hs"]


def _q(ns: str, tag: str) -> str:
    return "{%s}%s" % (ns, tag)


def _local(el: etree._Element) -> str:
    return etree.QName(el).localname


def _f(v: str | None, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    try:
        return float(v)
    except ValueError:
        return default


def _i(v: str | None, default: int = 0) -> int:
    if v is None or v == "":
        return default
    try:
        return int(float(v))
    except ValueError:
        return default


# --------------------------------------------------------------------------
# header.xml - 스타일 테이블
# --------------------------------------------------------------------------


@dataclass
class CharStyle:
    """글자 모양(charPr)."""

    size_pt: float = 10.0
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: Optional[tuple[int, int, int]] = None
    highlight: Optional[tuple[int, int, int]] = None
    font: Optional[str] = None
    #: 위/아래 첨자
    superscript: bool = False
    subscript: bool = False


@dataclass
class ParaStyle:
    """문단 모양(paraPr)."""

    align: Align = Align.LEFT
    indent_pt: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    line_spacing: Optional[float] = None
    outline_level: int = 0
    has_bullet: bool = False
    has_number: bool = False


@dataclass
class FillStyle:
    """borderFill 항목: 셀 배경 + 4방향 테두리."""

    background: Optional[tuple[int, int, int]] = None
    borders: Optional[Borders] = None


@dataclass
class StyleTable:
    chars: dict[int, CharStyle] = field(default_factory=dict)
    paras: dict[int, ParaStyle] = field(default_factory=dict)
    #: styleIDRef -> (스타일 이름, paraPrIDRef, charPrIDRef)
    styles: dict[int, tuple[str, int, int]] = field(default_factory=dict)
    #: 언어별 글꼴 이름 (HANGUL 우선)
    fonts: dict[int, str] = field(default_factory=dict)
    #: borderFillIDRef -> 배경/테두리
    fills: dict[int, FillStyle] = field(default_factory=dict)

    def fill(self, idref: str | int | None) -> FillStyle:
        return self.fills.get(
            _i(str(idref) if idref is not None else None, -1), FillStyle()
        )

    def char(self, idref: str | int | None) -> CharStyle:
        return self.chars.get(
            _i(str(idref) if idref is not None else None, -1), CharStyle()
        )

    def para(self, idref: str | int | None) -> ParaStyle:
        return self.paras.get(
            _i(str(idref) if idref is not None else None, -1), ParaStyle()
        )

    def style_name(self, idref: str | int | None) -> Optional[str]:
        s = self.styles.get(_i(str(idref) if idref is not None else None, -1))
        return s[0] if s else None


_ALIGN_MAP = {
    "LEFT": Align.LEFT,
    "CENTER": Align.CENTER,
    "RIGHT": Align.RIGHT,
    "JUSTIFY": Align.JUSTIFY,
    "DISTRIBUTE": Align.DISTRIBUTE,
    "DISTRIBUTE_SPACE": Align.DISTRIBUTE,
}

# 서식 적용 여부는 **알려진 값일 때만 켠다**(화이트리스트).
#
# "NONE이 아니면 적용"으로 판정하면 모르는 값 하나에 문서 전체가 망가진다.
# 실제로 어떤 한/글 저장본은 취소선 없음을 `shape="NONE"` 이 아니라
# `shape="3D"` 로 적어 두는데, 블랙리스트 방식이면 **모든 글자에 취소선**이
# 그어진다. 스펙에 없는 값이 나와도 서식을 켜지 않는 쪽이 안전하다.

#: 밑줄을 실제로 긋는 위치 값 (LineType/위치 열거)
_UNDERLINE_ON = frozenset({"BOTTOM", "CENTER", "TOP"})

#: 취소선을 실제로 긋는 선 모양 값 (LineType2 열거)
_STRIKE_ON = frozenset(
    {
        "SOLID",
        "DASH",
        "DOT",
        "DASH_DOT",
        "DASH_DOT_DOT",
        "LONG_DASH",
        "CIRCLE",
        "DOUBLE_SLIM",
        "SLIM_THICK",
        "THICK_SLIM",
        "SLIM_THICK_SLIM",
    }
)


def _parse_header(root: etree._Element) -> StyleTable:
    tbl = StyleTable()

    # -- 글꼴: 한글 계열을 우선하고, 없으면 처음 만난 것을 쓴다 ------------
    for ff in root.iter(_q(HH, "fontface")):
        lang = (ff.get("lang") or "").upper()
        prefer = lang in ("HANGUL", "")
        for fnt in ff.iter(_q(HH, "font")):
            fid = _i(fnt.get("id"), -1)
            face = fnt.get("face")
            if fid < 0 or not face:
                continue
            if prefer or fid not in tbl.fonts:
                tbl.fonts[fid] = face

    # -- 글자 모양 --------------------------------------------------------
    for cp in root.iter(_q(HH, "charPr")):
        cid = _i(cp.get("id"), -1)
        if cid < 0:
            continue
        cs = CharStyle()
        # height는 1/100 pt
        cs.size_pt = _f(cp.get("height"), 1000.0) / 100.0
        cs.color = units.hwp_color_to_rgb(cp.get("textColor"))
        shade = cp.get("shadeColor")
        if shade and shade.lower() not in ("none", "#ffffff"):
            cs.highlight = units.hwp_color_to_rgb(shade)

        for ch in cp:
            if not isinstance(ch.tag, str):
                continue
            name = _local(ch)
            if name == "bold":
                cs.bold = True
            elif name == "italic":
                cs.italic = True
            elif name == "underline":
                cs.underline = (ch.get("type") or "").upper() in _UNDERLINE_ON
            elif name == "strikeout":
                cs.strike = (ch.get("shape") or "").upper() in _STRIKE_ON
            elif name == "fontRef":
                fid = _i(ch.get("hangul"), _i(ch.get("latin"), -1))
                cs.font = tbl.fonts.get(fid)
            elif name == "offset":
                # 위/아래 첨자는 offset의 부호로 표현되기도 한다.
                off = _i(ch.get("hangul"), 0)
                if off > 0:
                    cs.superscript = True
                elif off < 0:
                    cs.subscript = True
        # 명시적 sub/sup 속성
        sub = (cp.get("subscript") or "0") in ("1", "true")
        sup = (cp.get("superscript") or "0") in ("1", "true")
        cs.subscript = cs.subscript or sub
        cs.superscript = cs.superscript or sup
        tbl.chars[cid] = cs

    # -- 문단 모양 --------------------------------------------------------
    for pp in root.iter(_q(HH, "paraPr")):
        pid = _i(pp.get("id"), -1)
        if pid < 0:
            continue
        ps = ParaStyle()

        al = pp.find(_q(HH, "align"))
        if al is not None:
            ps.align = _ALIGN_MAP.get((al.get("horizontal") or "").upper(), Align.LEFT)

        hd = pp.find(_q(HH, "heading"))
        if hd is not None:
            htype = (hd.get("type") or "NONE").upper()
            lvl = _i(hd.get("level"), 0)
            if htype == "OUTLINE":
                ps.outline_level = lvl + 1
            elif htype == "BULLET":
                ps.has_bullet = True
            elif htype == "NUMBER":
                ps.has_number = True

        # margin / lineSpacing 은 switch>default 안에 있을 수 있다.
        margin = _find_in_switch(pp, HH, "margin")
        if margin is not None:
            ps.indent_pt = units.hwp_to_pt(_first_attr(margin, HC, "left", "value"))
            ps.space_before_pt = units.hwp_to_pt(
                _first_attr(margin, HC, "prev", "value")
            )
            ps.space_after_pt = units.hwp_to_pt(
                _first_attr(margin, HC, "next", "value")
            )

        ls = _find_in_switch(pp, HH, "lineSpacing")
        if ls is not None:
            lstype = (ls.get("type") or "").upper()
            val = _f(ls.get("value"), 0.0)
            if lstype == "PERCENT" and val > 0:
                ps.line_spacing = val / 100.0
        tbl.paras[pid] = ps

    # -- 테두리/배경 ------------------------------------------------------
    for bf in root.iter(_q(HH, "borderFill")):
        bid = _i(bf.get("id"), -1)
        if bid < 0:
            continue
        tbl.fills[bid] = _parse_border_fill(bf)

    # -- 스타일 -----------------------------------------------------------
    for st in root.iter(_q(HH, "style")):
        sid = _i(st.get("id"), -1)
        if sid < 0:
            continue
        tbl.styles[sid] = (
            st.get("name") or "",
            _i(st.get("paraPrIDRef"), -1),
            _i(st.get("charPrIDRef"), -1),
        )
    return tbl


#: 한/글 테두리 종류 -> IR 표현. 실선 변형(이중/굵은 등)은 가장 가까운 것으로 접는다.
_BORDER_STYLE_MAP = {
    "NONE": BorderStyle.NONE,
    "SOLID": BorderStyle.SOLID,
    "THICK": BorderStyle.SOLID,
    "DASH": BorderStyle.DASHED,
    "DASHED": BorderStyle.DASHED,
    "DOT": BorderStyle.DOTTED,
    "DOTTED": BorderStyle.DOTTED,
    "DASH_DOT": BorderStyle.DASHED,
    "DASH_DOT_DOT": BorderStyle.DASHED,
    "LONG_DASH": BorderStyle.DASHED,
    "CIRCLE": BorderStyle.DOTTED,
    "DOUBLE_SLIM": BorderStyle.DOUBLE,
    "SLIM_THICK": BorderStyle.DOUBLE,
    "THICK_SLIM": BorderStyle.DOUBLE,
    "SLIM_THICK_SLIM": BorderStyle.DOUBLE,
    # 테두리에서는 3D 계열이 유효한 선 종류다(글자 취소선의 "3D"와는 뜻이 다르다).
    "3D": BorderStyle.SOLID,
    "3D_INSET": BorderStyle.SOLID,
    "THICK_3D": BorderStyle.SOLID,
    "THICK_3D_INSET": BorderStyle.SOLID,
    "WAVE": BorderStyle.DASHED,
    "DOUBLE_WAVE": BorderStyle.DOUBLE,
}

_BORDER_SIDES = ("leftBorder", "topBorder", "rightBorder", "bottomBorder")


def _parse_border_fill(bf: etree._Element) -> FillStyle:
    """borderFill 항목에서 배경색과 4방향 테두리를 뽑는다.

    HWPX는 셀에 서식을 직접 쓰지 않고 ``borderFillIDRef`` 로 이 테이블을
    참조한다. 여기를 읽지 않으면 배경 음영과 테두리가 통째로 사라져,
    변환 결과가 원본과 전혀 다른 인상을 준다.
    """
    out = FillStyle()

    # 배경: fillBrush > winBrush 의 faceColor
    for wb in bf.iter():
        if isinstance(wb.tag, str) and _local(wb) == "winBrush":
            face = wb.get("faceColor")
            if face and face.lower() not in ("none", "#ffffff"):
                out.background = units.hwp_color_to_rgb(face)
            break

    borders: list[Border] = []
    any_set = False
    for side in _BORDER_SIDES:
        el = bf.find(_q(HH, side))
        if el is None:
            borders.append(Border())
            continue
        any_set = True
        style = _BORDER_STYLE_MAP.get(
            (el.get("type") or "NONE").upper(), BorderStyle.SOLID
        )
        color = units.hwp_color_to_rgb(el.get("color"))
        borders.append(
            Border(style=style, width_pt=_border_width_pt(el.get("width")), color=color)
        )
    if any_set:
        out.borders = (borders[0], borders[1], borders[2], borders[3])
    return out


def _border_width_pt(v):
    """'0.4 mm' / '0.12 mm' 같은 표기를 pt로. 값이 없으면 가는 실선으로 본다."""
    if not v:
        return 0.4
    s = v.strip().lower()
    try:
        if s.endswith("mm"):
            return units.mm_to_pt(float(s[:-2].strip()))
        if s.endswith("pt"):
            return float(s[:-2].strip())
        return units.mm_to_pt(float(s))
    except ValueError:
        return 0.4


def _find_in_switch(
    parent: etree._Element, ns: str, tag: str
) -> Optional[etree._Element]:
    """직속 자식에서 찾고, 없으면 hp:switch > (hp:default | hp:case) 안을 본다.

    ``default`` 를 ``case`` 보다 우선한다. case는 특정 스키마 확장을 지원하는
    구현만 해석하도록 의도된 분기이므로, 범용 리더는 default를 따르는 것이
    안전하다.
    """
    direct = parent.find(_q(ns, tag))
    if direct is not None:
        return direct
    sw = parent.find(_q(HP, "switch"))
    if sw is None:
        return None
    for branch_tag in ("default", "case"):
        br = sw.find(_q(HP, branch_tag))
        if br is not None:
            found = br.find(_q(ns, tag))
            if found is not None:
                return found
    return None


def _first_attr(parent: etree._Element, ns: str, tag: str, attr: str) -> Optional[str]:
    el = parent.find(_q(ns, tag))
    if el is None:  # 네임스페이스가 hc가 아닐 수도 있어 로컬명으로 재시도
        for ch in parent:
            if isinstance(ch.tag, str) and _local(ch) == tag:
                el = ch
                break
    return el.get(attr) if el is not None else None


# --------------------------------------------------------------------------
# content.hpf - 메타데이터 + 바이너리 매니페스트
# --------------------------------------------------------------------------


def _parse_hpf(root: etree._Element) -> tuple[DocMeta, dict[str, str]]:
    meta = DocMeta(source_format="hwpx")
    manifest: dict[str, str] = {}

    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        name = _local(el)
        text = (el.text or "").strip()
        if name == "title" and text:
            meta.title = text
        elif name == "creator" and text:
            meta.author = text
        elif name == "subject" and text:
            meta.subject = text
        elif name == "meta":
            key = (el.get("name") or "").lower()
            val = text or (el.get("content") or "")
            if val in ("text", ""):
                val = text
            if key == "creator" and val and val != "text":
                meta.author = meta.author or val
            elif key == "lastsaveby" and val and val != "text":
                meta.extra["last_saved_by"] = val
                meta.author = meta.author or val
            elif key == "createddate" and val:
                meta.created = val
            elif key == "modifieddate" and val:
                meta.modified = val
            elif key == "keywords" and val:
                meta.keywords = val
        elif name == "item":
            iid = el.get("id")
            href = el.get("href")
            if iid and href:
                manifest[iid] = href
    return meta, manifest


# --------------------------------------------------------------------------
# 본문 파싱
# --------------------------------------------------------------------------

#: 확장자 -> IR 이미지 포맷
_IMG_EXT = {
    ".png": "png",
    ".jpg": "jpeg",
    ".jpeg": "jpeg",
    ".gif": "gif",
    ".bmp": "bmp",
    ".tif": "tiff",
    ".tiff": "tiff",
    ".emf": "emf",
    ".wmf": "wmf",
    ".svg": "svg",
}


class HwpxReader:
    """HWPX 파일을 Document IR로 변환한다."""

    def __init__(self, path: Path | str) -> None:
        self.path = Path(path)
        self.warnings: list[str] = []

    # -- 공개 API ---------------------------------------------------------

    def read(self) -> Document:
        with SafeZip(self.path) as z:
            self.warnings.extend(z.report.warnings)

            header_raw = z.read_optional("Contents/header.xml")
            styles = (
                _parse_header(z.read_xml("Contents/header.xml"))
                if header_raw
                else StyleTable()
            )

            hpf_raw = z.read_optional("Contents/content.hpf")
            if hpf_raw:
                meta, manifest = _parse_hpf(z.read_xml("Contents/content.hpf"))
            else:
                meta, manifest = DocMeta(source_format="hwpx"), {}

            sections = sorted(
                (n for n in z.namelist() if _is_section(n)),
                key=_section_index,
            )
            if not sections:
                raise CorruptFileError(
                    "HWPX 본문(Contents/sectionN.xml)을 찾을 수 없습니다.",
                    detail=str(self.path),
                )

            doc = Document(meta=meta, kind=DocKind.DOCUMENT)
            self._styles = styles
            self._manifest = manifest
            self._zip = z
            self._page_set = False

            for idx, name in enumerate(sections):
                if idx > 0:
                    doc.blocks.append(PageBreak())
                root = z.read_xml(name)
                doc.blocks.extend(self._parse_section(root, doc))

            doc.compact()
            if not meta.title:
                meta.title = self._infer_title(doc)
            return doc

    # -- 내부 -------------------------------------------------------------

    def _parse_section(self, root: etree._Element, doc: Document) -> list[Block]:
        blocks: list[Block] = []
        # hs:sec 의 직속 문단만 순회한다. 표 안의 문단은 셀에서 따로 처리한다.
        for p in root.findall(_q(HP, "p")):
            blocks.extend(self._parse_paragraph(p, doc))
        return blocks

    def _parse_paragraph(self, p: etree._Element, doc: Document) -> Iterator[Block]:
        """문단 하나를 0개 이상의 블록으로 변환한다.

        문단 안에 표/그림이 있으면 텍스트 문단과 분리해 별도 블록으로 낸다.
        한/글은 표를 문단 안의 run 자식으로 넣기 때문이다.
        """
        ps = self._styles.para(p.get("paraPrIDRef"))
        style_name = self._styles.style_name(p.get("styleIDRef"))

        para = Paragraph(
            align=ps.align,
            indent_pt=ps.indent_pt,
            space_before_pt=ps.space_before_pt,
            space_after_pt=ps.space_after_pt,
            line_spacing=ps.line_spacing,
            style_name=style_name,
        )
        para.heading = _heading_from_style(style_name, ps.outline_level)
        if ps.has_bullet:
            para.list_kind = ListKind.BULLET
        elif ps.has_number:
            para.list_kind = ListKind.NUMBER
        para.list_level = max(0, ps.outline_level - 1)

        pending: list[Block] = []

        for run in p.findall(_q(HP, "run")):
            cs = self._styles.char(run.get("charPrIDRef"))
            for child in run:
                if not isinstance(child.tag, str):
                    continue
                tag = _local(child)
                if tag == "t":
                    for r in self._runs_from_t(child, cs):
                        para.runs.append(r)
                elif tag == "tbl":
                    tb = self._parse_table(child, doc)
                    if tb is not None:
                        pending.append(tb)
                elif tag in ("pic", "container"):
                    img = self._parse_pic(child)
                    if img is not None:
                        pending.append(img)
                elif tag == "secPr":
                    self._apply_secpr(child, doc)
                elif tag == "ctrl":
                    self._apply_ctrl(child, doc)
                elif tag == "tab":
                    para.runs.append(Run(text="\t"))
                elif tag == "lineBreak":
                    para.runs.append(Run(text="\n"))
                # ctrl(단 설정, 쪽번호 등)과 그 외는 텍스트에 기여하지 않는다.

        if (p.get("pageBreak") or "0") == "1":
            yield PageBreak()

        para.merge_runs()
        if para.runs:
            yield para
        yield from pending

    def _runs_from_t(self, t: etree._Element, cs: CharStyle) -> Iterator[Run]:
        """hp:t 를 Run 목록으로. 중간에 낀 탭/줄바꿈 마커까지 반영한다."""

        def mk(text: str) -> Run:
            return Run(
                text=text,
                bold=cs.bold,
                italic=cs.italic,
                underline=cs.underline,
                strike=cs.strike,
                superscript=cs.superscript,
                subscript=cs.subscript,
                size_pt=cs.size_pt,
                font=cs.font,
                color=cs.color,
                highlight=cs.highlight,
            )

        if t.text:
            yield mk(t.text)
        for ch in t:
            if isinstance(ch.tag, str):
                name = _local(ch)
                if name == "tab":
                    yield mk("\t")
                elif name in ("lineBreak", "nbSpace"):
                    yield mk("\n" if name == "lineBreak" else " ")
                elif name == "hyphen":
                    yield mk("-")
            if ch.tail:
                yield mk(ch.tail)

    # -- 표 ---------------------------------------------------------------

    def _parse_table(self, tbl: etree._Element, doc: Document) -> Optional[Table]:
        """표를 그리드로 역산한다.

        병합된 칸은 ``<tc>`` 가 아예 없으므로, 선언된 rowCnt x colCnt 격자를
        만들고 각 셀을 ``cellAddr`` 좌표에 배치한 뒤 ``cellSpan`` 이 덮는
        칸을 placeholder로 채운다.
        """
        rows_el = tbl.findall(_q(HP, "tr"))
        if not rows_el:
            return None

        n_rows = _i(tbl.get("rowCnt"), len(rows_el)) or len(rows_el)
        n_cols = _i(tbl.get("colCnt"), 0)
        if n_cols <= 0:
            n_cols = max(
                (
                    max(
                        (
                            _i(_cell_attr(tc, "cellAddr", "colAddr"), 0)
                            + _i(_cell_attr(tc, "cellSpan", "colSpan"), 1)
                            for tc in tr.findall(_q(HP, "tc"))
                        ),
                        default=0,
                    )
                    for tr in rows_el
                ),
                default=0,
            )
        if n_cols <= 0:
            return None
        n_rows = max(n_rows, len(rows_el))

        grid: list[list[Optional[Cell]]] = [[None] * n_cols for _ in range(n_rows)]
        col_w: list[Optional[float]] = [None] * n_cols

        for r_idx, tr in enumerate(rows_el):
            for tc in tr.findall(_q(HP, "tc")):
                ca = tc.find(_q(HP, "cellAddr"))
                row = _i(ca.get("rowAddr"), r_idx) if ca is not None else r_idx
                col = _i(ca.get("colAddr"), 0) if ca is not None else 0
                span = tc.find(_q(HP, "cellSpan"))
                cspan = max(1, _i(span.get("colSpan"), 1) if span is not None else 1)
                rspan = max(1, _i(span.get("rowSpan"), 1) if span is not None else 1)

                cell = Cell(col_span=cspan, row_span=rspan)
                cell.blocks = list(self._parse_cell_blocks(tc, doc))
                # 셀의 배경/테두리는 borderFillIDRef 로 header.xml 을 참조한다.
                fill = self._styles.fill(tc.get("borderFillIDRef"))
                cell.background = fill.background
                cell.borders = fill.borders

                sz = tc.find(_q(HP, "cellSz"))
                if sz is not None and col < n_cols and col_w[col] is None:
                    w = units.hwp_to_pt(sz.get("width"))
                    if w > 0:
                        col_w[col] = w

                # 격자를 벗어나면 확장한다(선언값이 틀린 파일 방어).
                while row >= len(grid):
                    grid.append([None] * n_cols)
                if col >= n_cols:
                    continue

                grid[row][col] = cell
                for dr in range(rspan):
                    for dc in range(cspan):
                        if dr == 0 and dc == 0:
                            continue
                        rr, cc = row + dr, col + dc
                        while rr >= len(grid):
                            grid.append([None] * n_cols)
                        if cc < n_cols and grid[rr][cc] is None:
                            grid[rr][cc] = Cell(merged_placeholder=True)

        out = Table(
            rows=[[c if c is not None else Cell() for c in row] for row in grid],
            col_widths_pt=[w for w in col_w if w is not None] or None,
        )
        # 빈 행 제거(선언 rowCnt가 과대한 경우)
        out.rows = [
            r for r in out.rows if any(not c.merged_placeholder or c.text for c in r)
        ] or out.rows
        out.header_row = _guess_header_row(tbl)
        return out

    def _parse_cell_blocks(self, tc: etree._Element, doc: Document) -> Iterator[Block]:
        sub = tc.find(_q(HP, "subList"))
        if sub is None:
            return
        for p in sub.findall(_q(HP, "p")):
            yield from self._parse_paragraph(p, doc)

    # -- 그림 -------------------------------------------------------------

    def _parse_pic(self, el: etree._Element) -> Optional[Image]:
        img_el = el if _local(el) == "img" else el.find(".//" + _q(HC, "img"))
        if img_el is None:
            img_el = el.find(".//" + _q(HP, "img"))
        if img_el is None:
            return None
        ref = img_el.get("binaryItemIDRef") or img_el.get("src")
        if not ref:
            return None

        href = self._manifest.get(ref, ref)
        data = None
        for cand in (href, f"BinData/{ref}", f"Contents/{href}"):
            try:
                data = self._zip.read(cand)
                href = cand
                break
            except KeyError:
                continue
        if data is None:
            # 확장자를 모를 때는 BinData 안에서 접두사로 찾는다.
            for name in self._zip.iter_names(prefix="BinData/"):
                if Path(name).stem == ref:
                    data = self._zip.read(name)
                    href = name
                    break
        if data is None:
            self.warnings.append(f"그림 데이터를 찾지 못했습니다: {ref}")
            return None

        ext = Path(href).suffix.lower()
        w = h = None
        sz = el.find(".//" + _q(HP, "sz"))
        if sz is not None:
            w = units.hwp_to_pt(sz.get("width")) or None
            h = units.hwp_to_pt(sz.get("height")) or None
        return Image(
            data=data,
            fmt=_IMG_EXT.get(ext, "png"),
            width_pt=w,
            height_pt=h,
            name=Path(href).name,
        )

    def _apply_ctrl(self, ctrl: etree._Element, doc: Document) -> None:
        """구역 컨트롤에서 쪽 번호 설정을 가져온다.

        한/글은 쪽 번호를 본문 텍스트가 아니라 ctrl 요소에 둔다. 이걸 읽지
        않으면 변환 결과에 쪽 번호가 사라진다.
        """
        pn = ctrl.find(_q(HP, "pageNum"))
        if pn is None or doc.page_number is not None:
            return
        doc.page_number = PageNumberSpec(
            position=(pn.get("pos") or "BOTTOM_CENTER").upper(),
            side_char=(pn.get("sideChar") or "").strip(),
        )
        nn = ctrl.find(_q(HP, "newNum"))
        if nn is not None:
            doc.page_number.start = _i(nn.get("num"), 1)

    # -- 용지 -------------------------------------------------------------

    def _apply_secpr(self, secpr: etree._Element, doc: Document) -> None:
        if self._page_set:
            return
        pg = secpr.find(".//" + _q(HP, "pagePr"))
        if pg is None:
            return
        w = units.hwp_to_pt(pg.get("width"))
        h = units.hwp_to_pt(pg.get("height"))
        if w > 0 and h > 0:
            doc.page_width_pt, doc.page_height_pt = w, h
        mg = pg.find(_q(HP, "margin"))
        if mg is not None:
            doc.margin_pt = (
                units.hwp_to_pt(mg.get("left")),
                units.hwp_to_pt(mg.get("top")),
                units.hwp_to_pt(mg.get("right")),
                units.hwp_to_pt(mg.get("bottom")),
            )
        self._page_set = True

    # -- 보조 -------------------------------------------------------------

    @staticmethod
    def _infer_title(doc: Document) -> Optional[str]:
        """메타데이터에 제목이 없으면 첫 유의미 텍스트를 제목으로 삼는다.

        한/글 양식 문서는 본문 문단 없이 전체가 표 안에 있는 경우가 흔하다
        (기획서/보고서 서식). 그래서 문단만 훑으면 제목을 못 찾는다.
        """
        for b in doc.blocks:
            if isinstance(b, Paragraph) and not b.is_blank():
                return b.text.strip()[:120] or None
            if isinstance(b, Table):
                for row in b.rows:
                    for cell in row:
                        t = cell.text.strip()
                        if t:
                            return t[:120]
        return None


def _cell_attr(tc: etree._Element, tag: str, attr: str) -> Optional[str]:
    el = tc.find(_q(HP, tag))
    return el.get(attr) if el is not None else None


def _guess_header_row(tbl: etree._Element) -> bool:
    """첫 행의 모든 셀이 header=1이면 머리글 행으로 본다."""
    tr = tbl.find(_q(HP, "tr"))
    if tr is None:
        return False
    tcs = tr.findall(_q(HP, "tc"))
    if not tcs:
        return False
    return all((tc.get("header") or "0") == "1" for tc in tcs)


_HEADING_NAMES = {
    "title": 1,
    "subtitle": 2,
}


def _heading_from_style(style_name: Optional[str], outline_level: int) -> int:
    """스타일 이름과 개요 수준으로 제목 수준을 결정한다.

    한/글은 'heading 1'(영문 템플릿)과 '개요 1'(한글 템플릿)을 모두 쓴다.
    """
    if style_name:
        low = style_name.strip().lower()
        if low in _HEADING_NAMES:
            return _HEADING_NAMES[low]
        for prefix in ("heading ", "개요 ", "제목 "):
            if low.startswith(prefix):
                try:
                    return max(1, min(9, int(low[len(prefix) :].strip())))
                except ValueError:
                    pass
    if outline_level > 0:
        return max(1, min(9, outline_level))
    return 0


def _is_section(name: str) -> bool:
    return (
        name.startswith("Contents/section")
        and name.endswith(".xml")
        and name[len("Contents/section") : -4].isdigit()
    )


def _section_index(name: str) -> int:
    return int(name[len("Contents/section") : -4])


def read(path: Path | str) -> Document:
    """모듈 수준 진입점. registry가 이 이름을 찾는다."""
    r = HwpxReader(path)
    return _attach(r.read(), r)


def _attach(doc, reader):
    """Reader가 모은 경고를 Document에 실어 파이프라인까지 전달한다."""
    if getattr(reader, "warnings", None):
        doc.meta.extra.setdefault("warnings", []).extend(reader.warnings)
    return doc
