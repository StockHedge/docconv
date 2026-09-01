"""HWPX(OWPML) Writer.

한/글이 실제로 열 수 있는 HWPX를 생성한다. 구형 .hwp 쓰기는 지원하지 않으므로
"한글로 내보내기"의 목적지는 항상 이쪽이다. 사용자는 한/글에서 열어 .hwp로
다시 저장할 수 있다.

동작 원리
---------
HWPX 본문은 서식을 인라인으로 쓰지 않고 `charPrIDRef` / `paraPrIDRef` 로
header.xml의 테이블을 참조한다. 그래서 Writer는 2패스로 돈다.

1. IR을 훑어 등장하는 서식 조합을 모으고 ID를 매긴다.
2. 그 테이블로 header.xml을 만들고, 본문은 ID만 참조하게 직렬화한다.

컨테이너 구성(원본 파일에서 확인한 그대로)::

    mimetype                    application/hwp+zip
    version.xml                 HCFVersion
    META-INF/container.xml      rootfile -> Contents/content.hpf
    META-INF/manifest.xml       (빈 매니페스트)
    settings.xml                애플리케이션 설정
    Contents/content.hpf        OPF 매니페스트 + 메타데이터
    Contents/header.xml         참조 테이블
    Contents/section0.xml       본문
    BinData/*                   이미지
"""

from __future__ import annotations

import datetime as _dt
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional
from xml.sax.saxutils import escape as _xesc
from xml.sax.saxutils import quoteattr as _xattr

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
from ..util import units

# 모든 XML 파일에 붙는 공통 네임스페이스 선언 (원본과 동일하게 맞춘다)
_NSDECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

#: 언어 슬롯. charPr의 fontRef 등은 7개 언어 속성을 모두 요구한다.
_LANGS = ("HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER")
_LANG_ATTRS = "hangul latin hanja japanese other symbol user".split()

_DEFAULT_FONTS = ("함초롬바탕", "맑은 고딕", "굴림", "돋움", "바탕")

#: borderFill ID 규약
BF_NONE = 1  # 테두리 없음
BF_CELL = 2  # 표 셀 실선
BF_HEADER = 3  # 표 머리글(회색 배경)


# --------------------------------------------------------------------------
# 서식 테이블 수집
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class CharKey:
    size_pt: float
    bold: bool
    italic: bool
    underline: bool
    strike: bool
    color: Optional[tuple[int, int, int]]
    font: Optional[str]
    superscript: bool = False
    subscript: bool = False


@dataclass(frozen=True)
class ParaKey:
    align: Align
    indent_pt: float
    space_before_pt: float
    space_after_pt: float
    line_spacing: float
    heading: int


class _Tables:
    """본문에서 쓰인 서식 조합을 모아 ID를 매긴다."""

    def __init__(self, base_size: float) -> None:
        self.base_size = base_size
        self.chars: dict[CharKey, int] = {}
        self.paras: dict[ParaKey, int] = {}
        self.fonts: list[str] = list(_DEFAULT_FONTS)
        # 기본 항목을 0번으로 고정해 두면 참조가 없을 때도 안전하다.
        self.char_id(CharKey(base_size, False, False, False, False, None, None))
        self.para_id(ParaKey(Align.JUSTIFY, 0.0, 0.0, 0.0, 1.6, 0))

    def font_id(self, name: Optional[str]) -> int:
        if not name:
            return 0
        if name not in self.fonts:
            self.fonts.append(name)
        return self.fonts.index(name)

    def char_id(self, k: CharKey) -> int:
        if k not in self.chars:
            self.chars[k] = len(self.chars)
        return self.chars[k]

    def para_id(self, k: ParaKey) -> int:
        if k not in self.paras:
            self.paras[k] = len(self.paras)
        return self.paras[k]

    def key_of_run(self, r: Run, opts: ConvertOptions) -> CharKey:
        if not opts.keep_formatting:
            return CharKey(self.base_size, False, False, False, False, None, None)
        size = r.size_pt or self.base_size
        return CharKey(
            size_pt=round(max(1.0, min(400.0, size)), 2),
            bold=r.bold,
            italic=r.italic,
            underline=r.underline,
            strike=r.strike,
            color=r.color,
            font=r.font,
            superscript=r.superscript,
            subscript=r.subscript,
        )

    def key_of_para(self, p: Paragraph, opts: ConvertOptions) -> ParaKey:
        if not opts.keep_formatting:
            return ParaKey(Align.JUSTIFY, 0.0, 0.0, 0.0, 1.6, p.heading)
        return ParaKey(
            align=p.align,
            indent_pt=round(p.indent_pt, 1),
            space_before_pt=round(p.space_before_pt, 1),
            space_after_pt=round(p.space_after_pt, 1),
            line_spacing=round(p.line_spacing or 1.6, 2),
            heading=p.heading,
        )


# --------------------------------------------------------------------------
# 진입점
# --------------------------------------------------------------------------


def write_hwpx(doc: Document, path: Path | str, opts: ConvertOptions) -> None:
    tables = _Tables(base_size=10.0)
    images: dict[int, tuple[str, bytes]] = {}
    if opts.include_images:
        _collect_images(doc, images)

    # 1패스: 서식 수집 + 본문 생성 (본문 생성 중에 ID가 확정된다)
    body = _section_xml(doc, tables, opts, images)
    header = _header_xml(tables, doc)
    hpf = _content_hpf(doc, images)

    try:
        with zipfile.ZipFile(
            str(path), "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
        ) as z:
            z.writestr("mimetype", "application/hwp+zip")
            z.writestr("version.xml", _version_xml())
            z.writestr("META-INF/container.xml", _container_xml())
            z.writestr("META-INF/manifest.xml", _manifest_xml())
            z.writestr("settings.xml", _settings_xml())
            z.writestr("Contents/content.hpf", hpf)
            z.writestr("Contents/header.xml", header)
            z.writestr("Contents/section0.xml", body)
            for name, data in images.values():
                z.writestr(f"BinData/{name}", data)
    except OSError as e:
        raise WriteError(f"파일을 저장할 수 없습니다: {path}", detail=str(e)) from e


def _collect_images(doc: Document, out: dict[int, tuple[str, bytes]]) -> None:
    idx = 0
    for b in _walk(doc.blocks):
        if isinstance(b, Image) and b.data:
            ext = "jpg" if b.fmt == "jpeg" else (b.fmt or "png")
            idx += 1
            out[id(b)] = (f"image{idx}.{ext}", b.data)


def _walk(blocks: Iterable[Block]) -> Iterable[Block]:
    for b in blocks:
        yield b
        if isinstance(b, Table):
            for row in b.rows:
                for c in row:
                    yield from _walk(c.blocks)


# --------------------------------------------------------------------------
# 고정 파일들
# --------------------------------------------------------------------------


def _version_xml() -> str:
    return (
        _XML_DECL
        + '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" '
        'tagetApplication="WORDPROCESSOR" major="5" minor="0" micro="5" '
        'buildNumber="0" os="1" xmlVersion="1.4" application="docconv" '
        'appVersion="1.0"/>'
    )


def _container_xml() -> str:
    return (
        _XML_DECL
        + '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
        'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles>'
        '<ocf:rootfile full-path="Contents/content.hpf" '
        'media-type="application/hwpml-package+xml"/>'
        "</ocf:rootfiles></ocf:container>"
    )


def _manifest_xml() -> str:
    return (
        _XML_DECL
        + '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
    )


def _settings_xml() -> str:
    return (
        _XML_DECL
        + '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
        'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
        '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
        "</ha:HWPApplicationSetting>"
    )


def _content_hpf(doc: Document, images: dict[int, tuple[str, bytes]]) -> str:
    m = doc.meta
    now = _dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    items = [
        '<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>',
        '<opf:item id="settings" href="settings.xml" media-type="application/xml"/>',
        '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>',
    ]
    for i, (name, _) in enumerate(images.values(), start=1):
        mt = _media_type(name)
        items.append(
            f'<opf:item id="image{i}" href="BinData/{_xesc(name)}" media-type="{mt}"/>'
        )
    return (
        _XML_DECL + f'<opf:package {_NSDECL} version="" unique-identifier="" id="">'
        "<opf:metadata>"
        f"<opf:title>{_xesc(m.title or '')}</opf:title>"
        "<opf:language>ko</opf:language>"
        f'<opf:meta name="creator" content="text">{_xesc(m.author or "")}</opf:meta>'
        f'<opf:meta name="subject" content="text">{_xesc(m.subject or "")}</opf:meta>'
        '<opf:meta name="description" content="text"/>'
        f'<opf:meta name="lastsaveby" content="text">{_xesc(m.author or "")}</opf:meta>'
        f'<opf:meta name="CreatedDate" content="text">{_xesc(m.created or now)}</opf:meta>'
        f'<opf:meta name="ModifiedDate" content="text">{now}</opf:meta>'
        f'<opf:meta name="keyword" content="text">{_xesc(m.keywords or "")}</opf:meta>'
        "</opf:metadata>"
        "<opf:manifest>" + "".join(items) + "</opf:manifest>"
        "<opf:spine>"
        '<opf:itemref idref="header"/>'
        '<opf:itemref idref="section0" linear="yes"/>'
        "</opf:spine>"
        "</opf:package>"
    )


def _media_type(name: str) -> str:
    ext = Path(name).suffix.lower().lstrip(".")
    return {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tiff": "image/tiff",
    }.get(ext, "application/octet-stream")


# --------------------------------------------------------------------------
# header.xml
# --------------------------------------------------------------------------


def _header_xml(t: _Tables, doc: Document) -> str:
    p: list[str] = [
        _XML_DECL,
        f'<hh:head {_NSDECL} version="1.4" secCnt="1">',
        '<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>',
        "<hh:refList>",
        _fontfaces(t),
        _borderfills(),
        _char_properties(t),
        '<hh:tabProperties itemCnt="1">'
        '<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
        "</hh:tabProperties>",
        _numberings(),
        _bullets(),
        _para_properties(t),
        _styles(t),
        "</hh:refList>",
        '<hh:compatibleDocument targetProgram="HWP201X">'
        "<hh:layoutCompatibility/></hh:compatibleDocument>",
        "</hh:head>",
    ]
    return "".join(p)


def _fontfaces(t: _Tables) -> str:
    """7개 언어 슬롯 각각에 같은 글꼴 목록을 등록한다."""
    blocks: list[str] = [f'<hh:fontfaces itemCnt="{len(_LANGS)}">']
    for lang in _LANGS:
        blocks.append(f'<hh:fontface lang="{lang}" fontCnt="{len(t.fonts)}">')
        for i, face in enumerate(t.fonts):
            blocks.append(
                f'<hh:font id="{i}" face={_xattr(face)} type="TTF" isEmbedded="0">'
                '<hh:typeInfo familyType="FCAT_GOTHIC" weight="5" proportion="3" '
                'contrast="0" strokeVariation="0" armStyle="0" letterform="0" '
                'midline="0" xHeight="0"/></hh:font>'
            )
        blocks.append("</hh:fontface>")
    blocks.append("</hh:fontfaces>")
    return "".join(blocks)


def _border(kind: str, style: str, color: str = "#000000") -> str:
    return f'<hh:{kind} type="{style}" width="0.12 mm" color="{color}"/>'


def _borderfills() -> str:
    def fill(bg: Optional[str]) -> str:
        if not bg:
            return ""
        return (
            "<hc:fillBrush><hc:winBrush "
            f'faceColor="{bg}" hatchColor="#999999" alpha="0"/></hc:fillBrush>'
        )

    def bf(idx: int, style: str, bg: Optional[str] = None) -> str:
        return (
            f'<hh:borderFill id="{idx}" threeD="0" shadow="0" centerLine="NONE" '
            'breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            + _border("leftBorder", style)
            + _border("rightBorder", style)
            + _border("topBorder", style)
            + _border("bottomBorder", style)
            + '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            + fill(bg)
            + "</hh:borderFill>"
        )

    return (
        '<hh:borderFills itemCnt="3">'
        + bf(BF_NONE, "NONE")
        + bf(BF_CELL, "SOLID")
        + bf(BF_HEADER, "SOLID", "#F2F2F2")
        + "</hh:borderFills>"
    )


def _per_lang(tag: str, value: str) -> str:
    """7개 언어 슬롯에 같은 값을 채운 요소를 만든다.

    charPr의 fontRef/ratio/spacing/relSz/offset은 hangul..user 7개 속성을
    모두 요구한다. 하나라도 빠지면 한/글이 파일을 거부한다.
    """
    attrs = " ".join('%s="%s"' % (a, value) for a in _LANG_ATTRS)
    return "<hh:%s %s/>" % (tag, attrs)


def _char_properties(t: _Tables) -> str:
    items: list[str] = [f'<hh:charProperties itemCnt="{len(t.chars)}">']
    for key, cid in sorted(t.chars.items(), key=lambda kv: kv[1]):
        fid = t.font_id(key.font)
        height = int(round(key.size_pt * 100))
        color = units.rgb_to_hex(key.color) or "#000000"
        ul = "BOTTOM" if key.underline else "NONE"
        strike = "SOLID" if key.strike else "NONE"

        offset = "0"
        if key.superscript:
            offset = "30"
        elif key.subscript:
            offset = "-30"

        items.append(
            f'<hh:charPr id="{cid}" height="{height}" textColor="{color}" '
            f'shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" '
            f'borderFillIDRef="{BF_NONE}">'
            + _per_lang("fontRef", str(fid))
            + _per_lang("ratio", "100")
            + _per_lang("spacing", "0")
            + _per_lang("relSz", "100")
            + _per_lang("offset", offset)
            + ("<hh:bold/>" if key.bold else "")
            + ("<hh:italic/>" if key.italic else "")
            + f'<hh:underline type="{ul}" shape="SOLID" color="{color}"/>'
            + f'<hh:strikeout shape="{strike}" color="{color}"/>'
            + '<hh:outline type="NONE"/>'
            + '<hh:shadow type="NONE" color="#B2B2B2" offsetX="10" offsetY="10"/>'
            + "</hh:charPr>"
        )
    items.append("</hh:charProperties>")
    return "".join(items)


_ALIGN_OUT = {
    Align.LEFT: "LEFT",
    Align.CENTER: "CENTER",
    Align.RIGHT: "RIGHT",
    Align.JUSTIFY: "JUSTIFY",
    Align.DISTRIBUTE: "DISTRIBUTE",
}


def _para_properties(t: _Tables) -> str:
    items: list[str] = [f'<hh:paraProperties itemCnt="{len(t.paras)}">']
    for key, pid in sorted(t.paras.items(), key=lambda kv: kv[1]):
        align = _ALIGN_OUT.get(key.align, "JUSTIFY")
        left = units.pt_to_hwp(key.indent_pt)
        prev = units.pt_to_hwp(key.space_before_pt)
        nxt = units.pt_to_hwp(key.space_after_pt)
        spacing = int(round(key.line_spacing * 100))
        heading = (
            f'<hh:heading type="OUTLINE" idRef="0" level="{min(9, key.heading) - 1}"/>'
            if key.heading
            else '<hh:heading type="NONE" idRef="0" level="0"/>'
        )
        items.append(
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0" fontLineHeight="0" '
            'snapToGrid="1" suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            + heading
            + '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" '
            'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" '
            'lineWrap="BREAK"/>'
            '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            "<hh:margin>"
            '<hc:intent value="0" unit="HWPUNIT"/>'
            f'<hc:left value="{left}" unit="HWPUNIT"/>'
            '<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="{prev}" unit="HWPUNIT"/>'
            f'<hc:next value="{nxt}" unit="HWPUNIT"/>'
            "</hh:margin>"
            f'<hh:lineSpacing type="PERCENT" value="{spacing}" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="{BF_NONE}" offsetLeft="0" offsetRight="0" '
            'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            "</hh:paraPr>"
        )
    items.append("</hh:paraProperties>")
    return "".join(items)


def _numberings() -> str:
    levels = "".join(
        f'<hh:paraHead start="1" level="{i}" align="LEFT" useInstWidth="1" '
        f'autoIndent="1" widthAdjust="0" textOffsetType="PERCENT" textOffset="50" '
        f'numFormat="DIGIT" charPrIDRef="4294967295" checkable="0">^{i}.</hh:paraHead>'
        for i in range(1, 8)
    )
    return f'<hh:numberings itemCnt="1"><hh:numbering id="1" start="1">{levels}</hh:numbering></hh:numberings>'


def _bullets() -> str:
    return (
        '<hh:bullets itemCnt="1">'
        '<hh:bullet id="1" char="&#8226;" checkedChar="&#8226;" useImage="0">'
        '<hh:paraHead start="1" level="0" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" '
        'charPrIDRef="4294967295" checkable="0"/>'
        "</hh:bullet></hh:bullets>"
    )


def _styles(t: _Tables) -> str:
    """바탕글 + 개요 1~9. 본문에서 heading을 쓰면 이 스타일을 참조한다."""
    rows = [
        '<hh:style id="0" type="PARA" name="바탕글" engName="Normal" '
        'paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
    ]
    for lvl in range(1, 10):
        pid = t.para_id(ParaKey(Align.LEFT, 0.0, 6.0, 3.0, 1.6, lvl))
        cid = t.char_id(
            CharKey(round(_heading_size(lvl), 2), True, False, False, False, None, None)
        )
        rows.append(
            f'<hh:style id="{lvl}" type="PARA" name="개요 {lvl}" engName="Outline {lvl}" '
            f'paraPrIDRef="{pid}" charPrIDRef="{cid}" nextStyleIDRef="0" '
            'langID="1042" lockForm="0"/>'
        )
    return f'<hh:styles itemCnt="{len(rows)}">' + "".join(rows) + "</hh:styles>"


def _heading_size(level: int) -> float:
    return {1: 18.0, 2: 15.0, 3: 13.0, 4: 12.0, 5: 11.0}.get(level, 10.5)


# --------------------------------------------------------------------------
# section0.xml
# --------------------------------------------------------------------------


def _section_xml(
    doc: Document,
    t: _Tables,
    opts: ConvertOptions,
    images: dict[int, tuple[str, bytes]],
) -> str:
    out: list[str] = [_XML_DECL, f"<hs:sec {_NSDECL}>"]

    blocks = list(doc.blocks) or [Paragraph.of("")]
    first = True
    counter = _Counter()

    for b in blocks:
        if isinstance(b, Paragraph):
            out.append(
                _para_xml(
                    b, t, opts, secpr=_secpr(doc) if first else None, counter=counter
                )
            )
            first = False
        elif isinstance(b, Table):
            # 표는 문단 안의 run에 들어가야 한다.
            out.append(
                _para_xml(
                    Paragraph(),
                    t,
                    opts,
                    secpr=_secpr(doc) if first else None,
                    inner=_table_xml(b, t, opts, images, counter),
                    counter=counter,
                )
            )
            first = False
        elif isinstance(b, Image) and opts.include_images:
            name = images.get(id(b))
            inner = _pic_xml(b, name[0], counter) if name else ""
            out.append(
                _para_xml(
                    Paragraph(),
                    t,
                    opts,
                    secpr=_secpr(doc) if first else None,
                    inner=inner,
                    counter=counter,
                )
            )
            first = False
        elif isinstance(b, PageBreak) and opts.keep_page_breaks:
            out.append(
                _para_xml(
                    Paragraph(),
                    t,
                    opts,
                    secpr=_secpr(doc) if first else None,
                    page_break=True,
                    counter=counter,
                )
            )
            first = False

    out.append("</hs:sec>")
    return "".join(out)


class _Counter:
    """개체 ID 발급기. 한/글은 문서 내 유일한 ID를 요구한다."""

    def __init__(self) -> None:
        self.n = 1000000000

    def next(self) -> int:
        self.n += 1
        return self.n


def _para_xml(
    p: Paragraph,
    t: _Tables,
    opts: ConvertOptions,
    *,
    secpr: Optional[str] = None,
    inner: str = "",
    page_break: bool = False,
    counter: Optional[_Counter] = None,
) -> str:
    pid = t.para_id(t.key_of_para(p, opts))
    style_id = min(9, p.heading) if p.heading else 0

    runs: list[str] = []
    if secpr:
        cid0 = t.char_id(t.key_of_run(Run(), opts))
        runs.append(f'<hp:run charPrIDRef="{cid0}">{secpr}</hp:run>')
    if inner:
        cid0 = t.char_id(t.key_of_run(Run(), opts))
        runs.append(f'<hp:run charPrIDRef="{cid0}">{inner}</hp:run>')

    for r in p.runs:
        cid = t.char_id(t.key_of_run(r, opts))
        runs.append(f'<hp:run charPrIDRef="{cid}">{_text_xml(r.text)}</hp:run>')

    if not runs:
        cid0 = t.char_id(t.key_of_run(Run(), opts))
        runs.append(f'<hp:run charPrIDRef="{cid0}"><hp:t/></hp:run>')

    pb = "1" if page_break else "0"
    return (
        f'<hp:p id="0" paraPrIDRef="{pid}" styleIDRef="{style_id}" '
        f'pageBreak="{pb}" columnBreak="0" merged="0">' + "".join(runs) + "</hp:p>"
    )


def _text_xml(text: str) -> str:
    """텍스트를 hp:t로. 탭과 줄바꿈은 전용 요소로 바꾼다."""
    if not text:
        return "<hp:t/>"
    parts: list[str] = ["<hp:t>"]
    buf: list[str] = []
    for ch in text:
        if ch == "\t":
            parts.append(_xesc("".join(buf)))
            buf.clear()
            parts.append("<hp:tab/>")
        elif ch in ("\n", "\r"):
            parts.append(_xesc("".join(buf)))
            buf.clear()
            parts.append("<hp:lineBreak/>")
        elif ord(ch) < 0x20 and ch not in ("\t",):
            continue  # 제어문자는 버린다. XML에 넣을 수 없다.
        else:
            buf.append(ch)
    parts.append(_xesc("".join(buf)))
    parts.append("</hp:t>")
    return "".join(parts)


def _secpr(doc: Document) -> str:
    w = units.pt_to_hwp(doc.page_width_pt)
    h = units.pt_to_hwp(doc.page_height_pt)
    left, top, right, bottom = (units.pt_to_hwp(v) for v in doc.margin_pt)
    landscape = "NARROWLY" if doc.page_width_pt > doc.page_height_pt else "WIDELY"
    return (
        '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" '
        'tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" '
        'textVerticalWidthHead="0" masterPageCnt="0">'
        '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0" strtnum="0"/>'
        '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
        '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" '
        'border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" '
        'showLineNumber="0"/>'
        '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
        f'<hp:pagePr landscape="{landscape}" width="{w}" height="{h}" gutterType="LEFT_ONLY">'
        f'<hp:margin header="4252" footer="4252" gutter="0" left="{left}" right="{right}" '
        f'top="{top}" bottom="{bottom}"/>'
        "</hp:pagePr>"
        '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" '
        'suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" '
        'color="#000000"/><hp:noteSpacing betweenNotes="850" belowLine="567" aboveLine="850"/>'
        '<hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" '
        'beneathText="0"/></hp:footNotePr>'
        '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" '
        'suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" '
        'width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="567" '
        'aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/>'
        '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
        + "".join(
            f'<hp:pageBorderFill type="{k}" borderFillIDRef="{BF_NONE}" textBorder="PAPER" '
            'headerInside="0" footerInside="0" fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            "</hp:pageBorderFill>"
            for k in ("BOTH", "EVEN", "ODD")
        )
        + "</hp:secPr>"
    )


# --------------------------------------------------------------------------
# 표 / 그림
# --------------------------------------------------------------------------

#: 표의 기본 셀 안쪽 여백 (HWPUNIT)
_CELL_MARGIN = "510"
_CELL_MARGIN_V = "141"


def _table_xml(
    t: Table,
    tables: _Tables,
    opts: ConvertOptions,
    images: dict[int, tuple[str, bytes]],
    counter: _Counter,
) -> str:
    t.normalize()
    if not t.rows or not t.n_cols:
        return ""

    n_cols = t.n_cols
    # 열 너비: 지정이 없으면 본문 폭을 균등 분할한다.
    body_w = 42520  # A4 - 기본 여백 (HWPUNIT)
    if t.col_widths_pt and len(t.col_widths_pt) >= n_cols:
        widths = [max(1000, units.pt_to_hwp(w)) for w in t.col_widths_pt[:n_cols]]
    else:
        widths = [body_w // n_cols] * n_cols
    total_w = sum(widths)
    row_h = 2000

    rows_xml: list[str] = []
    for ri, row in enumerate(t.rows):
        cells_xml: list[str] = []
        for ci, c in enumerate(row):
            if c.merged_placeholder:
                continue
            bf = BF_HEADER if (t.header_row and ri == 0) else BF_CELL
            sub = "".join(
                _cell_block_xml(b, tables, opts, images, counter) for b in c.blocks
            )
            if not sub:
                sub = _para_xml(Paragraph(), tables, opts, counter=counter)
            cw = sum(widths[ci : ci + c.col_span]) or widths[min(ci, n_cols - 1)]
            cells_xml.append(
                f'<hp:tc name="" header="{"1" if (t.header_row and ri == 0) else "0"}" '
                f'hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="{bf}">'
                '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
                'vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" '
                'textHeight="0" hasTextRef="0" hasNumRef="0">' + sub + "</hp:subList>"
                f'<hp:cellAddr colAddr="{ci}" rowAddr="{ri}"/>'
                f'<hp:cellSpan colSpan="{c.col_span}" rowSpan="{c.row_span}"/>'
                f'<hp:cellSz width="{cw}" height="{row_h}"/>'
                f'<hp:cellMargin left="{_CELL_MARGIN}" right="{_CELL_MARGIN}" '
                f'top="{_CELL_MARGIN_V}" bottom="{_CELL_MARGIN_V}"/>'
                "</hp:tc>"
            )
        rows_xml.append("<hp:tr>" + "".join(cells_xml) + "</hp:tr>")

    tbl_id = counter.next()
    return (
        f'<hp:tbl id="{tbl_id}" zOrder="0" numberingType="TABLE" '
        'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" '
        f'pageBreak="CELL" repeatHeader="{"1" if t.header_row else "0"}" '
        f'rowCnt="{t.n_rows}" colCnt="{n_cols}" cellSpacing="0" '
        f'borderFillIDRef="{BF_CELL}" noAdjust="0">'
        f'<hp:sz width="{total_w}" widthRelTo="ABSOLUTE" height="{row_h * t.n_rows}" '
        'heightRelTo="ABSOLUTE" protect="0"/>'
        '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
        'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        '<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:inMargin left="{_CELL_MARGIN}" right="{_CELL_MARGIN}" '
        f'top="{_CELL_MARGIN_V}" bottom="{_CELL_MARGIN_V}"/>'
        + "".join(rows_xml)
        + "</hp:tbl>"
    )


def _cell_block_xml(
    b: Block,
    tables: _Tables,
    opts: ConvertOptions,
    images: dict[int, tuple[str, bytes]],
    counter: _Counter,
) -> str:
    if isinstance(b, Paragraph):
        return _para_xml(b, tables, opts, counter=counter)
    if isinstance(b, Table):
        return _para_xml(
            Paragraph(),
            tables,
            opts,
            inner=_table_xml(b, tables, opts, images, counter),
            counter=counter,
        )
    if isinstance(b, Image) and opts.include_images:
        nm = images.get(id(b))
        if nm:
            return _para_xml(
                Paragraph(),
                tables,
                opts,
                inner=_pic_xml(b, nm[0], counter),
                counter=counter,
            )
    return ""


def _pic_xml(img: Image, name: str, counter: _Counter) -> str:
    w = units.pt_to_hwp(img.width_pt or 200.0)
    h = units.pt_to_hwp(img.height_pt or 150.0)
    pid = counter.next()
    return (
        f'<hp:pic id="{pid}" zOrder="0" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" '
        'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" '
        'instid="0" reverse="0">'
        f'<hp:sz width="{w}" widthRelTo="ABSOLUTE" height="{h}" heightRelTo="ABSOLUTE" '
        'protect="0"/>'
        '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
        'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        '<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{w}" y="0"/>'
        f'<hc:pt2 x="{w}" y="{h}"/><hc:pt3 x="0" y="{h}"/></hp:imgRect>'
        f'<hp:imgClip left="0" right="{w}" top="0" bottom="{h}"/>'
        '<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgDim dimwidth="{w}" dimheight="{h}"/>'
        f'<hc:img binaryItemIDRef={_xattr(name)} bright="0" contrast="0" effect="REAL_PIC" '
        'alpha="0"/>'
        "<hp:effects/>"
        "</hp:pic>"
    )
