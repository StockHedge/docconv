"""문서 중간표현(Intermediate Representation).

모든 Reader는 임의의 입력 포맷을 이 IR로 환원하고, 모든 Writer는 이 IR만을
소비한다. 덕분에 N개 포맷을 지원하는 데 N*(N-1)개가 아니라 2N개의 모듈만
필요하다.

단위 규약
---------
길이는 전부 **pt(포인트, 1/72인치)** 로 통일한다. 각 포맷의 고유 단위
(HWPUNIT=1/7200인치, EMU=1/914400인치, twip=1/1440인치)는 Reader/Writer
경계에서만 환산한다. `docconv.util.units` 의 헬퍼를 쓸 것.

색은 (r, g, b) 0~255 튜플로 통일한다. None은 "지정 없음"(상속)을 뜻하며
(0, 0, 0)(검정)과 구분된다.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Any, Iterator, Optional, Sequence

RGB = tuple[int, int, int]


# --------------------------------------------------------------------------
# 인라인 수준
# --------------------------------------------------------------------------


@dataclass
class Run:
    """동일한 서식을 갖는 연속된 텍스트 조각."""

    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    superscript: bool = False
    subscript: bool = False
    size_pt: Optional[float] = None
    font: Optional[str] = None
    color: Optional[RGB] = None
    highlight: Optional[RGB] = None
    #: 하이퍼링크 대상. 없으면 None.
    href: Optional[str] = None

    def is_blank(self) -> bool:
        return not self.text.strip()

    def same_style(self, other: "Run") -> bool:
        """텍스트를 제외한 서식이 동일한지. Run 병합 최적화에 쓴다."""
        return (
            self.bold == other.bold
            and self.italic == other.italic
            and self.underline == other.underline
            and self.strike == other.strike
            and self.superscript == other.superscript
            and self.subscript == other.subscript
            and self.size_pt == other.size_pt
            and self.font == other.font
            and self.color == other.color
            and self.highlight == other.highlight
            and self.href == other.href
        )


class Align(enum.Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"
    DISTRIBUTE = "distribute"


class ListKind(enum.Enum):
    NONE = "none"
    BULLET = "bullet"
    NUMBER = "number"


# --------------------------------------------------------------------------
# 블록 수준
# --------------------------------------------------------------------------


@dataclass
class Block:
    """모든 블록의 기반. 식별은 isinstance로 한다."""


@dataclass
class Paragraph(Block):
    runs: list[Run] = field(default_factory=list)
    align: Align = Align.LEFT
    #: 0이면 본문, 1~9면 제목 수준.
    heading: int = 0
    list_kind: ListKind = ListKind.NONE
    list_level: int = 0
    indent_pt: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    #: 줄간격 배수(1.0=단일). None이면 포맷 기본값.
    line_spacing: Optional[float] = None
    style_name: Optional[str] = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def is_blank(self) -> bool:
        return not self.text.strip()

    def merge_runs(self) -> None:
        """서식이 같은 인접 Run을 합쳐 출력 파일 크기를 줄인다."""
        if len(self.runs) < 2:
            return
        merged: list[Run] = [self.runs[0]]
        for r in self.runs[1:]:
            if merged[-1].same_style(r):
                merged[-1].text += r.text
            else:
                merged.append(r)
        self.runs = merged

    @classmethod
    def of(cls, text: str, **kw: Any) -> "Paragraph":
        """단일 Run 문단을 만드는 지름길."""
        run_kw = {}
        for key in ("bold", "italic", "size_pt", "font", "color", "href"):
            if key in kw:
                run_kw[key] = kw.pop(key)
        return cls(runs=[Run(text=text, **run_kw)], **kw)


class BorderStyle(enum.Enum):
    NONE = "none"
    SOLID = "solid"
    DASHED = "dashed"
    DOTTED = "dotted"
    DOUBLE = "double"


@dataclass
class Border:
    """셀/표의 한 방향 테두리.

    한/글 문서는 셀마다 테두리 유무가 다른 경우가 흔하다(제목 칸만 상자,
    본문 칸은 좌우선만 등). 이 정보를 버리면 변환 결과가 균일한 격자로 보여
    원본과 인상이 크게 달라진다.
    """

    style: BorderStyle = BorderStyle.NONE
    width_pt: float = 0.0
    color: Optional[RGB] = None

    @property
    def visible(self) -> bool:
        return self.style is not BorderStyle.NONE and self.width_pt > 0

    def css(self) -> str:
        if not self.visible:
            return "none"
        c = "#%02X%02X%02X" % (self.color or (0, 0, 0))
        return f"{max(0.3, self.width_pt):.2f}pt {self.style.value} {c}"


#: (좌, 상, 우, 하) 순서. Document.margin_pt 와 같은 규약이다.
Borders = tuple[Border, Border, Border, Border]


def no_borders() -> Borders:
    return (Border(), Border(), Border(), Border())


@dataclass
class Cell:
    """표의 셀. 내부에 블록을 담을 수 있어 중첩 표/여러 문단을 표현한다."""

    blocks: list[Block] = field(default_factory=list)
    col_span: int = 1
    row_span: int = 1
    #: 병합으로 가려진 자리(placeholder). 그리드 정합을 위해 남긴다.
    merged_placeholder: bool = False
    background: Optional[RGB] = None
    align: Optional[Align] = None
    #: (좌, 상, 우, 하) 테두리. None이면 "지정 없음"이라 Writer 기본값을 쓴다.
    borders: Optional[Borders] = None
    #: 스프레드시트 원본 값(숫자/날짜 등). 표형 -> 표형 변환 시 타입 보존용.
    raw_value: Any = None
    number_format: Optional[str] = None

    def has_any_border(self) -> bool:
        return self.borders is not None and any(b.visible for b in self.borders)

    @property
    def text(self) -> str:
        parts: list[str] = []
        for b in self.blocks:
            if isinstance(b, Paragraph):
                parts.append(b.text)
            elif isinstance(b, Table):
                parts.append(b.to_text())
        return "\n".join(p for p in parts if p)

    @classmethod
    def of(cls, text: str, **kw: Any) -> "Cell":
        blocks = [Paragraph.of(text)] if text != "" else []
        return cls(blocks=blocks, **kw)


@dataclass
class Table(Block):
    rows: list[list[Cell]] = field(default_factory=list)
    #: 스프레드시트에서 왔다면 시트 이름.
    name: Optional[str] = None
    #: 첫 행이 머리글인지. Writer가 스타일링에 활용한다.
    header_row: bool = False
    col_widths_pt: Optional[list[float]] = None

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.rows), default=0)

    def to_text(self, sep: str = "\t") -> str:
        return "\n".join(sep.join(c.text for c in row) for row in self.rows)

    def to_grid(self) -> list[list[str]]:
        """문자열 2차원 배열. 짧은 행은 빈 칸으로 패딩한다."""
        w = self.n_cols
        return [[c.text for c in row] + [""] * (w - len(row)) for row in self.rows]

    def normalize(self) -> None:
        """모든 행의 길이를 n_cols로 맞춘다. Writer의 전제 조건."""
        w = self.n_cols
        for row in self.rows:
            while len(row) < w:
                row.append(Cell())


@dataclass
class Image(Block):
    data: bytes = b""
    #: 'png' | 'jpeg' | 'gif' | 'bmp' | 'emf' ...
    fmt: str = "png"
    width_pt: Optional[float] = None
    height_pt: Optional[float] = None
    alt_text: Optional[str] = None
    #: 원본 파일 내 이름(디버깅/추출용).
    name: Optional[str] = None


@dataclass
class PageBreak(Block):
    pass


@dataclass
class SectionBreak(Block):
    """구역 분리. 용지 방향/여백이 바뀌는 지점."""

    page_width_pt: Optional[float] = None
    page_height_pt: Optional[float] = None
    landscape: bool = False


# --------------------------------------------------------------------------
# 문서 수준
# --------------------------------------------------------------------------


class DocKind(enum.Enum):
    """IR이 본질적으로 무엇이었는지에 대한 힌트.

    Writer가 레이아웃을 고를 때 참고한다. 예를 들어 WORKBOOK을 docx로 쓸 때는
    시트마다 제목 문단을 넣어주는 편이 읽기 좋다.
    """

    DOCUMENT = "document"
    WORKBOOK = "workbook"


@dataclass
class DocMeta:
    title: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    keywords: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None
    #: 원본 포맷 확장자(점 없음). 변환 이력 추적용.
    source_format: Optional[str] = None
    #: Reader가 남기는 자유 형식 부가정보.
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class PageNumberSpec:
    """쪽 번호 표기.

    한/글은 이 정보를 본문이 아니라 구역 설정(ctrl)에 둔다. 옮기지 않으면
    변환 결과에 쪽 번호가 통째로 사라져 원본과 인상이 달라진다.
    """

    #: 'BOTTOM_CENTER' | 'BOTTOM_LEFT' | 'BOTTOM_RIGHT' | 'TOP_*'
    position: str = "BOTTOM_CENTER"
    #: 번호 양옆에 붙는 장식 문자. '-' 이면 `- 1 -` 형태가 된다.
    side_char: str = ""
    start: int = 1


# A4 기준 기본값(pt)
A4_WIDTH_PT = 595.28
A4_HEIGHT_PT = 841.89


@dataclass
class Document:
    blocks: list[Block] = field(default_factory=list)
    meta: DocMeta = field(default_factory=DocMeta)
    kind: DocKind = DocKind.DOCUMENT
    page_width_pt: float = A4_WIDTH_PT
    page_height_pt: float = A4_HEIGHT_PT
    #: (좌, 상, 우, 하) 여백 pt
    margin_pt: tuple[float, float, float, float] = (56.7, 56.7, 56.7, 56.7)
    #: 쪽 번호 표기. None이면 넣지 않는다.
    page_number: Optional["PageNumberSpec"] = None

    # -- 편의 접근자 -------------------------------------------------------

    def paragraphs(self) -> Iterator[Paragraph]:
        for b in self.blocks:
            if isinstance(b, Paragraph):
                yield b

    def tables(self) -> Iterator[Table]:
        """중첩 표까지 포함해 모든 Table을 순회한다."""
        yield from _walk_tables(self.blocks)

    def images(self) -> Iterator[Image]:
        for b in self.blocks:
            if isinstance(b, Image):
                yield b

    def to_text(self, table_sep: str = "\t") -> str:
        out: list[str] = []
        for b in self.blocks:
            if isinstance(b, Paragraph):
                out.append(b.text)
            elif isinstance(b, Table):
                if b.name:
                    out.append("# " + b.name)
                out.append(b.to_text(table_sep))
            elif isinstance(b, PageBreak):
                out.append("\f")
        return "\n".join(out)

    def stats(self) -> dict[str, int]:
        n_par = n_tbl = n_img = n_chr = 0
        for b in self.blocks:
            if isinstance(b, Paragraph):
                n_par += 1
                n_chr += len(b.text)
            elif isinstance(b, Table):
                n_tbl += 1
            elif isinstance(b, Image):
                n_img += 1
        return {
            "blocks": len(self.blocks),
            "paragraphs": n_par,
            "tables": n_tbl,
            "images": n_img,
            "characters": n_chr,
        }

    def is_empty(self) -> bool:
        return not any(
            (isinstance(b, Paragraph) and not b.is_blank())
            or isinstance(b, (Table, Image))
            for b in self.blocks
        )

    def compact(self) -> None:
        """연속된 빈 문단을 하나로 줄이고 Run을 병합한다."""
        out: list[Block] = []
        blank_streak = 0
        for b in self.blocks:
            if isinstance(b, Paragraph):
                b.merge_runs()
                if b.is_blank():
                    blank_streak += 1
                    if blank_streak > 1:
                        continue
                else:
                    blank_streak = 0
            else:
                blank_streak = 0
            out.append(b)
        self.blocks = out


def _walk_tables(blocks: Sequence[Block]) -> Iterator[Table]:
    for b in blocks:
        if isinstance(b, Table):
            yield b
            for row in b.rows:
                for cell in row:
                    yield from _walk_tables(cell.blocks)
