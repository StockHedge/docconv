"""IR -> 격자(2차원 표) 평탄화.

문서형(문단 + 표)을 CSV/XLSX 같은 격자 포맷으로 내보내려면 어딘가에서
"문단을 행으로 바꾸는" 결정을 내려야 한다. 그 결정을 Writer마다 중복 구현하지
않도록 이 모듈에 모았다.

병합 셀 처리 정책도 여기서 한 번에 적용한다.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

from ..ir import Block, Cell, DocKind, Document, Image, PageBreak, Paragraph, Table
from ..options import ConvertOptions, MergedCellPolicy, SheetLayout, TableFlatten


@dataclass
class Sheet:
    """평탄화 결과 한 장."""

    name: str
    #: 화면 표시용 문자열 격자
    rows: list[list[str]]
    #: 원본 Cell (타입/서식 보존용). rows와 같은 모양.
    cells: list[list[Optional[Cell]]]
    header_row: bool = False
    col_widths_pt: Optional[list[float]] = None
    #: (row, col) -> (rowspan, colspan). KEEP 정책에서만 채운다.
    merges: dict[tuple[int, int], tuple[int, int]] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.merges is None:
            self.merges = {}

    @property
    def n_cols(self) -> int:
        return max((len(r) for r in self.rows), default=0)

    def pad(self) -> None:
        w = self.n_cols
        for r in self.rows:
            r.extend([""] * (w - len(r)))
        for c in self.cells:
            c.extend([None] * (w - len(c)))


def to_sheets(doc: Document, opts: ConvertOptions) -> list[Sheet]:
    """Document를 시트 목록으로 평탄화한다."""
    tables = [b for b in doc.blocks if isinstance(b, Table)]

    if doc.kind is DocKind.WORKBOOK or (
        opts.table_flatten is TableFlatten.TABLES_ONLY and tables
    ):
        sheets = [
            _table_to_sheet(t, opts, _sheet_name(t, i)) for i, t in enumerate(tables)
        ]
    else:
        sheets = _document_to_sheets(doc, opts)

    if not sheets:
        sheets = [Sheet(name="Sheet1", rows=[], cells=[])]

    if opts.sheet_layout is SheetLayout.SINGLE and len(sheets) > 1:
        sheets = [_concat(sheets)]

    for s in sheets:
        s.pad()
    _dedupe_names(sheets)
    return sheets


# --------------------------------------------------------------------------
# 문서형
# --------------------------------------------------------------------------


def pick_primary_sheet(sheets: list["Sheet"]) -> int:
    """여러 시트 중 '주 출력물'로 삼을 시트의 인덱스를 고른다.

    왜 필요한가
    -----------
    CSV는 한 파일에 시트 하나만 담을 수 있다. 그래서 문서에서 시트가 여러 장
    나오면 첫 시트가 ``이름.csv`` 가 되고 나머지는 ``이름_표1.csv`` 처럼
    갈라진다. 문제는 **첫 시트가 항상 알맹이는 아니라는 점**이다.

    실제 사례: 한/글 기획서 양식은 문서 전체가 표 한 장 안에 들어 있다. 이때
    '본문' 시트에는 ``[표 1: ...]`` 같은 자리표시자 두 줄만 남아 33바이트짜리
    CSV가 주 출력물이 되고, 정작 12,000자짜리 내용은 ``_표2.csv`` 로 밀린다.

    채택한 규칙: **첫 시트를 존중하되, 그 시트가 사실상 비었을 때만 넘긴다.**

    후보가 여럿이었다.

    * 글자 수 최대 - 이번 사례는 해결하지만, 표가 여러 장인 진짜 통합문서에서
      주 파일이 세 번째 시트가 되는 식으로 순서가 뒤집혀 예측하기 어렵다.
    * 행 수 최대 - 데이터 표에는 맞지만 문서형에는 잘 안 맞는다.
    * 첫 시트 고정 - 예측은 쉽지만 33바이트짜리 껍데기가 주 파일이 된다.

    첫 시트를 기본으로 두면 "원본에서 먼저 나온 것이 주 파일"이라는 규칙이
    유지되어 결과를 예상하기 쉽다. 다만 그 시트에 사실상 내용이 없다면 그건
    사용자가 원한 결과가 아니므로, 그때만 가장 알찬 시트로 넘긴다.

    :param sheets: 평탄화된 시트 목록. 비어 있지 않다.
    :return: 주 출력으로 쓸 시트의 인덱스
    """

    def weight(s: "Sheet") -> int:
        return sum(len(c.strip()) for row in s.rows for c in row)

    #: 이보다 적으면 "내용이 없는 껍데기"로 본다. 자리표시자 몇 줄 수준이다.
    min_meaningful = 50

    if not sheets:
        return 0
    if weight(sheets[0]) >= min_meaningful:
        return 0
    richest = max(range(len(sheets)), key=lambda i: weight(sheets[i]))
    return richest if weight(sheets[richest]) > weight(sheets[0]) else 0



def _document_to_sheets(doc: Document, opts: ConvertOptions) -> list[Sheet]:
    """문단은 1열 행으로, 표는 그대로. 표는 별도 시트로 분리한다.

    표를 본문 흐름에 그대로 끼워 넣으면 열 수가 뒤죽박죽이 되어 스프레드시트
    에서 다루기 어렵다. 그래서 기본은 '본문 시트 + 표별 시트'다.
    """
    body_rows: list[list[str]] = []
    body_cells: list[list[Optional[Cell]]] = []
    table_sheets: list[Sheet] = []
    t_idx = 0

    for b in doc.blocks:
        if isinstance(b, Paragraph):
            text = b.text
            if b.heading:
                text = f"{'#' * b.heading} {text}".strip()
            body_rows.append([text])
            body_cells.append([None])
        elif isinstance(b, Table):
            name = _sheet_name(b, t_idx)
            table_sheets.append(_table_to_sheet(b, opts, name))
            body_rows.append([f"[표 {t_idx + 1}: {name}]"])
            body_cells.append([None])
            t_idx += 1
        elif isinstance(b, PageBreak) and opts.keep_page_breaks:
            body_rows.append([""])
            body_cells.append([None])
        elif isinstance(b, Image):
            body_rows.append([f"[그림: {b.name or 'image'}]"])
            body_cells.append([None])

    out: list[Sheet] = []
    if any(r and r[0].strip() for r in body_rows):
        out.append(Sheet(name="본문", rows=body_rows, cells=body_cells))
    out.extend(table_sheets)
    return out


# --------------------------------------------------------------------------
# 표
# --------------------------------------------------------------------------


def _table_to_sheet(t: Table, opts: ConvertOptions, name: str) -> Sheet:
    t.normalize()
    rows: list[list[str]] = []
    cells: list[list[Optional[Cell]]] = []
    merges: dict[tuple[int, int], tuple[int, int]] = {}

    # 1차: 값 그대로 채우고 병합 정보를 모은다.
    for ri, row in enumerate(t.rows):
        line_txt: list[str] = []
        line_cell: list[Optional[Cell]] = []
        for ci, c in enumerate(row):
            line_txt.append("" if c.merged_placeholder else c.text)
            line_cell.append(None if c.merged_placeholder else c)
            if not c.merged_placeholder and (c.col_span > 1 or c.row_span > 1):
                merges[(ri, ci)] = (c.row_span, c.col_span)
        rows.append(line_txt)
        cells.append(line_cell)

    if opts.merged_cells is MergedCellPolicy.DUPLICATE:
        _duplicate_merged(rows, cells, merges)
        merges = {}
    elif opts.merged_cells is MergedCellPolicy.PLACEHOLDER:
        merges = {}
    # KEEP이면 merges를 그대로 넘겨 XLSX Writer가 실제 병합을 만든다.

    return Sheet(
        name=name,
        rows=rows,
        cells=cells,
        header_row=t.header_row,
        col_widths_pt=t.col_widths_pt,
        merges=merges,
    )


def _duplicate_merged(
    rows: list[list[str]],
    cells: list[list[Optional[Cell]]],
    merges: dict[tuple[int, int], tuple[int, int]],
) -> None:
    """병합 영역 전체에 좌상단 값을 복제한다.

    피벗테이블/필터/정렬을 걸 때는 빈 칸보다 이쪽이 훨씬 쓸모 있다.
    """
    for (r0, c0), (rs, cs) in merges.items():
        try:
            val = rows[r0][c0]
            src = cells[r0][c0]
        except IndexError:
            continue
        for dr in range(rs):
            for dc in range(cs):
                r, c = r0 + dr, c0 + dc
                if (dr or dc) and r < len(rows) and c < len(rows[r]):
                    rows[r][c] = val
                    cells[r][c] = src


# --------------------------------------------------------------------------
# 보조
# --------------------------------------------------------------------------


def _sheet_name(t: Table, idx: int) -> str:
    raw = (t.name or "").strip()
    if not raw:
        raw = f"표{idx + 1}"
    return sanitize_sheet_name(raw)


#: Excel 시트 이름에 쓸 수 없는 문자
_BAD_SHEET_CHARS = set(r"[]:*?/\\")


def sanitize_sheet_name(name: str) -> str:
    """Excel 규칙(31자, 금지문자, 작은따옴표로 시작/끝 불가)에 맞춘다."""
    s = "".join("_" if ch in _BAD_SHEET_CHARS else ch for ch in name)
    s = s.replace("\n", " ").replace("\r", " ").strip()
    s = s.strip("'")
    if not s:
        s = "Sheet"
    return s[:31]


def _dedupe_names(sheets: list[Sheet]) -> None:
    seen: dict[str, int] = {}
    for s in sheets:
        base = s.name
        if base not in seen:
            seen[base] = 1
            continue
        seen[base] += 1
        suffix = f"_{seen[base]}"
        s.name = (
            (base[: 31 - len(suffix)] + suffix)
            if len(base) + len(suffix) > 31
            else base + suffix
        )


def concat_rows(sheets: list[Sheet]) -> list[list[str]]:
    return _concat(sheets).rows


def _concat(sheets: list[Sheet]) -> Sheet:
    """여러 시트를 한 장으로. 시트 사이에 이름 행과 빈 행을 넣는다."""
    rows: list[list[str]] = []
    cells: list[list[Optional[Cell]]] = []
    for i, s in enumerate(sheets):
        if i:
            rows.append([""])
            cells.append([None])
        if len(sheets) > 1:
            rows.append([f"# {s.name}"])
            cells.append([None])
        rows.extend(s.rows)
        cells.extend(s.cells)
    out = Sheet(name=sheets[0].name if sheets else "Sheet1", rows=rows, cells=cells)
    out.pad()
    return out
