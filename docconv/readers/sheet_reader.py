"""스프레드시트 Reader: XLSX / XLSM / XLS / CSV / TSV.

세 포맷을 한 모듈에 두는 이유는 IR로의 환원 방식이 동일하기 때문이다.
시트 하나가 Table 블록 하나가 되고, 문서 종류는 WORKBOOK이 된다.

한국어 환경에서 CSV는 CP949(엑셀 기본 저장)와 UTF-8이 뒤섞여 돌아다닌다.
인코딩을 잘못 잡으면 조용히 깨진 글자가 나오므로, BOM -> 후보 순차 시도 ->
한글 음절 비율 채점 순으로 판별한다.
"""

from __future__ import annotations

import csv
import datetime as _dt
import io
from pathlib import Path
from typing import Any, Iterable, Optional

from ..errors import CorruptFileError, ReadError
from ..ir import Cell, DocKind, DocMeta, Document, Paragraph, Table
from ..util import units

# --------------------------------------------------------------------------
# XLSX / XLSM
# --------------------------------------------------------------------------


class XlsxReader:
    def __init__(self, path: Path | str, *, max_rows: int = 200_000) -> None:
        self.path = Path(path)
        self.max_rows = max_rows
        self.warnings: list[str] = []

    def read(self) -> Document:
        try:
            import openpyxl
        except ImportError as e:  # pragma: no cover
            raise ReadError("openpyxl이 필요합니다. pip install openpyxl") from e

        try:
            wb = openpyxl.load_workbook(str(self.path), data_only=True, rich_text=False)
        except TypeError:
            wb = openpyxl.load_workbook(str(self.path), data_only=True)
        except Exception as e:
            raise CorruptFileError(
                "Excel 파일을 열 수 없습니다.", detail=f"{self.path}: {e}"
            ) from e

        doc = Document(kind=DocKind.WORKBOOK, meta=self._meta(wb))
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                self.warnings.append(f"숨겨진 시트를 건너뜁니다: {ws.title}")
                continue
            t = self._sheet(ws)
            if t is not None:
                doc.blocks.append(t)
        if not doc.blocks:
            self.warnings.append("읽을 수 있는 시트가 없습니다.")
        wb.close()
        return doc

    def _meta(self, wb) -> DocMeta:
        p = wb.properties
        return DocMeta(
            title=p.title or None,
            author=p.creator or None,
            subject=p.subject or None,
            keywords=p.keywords or None,
            created=p.created.isoformat() if p.created else None,
            modified=p.modified.isoformat() if p.modified else None,
            source_format="xlsx",
        )

    def _sheet(self, ws) -> Optional[Table]:
        dims = ws.calculate_dimension()
        rows = list(ws.iter_rows(values_only=False))
        if not rows:
            return None

        # 실제 데이터가 있는 범위로 잘라낸다. 엑셀은 빈 셀도 차원에 포함시킨다.
        last_row = 0
        last_col = 0
        for ri, row in enumerate(rows):
            for ci, c in enumerate(row):
                if c.value is not None and str(c.value).strip() != "":
                    last_row = max(last_row, ri)
                    last_col = max(last_col, ci)
        if last_row == 0 and last_col == 0:
            first = rows[0][0].value if rows and rows[0] else None
            if first is None:
                return Table(rows=[], name=ws.title)

        rows = rows[: last_row + 1]
        if len(rows) > self.max_rows:
            self.warnings.append(
                f"'{ws.title}' 시트가 {len(rows)}행이라 {self.max_rows}행까지만 읽습니다."
            )
            rows = rows[: self.max_rows]

        # 병합 정보: (row, col) -> (rowspan, colspan) / placeholder 집합
        anchors: dict[tuple[int, int], tuple[int, int]] = {}
        covered: set[tuple[int, int]] = set()
        for rng in ws.merged_cells.ranges:
            r0, c0 = rng.min_row - 1, rng.min_col - 1
            anchors[(r0, c0)] = (
                rng.max_row - rng.min_row + 1,
                rng.max_col - rng.min_col + 1,
            )
            for r in range(rng.min_row - 1, rng.max_row):
                for c in range(rng.min_col - 1, rng.max_col):
                    if (r, c) != (r0, c0):
                        covered.add((r, c))

        out = Table(name=ws.title)
        for ri, row in enumerate(rows):
            line: list[Cell] = []
            for ci, c in enumerate(row[: last_col + 1]):
                if (ri, ci) in covered:
                    line.append(Cell(merged_placeholder=True))
                    continue
                cell = Cell.of(_fmt_value(c.value))
                cell.raw_value = c.value
                cell.number_format = (
                    c.number_format if c.number_format != "General" else None
                )
                rs, cs = anchors.get((ri, ci), (1, 1))
                cell.row_span, cell.col_span = rs, cs
                cell.background = _fill_rgb(c)
                line.append(cell)
            out.rows.append(line)

        out.col_widths_pt = _col_widths(ws, last_col + 1)
        out.header_row = _guess_header(out)
        return out


def _fill_rgb(c) -> Optional[tuple[int, int, int]]:
    try:
        f = c.fill
        if f is None or f.fill_type != "solid":
            return None
        rgb = getattr(f.start_color, "rgb", None)
        if not isinstance(rgb, str):
            return None
        v = units.hex_to_rgb(rgb)
        return None if v in ((255, 255, 255), (0, 0, 0)) else v
    except Exception:
        return None


def _col_widths(ws, n: int) -> Optional[list[float]]:
    out: list[float] = []
    for i in range(1, n + 1):
        try:
            from openpyxl.utils import get_column_letter

            d = ws.column_dimensions.get(get_column_letter(i))
            out.append(units.excel_col_width_to_pt(d.width) if d and d.width else 0.0)
        except Exception:
            out.append(0.0)
    return out if any(out) else None


def _fmt_value(v: Any) -> str:
    """셀 값을 사람이 읽는 문자열로. 부동소수 잡음(1.0000000002)을 제거한다."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return repr(round(v, 10))
    if isinstance(v, _dt.datetime):
        if v.hour or v.minute or v.second:
            return v.strftime("%Y-%m-%d %H:%M:%S")
        return v.strftime("%Y-%m-%d")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, _dt.time):
        return v.strftime("%H:%M:%S")
    return str(v)


def _guess_header(t: Table) -> bool:
    """첫 행이 전부 문자열이고 둘째 행에 숫자가 있으면 머리글로 본다."""
    if len(t.rows) < 2:
        return False
    first, second = t.rows[0], t.rows[1]
    if not any(c.text.strip() for c in first):
        return False
    if any(
        isinstance(c.raw_value, (int, float)) and not isinstance(c.raw_value, bool)
        for c in first
    ):
        return False
    return any(
        isinstance(c.raw_value, (int, float, _dt.date, _dt.datetime))
        and not isinstance(c.raw_value, bool)
        for c in second
    )


# --------------------------------------------------------------------------
# XLS (Excel 97-2003)
# --------------------------------------------------------------------------


class XlsReader:
    def __init__(self, path: Path | str) -> None:
        self.path = Path(path)
        self.warnings: list[str] = []

    def read(self) -> Document:
        try:
            import xlrd
        except ImportError as e:
            raise ReadError(
                "구형 .xls를 읽으려면 xlrd가 필요합니다.\n"
                "  pip install xlrd\n"
                "(xlrd 2.x는 .xls 전용입니다. .xlsx는 openpyxl이 처리합니다.)"
            ) from e

        try:
            book = xlrd.open_workbook(str(self.path), formatting_info=False)
        except Exception as e:
            raise CorruptFileError(
                "XLS 파일을 열 수 없습니다.", detail=f"{self.path}: {e}"
            ) from e

        doc = Document(kind=DocKind.WORKBOOK, meta=DocMeta(source_format="xls"))
        for sh in book.sheets():
            t = Table(name=sh.name)
            for ri in range(sh.nrows):
                line: list[Cell] = []
                for ci in range(sh.ncols):
                    cv = sh.cell(ri, ci)
                    val = _xlrd_value(cv, book.datemode)
                    c = Cell.of(_fmt_value(val))
                    c.raw_value = val
                    line.append(c)
                t.rows.append(line)
            t.header_row = _guess_header(t)
            doc.blocks.append(t)
        return doc


def _xlrd_value(cv, datemode: int) -> Any:
    import xlrd

    if cv.ctype == xlrd.XL_CELL_DATE:
        try:
            y, mo, d, h, mi, s = xlrd.xldate_as_tuple(cv.value, datemode)
            if (y, mo, d) == (0, 0, 0):
                return _dt.time(h, mi, s)
            return _dt.datetime(y, mo, d, h, mi, s)
        except Exception:
            return cv.value
    if cv.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cv.value)
    if cv.ctype == xlrd.XL_CELL_EMPTY:
        return None
    if cv.ctype == xlrd.XL_CELL_ERROR:
        return "#ERROR"
    return cv.value


# --------------------------------------------------------------------------
# CSV / TSV
# --------------------------------------------------------------------------

#: 시도 순서. 한국 환경에서는 CP949 저장본이 매우 흔하다.
ENCODING_CANDIDATES = (
    "utf-8-sig",
    "utf-8",
    "cp949",  # = ms949, euc-kr 상위호환
    "euc-kr",
    "utf-16",
    "utf-16-le",
    "latin-1",  # 최후 수단: 절대 실패하지 않는다
)


def detect_encoding(raw: bytes) -> str:
    """BOM -> 디코딩 성공 -> 한글 비율 점수 순으로 인코딩을 고른다.

    latin-1은 어떤 바이트열도 디코딩에 성공하므로 '성공'만으로는 판별이 안 된다.
    그래서 디코딩된 문자열에서 한글/ASCII 비율과 깨진 문자 비율을 점수화한다.
    """
    if raw.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return "utf-16"

    sample = raw[:262_144]
    best, best_score = "utf-8", float("-inf")
    for enc in ENCODING_CANDIDATES:
        try:
            text = sample.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
        score = _score_text(text)
        # 엄격한 인코딩이 성공하면 가산점을 준다(latin-1 남용 방지).
        if enc.startswith("utf-8"):
            score += 12
        elif enc in ("cp949", "euc-kr"):
            score += 8
        elif enc == "latin-1":
            score -= 40
        if score > best_score:
            best, best_score = enc, score
    return best


def _score_text(s: str) -> float:
    if not s:
        return 0.0
    hangul = sum(1 for ch in s if "가" <= ch <= "힣")
    ascii_print = sum(1 for ch in s if 32 <= ord(ch) < 127)
    # 제어문자/사용영역/치환문자는 잘못된 디코딩의 신호
    bad = sum(
        1
        for ch in s
        if ch == "�" or "" <= ch <= "" or (ord(ch) < 32 and ch not in "\r\n\t")
    )
    n = len(s)
    return (hangul * 2 + ascii_print) / n * 100 - bad / n * 400


def detect_delimiter(text: str, path: Path) -> str:
    if path.suffix.lower() == ".tsv":
        return "\t"
    head = "\n".join(text.splitlines()[:40])
    try:
        return csv.Sniffer().sniff(head, delimiters=",;\t|").delimiter
    except csv.Error:
        pass
    # Sniffer 실패 시 빈도로 결정한다.
    counts = {d: head.count(d) for d in (",", ";", "\t", "|")}
    best = max(counts, key=lambda k: counts[k])
    return best if counts[best] else ","


class CsvReader:
    def __init__(
        self,
        path: Path | str,
        *,
        encoding: Optional[str] = None,
        delimiter: Optional[str] = None,
        has_header: Optional[bool] = None,
    ) -> None:
        self.path = Path(path)
        self.encoding = encoding
        self.delimiter = delimiter
        self.has_header = has_header
        self.warnings: list[str] = []

    def read(self) -> Document:
        raw = self.path.read_bytes()
        enc = self.encoding or detect_encoding(raw)
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            self.warnings.append(f"{enc} 디코딩 실패 - 손상 문자를 대체합니다.")
            text = raw.decode(enc, errors="replace")
        if enc not in ("utf-8", "utf-8-sig"):
            self.warnings.append(f"인코딩을 {enc}로 판별했습니다.")

        delim = self.delimiter or detect_delimiter(text, self.path)
        rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))
        # 뒤쪽 빈 행 제거
        while rows and not any(c.strip() for c in rows[-1]):
            rows.pop()

        t = Table(name=self.path.stem)
        width = max((len(r) for r in rows), default=0)
        for r in rows:
            line = [Cell.of(v) for v in r] + [Cell() for _ in range(width - len(r))]
            for c, v in zip(line, r):
                c.raw_value = _coerce(v)
            t.rows.append(line)
        t.header_row = (
            self.has_header if self.has_header is not None else _guess_header(t)
        )

        doc = Document(
            kind=DocKind.WORKBOOK,
            meta=DocMeta(
                title=self.path.stem,
                source_format="csv",
                extra={"encoding": enc, "delimiter": delim},
            ),
            blocks=[t],
        )
        return doc


def _coerce(s: str) -> Any:
    """CSV 문자열을 가능한 경우 숫자로 바꾼다. XLSX로 내보낼 때 타입을 살린다."""
    v = s.strip()
    if not v:
        return None
    if v.lower() in ("true", "false"):
        return v.lower() == "true"
    neg = v.startswith("(") and v.endswith(")")
    body = v[1:-1] if neg else v
    body = body.replace(",", "")
    try:
        if body.lstrip("+-").isdigit():
            n = int(body)
            return -n if neg else n
        f = float(body)
        return -f if neg else f
    except ValueError:
        return s


# --------------------------------------------------------------------------
# 진입점
# --------------------------------------------------------------------------


def read_xlsx(path: Path | str) -> Document:
    r = XlsxReader(path)
    return _attach(r.read(), r)


def read_xls(path: Path | str) -> Document:
    r = XlsReader(path)
    return _attach(r.read(), r)


def read_csv(path: Path | str, **kw: Any) -> Document:
    r = CsvReader(path, **kw)
    return _attach(r.read(), r)


def _attach(doc, reader):
    """Reader가 모은 경고를 Document에 실어 파이프라인까지 전달한다."""
    if getattr(reader, "warnings", None):
        doc.meta.extra.setdefault("warnings", []).extend(reader.warnings)
    return doc
