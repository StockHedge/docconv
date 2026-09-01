"""지원 포맷 정의와 파일 형식 판별.

확장자만 믿지 않는다. 매직 바이트로 실제 컨테이너를 확인해 "이름은 .hwp인데
내용은 HWPX ZIP" 같은 흔한 사고를 잡아낸다.
"""

from __future__ import annotations

import enum
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


class Family(enum.Enum):
    """포맷의 성격. 라우팅과 UI 그룹핑에 쓴다."""

    DOCUMENT = "document"  # 흐름 문서
    SPREADSHEET = "spreadsheet"  # 격자 데이터
    FIXED = "fixed"  # 고정 레이아웃(PDF)
    TEXT = "text"  # 평문/마크업


@dataclass(frozen=True)
class FormatSpec:
    key: str
    ext: tuple[str, ...]
    label: str
    family: Family
    can_read: bool
    can_write: bool
    note: str = ""

    @property
    def primary_ext(self) -> str:
        return self.ext[0]


FORMATS: dict[str, FormatSpec] = {
    "pdf": FormatSpec(
        "pdf",
        (".pdf",),
        "PDF 문서",
        Family.FIXED,
        True,
        True,
        "읽기: 텍스트/표/이미지 추출. 쓰기: 텍스트 레이아웃 재구성",
    ),
    "docx": FormatSpec(
        "docx",
        (".docx",),
        "Word 문서",
        Family.DOCUMENT,
        True,
        True,
    ),
    "doc": FormatSpec(
        "doc",
        (".doc",),
        "Word 97-2003 문서",
        Family.DOCUMENT,
        True,
        True,
        "외부 엔진(LibreOffice 또는 MS Word) 필요",
    ),
    "hwpx": FormatSpec(
        "hwpx",
        (".hwpx",),
        "한글 문서 (OWPML)",
        Family.DOCUMENT,
        True,
        True,
    ),
    "hwp": FormatSpec(
        "hwp",
        (".hwp",),
        "한글 문서 (HWP 5.0)",
        Family.DOCUMENT,
        True,
        False,
        "읽기 전용. 쓰기는 HWPX로 대체하세요",
    ),
    "xlsx": FormatSpec(
        "xlsx",
        (".xlsx", ".xlsm"),
        "Excel 통합 문서",
        Family.SPREADSHEET,
        True,
        True,
    ),
    "xls": FormatSpec(
        "xls",
        (".xls",),
        "Excel 97-2003 통합 문서",
        Family.SPREADSHEET,
        True,
        True,
        "읽기: 네이티브. 쓰기: 외부 엔진 필요",
    ),
    "csv": FormatSpec(
        "csv",
        (".csv", ".tsv"),
        "CSV / TSV",
        Family.SPREADSHEET,
        True,
        True,
    ),
    # 보너스 포맷 - 파이프라인이 IR 기반이므로 거의 공짜로 얻는다.
    "txt": FormatSpec(
        "txt",
        (".txt",),
        "일반 텍스트",
        Family.TEXT,
        True,
        True,
    ),
    "md": FormatSpec(
        "md",
        (".md", ".markdown"),
        "Markdown",
        Family.TEXT,
        True,
        True,
    ),
    "html": FormatSpec(
        "html",
        (".html", ".htm"),
        "HTML",
        Family.TEXT,
        True,
        True,
    ),
    "rtf": FormatSpec(
        "rtf",
        (".rtf",),
        "서식 있는 텍스트",
        Family.DOCUMENT,
        True,
        False,
        "읽기 전용(striprtf)",
    ),
    "odt": FormatSpec(
        "odt",
        (".odt",),
        "OpenDocument 텍스트",
        Family.DOCUMENT,
        True,
        True,
    ),
    "json": FormatSpec(
        "json",
        (".json",),
        "JSON (IR 덤프)",
        Family.TEXT,
        False,
        True,
        "디버깅/파이프라인 연계용",
    ),
}

#: 확장자 -> 포맷 키
_EXT_MAP: dict[str, str] = {e: spec.key for spec in FORMATS.values() for e in spec.ext}

#: 사용자가 주로 쓰는 8종. GUI에서 우선 노출한다.
CORE_FORMATS: tuple[str, ...] = (
    "pdf",
    "docx",
    "hwp",
    "hwpx",
    "doc",
    "csv",
    "xlsx",
    "xls",
)


def readable_formats() -> list[str]:
    return [k for k, s in FORMATS.items() if s.can_read]


def writable_formats() -> list[str]:
    return [k for k, s in FORMATS.items() if s.can_write]


def format_of_ext(ext: str) -> Optional[str]:
    return _EXT_MAP.get(ext.lower() if ext.startswith(".") else "." + ext.lower())


def spec(key: str) -> FormatSpec:
    return FORMATS[key]


# --------------------------------------------------------------------------
# 실제 내용 기반 판별
# --------------------------------------------------------------------------

_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # OLE2 복합 문서
_ZIP_MAGIC = b"PK\x03\x04"
_PDF_MAGIC = b"%PDF-"
_RTF_MAGIC = b"{\\rt"


def sniff(path: Path) -> Optional[str]:
    """파일 내용으로 포맷을 추정한다. 판단 불가면 None.

    확장자보다 이쪽을 신뢰한다. 다만 CSV/TXT처럼 매직이 없는 포맷은
    확장자로만 판별 가능하므로 None을 돌려준다.
    """
    try:
        with path.open("rb") as f:
            head = f.read(8)
    except OSError:
        return None

    if head.startswith(_PDF_MAGIC):
        return "pdf"
    if head.startswith(_RTF_MAGIC):
        return "rtf"
    if head.startswith(_CFB_MAGIC):
        return _sniff_cfb(path)
    if head.startswith(_ZIP_MAGIC):
        return _sniff_zip(path)
    return None


def _sniff_cfb(path: Path) -> Optional[str]:
    """OLE2 컨테이너 안의 스트림 이름으로 hwp / doc / xls를 구분한다."""
    try:
        import olefile
    except ImportError:  # pragma: no cover
        return None
    try:
        if not olefile.isOleFile(str(path)):
            return None
        with olefile.OleFileIO(str(path)) as ole:
            names = {"/".join(p) for p in ole.listdir()}
            if "FileHeader" in names or "HwpSummaryInformation" in names:
                return "hwp"
            if "WordDocument" in names:
                return "doc"
            if "Workbook" in names or "Book" in names:
                return "xls"
    except Exception:
        return None
    return None


def _sniff_zip(path: Path) -> Optional[str]:
    """ZIP 컨테이너의 mimetype/구성으로 hwpx / docx / xlsx / odt를 구분한다."""
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "mimetype" in names:
                try:
                    mt = z.read("mimetype").decode("ascii", "ignore").strip()
                except Exception:
                    mt = ""
                if "hwp" in mt:
                    return "hwpx"
                if mt == "application/vnd.oasis.opendocument.text":
                    return "odt"
                if mt == "application/vnd.oasis.opendocument.spreadsheet":
                    return "ods"
            # 한/글은 mimetype 없이 배포되는 경우도 있다.
            if any(n.startswith("Contents/section") for n in names):
                return "hwpx"
            if "word/document.xml" in names:
                return "docx"
            if "xl/workbook.xml" in names:
                return "xlsx"
            if "ppt/presentation.xml" in names:
                return "pptx"
    except (zipfile.BadZipFile, OSError):
        return None
    return None


def detect(path: Path | str) -> str:
    """확장자 + 매직 바이트로 포맷 키를 결정한다.

    내용 판별이 성공하면 그쪽을 우선한다. 단 xlsx/xlsm처럼 같은 컨테이너를
    공유하는 경우엔 확장자가 더 구체적이므로 확장자를 남긴다.
    """
    from .errors import UnsupportedFormatError

    p = Path(path)
    by_ext = format_of_ext(p.suffix)
    by_content = sniff(p)

    if by_content is None:
        if by_ext is None:
            raise UnsupportedFormatError(
                f"지원하지 않는 파일 형식입니다: {p.suffix or '(확장자 없음)'}",
                detail=str(p),
            )
        return by_ext

    if by_ext == by_content:
        return by_ext
    # 같은 계열 안에서의 차이는 확장자를 존중한다(xlsx vs xlsm 등).
    if by_ext and FORMATS.get(by_ext) and FORMATS.get(by_content):
        if FORMATS[by_ext].family == FORMATS[by_content].family and by_ext in (
            "xlsx",
            "csv",
            "md",
            "html",
            "txt",
        ):
            return by_ext
    if by_content not in FORMATS:
        if by_ext:
            return by_ext
        raise UnsupportedFormatError(
            f"인식했으나 지원하지 않는 형식입니다: {by_content}", detail=str(p)
        )
    return by_content
