"""docconv - 로컬 문서 변환기.

PDF / DOCX / DOC / HWP / HWPX / XLSX / XLS / CSV 를 서로 변환한다.
외부 서비스에 파일을 올리지 않고 전부 이 컴퓨터에서 처리한다.

기본 사용법::

    from docconv import convert
    r = convert("기획서.hwpx", "pdf")
    print(r.summary())

여러 개를 한 번에::

    from docconv import convert_many
    results = convert_many(["a.hwp", "b.docx"], "pdf", out_dir="결과")

구조::

    ir.py          중간표현(Document IR) - 모든 변환의 중심
    formats.py     포맷 정의와 판별
    options.py     변환 옵션
    pipeline.py    엔진 선택, 브리지 경로 탐색, 실행
    readers/       각 포맷 -> IR
    writers/       IR -> 각 포맷
    engines/       native / libreoffice / msoffice
    gui/           tkinter 화면
"""

from __future__ import annotations

__version__ = "1.0.0"

from .errors import (
    ConversionError,
    CorruptFileError,
    EncryptedFileError,
    EngineUnavailableError,
    ReadError,
    UnsupportedFormatError,
    UnsupportedRouteError,
    WriteError,
)
from .formats import CORE_FORMATS, FORMATS, detect, readable_formats, writable_formats
from .options import (
    ConvertOptions,
    MergedCellPolicy,
    OnConflict,
    SheetLayout,
    TableFlatten,
)
from .pipeline import (
    ConvertResult,
    available_engines,
    convert,
    convert_many,
    diagnose,
    supported_targets,
)

__all__ = [
    "__version__",
    # 핵심 API
    "convert",
    "convert_many",
    "ConvertResult",
    "ConvertOptions",
    # 옵션 열거형
    "OnConflict",
    "MergedCellPolicy",
    "TableFlatten",
    "SheetLayout",
    # 포맷
    "FORMATS",
    "CORE_FORMATS",
    "detect",
    "readable_formats",
    "writable_formats",
    "supported_targets",
    # 진단
    "diagnose",
    "available_engines",
    # 예외
    "ConversionError",
    "UnsupportedFormatError",
    "UnsupportedRouteError",
    "EngineUnavailableError",
    "ReadError",
    "EncryptedFileError",
    "CorruptFileError",
    "WriteError",
]
