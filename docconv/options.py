"""변환 옵션.

Reader/Writer/엔진이 공유하는 단일 설정 객체다. CLI 플래그와 GUI 위젯이
모두 이 하나로 수렴하므로, 새 옵션을 추가할 때 여기만 고치면 된다.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Optional


class OnConflict(enum.Enum):
    """출력 파일이 이미 있을 때의 처리."""

    RENAME = "rename"  # foo.pdf -> foo (1).pdf
    OVERWRITE = "overwrite"
    SKIP = "skip"
    ERROR = "error"


class MergedCellPolicy(enum.Enum):
    """병합된 셀을 격자 포맷(CSV/XLSX)으로 내보낼 때의 처리.

    PLACEHOLDER : 좌상단에만 값, 나머지는 빈 칸. 원본 모양에 가깝다.
    DUPLICATE   : 병합 영역 전체에 같은 값을 채운다. 필터/피벗에 유리하다.
    KEEP        : 실제 병합을 유지한다(XLSX만 가능. CSV는 PLACEHOLDER와 동일).
    """

    PLACEHOLDER = "placeholder"
    DUPLICATE = "duplicate"
    KEEP = "keep"


class TableFlatten(enum.Enum):
    """문서형(문단+표)을 격자 포맷으로 내보낼 때 무엇을 담을지.

    ALL         : 문단은 한 줄씩, 표는 그대로. 문서 전체를 보존한다.
    TABLES_ONLY : 표만 추출한다. 데이터 분석용.
    """

    ALL = "all"
    TABLES_ONLY = "tables_only"


class SheetLayout(enum.Enum):
    """표형(여러 시트)을 XLSX로 내보낼 때의 배치."""

    PER_SHEET = "per_sheet"  # 표 하나당 시트 하나
    SINGLE = "single"  # 한 시트에 이어 붙이기


@dataclass
class ConvertOptions:
    # -- 공통 -------------------------------------------------------------
    on_conflict: OnConflict = OnConflict.RENAME
    #: 원본 서식(굵게/크기/색)을 최대한 보존할지. False면 텍스트만 옮긴다.
    keep_formatting: bool = True
    include_images: bool = True
    #: 변환 후 결과를 다시 열어 검증할지(느리지만 안전)
    verify_output: bool = True

    # -- CSV --------------------------------------------------------------
    #: 기본값이 utf-8-sig인 이유: Excel은 BOM 없는 UTF-8 CSV를 CP949로
    #: 오독해 한글을 깨뜨린다. BOM을 붙이면 Excel/메모장 모두 올바로 연다.
    csv_encoding: str = "utf-8-sig"
    csv_delimiter: str = ","
    csv_line_terminator: str = "\r\n"
    #: 읽을 때 강제할 인코딩. None이면 자동 판별.
    csv_input_encoding: Optional[str] = None
    csv_input_delimiter: Optional[str] = None

    # -- 표 ---------------------------------------------------------------
    merged_cells: MergedCellPolicy = MergedCellPolicy.PLACEHOLDER
    table_flatten: TableFlatten = TableFlatten.ALL
    sheet_layout: SheetLayout = SheetLayout.PER_SHEET
    #: 표를 텍스트 포맷으로 내보낼 때의 열 구분자
    table_text_sep: str = "\t"

    # -- PDF --------------------------------------------------------------
    #: 사용할 한글 글꼴 파일. None이면 시스템에서 자동 탐색.
    pdf_font_path: Optional[str] = None
    pdf_base_size_pt: float = 10.5
    #: 텍스트가 없는 PDF(스캔본)를 만났을 때 경고만 하고 계속할지
    pdf_allow_empty: bool = True

    # -- 문서 -------------------------------------------------------------
    #: 표형 -> 문서형 변환 시 시트 이름을 제목 문단으로 넣을지
    sheet_name_as_heading: bool = True
    #: 페이지 나누기 유지
    keep_page_breaks: bool = True

    # -- 엔진 -------------------------------------------------------------
    #: 사용할 엔진을 강제한다. None이면 자동 선택.
    force_engine: Optional[str] = None
    #: 외부 엔진(LibreOffice 등) 사용을 허용할지
    allow_external_engines: bool = True
    #: 외부 엔진 타임아웃(초)
    engine_timeout: float = 180.0

    # -- 입력 -------------------------------------------------------------
    password: Optional[str] = None
    #: PDF에서 읽을 페이지 범위 (0-base, [start, end))
    page_range: Optional[tuple[int, int]] = None

    def replace(self, **kw: object) -> "ConvertOptions":
        from dataclasses import replace as _replace

        return _replace(self, **kw)  # type: ignore[arg-type]
