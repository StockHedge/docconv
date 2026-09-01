"""길이 단위 환산.

IR 내부 표준은 pt(1/72인치)다. 각 포맷의 고유 단위는 Reader/Writer 경계에서만
이 모듈을 통해 환산한다.

    HWPUNIT : 1/7200 인치 (한/글). 1pt = 100 HWPUNIT
    EMU     : 1/914400 인치 (OOXML DrawingML). 1pt = 12700 EMU
    twip    : 1/1440 인치 (Word 문단/여백). 1pt = 20 twip
    px      : 화면 픽셀. 96dpi 가정. 1pt = 4/3 px
"""

from __future__ import annotations

INCH_PT = 72.0
HWPUNIT_PER_PT = 100.0  # 7200 / 72
EMU_PER_PT = 12700.0  # 914400 / 72
TWIP_PER_PT = 20.0  # 1440 / 72
PX_PER_PT_96DPI = 96.0 / 72.0
MM_PER_INCH = 25.4


def hwp_to_pt(v: float | int | str | None) -> float:
    if v is None or v == "":
        return 0.0
    return float(v) / HWPUNIT_PER_PT


def pt_to_hwp(v: float) -> int:
    return int(round(v * HWPUNIT_PER_PT))


def emu_to_pt(v: float | int | None) -> float:
    return 0.0 if v is None else float(v) / EMU_PER_PT


def pt_to_emu(v: float) -> int:
    return int(round(v * EMU_PER_PT))


def twip_to_pt(v: float | int | None) -> float:
    return 0.0 if v is None else float(v) / TWIP_PER_PT


def pt_to_twip(v: float) -> int:
    return int(round(v * TWIP_PER_PT))


def px_to_pt(v: float, dpi: float = 96.0) -> float:
    return v * INCH_PT / dpi


def pt_to_px(v: float, dpi: float = 96.0) -> float:
    return v * dpi / INCH_PT


def mm_to_pt(v: float) -> float:
    return v / MM_PER_INCH * INCH_PT


def pt_to_mm(v: float) -> float:
    return v / INCH_PT * MM_PER_INCH


def excel_col_width_to_pt(width: float | None) -> float:
    """openpyxl의 열 너비(기본 글꼴 '0' 문자 개수)를 pt로 근사 변환.

    Excel의 열 너비 단위는 '표준 글꼴에서 숫자 0의 폭'이다. Calibri 11pt 기준
    1단위 = 7px = 5.25pt 로 근사한다. 정확한 값은 글꼴 메트릭에 의존하지만,
    변환 결과물의 열 비율을 보존하는 데는 충분하다.
    """
    if width is None:
        return 0.0
    return float(width) * 7.0 * INCH_PT / 96.0


def pt_to_excel_col_width(pt: float) -> float:
    if pt <= 0:
        return 8.43  # Excel 기본값
    return pt * 96.0 / INCH_PT / 7.0


def excel_row_height_to_pt(height: float | None) -> float:
    """Excel 행 높이는 이미 pt 단위다."""
    return 0.0 if height is None else float(height)


# ---- 색 -------------------------------------------------------------------


def hex_to_rgb(s: str | None) -> tuple[int, int, int] | None:
    """'#RRGGBB' / 'RRGGBB' / 'AARRGGBB' 를 (r,g,b)로. 실패 시 None."""
    if not s:
        return None
    t = s.strip().lstrip("#")
    if len(t) == 8:  # ARGB (openpyxl 스타일)
        t = t[2:]
    if len(t) != 6:
        return None
    try:
        return (int(t[0:2], 16), int(t[2:4], 16), int(t[4:6], 16))
    except ValueError:
        return None


def rgb_to_hex(rgb: tuple[int, int, int] | None, prefix: str = "#") -> str | None:
    if rgb is None:
        return None
    r, g, b = (max(0, min(255, int(c))) for c in rgb)
    return "%s%02X%02X%02X" % (prefix, r, g, b)


def hwp_color_to_rgb(v: str | int | None) -> tuple[int, int, int] | None:
    """한/글 색상 표기를 (r,g,b)로.

    HWPX는 '#RRGGBB' 문자열을 쓰고, HWP 바이너리는 리틀엔디언 정수(0x00BBGGRR)를
    쓴다. 둘 다 받아 처리한다.
    """
    if v is None or v == "":
        return None
    if isinstance(v, str):
        s = v.strip()
        if s.startswith("#"):
            return hex_to_rgb(s)
        try:
            v = int(s, 0)
        except ValueError:
            return hex_to_rgb(s)
    n = int(v)
    if n < 0:
        return None
    return (n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)


def rgb_to_hwp_color(rgb: tuple[int, int, int] | None) -> int:
    """(r,g,b) -> HWP 바이너리 정수(0x00BBGGRR)."""
    if rgb is None:
        return 0
    r, g, b = (max(0, min(255, int(c))) for c in rgb)
    return r | (g << 8) | (b << 16)
