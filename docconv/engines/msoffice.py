"""MS Office 엔진 (Windows COM 전용).

이 시스템에 MS Office가 설치되어 있으면 `.doc` / `.xls` 처리에 쓸 수 있다.
LibreOffice보다 원본 재현도가 높지만 제약이 크다.

* **Windows에서만 동작한다.** macOS/Linux에서는 항상 비활성이다.
  (macOS의 Word는 AppleScript로 제어할 수 있으나 신뢰성이 낮아 넣지 않았다.)
* pywin32가 필요하다.
* Office 애플리케이션 프로세스를 띄우므로 느리고, 사용자가 같은 앱을 쓰고
  있으면 간섭이 생길 수 있다.

그래서 우선순위를 LibreOffice보다 뒤에 둔다.
"""

from __future__ import annotations

import os
import platform
from pathlib import Path
from typing import Optional

from ..errors import ConversionError, EngineUnavailableError
from ..options import ConvertOptions
from .base import Engine, EngineResult

IS_WINDOWS = platform.system() == "Windows"

# Word의 WdSaveFormat
# 이름만 보고 넘겨짚으면 안 되는 상수다. 특히 7은 wdFormatText(평문)가 아니라
# **wdFormatUnicodeText(UTF-16)** 다. 이걸 그냥 쓰면 결과가 UTF-16LE로 저장되어
# UTF-8로 읽는 쪽에서 전부 깨진다(실측: 깨진 문자 10,694개).
# 그래서 저장 뒤 _normalize_text_encoding() 으로 UTF-8로 되돌린다.
_WD_FORMAT = {
    "docx": 16,  # wdFormatDocumentDefault (= .docx)
    "doc": 0,  # wdFormatDocument97
    "pdf": 17,  # wdFormatPDF
    "rtf": 6,  # wdFormatRTF
    "txt": 7,  # wdFormatUnicodeText - Encoding 인자로 UTF-8을 요청한다
    "html": 8,  # wdFormatHTML
    "odt": 23,  # wdFormatOpenDocumentText
}

#: msoCharacterSet - UTF-8
_MSO_UTF8 = 65001

# Excel의 XlFileFormat
_XL_FORMAT = {
    "xlsx": 51,  # xlOpenXMLWorkbook
    "xls": 56,  # xlExcel8
    "csv": 62,  # xlCSVUTF8 (Excel 2016+). 미지원 시 6(xlCSV)로 폴백
    "pdf": None,  # ExportAsFixedFormat 사용
    "html": 44,  # xlHtml
}

_WORD_READ = {"doc", "docx", "rtf", "txt", "html", "odt"}
_WORD_WRITE = {"docx", "doc", "pdf", "rtf", "txt", "html", "odt"}
_EXCEL_READ = {"xls", "xlsx", "csv"}
_EXCEL_WRITE = {"xlsx", "xls", "csv", "pdf", "html"}


class MsOfficeEngine(Engine):
    priority = 70
    name = "msoffice"
    description = "MS Office COM (Windows 전용, .doc/.xls 최종 폴백)"

    def __init__(self) -> None:
        self._available: Optional[bool] = None

    # -- 가용성 -----------------------------------------------------------

    def is_available(self) -> bool:
        if self._available is not None:
            return self._available
        if not IS_WINDOWS:
            self._available = False
            return False
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            self._available = False
            return False
        self._available = _office_installed()
        return self._available

    def supports(self, src: str, dst: str) -> bool:
        if not IS_WINDOWS:
            return False
        word = src in _WORD_READ and dst in _WORD_WRITE
        excel = src in _EXCEL_READ and dst in _EXCEL_WRITE
        return word or excel

    # -- 변환 -------------------------------------------------------------

    def convert(
        self,
        src: Path,
        dst: Path,
        src_fmt: str,
        dst_fmt: str,
        opts: ConvertOptions,
    ) -> EngineResult:
        if not self.is_available():
            raise EngineUnavailableError(
                self.name,
                hint="Windows + MS Office + pywin32(pip install pywin32)가 필요합니다.",
            )
        dst.parent.mkdir(parents=True, exist_ok=True)
        if src_fmt in _EXCEL_READ:
            _excel_convert(src, dst, dst_fmt)
        else:
            _word_convert(src, dst, dst_fmt)
        return EngineResult(output=dst)


# --------------------------------------------------------------------------
# 구현
# --------------------------------------------------------------------------


def _office_installed() -> bool:
    try:
        import winreg
    except ImportError:
        return False
    for key in (r"Word.Application", r"Excel.Application"):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key):
                return True
        except OSError:
            continue
    return False


def _word_convert(src: Path, dst: Path, dst_fmt: str) -> None:
    import pythoncom
    import win32com.client

    fmt = _WD_FORMAT.get(dst_fmt)
    if fmt is None:
        raise ConversionError(f"Word가 {dst_fmt} 출력을 지원하지 않습니다.")

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(
            str(src.resolve()), ReadOnly=True, AddToRecentFiles=False, Visible=False
        )
        if dst_fmt == "pdf":
            doc.ExportAsFixedFormat(str(dst.resolve()), 17)
        elif dst_fmt == "txt":
            # Encoding 인자로 UTF-8을 요청하지만, Word 버전과 로캘에 따라
            # 무시되고 UTF-16이나 CP949로 저장되는 일이 잦다. 그래서 저장 뒤
            # 아래에서 인코딩을 다시 확인해 UTF-8로 맞춘다.
            doc.SaveAs2(str(dst.resolve()), FileFormat=fmt, Encoding=_MSO_UTF8)
        else:
            doc.SaveAs2(str(dst.resolve()), FileFormat=fmt)
    except Exception as e:
        raise ConversionError("Word 변환에 실패했습니다.", detail=str(e)) from e
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    if dst_fmt in ("txt", "html"):
        _normalize_text_encoding(dst)


def _normalize_text_encoding(path: Path) -> None:
    """Office가 만든 텍스트 파일을 UTF-8로 되돌린다.

    Word/Excel은 요청한 인코딩을 무시하고 UTF-16이나 시스템 코드페이지(한국어
    Windows는 CP949)로 저장하는 일이 잦다. 그대로 두면 다음 단계에서 한글이
    통째로 깨진다. 이 프로그램의 텍스트 출력은 항상 UTF-8이어야 하므로
    여기서 한 번 정규화한다.
    """
    from ..readers.sheet_reader import detect_encoding

    try:
        raw = path.read_bytes()
    except OSError:
        return
    if not raw:
        return

    enc = detect_encoding(raw)
    if enc in ("utf-8", "utf-8-sig"):
        return
    try:
        text = raw.decode(enc, errors="replace")
    except (LookupError, UnicodeDecodeError):
        return
    try:
        path.write_text(text, encoding="utf-8", newline="")
    except OSError:
        pass


def _excel_convert(src: Path, dst: Path, dst_fmt: str) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(str(src.resolve()), ReadOnly=True, UpdateLinks=0)
        if dst_fmt == "pdf":
            wb.ExportAsFixedFormat(0, str(dst.resolve()))
        elif dst_fmt == "csv":
            try:
                wb.SaveAs(str(dst.resolve()), FileFormat=62)  # xlCSVUTF8
            except Exception:
                wb.SaveAs(str(dst.resolve()), FileFormat=6)  # xlCSV (로캘 인코딩)
        else:
            fmt = _XL_FORMAT.get(dst_fmt)
            if fmt is None:
                raise ConversionError(f"Excel이 {dst_fmt} 출력을 지원하지 않습니다.")
            wb.SaveAs(str(dst.resolve()), FileFormat=fmt)
    except Exception as e:
        raise ConversionError("Excel 변환에 실패했습니다.", detail=str(e)) from e
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    if dst_fmt in ("csv", "html"):
        _normalize_text_encoding(dst)
