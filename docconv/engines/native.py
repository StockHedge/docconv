"""네이티브 엔진 - 순수 파이썬 Reader/Writer 조합.

외부 프로그램 없이 동작하므로 1순위다. Reader가 IR을 만들고 Writer가 IR을
소비하는 구조라, 지원 여부는 단순히 "그 포맷의 Reader/Writer가 등록되어
있는가"로 결정된다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

from ..errors import UnsupportedRouteError
from ..ir import Document
from ..options import ConvertOptions
from .base import Engine, EngineResult

# --------------------------------------------------------------------------
# 레지스트리
# --------------------------------------------------------------------------

ReaderFn = Callable[..., Document]
WriterFn = Callable[..., object]


def _readers() -> dict[str, ReaderFn]:
    """지연 임포트. 쓰지 않는 포맷의 의존성까지 끌어오지 않기 위해서다."""
    from ..readers import (
        docx_reader,
        hwp_reader,
        hwpx_reader,
        pdf_reader,
        sheet_reader,
        text_reader,
    )

    return {
        "pdf": pdf_reader.read,
        "docx": docx_reader.read,
        "hwpx": hwpx_reader.read,
        "hwp": hwp_reader.read,
        "xlsx": sheet_reader.read_xlsx,
        "xls": sheet_reader.read_xls,
        "csv": sheet_reader.read_csv,
        "txt": text_reader.read_txt,
        "md": text_reader.read_md,
        "html": text_reader.read_html,
        "rtf": text_reader.read_rtf,
        "odt": text_reader.read_odt,
    }


def _writers() -> dict[str, WriterFn]:
    from ..writers import (
        docx_writer,
        hwpx_writer,
        pdf_writer,
        sheet_writer,
        text_writer,
    )

    return {
        "pdf": pdf_writer.write_pdf,
        "docx": docx_writer.write_docx,
        "hwpx": hwpx_writer.write_hwpx,
        "xlsx": sheet_writer.write_xlsx,
        "csv": sheet_writer.write_csv,
        "txt": text_writer.write_txt,
        "md": text_writer.write_md,
        "html": text_writer.write_html,
        "json": text_writer.write_json,
    }


#: 모듈 로드 시점이 아니라 최초 사용 시점에 채운다.
_READER_CACHE: Optional[dict[str, ReaderFn]] = None
_WRITER_CACHE: Optional[dict[str, WriterFn]] = None


def readers() -> dict[str, ReaderFn]:
    global _READER_CACHE
    if _READER_CACHE is None:
        _READER_CACHE = _readers()
    return _READER_CACHE


def writers() -> dict[str, WriterFn]:
    global _WRITER_CACHE
    if _WRITER_CACHE is None:
        _WRITER_CACHE = _writers()
    return _WRITER_CACHE


def native_readable() -> list[str]:
    return sorted(readers())


def native_writable() -> list[str]:
    return sorted(writers())


# --------------------------------------------------------------------------
# 엔진
# --------------------------------------------------------------------------


class NativeEngine(Engine):
    priority = 10
    name = "native"
    description = "순수 파이썬 (외부 프로그램 불필요)"

    def is_available(self) -> bool:
        return True

    def supports(self, src: str, dst: str) -> bool:
        return src in readers() and dst in writers()

    def convert(
        self,
        src: Path,
        dst: Path,
        src_fmt: str,
        dst_fmt: str,
        opts: ConvertOptions,
    ) -> EngineResult:
        doc, warnings = self.read(src, src_fmt, opts)
        extra = self.write(doc, dst, dst_fmt, opts)
        return EngineResult(output=dst, warnings=warnings, extra_outputs=extra)

    # -- 분리된 단계(파이프라인이 IR을 재사용할 수 있도록 공개) -----------

    @staticmethod
    def read(
        src: Path, src_fmt: str, opts: ConvertOptions
    ) -> tuple[Document, list[str]]:
        fn = readers().get(src_fmt)
        if fn is None:
            raise UnsupportedRouteError(
                src_fmt, "?", reason="읽기를 지원하지 않는 형식"
            )

        kwargs: dict[str, object] = {}
        if src_fmt == "pdf":
            if opts.password:
                kwargs["password"] = opts.password
            if opts.page_range:
                kwargs["page_range"] = opts.page_range
        elif src_fmt == "csv":
            if opts.csv_input_encoding:
                kwargs["encoding"] = opts.csv_input_encoding
            if opts.csv_input_delimiter:
                kwargs["delimiter"] = opts.csv_input_delimiter

        # Reader 인스턴스가 경고를 모으는 경우를 위해 클래스 경로도 지원한다.
        doc = fn(src, **kwargs) if kwargs else fn(src)
        warnings = list(getattr(doc, "_warnings", []) or [])
        extra_warn = doc.meta.extra.get("warnings") if doc.meta.extra else None
        if extra_warn:
            warnings.extend(extra_warn)
        return doc, warnings

    @staticmethod
    def write(
        doc: Document, dst: Path, dst_fmt: str, opts: ConvertOptions
    ) -> list[Path]:
        fn = writers().get(dst_fmt)
        if fn is None:
            raise UnsupportedRouteError(
                "?", dst_fmt, reason="쓰기를 지원하지 않는 형식"
            )
        result = fn(doc, dst, opts)
        if isinstance(result, list):
            return [Path(p) for p in result if Path(p) != dst]
        return []
