"""예외 계층.

호출자가 "무엇을 할 수 있는가"에 따라 예외를 나눈다.

    ConversionError                 모든 변환 실패의 기반
    +- UnsupportedFormatError       확장자 자체를 모른다 -> 사용자에게 알린다
    +- UnsupportedRouteError        포맷은 알지만 그 경로를 지원하지 않는다
    +- EngineUnavailableError       엔진 미설치/미탐지 -> 설치 안내가 유효
    +- ReadError                    입력을 못 읽는다(손상/암호/버전)
    |   +- EncryptedFileError       암호 걸림 -> 비밀번호를 받으면 재시도 가능
    |   +- CorruptFileError
    +- WriteError                   출력 생성 실패(권한/디스크/잠김)
"""

from __future__ import annotations

from typing import Optional


class ConversionError(Exception):
    """변환 실패의 기반 예외."""

    def __init__(self, message: str, *, detail: Optional[str] = None) -> None:
        super().__init__(message)
        self.message = message
        self.detail = detail

    def __str__(self) -> str:  # pragma: no cover - 표현용
        if self.detail:
            return f"{self.message} ({self.detail})"
        return self.message


class UnsupportedFormatError(ConversionError):
    """확장자를 인식하지 못했다."""


class UnsupportedRouteError(ConversionError):
    """src -> dst 경로를 처리할 엔진이 등록되어 있지 않다."""

    def __init__(self, src: str, dst: str, *, reason: Optional[str] = None) -> None:
        super().__init__(
            f"{src.upper()} -> {dst.upper()} 변환은 지원하지 않습니다.",
            detail=reason,
        )
        self.src = src
        self.dst = dst


class EngineUnavailableError(ConversionError):
    """엔진이 이 시스템에 없다. 설치하면 해결될 수 있다."""

    def __init__(self, engine: str, *, hint: Optional[str] = None) -> None:
        super().__init__(f"'{engine}' 엔진을 사용할 수 없습니다.", detail=hint)
        self.engine = engine
        self.hint = hint


class ReadError(ConversionError):
    """입력 파일을 읽지 못했다."""


class EncryptedFileError(ReadError):
    """암호로 보호된 파일."""

    def __init__(self, path: str) -> None:
        super().__init__(
            "암호로 보호된 파일입니다. 비밀번호를 지정하거나 보호를 해제해 주세요.",
            detail=path,
        )


class CorruptFileError(ReadError):
    """구조가 손상되었거나 형식이 선언과 다르다."""


class WriteError(ConversionError):
    """출력 파일을 만들지 못했다."""
