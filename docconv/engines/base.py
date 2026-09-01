"""엔진 인터페이스.

엔진은 "이 포맷 쌍을 실제로 변환할 수 있는 주체"다. 세 종류가 있다.

* NativeEngine       순수 파이썬. 의존성이 가볍고 어디서나 돈다.
* LibreOfficeEngine  soffice --headless. 크로스 플랫폼, 커버리지가 넓다.
* MsOfficeEngine     Windows COM. .doc/.xls 최종 폴백.

파이프라인이 (src, dst)에 대해 사용 가능한 엔진을 우선순위대로 시도한다.
"""

from __future__ import annotations

import abc
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from ..options import ConvertOptions


@dataclass
class EngineResult:
    """엔진 실행 결과."""

    output: Path
    warnings: list[str] = field(default_factory=list)
    #: 엔진이 부수적으로 만든 파일들(예: 시트별 CSV)
    extra_outputs: list[Path] = field(default_factory=list)


class Engine(abc.ABC):
    """변환 엔진의 공통 계약."""

    #: 낮을수록 먼저 시도한다.
    priority: int = 100
    name: str = "engine"
    #: 사람이 읽는 설명. 진단 출력에 쓴다.
    description: str = ""

    @abc.abstractmethod
    def is_available(self) -> bool:
        """이 시스템에서 실제로 쓸 수 있는지."""

    @abc.abstractmethod
    def supports(self, src: str, dst: str) -> bool:
        """src -> dst 직접 변환이 가능한지."""

    @abc.abstractmethod
    def convert(
        self,
        src: Path,
        dst: Path,
        src_fmt: str,
        dst_fmt: str,
        opts: ConvertOptions,
    ) -> EngineResult:
        """실제 변환. 실패하면 ConversionError 계열을 던진다."""

    # -- 진단 -------------------------------------------------------------

    def status(self) -> str:
        mark = "사용 가능" if self.is_available() else "사용 불가"
        return f"{self.name:14s} {mark:9s} {self.description}"

    def __repr__(self) -> str:  # pragma: no cover
        return f"<{type(self).__name__} {self.name}>"
