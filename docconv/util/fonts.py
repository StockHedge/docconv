"""한글 지원 폰트 탐색.

PDF를 만들 때 가장 흔한 사고가 한글이 전부 네모(□)로 나오는 것이다. PDF의
표준 14폰트(Helvetica, Times, Courier)에는 한글 글리프가 아예 없어서, CJK
문자를 쓰려면 **폰트 파일을 문서에 임베딩**해야 한다.

이 모듈은 OS별 표준 위치에서 한글 폰트를 찾아 (본문, 굵게) 한 쌍을 돌려준다.
사용자 환경마다 설치 폰트가 다르므로 선호 순서대로 훑고, 하나도 없으면
호출자가 명확한 오류를 낼 수 있도록 None을 돌려준다.

폰트 파일이 없는 환경(도커, 최소 설치 리눅스)을 위해 CID 폰트 대체 경로도
안내한다.
"""

from __future__ import annotations

import functools
import platform
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional


@dataclass(frozen=True)
class FontPair:
    """본문/굵게 한 쌍. 굵은 파일이 없으면 bold는 regular와 같다."""

    name: str
    regular: Path
    bold: Path
    #: TTC(컬렉션) 안의 인덱스. TTF면 0.
    regular_index: int = 0
    bold_index: int = 0

    @property
    def is_collection(self) -> bool:
        return self.regular.suffix.lower() in (".ttc", ".otc")


# 선호 순서: 가독성이 좋고 폭이 안정적인 고딕 계열 우선.
_WINDOWS_CANDIDATES: tuple[tuple[str, str, str], ...] = (
    ("맑은 고딕", "malgun.ttf", "malgunbd.ttf"),
    ("나눔고딕", "NanumGothic.ttf", "NanumGothicBold.ttf"),
    ("굴림", "gulim.ttc", "gulim.ttc"),
    ("바탕", "batang.ttc", "batang.ttc"),
    ("돋움", "dotum.ttc", "dotum.ttc"),
)

_MACOS_CANDIDATES: tuple[tuple[str, str, str], ...] = (
    ("Apple SD Gothic Neo", "AppleSDGothicNeo.ttc", "AppleSDGothicNeo.ttc"),
    ("나눔고딕", "NanumGothic.ttf", "NanumGothicBold.ttf"),
    ("AppleGothic", "AppleGothic.ttf", "AppleGothic.ttf"),
    ("Noto Sans KR", "NotoSansKR-Regular.otf", "NotoSansKR-Bold.otf"),
)

_LINUX_CANDIDATES: tuple[tuple[str, str, str], ...] = (
    ("나눔고딕", "NanumGothic.ttf", "NanumGothicBold.ttf"),
    ("Noto Sans CJK KR", "NotoSansCJK-Regular.ttc", "NotoSansCJK-Bold.ttc"),
    ("Noto Sans KR", "NotoSansKR-Regular.otf", "NotoSansKR-Bold.otf"),
    ("은 돋움", "UnDotum.ttf", "UnDotumBold.ttf"),
)


def _font_dirs() -> list[Path]:
    sysname = platform.system()
    dirs: list[Path] = []
    home = Path.home()

    if sysname == "Windows":
        import os

        win = Path(os.environ.get("SystemRoot", r"C:\Windows"))
        dirs += [
            win / "Fonts",
            home / "AppData" / "Local" / "Microsoft" / "Windows" / "Fonts",
        ]
    elif sysname == "Darwin":
        dirs += [
            Path("/System/Library/Fonts"),
            Path("/System/Library/Fonts/Supplemental"),
            Path("/Library/Fonts"),
            home / "Library" / "Fonts",
        ]
    else:
        dirs += [
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            home / ".local" / "share" / "fonts",
            home / ".fonts",
        ]
    return [d for d in dirs if d.is_dir()]


def _candidates() -> tuple[tuple[str, str, str], ...]:
    s = platform.system()
    if s == "Windows":
        return _WINDOWS_CANDIDATES
    if s == "Darwin":
        return _MACOS_CANDIDATES
    return _LINUX_CANDIDATES


def _find_file(dirs: Iterable[Path], filename: str) -> Optional[Path]:
    lower = filename.lower()
    for d in dirs:
        direct = d / filename
        if direct.is_file():
            return direct
        # 리눅스는 하위 디렉터리에 흩어져 있다.
        try:
            for p in d.rglob(filename):
                if p.is_file():
                    return p
            for p in d.rglob("*"):
                if p.is_file() and p.name.lower() == lower:
                    return p
        except (OSError, PermissionError):
            continue
    return None


@functools.lru_cache(maxsize=1)
def find_korean_font() -> Optional[FontPair]:
    """시스템에서 한글 폰트 한 쌍을 찾는다. 결과는 캐시된다."""
    dirs = _font_dirs()
    if not dirs:
        return None
    for name, reg, bold in _candidates():
        rp = _find_file(dirs, reg)
        if rp is None:
            continue
        bp = _find_file(dirs, bold) or rp
        return FontPair(name=name, regular=rp, bold=bp)
    # 이름으로 못 찾으면 CJK로 보이는 아무 폰트나 찾는다.
    return _fallback_scan(dirs)


def _fallback_scan(dirs: Iterable[Path]) -> Optional[FontPair]:
    keywords = (
        "gothic",
        "gulim",
        "batang",
        "dotum",
        "malgun",
        "nanum",
        "noto",
        "cjk",
        "korea",
    )
    for d in dirs:
        try:
            for p in sorted(d.iterdir()):
                if not p.is_file() or p.suffix.lower() not in (
                    ".ttf",
                    ".ttc",
                    ".otf",
                    ".otc",
                ):
                    continue
                if any(k in p.name.lower() for k in keywords):
                    return FontPair(name=p.stem, regular=p, bold=p)
        except (OSError, PermissionError):
            continue
    return None


def korean_font_or_raise() -> FontPair:
    from ..errors import WriteError

    fp = find_korean_font()
    if fp is None:
        raise WriteError(
            "한글을 표시할 수 있는 글꼴을 찾지 못했습니다. "
            "PDF에 한글을 넣으려면 글꼴 파일이 필요합니다.",
            detail=_install_hint(),
        )
    return fp


def _install_hint() -> str:
    s = platform.system()
    if s == "Windows":
        return "일반적으로 C:\\Windows\\Fonts\\malgun.ttf 가 있어야 합니다."
    if s == "Darwin":
        return "일반적으로 /System/Library/Fonts/AppleSDGothicNeo.ttc 가 있어야 합니다."
    return "예: sudo apt install fonts-nanum  또는  fonts-noto-cjk"


def describe() -> str:
    """진단 출력용 한 줄 요약."""
    fp = find_korean_font()
    if fp is None:
        return "한글 글꼴: 없음 (PDF 출력 시 한글이 깨질 수 있음)"
    same = " (굵게 동일)" if fp.bold == fp.regular else ""
    return f"한글 글꼴: {fp.name} <- {fp.regular}{same}"


if __name__ == "__main__":  # pragma: no cover
    print(describe())
    for d in _font_dirs():
        print(" 검색 경로:", d)
