"""LibreOffice headless 엔진.

크로스 플랫폼(Windows/macOS/Linux)이고 커버리지가 넓어, 네이티브가 못 하는
경로(.doc 읽기, .xls 쓰기 등)의 주 폴백이다. 설치되어 있지 않으면 조용히
비활성화되고 파이프라인이 다른 경로를 찾는다.

주의점
------
* LibreOffice는 **같은 사용자 프로필을 두 프로세스가 동시에 쓰지 못한다.**
  사용자가 LibreOffice를 켜 둔 상태로 변환하면 실패하거나 GUI가 뜬다. 그래서
  변환마다 임시 프로필(`-env:UserInstallation`)을 지정한다.
* 출력 파일 이름을 지정할 수 없다. `--outdir` 에 원본 이름 + 새 확장자로
  떨어지므로, 임시 디렉터리에 만든 뒤 목적지로 옮긴다.
* `--convert-to` 에 필터 이름을 붙여야 원하는 결과가 나오는 경우가 있다
  (예: CSV의 UTF-8 인코딩 지정).
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from ..errors import ConversionError, EngineUnavailableError
from ..options import ConvertOptions
from .base import Engine, EngineResult

#: LibreOffice가 읽을 수 있는 포맷
_READ = {
    "doc",
    "docx",
    "odt",
    "rtf",
    "txt",
    "html",
    "hwp",
    "hwpx",
    "xls",
    "xlsx",
    "csv",
    "ods",
    "pdf",
}
#: 쓸 수 있는 포맷
_WRITE = {
    "pdf",
    "docx",
    "doc",
    "odt",
    "rtf",
    "txt",
    "html",
    "xlsx",
    "xls",
    "csv",
    "ods",
}

#: 확장자 -> LibreOffice 변환 대상 문자열(필터 포함)
_TARGET = {
    "pdf": "pdf",
    "docx": "docx:MS Word 2007 XML",
    "doc": "doc:MS Word 97",
    "odt": "odt",
    "rtf": "rtf",
    "txt": "txt:Text (encoded):UTF8",
    "html": "html",
    "xlsx": "xlsx:Calc MS Excel 2007 XML",
    "xls": "xls:MS Excel 97",
    # 44=UTF-8, 9=구분자(TAB) 대신 44(,) / 필드구분자,텍스트구분자,인코딩
    "csv": "csv:Text - txt - csv (StarCalc):44,34,76",
    "ods": "ods",
}


def _candidate_paths() -> list[Path]:
    s = platform.system()
    out: list[Path] = []
    if s == "Windows":
        for base in (
            os.environ.get("ProgramFiles", r"C:\Program Files"),
            os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
        ):
            out.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
            out.append(Path(base) / "LibreOffice 7" / "program" / "soffice.exe")
    elif s == "Darwin":
        out += [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    else:
        out += [
            Path("/usr/bin/soffice"),
            Path("/usr/local/bin/soffice"),
            Path("/opt/libreoffice/program/soffice"),
            Path("/snap/bin/libreoffice"),
        ]
    return out


def find_soffice() -> Optional[Path]:
    """PATH와 표준 설치 위치에서 soffice 실행 파일을 찾는다."""
    env = os.environ.get("DOCCONV_SOFFICE")
    if env and Path(env).is_file():
        return Path(env)
    for name in ("soffice", "libreoffice", "soffice.exe"):
        w = shutil.which(name)
        if w:
            return Path(w)
    for p in _candidate_paths():
        if p.is_file():
            return p
    return None


class LibreOfficeEngine(Engine):
    priority = 50
    name = "libreoffice"
    description = "LibreOffice headless (.doc/.xls 등 폭넓은 폴백)"

    def __init__(self) -> None:
        self._exe: Optional[Path] = None
        self._checked = False

    # -- 가용성 -----------------------------------------------------------

    @property
    def exe(self) -> Optional[Path]:
        if not self._checked:
            self._exe = find_soffice()
            self._checked = True
        return self._exe

    def is_available(self) -> bool:
        return self.exe is not None

    def supports(self, src: str, dst: str) -> bool:
        return src in _READ and dst in _WRITE and dst in _TARGET

    # -- 변환 -------------------------------------------------------------

    def convert(
        self,
        src: Path,
        dst: Path,
        src_fmt: str,
        dst_fmt: str,
        opts: ConvertOptions,
    ) -> EngineResult:
        exe = self.exe
        if exe is None:
            raise EngineUnavailableError(self.name, hint=install_hint())

        target = _TARGET.get(dst_fmt)
        if target is None:
            raise ConversionError(f"LibreOffice가 {dst_fmt} 출력을 지원하지 않습니다.")

        warnings: list[str] = []
        with tempfile.TemporaryDirectory(prefix="docconv-lo-") as tmp:
            tmpdir = Path(tmp)
            profile = (tmpdir / "profile").as_uri()
            cmd = [
                str(exe),
                "--headless",
                "--norestore",
                "--nolockcheck",
                "--nodefault",
                "--nofirststartwizard",
                f"-env:UserInstallation={profile}",
                "--convert-to",
                target,
                "--outdir",
                str(tmpdir),
                str(src),
            ]
            try:
                proc = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=opts.engine_timeout,
                    check=False,
                    creationflags=_no_window_flag(),
                )
            except subprocess.TimeoutExpired as e:
                raise ConversionError(
                    f"LibreOffice 변환이 {opts.engine_timeout:.0f}초 안에 끝나지 "
                    "않았습니다."
                ) from e
            except OSError as e:
                raise ConversionError(
                    "LibreOffice 실행에 실패했습니다.", detail=str(e)
                ) from e

            produced = _find_output(tmpdir, src.stem, dst_fmt)
            if produced is None:
                err = (proc.stderr or b"").decode("utf-8", "replace").strip()
                out = (proc.stdout or b"").decode("utf-8", "replace").strip()
                raise ConversionError(
                    "LibreOffice가 결과 파일을 만들지 못했습니다.",
                    detail=(err or out or f"exit={proc.returncode}")[:500],
                )

            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(produced), str(dst))

        return EngineResult(output=dst, warnings=warnings)


def _find_output(tmpdir: Path, stem: str, dst_fmt: str) -> Optional[Path]:
    """LibreOffice가 만든 파일을 찾는다. 확장자가 예상과 다를 수 있다."""
    exact = tmpdir / f"{stem}.{dst_fmt}"
    if exact.is_file():
        return exact
    candidates = [p for p in tmpdir.iterdir() if p.is_file() and p.name != "profile"]
    if len(candidates) == 1:
        return candidates[0]
    for p in candidates:
        if p.stem == stem:
            return p
    return candidates[0] if candidates else None


def _no_window_flag() -> int:
    """Windows에서 콘솔 창이 깜빡이지 않게 한다."""
    if platform.system() == "Windows":
        return getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return 0


def install_hint() -> str:
    s = platform.system()
    if s == "Windows":
        return (
            "LibreOffice를 설치하면 .doc/.xls 등 더 많은 형식을 변환할 수 있습니다.\n"
            "  https://ko.libreoffice.org/download/  (또는 winget install "
            "TheDocumentFoundation.LibreOffice)"
        )
    if s == "Darwin":
        return (
            "LibreOffice를 설치하면 .doc/.xls 등 더 많은 형식을 변환할 수 있습니다.\n"
            "  brew install --cask libreoffice"
        )
    return "sudo apt install libreoffice  (또는 배포판 패키지 관리자)"
