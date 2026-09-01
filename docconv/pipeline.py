"""변환 파이프라인 - 엔진 선택, 경로 탐색, 실행.

`convert()` 하나가 공개 진입점이다. 내부적으로 이렇게 동작한다.

1. 입력 포맷을 확정한다(확장자 + 매직 바이트).
2. 출력 경로를 정하고 충돌 정책을 적용한다.
3. (src, dst)를 직접 처리할 수 있는 엔진을 우선순위대로 시도한다.
4. 직접 경로가 없으면 **중간 포맷을 경유하는 2단계 경로**를 찾는다.
   예: doc -> [LibreOffice] -> docx -> [Native] -> hwpx
5. 결과와 경고를 담은 ConvertResult를 돌려준다.
"""

from __future__ import annotations

import shutil
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence

from . import formats
from .engines.base import Engine, EngineResult
from .engines.libreoffice import LibreOfficeEngine
from .engines.msoffice import MsOfficeEngine
from .engines.native import NativeEngine
from .errors import (
    ConversionError,
    UnsupportedFormatError,
    UnsupportedRouteError,
    WriteError,
)
from .options import ConvertOptions, OnConflict

#: 경유지로 쓸 중간 포맷. IR 왕복 손실이 적은 순서로 둔다.
BRIDGE_FORMATS: tuple[str, ...] = ("docx", "xlsx", "html", "csv", "odt")

#: 경유 최대 단계 수. 3이면 A→B→C→D 까지 찾는다.
MAX_HOPS = 3


# --------------------------------------------------------------------------
# 결과
# --------------------------------------------------------------------------


@dataclass
class ConvertResult:
    source: Path
    output: Optional[Path]
    src_format: str
    dst_format: str
    ok: bool = True
    engine: str = ""
    #: 브리지를 거쳤다면 경유 포맷
    via: Optional[str] = None
    elapsed_sec: float = 0.0
    warnings: list[str] = field(default_factory=list)
    extra_outputs: list[Path] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def route(self) -> str:
        if self.via:
            return f"{self.src_format} → {self.via} → {self.dst_format}"
        return f"{self.src_format} → {self.dst_format}"

    def summary(self) -> str:
        if not self.ok:
            return f"실패  {self.source.name}  ({self.route})  {self.error}"
        extra = f" +{len(self.extra_outputs)}개" if self.extra_outputs else ""
        return (
            f"완료  {self.source.name} → {self.output.name if self.output else '?'}"
            f"{extra}  [{self.engine}] {self.elapsed_sec:.2f}s"
        )


ProgressFn = Callable[[str], None]


# --------------------------------------------------------------------------
# 엔진 레지스트리
# --------------------------------------------------------------------------

_ENGINES: Optional[list[Engine]] = None


def engines() -> list[Engine]:
    """우선순위 순 엔진 목록. 프로세스당 한 번만 만든다."""
    global _ENGINES
    if _ENGINES is None:
        _ENGINES = sorted(
            [NativeEngine(), LibreOfficeEngine(), MsOfficeEngine()],
            key=lambda e: e.priority,
        )
    return _ENGINES


def available_engines() -> list[Engine]:
    return [e for e in engines() if e.is_available()]


def engine_by_name(name: str) -> Optional[Engine]:
    for e in engines():
        if e.name == name:
            return e
    return None


def supported_targets(src_fmt: str, opts: Optional[ConvertOptions] = None) -> list[str]:
    """이 입력 포맷에서 만들어낼 수 있는 출력 포맷 목록(브리지 포함)."""
    opts = opts or ConvertOptions()
    out: set[str] = set()
    for dst in formats.writable_formats():
        if dst == src_fmt:
            continue
        if _direct_engine(src_fmt, dst, opts) is not None:
            out.add(dst)
        elif _bridge_route(src_fmt, dst, opts) is not None:
            out.add(dst)
    return sorted(out)


# --------------------------------------------------------------------------
# 경로 탐색
# --------------------------------------------------------------------------


def _usable(e: Engine, opts: ConvertOptions) -> bool:
    if not e.is_available():
        return False
    if opts.force_engine and e.name != opts.force_engine:
        return False
    if not opts.allow_external_engines and e.name != "native":
        return False
    return True


def _direct_engine(src: str, dst: str, opts: ConvertOptions) -> Optional[Engine]:
    for e in engines():
        if _usable(e, opts) and e.supports(src, dst):
            return e
    return None


def _bridge_route(
    src: str, dst: str, opts: ConvertOptions, max_hops: int = MAX_HOPS
) -> Optional[list[tuple[Engine, str]]]:
    """경유 경로를 너비 우선으로 찾는다. 결과는 [(엔진, 목적형식), ...].

    2단계 고정으로 두면 메워지지 않는 칸이 남는다. 예를 들어 ``doc -> xls`` 는
    어떤 2단계 조합으로도 불가능하다 - ``doc->docx`` 는 되지만 ``docx->xls`` 가
    안 되고(Excel은 docx를 못 읽는다), ``csv->xls`` 는 되지만 ``doc->csv`` 가
    안 된다(Word는 csv를 못 만든다). 실제로 필요한 경로는
    ``doc -> docx -> xlsx -> xls`` 3단계다.

    너비 우선이므로 항상 **가장 짧은 경로**를 고른다. 경유가 늘수록 서식 손실이
    쌓이므로 이 성질이 중요하다.
    """
    if max_hops < 1:
        return None

    from collections import deque

    queue: deque[tuple[str, list[tuple[Engine, str]]]] = deque([(src, [])])
    visited: set[str] = {src}

    while queue:
        cur, path = queue.popleft()
        if len(path) >= max_hops:
            continue
        for mid in BRIDGE_FORMATS:
            if mid == cur or mid == dst or mid in visited:
                continue
            e1 = _direct_engine(cur, mid, opts)
            if e1 is None:
                continue
            step = path + [(e1, mid)]
            e2 = _direct_engine(mid, dst, opts)
            if e2 is not None:
                return step + [(e2, dst)]
            if len(step) < max_hops - 1:
                visited.add(mid)
                queue.append((mid, step))
    return None


# --------------------------------------------------------------------------
# 출력 경로
# --------------------------------------------------------------------------


def resolve_output(
    src: Path,
    dst_fmt: str,
    out_dir: Optional[Path],
    opts: ConvertOptions,
) -> Optional[Path]:
    """출력 경로를 정한다. SKIP 정책에서 건너뛸 경우 None."""
    ext = formats.spec(dst_fmt).primary_ext
    target_dir = Path(out_dir) if out_dir else src.parent
    target = target_dir / (src.stem + ext)

    if not target.exists():
        return target
    if target.resolve() == src.resolve():
        # 같은 파일을 덮어쓰는 사고를 막는다.
        return _numbered(target)
    if opts.on_conflict is OnConflict.OVERWRITE:
        return target
    if opts.on_conflict is OnConflict.SKIP:
        return None
    if opts.on_conflict is OnConflict.ERROR:
        raise WriteError(f"출력 파일이 이미 있습니다: {target}")
    return _numbered(target)


def _numbered(target: Path) -> Path:
    for i in range(1, 10_000):
        cand = target.with_name(f"{target.stem} ({i}){target.suffix}")
        if not cand.exists():
            return cand
    raise WriteError(f"사용 가능한 파일 이름을 찾지 못했습니다: {target}")


# --------------------------------------------------------------------------
# 실행
# --------------------------------------------------------------------------


def convert(
    src: Path | str,
    dst_fmt: str,
    *,
    out_dir: Optional[Path | str] = None,
    out_path: Optional[Path | str] = None,
    opts: Optional[ConvertOptions] = None,
    progress: Optional[ProgressFn] = None,
) -> ConvertResult:
    """파일 하나를 dst_fmt로 변환한다.

    out_path를 주면 그 경로에 정확히 쓴다(충돌 정책 무시). 아니면 out_dir
    (또는 원본과 같은 폴더)에 원본 이름 + 새 확장자로 쓴다.
    """
    opts = opts or ConvertOptions()
    src = Path(src)
    t0 = time.perf_counter()

    if not src.is_file():
        return ConvertResult(
            source=src,
            output=None,
            src_format="?",
            dst_format=dst_fmt,
            ok=False,
            error="파일을 찾을 수 없습니다.",
        )

    try:
        src_fmt = formats.detect(src)
    except UnsupportedFormatError as e:
        return ConvertResult(
            source=src,
            output=None,
            src_format="?",
            dst_format=dst_fmt,
            ok=False,
            error=str(e),
        )

    if dst_fmt not in formats.FORMATS or not formats.spec(dst_fmt).can_write:
        return ConvertResult(
            source=src,
            output=None,
            src_format=src_fmt,
            dst_format=dst_fmt,
            ok=False,
            error=f"'{dst_fmt}' 형식으로는 저장할 수 없습니다.",
        )

    if out_path is not None:
        target: Optional[Path] = Path(out_path)
    else:
        try:
            target = resolve_output(
                src, dst_fmt, Path(out_dir) if out_dir else None, opts
            )
        except WriteError as e:
            return ConvertResult(
                source=src,
                output=None,
                src_format=src_fmt,
                dst_format=dst_fmt,
                ok=False,
                error=str(e),
            )
    if target is None:
        return ConvertResult(
            source=src,
            output=None,
            src_format=src_fmt,
            dst_format=dst_fmt,
            ok=True,
            engine="skip",
            warnings=["이미 존재하여 건너뛰었습니다."],
            elapsed_sec=time.perf_counter() - t0,
        )

    if src_fmt == dst_fmt and out_path is None:
        # 같은 포맷이면 굳이 재조립하지 않고 복사한다(무손실).
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, target)
            return ConvertResult(
                source=src,
                output=target,
                src_format=src_fmt,
                dst_format=dst_fmt,
                engine="copy",
                elapsed_sec=time.perf_counter() - t0,
                warnings=["입력과 출력 형식이 같아 파일을 복사했습니다."],
            )
        except OSError as e:
            return ConvertResult(
                source=src,
                output=None,
                src_format=src_fmt,
                dst_format=dst_fmt,
                ok=False,
                error=f"복사 실패: {e}",
            )

    _say(progress, f"{src.name}: {src_fmt} → {dst_fmt} 변환 중...")

    errors: list[str] = []

    # 1) 직접 경로 - 가능한 모든 엔진을 순서대로 시도한다.
    for e in engines():
        if not _usable(e, opts) or not e.supports(src_fmt, dst_fmt):
            continue
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            res = e.convert(src, target, src_fmt, dst_fmt, opts)
            return _finish(src, res, src_fmt, dst_fmt, e.name, None, t0, opts, progress)
        except ConversionError as ex:
            errors.append(f"[{e.name}] {ex}")
            _say(progress, f"  {e.name} 실패 - 다음 방법을 시도합니다.")
        except Exception as ex:  # 엔진 내부의 예상 못한 오류도 폴백 대상
            errors.append(f"[{e.name}] {type(ex).__name__}: {ex}")
            _say(progress, f"  {e.name} 오류 - 다음 방법을 시도합니다.")

    # 2) 경유 경로 (최단 경로를 너비 우선으로 탐색)
    route = _bridge_route(src_fmt, dst_fmt, opts)
    if route:
        hops = " → ".join(f for _, f in route[:-1])
        engines_used = "+".join(e.name for e, _ in route)
        _say(progress, f"  {hops}을(를) 거쳐 변환합니다 ({engines_used}).")
        try:
            with tempfile.TemporaryDirectory(prefix="docconv-bridge-") as tmp:
                cur_path, cur_fmt = src, src_fmt
                warnings: list[str] = []
                last = None
                for idx, (eng, out_fmt) in enumerate(route):
                    if idx == len(route) - 1:
                        out_path = target
                        out_path.parent.mkdir(parents=True, exist_ok=True)
                    else:
                        out_path = Path(tmp) / (
                            f"step{idx}" + formats.spec(out_fmt).primary_ext
                        )
                    last = eng.convert(cur_path, out_path, cur_fmt, out_fmt, opts)
                    warnings.extend(last.warnings)
                    cur_path, cur_fmt = last.output, out_fmt
                if last is not None:
                    last.warnings = warnings
                    return _finish(
                        src,
                        last,
                        src_fmt,
                        dst_fmt,
                        engines_used,
                        hops,
                        t0,
                        opts,
                        progress,
                    )
        except ConversionError as ex:
            errors.append(f"[경유:{hops}] {ex}")
        except Exception as ex:
            errors.append(f"[경유:{hops}] {type(ex).__name__}: {ex}")

    msg = _no_route_message(src_fmt, dst_fmt, opts, errors)
    return ConvertResult(
        source=src,
        output=None,
        src_format=src_fmt,
        dst_format=dst_fmt,
        ok=False,
        error=msg,
        elapsed_sec=time.perf_counter() - t0,
    )


def _finish(
    src: Path,
    res: EngineResult,
    src_fmt: str,
    dst_fmt: str,
    engine_name: str,
    via: Optional[str],
    t0: float,
    opts: ConvertOptions,
    progress: Optional[ProgressFn],
) -> ConvertResult:
    warnings = list(res.warnings)
    if opts.verify_output:
        warnings.extend(_verify(res.output, dst_fmt))
    out = ConvertResult(
        source=src,
        output=res.output,
        src_format=src_fmt,
        dst_format=dst_fmt,
        engine=engine_name,
        via=via,
        elapsed_sec=time.perf_counter() - t0,
        warnings=warnings,
        extra_outputs=list(res.extra_outputs),
    )
    _say(progress, "  " + out.summary())
    return out


def _verify(path: Optional[Path], dst_fmt: str) -> list[str]:
    """결과물이 실제로 열리는지 가볍게 확인한다.

    변환이 '성공'했다고 보고했는데 파일이 0바이트이거나 컨테이너가 깨진 경우를
    잡아낸다. 전체 파싱은 비용이 크므로 헤더 수준만 본다.
    """
    if path is None or not path.is_file():
        return ["결과 파일이 생성되지 않았습니다."]
    size = path.stat().st_size
    if size == 0:
        return ["결과 파일이 비어 있습니다."]

    try:
        detected = formats.sniff(path)
    except Exception:
        return []
    if detected is None:
        return []
    if detected != dst_fmt and not _compatible(detected, dst_fmt):
        return [f"결과 파일 형식이 예상과 다릅니다(감지: {detected})."]
    return []


def _compatible(detected: str, expected: str) -> bool:
    groups = ({"xlsx", "xlsm"}, {"docx"}, {"html", "txt", "md", "csv"})
    return any(detected in g and expected in g for g in groups)


def _no_route_message(
    src_fmt: str, dst_fmt: str, opts: ConvertOptions, errors: list[str]
) -> str:
    from .engines.libreoffice import LibreOfficeEngine, install_hint

    lines = [f"{src_fmt.upper()} → {dst_fmt.upper()} 변환에 실패했습니다."]
    if errors:
        lines.append("시도한 방법:")
        lines.extend("  " + e for e in errors[:4])
    if dst_fmt == "hwp":
        lines.append(
            "참고: .hwp(구형) 저장은 지원하지 않습니다. .hwpx로 저장한 뒤 "
            "한/글에서 열어 .hwp로 다시 저장하세요."
        )
    elif (
        not LibreOfficeEngine().is_available()
        and src_fmt in ("doc",)
        or dst_fmt in ("doc", "xls")
    ):
        lines.append(install_hint())
    return "\n".join(lines)


def _say(progress: Optional[ProgressFn], msg: str) -> None:
    if progress is not None:
        progress(msg)


# --------------------------------------------------------------------------
# 일괄 변환
# --------------------------------------------------------------------------


def convert_many(
    sources: Sequence[Path | str],
    dst_fmt: str,
    *,
    out_dir: Optional[Path | str] = None,
    opts: Optional[ConvertOptions] = None,
    progress: Optional[ProgressFn] = None,
    on_result: Optional[Callable[[ConvertResult], None]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
) -> list[ConvertResult]:
    """여러 파일을 같은 목적 포맷으로 변환한다.

    한 파일이 실패해도 나머지를 계속 처리한다 - 20개를 걸어 두고 3번째에서
    멈추면 곤란하기 때문이다.
    """
    opts = opts or ConvertOptions()
    results: list[ConvertResult] = []
    for i, s in enumerate(sources, start=1):
        if should_stop is not None and should_stop():
            _say(progress, "사용자 요청으로 중단했습니다.")
            break
        _say(progress, f"[{i}/{len(sources)}] {Path(s).name}")
        r = convert(s, dst_fmt, out_dir=out_dir, opts=opts, progress=progress)
        results.append(r)
        if on_result is not None:
            on_result(r)
    return results


def diagnose() -> str:
    """현재 환경 진단. CLI의 --doctor와 GUI 정보 창에서 쓴다."""
    from .engines.native import native_readable, native_writable
    from .util import fonts

    lines = ["[엔진]"]
    for e in engines():
        lines.append("  " + e.status())
    lines.append("")
    lines.append("[네이티브 지원]")
    lines.append("  읽기: " + ", ".join(native_readable()))
    lines.append("  쓰기: " + ", ".join(native_writable()))
    lines.append("")
    lines.append("[글꼴]")
    lines.append("  " + fonts.describe())
    return "\n".join(lines)
