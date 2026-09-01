"""8종 전 조합(56가지) 전수 변환 테스트.

사용자가 실제로 쓰는 형식 8종만 대상으로, 가능한 모든 방향을 하나씩 시도한다.

    pdf, docx, hwp, hwpx, doc, csv, xlsx, xls   ->  8 x 7 = 56 조합

각 칸은 다음 중 하나로 판정한다.

    OK    변환 성공 + 표본 문자열이 살아남음
    부분   변환은 됐지만 표본 일부가 유실 (형식 특성상 예상되는 손실 포함)
    불가   원리적으로 불가능 (지원하지 않음)
    실패   가능해야 하는데 실패함  <- 이것만 문제다

실행::

    python -m tests.test_all_pairs
"""

from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docconv import formats, pipeline  # noqa: E402
from docconv.options import ConvertOptions  # noqa: E402

from tests.sample import PROBES, write_sample_hwpx  # noqa: E402

WORK = ROOT / "output" / "pairs"

SRC_ENV = "DOCCONV_TEST_SOURCE"


def resolve_source() -> Path:
    r"""시험에 쓸 원본을 정한다.

    기본은 **프로그램이 직접 만든 표본**이다. 특정 개인의 파일에 기대면 저장소를
    내려받은 사람이 시험을 돌릴 수 없기 때문이다. 손에 있는 실제 문서로 확인하고
    싶으면 환경 변수로 지정한다::

        set DOCCONV_TEST_SOURCE=C:\경로\내문서.hwpx    (Windows)
        export DOCCONV_TEST_SOURCE=~/문서/내문서.hwpx      (macOS/Linux)
    """
    import os

    override = os.environ.get(SRC_ENV, "").strip()
    if override:
        p = Path(override).expanduser()
        if p.is_file():
            return p
        print(f"경고: {SRC_ENV} 가 가리키는 파일이 없어 기본 표본을 씁니다: {p}")
    return write_sample_hwpx(ROOT / "output" / "시험표본.hwpx")


#: 사용자가 지목한 8종
KINDS = ("pdf", "docx", "hwp", "hwpx", "doc", "csv", "xlsx", "xls")

def read_text(path: Path) -> str:
    import tempfile

    from docconv.engines.native import readers
    from docconv.readers.sheet_reader import detect_encoding

    try:
        fmt = formats.detect(path)
    except Exception:
        return ""
    fn = readers().get(fmt)
    if fn is not None:
        try:
            return fn(path).to_text()
        except Exception:
            return ""
    try:
        with tempfile.TemporaryDirectory(prefix="dc-probe-") as tmp:
            t = Path(tmp) / (path.stem + ".txt")
            r = pipeline.convert(path, "txt", out_path=t, opts=ConvertOptions())
            if r.ok and r.output and r.output.is_file():
                raw = r.output.read_bytes()
                return raw.decode(detect_encoding(raw), errors="replace")
    except Exception:
        pass
    return ""


def hits(text: str) -> int:
    return sum(1 for p in PROBES if p in text)


def main() -> int:
    source = resolve_source()
    print(f"원본: {source}  ({source.stat().st_size:,}B)")
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    print("=" * 76)
    print(pipeline.diagnose())
    print("=" * 76)

    opts = ConvertOptions()

    # --- 준비: 원본 HWPX에서 8종의 '입력 표본'을 만든다 -------------------
    print("\n[준비] 각 형식의 입력 표본 생성")
    seeds: dict[str, Path] = {"hwpx": source}

    # .hwp 는 이 프로그램이 쓰지 못하므로(한컴오피스 없이는 불가) 명세대로
    # 조립한 표본을 쓴다. 읽기 경로 검증이 목적이다.
    try:
        from tests.make_hwp_sample import SAMPLE_PARAGRAPHS, write_hwp

        hwp_path = WORK / "seed" / "표본.hwp"
        hwp_path.parent.mkdir(parents=True, exist_ok=True)
        write_hwp(hwp_path, SAMPLE_PARAGRAPHS)
        seeds["hwp"] = hwp_path
        print(f"  hwp   OK    {hwp_path.stat().st_size:>9,d}B  [합성 표본]")
    except Exception as e:
        print(f"  hwp   표본 생성 실패({type(e).__name__}) - hwp 입력 행은 건너뜁니다")

    for k in KINDS:
        if k in seeds:
            continue
        if not formats.spec(k).can_write:
            print(
                f"  {k:<5s} 생성 불가 (쓰기 미지원) - 이 형식을 입력으로 쓰는 행은 건너뜁니다"
            )
            continue
        r = pipeline.convert(source, k, out_dir=WORK / "seed", opts=opts)
        if r.ok and r.output:
            seeds[k] = r.output
            print(f"  {k:<5s} OK  {r.output.stat().st_size:>9,d}B  [{r.engine}]")
        else:
            print(f"  {k:<5s} 실패: {(r.error or '').splitlines()[0][:50]}")

    # --- 본 시험: 56 조합 --------------------------------------------------
    print(
        f"\n[본시험] {len(KINDS)}x{len(KINDS) - 1} = {len(KINDS) * (len(KINDS) - 1)} 조합\n"
    )
    header = "src \\ dst │" + "".join(f"{d:>7s}" for d in KINDS)
    print(header)
    print("─" * len(header))

    grid: dict[tuple[str, str], str] = {}
    detail: list[str] = []

    for src in KINDS:
        row = [f"{src:<9s} │"]
        for dst in KINDS:
            if src == dst:
                row.append(f"{'―':>7s}")
                continue
            if src not in seeds:
                grid[(src, dst)] = "표본없음"
                row.append(f"{'-':>7s}")
                continue
            if not formats.spec(dst).can_write:
                grid[(src, dst)] = "불가"
                row.append(f"{'불가':>6s}")
                continue

            t0 = time.perf_counter()
            r = pipeline.convert(seeds[src], dst, out_dir=WORK / src, opts=opts)
            dt = time.perf_counter() - t0

            if not r.ok:
                grid[(src, dst)] = "실패"
                row.append(f"{'실패':>6s}")
                detail.append(
                    f"  실패  {src}→{dst}: {(r.error or '').splitlines()[0][:60]}"
                )
                continue

            # CSV는 한 파일에 시트 하나만 담기므로 여러 시트가 나오면 파일이
            # 갈라진다. 갈라진 파일까지 합쳐 봐야 실제 손실 여부를 알 수 있다.
            merged = read_text(r.output)
            for extra in r.extra_outputs:
                merged += "\n" + read_text(extra)
            n = hits(merged)
            if n == len(PROBES):
                grid[(src, dst)] = "OK"
                row.append(f"{'OK':>7s}")
            else:
                grid[(src, dst)] = f"부분{n}/{len(PROBES)}"
                row.append(f"{n}/{len(PROBES):<5d}"[:7].rjust(7))
                detail.append(
                    f"  부분  {src}→{dst}: 표본 {n}/{len(PROBES)} ({dt:.1f}s)"
                )
        print("".join(row))

    # --- 요약 -------------------------------------------------------------
    print()
    ok = sum(1 for v in grid.values() if v == "OK")
    part = sum(1 for v in grid.values() if v.startswith("부분"))
    imp = sum(1 for v in grid.values() if v == "불가")
    fail = sum(1 for v in grid.values() if v == "실패")
    noseed = sum(1 for v in grid.values() if v == "표본없음")

    print("=" * 76)
    print(
        f"OK {ok} / 부분 {part} / 원리적 불가 {imp} / 실패 {fail} / 표본없음 {noseed}"
    )
    if detail:
        print("\n[상세]")
        for d in detail:
            print(d)

    print("\n[원리적으로 불가능한 칸]")
    for (s, d), v in grid.items():
        if v == "불가":
            print(f"  {s} → {d}: {formats.spec(d).note or '쓰기 미지원'}")

    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
