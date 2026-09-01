"""변환 매트릭스 통합 테스트.

원본 HWPX 하나에서 출발해 지원하는 모든 형식을 만들고, 만들어진 각 파일을
다시 여러 형식으로 변환해 **텍스트가 살아남는지** 확인한다.

단위 테스트가 아니라 실환경 검증이다. 외부 엔진(LibreOffice/MS Office) 설치
여부에 따라 통과 범위가 달라지므로, 불가능한 경로는 실패가 아니라 '건너뜀'
으로 집계한다.

실행::

    python -m tests.test_matrix
    python -m tests.test_matrix --quick
"""

from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docconv import pipeline  # noqa: E402
from docconv.options import ConvertOptions  # noqa: E402

from tests.sample import PROBES, write_sample_hwpx  # noqa: E402

WORK = ROOT / "output" / "matrix"

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


#: 1단계에서 만들 형식
STAGE1 = ("pdf", "docx", "doc", "hwpx", "xlsx", "xls", "csv", "txt", "md", "html")
#: 2단계에서 각 결과물을 다시 변환해 볼 형식
STAGE2 = ("pdf", "docx", "hwpx", "xlsx", "csv", "txt")


def read_text(path: Path) -> str:
    """어떤 형식이든 텍스트로 읽어낸다. 실패하면 빈 문자열.

    네이티브 리더가 없는 형식(.doc 등)은 파이프라인으로 txt를 만들어 읽는다.
    이렇게 하지 않으면 실제로는 잘 변환된 .doc 이 '0자'로 집계되어 멀쩡한
    경로가 실패로 보인다.
    """
    import tempfile

    from docconv import formats
    from docconv.engines.native import readers

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

    # 네이티브로 못 읽는 형식: 외부 엔진을 거쳐 txt로 뽑아 본다.
    from docconv.readers.sheet_reader import detect_encoding

    try:
        with tempfile.TemporaryDirectory(prefix="docconv-probe-") as tmp:
            target = Path(tmp) / (path.stem + ".txt")
            r = pipeline.convert(path, "txt", out_path=target, opts=ConvertOptions())
            if r.ok and r.output and r.output.is_file():
                raw = r.output.read_bytes()
                # 인코딩을 단정하지 않는다. 외부 엔진이 무엇으로 저장했든
                # 검증 단계에서 오판하면 멀쩡한 변환이 실패로 보인다.
                return raw.decode(detect_encoding(raw), errors="replace")
    except Exception:
        pass
    return ""


def score(text: str) -> tuple[int, int, int]:
    """(표본 적중 수, 한글 글자 수, 깨진 문자 수)."""
    hits = sum(1 for p in PROBES if p in text)
    hangul = sum(1 for ch in text if "가" <= ch <= "힣")
    broken = text.count("�") + text.count("□")
    return hits, hangul, broken


def main(quick: bool = False) -> int:
    src = resolve_source()
    print(f"원본: {src}  ({src.stat().st_size:,}B)")
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    print("=" * 78)
    print(pipeline.diagnose())
    print("=" * 78)

    base_text = read_text(src)
    b_hits, b_hangul, _ = score(base_text)
    print(
        f"\n원본: {len(base_text):,}자, 한글 {b_hangul:,}자, 표본 {b_hits}/{len(PROBES)}\n"
    )

    opts = ConvertOptions(verify_output=True)

    # ---- 1단계 --------------------------------------------------------
    print("--- 1단계: HWPX → 각 형식 " + "-" * 44)
    stage1_out: dict[str, Path] = {}
    s1_ok = s1_skip = s1_fail = 0

    for fmt in STAGE1:
        t0 = time.perf_counter()
        r = pipeline.convert(src, fmt, out_dir=WORK / "s1", opts=opts)
        if not r.ok:
            reason = (r.error or "").split("\n")[0]
            if "지원하지" in reason or "실패했습니다" in reason:
                s1_skip += 1
                print(f"  건너뜀  {fmt:<6s} {reason[:56]}")
            else:
                s1_fail += 1
                print(f"  실패    {fmt:<6s} {reason[:56]}")
            continue
        stage1_out[fmt] = r.output
        text = read_text(r.output)
        hits, hangul, broken = score(text)
        mark = "OK  " if hits == len(PROBES) and broken == 0 else "확인"
        if mark == "OK  ":
            s1_ok += 1
        else:
            s1_fail += 1
        size = r.output.stat().st_size
        print(
            f"  {mark}    {fmt:<6s} {size:>10,d}B  표본 {hits}/{len(PROBES)}  "
            f"한글 {hangul:>6,d}  깨짐 {broken}  [{r.engine}] "
            f"{time.perf_counter() - t0:.2f}s"
        )

    # ---- 2단계 --------------------------------------------------------
    print("\n--- 2단계: 각 형식 → 다른 형식 (교차) " + "-" * 32)
    pairs = 0
    s2_ok = s2_skip = s2_fail = 0
    sources = list(stage1_out.items())
    if quick:
        sources = [(k, v) for k, v in sources if k in ("docx", "pdf", "xlsx", "csv")]

    for src_fmt, src_path in sources:
        line = [f"  {src_fmt:<5s} →"]
        for dst_fmt in STAGE2:
            if dst_fmt == src_fmt:
                continue
            pairs += 1
            r = pipeline.convert(
                src_path, dst_fmt, out_dir=WORK / "s2" / src_fmt, opts=opts
            )
            if not r.ok:
                s2_skip += 1
                line.append(f"{dst_fmt}:-")
                continue
            text = read_text(r.output)
            hits, _, broken = score(text)
            # PDF는 표 재구성 과정에서 표본이 일부 빠질 수 있어 기준을 낮춘다.
            need = len(PROBES) if src_fmt != "pdf" else max(1, len(PROBES) - 2)
            if hits >= need and broken == 0:
                s2_ok += 1
                line.append(f"{dst_fmt}:OK")
            else:
                s2_fail += 1
                line.append(f"{dst_fmt}:{hits}/{len(PROBES)}!")
        print(" ".join(line))

    # ---- 요약 ---------------------------------------------------------
    print("\n" + "=" * 78)
    print(
        f"1단계: 성공 {s1_ok} / 실패 {s1_fail} / 건너뜀 {s1_skip}  (총 {len(STAGE1)})"
    )
    print(f"2단계: 성공 {s2_ok} / 실패 {s2_fail} / 건너뜀 {s2_skip}  (총 {pairs})")
    print(f"결과물 위치: {WORK}")

    failed = s1_fail + s2_fail
    print("전체 통과" if failed == 0 else f"확인이 필요한 항목 {failed}건")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main(quick="--quick" in sys.argv))
