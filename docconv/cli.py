"""명령줄 인터페이스.

    docconv 파일1.hwpx 파일2.pdf --to docx
    docconv *.hwp --to pdf --out-dir 변환결과
    docconv 보고서.pdf --to xlsx --tables-only
    docconv --doctor
    docconv --gui

GUI 없이도 전 기능을 쓸 수 있게 하는 것이 목표다. 배치 스크립트나 작업
스케줄러에 걸어 쓰기 좋다.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable, Optional, Sequence

from . import formats, pipeline
from .options import (
    ConvertOptions,
    MergedCellPolicy,
    OnConflict,
    SheetLayout,
    TableFlatten,
)

PROG = "docconv"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog=PROG,
        description="로컬 문서 변환기 - PDF/DOCX/HWP/HWPX/DOC/CSV/XLSX/XLS 상호 변환",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=_epilog(),
    )
    p.add_argument(
        "inputs", nargs="*", type=str, help="변환할 파일 (여러 개 가능, 폴더도 가능)"
    )
    p.add_argument(
        "-t",
        "--to",
        dest="to",
        metavar="형식",
        help="목적 형식: " + ", ".join(formats.writable_formats()),
    )
    p.add_argument(
        "-o", "--out-dir", metavar="폴더", help="출력 폴더 (기본: 원본과 같은 폴더)"
    )
    p.add_argument("--out", metavar="파일", help="출력 파일 경로 (입력이 하나일 때만)")

    g = p.add_argument_group("동작")
    g.add_argument(
        "--on-conflict",
        choices=[e.value for e in OnConflict],
        default="rename",
        help="같은 이름 파일이 있을 때 (기본: rename)",
    )
    g.add_argument(
        "--no-format", action="store_true", help="서식을 버리고 텍스트만 옮긴다"
    )
    g.add_argument("--no-images", action="store_true", help="이미지를 포함하지 않는다")
    g.add_argument(
        "--no-verify", action="store_true", help="변환 후 결과 검사를 생략한다"
    )
    g.add_argument(
        "-r", "--recursive", action="store_true", help="폴더를 하위까지 훑는다"
    )
    g.add_argument(
        "-q", "--quiet", action="store_true", help="진행 상황을 출력하지 않는다"
    )

    t = p.add_argument_group("표 / 스프레드시트")
    t.add_argument(
        "--merged",
        choices=[e.value for e in MergedCellPolicy],
        default="placeholder",
        help="병합 셀 처리 (기본: placeholder)",
    )
    t.add_argument(
        "--tables-only", action="store_true", help="표만 추출한다(문단 무시)"
    )
    t.add_argument(
        "--single-sheet", action="store_true", help="여러 표를 한 시트로 합친다"
    )
    t.add_argument(
        "--csv-encoding",
        default="utf-8-sig",
        help="CSV 출력 인코딩 (기본: utf-8-sig, Excel 호환)",
    )
    t.add_argument("--csv-delimiter", default=",", help="CSV 구분자 (기본: ,)")
    t.add_argument("--csv-in-encoding", help="CSV 입력 인코딩 강제 (기본: 자동 판별)")

    d = p.add_argument_group("문서")
    d.add_argument("--font", metavar="파일", help="PDF에 쓸 글꼴 파일 경로")
    d.add_argument(
        "--base-size", type=float, default=10.5, help="PDF 기본 글자 크기(pt)"
    )
    d.add_argument("--password", help="암호로 보호된 입력 파일의 비밀번호")
    d.add_argument("--pages", metavar="N-M", help="PDF에서 읽을 쪽 범위 (예: 1-10)")

    e = p.add_argument_group("엔진")
    e.add_argument("--engine", help="사용할 엔진 강제 (native/libreoffice/msoffice)")
    e.add_argument("--native-only", action="store_true", help="외부 엔진을 쓰지 않는다")
    e.add_argument(
        "--timeout", type=float, default=180.0, help="외부 엔진 제한 시간(초)"
    )

    m = p.add_argument_group("기타")
    m.add_argument("--doctor", action="store_true", help="환경 진단 후 종료")
    m.add_argument(
        "--list-formats", action="store_true", help="지원 형식 표 출력 후 종료"
    )
    m.add_argument(
        "--targets", metavar="파일", help="이 파일에서 만들 수 있는 형식 목록"
    )
    m.add_argument("--gui", action="store_true", help="그래픽 화면을 연다")
    return p


def _epilog() -> str:
    return (
        "예시:\n"
        f"  {PROG} 기획서.hwpx --to pdf\n"
        f"  {PROG} *.hwp --to docx --out-dir 결과\n"
        f"  {PROG} 보고서.pdf --to xlsx --tables-only\n"
        f"  {PROG} 자료.xlsx --to csv --merged duplicate\n"
        f"  {PROG} --doctor\n"
        "\n"
        "참고: .hwp(구형) 저장은 지원하지 않습니다. .hwpx로 저장한 뒤\n"
        "      한/글에서 열어 .hwp로 다시 저장하세요.\n"
    )


# --------------------------------------------------------------------------


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.gui:
        return _launch_gui()
    if args.doctor:
        print(pipeline.diagnose())
        return 0
    if args.list_formats:
        print(format_table())
        return 0
    if args.targets:
        return _show_targets(Path(args.targets))

    if not args.inputs:
        parser.print_help()
        return 1
    if not args.to:
        print("오류: --to 로 목적 형식을 지정하세요.", file=sys.stderr)
        print("      " + ", ".join(formats.writable_formats()), file=sys.stderr)
        return 2

    dst = args.to.lower().lstrip(".")
    if dst not in formats.FORMATS:
        print(f"오류: 알 수 없는 형식 '{dst}'", file=sys.stderr)
        return 2
    if not formats.spec(dst).can_write:
        note = formats.spec(dst).note
        print(f"오류: '{dst}' 형식으로는 저장할 수 없습니다. {note}", file=sys.stderr)
        return 2

    files = list(_collect(args.inputs, args.recursive))
    if not files:
        print("오류: 변환할 파일을 찾지 못했습니다.", file=sys.stderr)
        return 1

    opts = _make_options(args)
    out_dir = Path(args.out_dir) if args.out_dir else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    progress = None if args.quiet else (lambda m: print(m, flush=True))

    if args.out and len(files) == 1:
        results = [
            pipeline.convert(
                files[0], dst, out_path=Path(args.out), opts=opts, progress=progress
            )
        ]
    else:
        if args.out and len(files) > 1:
            print(
                "경고: --out 은 입력이 하나일 때만 쓸 수 있어 무시합니다.",
                file=sys.stderr,
            )
        results = pipeline.convert_many(
            files, dst, out_dir=out_dir, opts=opts, progress=progress
        )

    return _report(results, quiet=args.quiet)


def _make_options(args) -> ConvertOptions:
    page_range = None
    if args.pages:
        page_range = _parse_pages(args.pages)

    return ConvertOptions(
        on_conflict=OnConflict(args.on_conflict),
        keep_formatting=not args.no_format,
        include_images=not args.no_images,
        verify_output=not args.no_verify,
        csv_encoding=args.csv_encoding,
        csv_delimiter=_unescape_delim(args.csv_delimiter),
        csv_input_encoding=args.csv_in_encoding,
        merged_cells=MergedCellPolicy(args.merged),
        table_flatten=TableFlatten.TABLES_ONLY
        if args.tables_only
        else TableFlatten.ALL,
        sheet_layout=SheetLayout.SINGLE if args.single_sheet else SheetLayout.PER_SHEET,
        pdf_font_path=args.font,
        pdf_base_size_pt=args.base_size,
        password=args.password,
        page_range=page_range,
        force_engine=args.engine,
        allow_external_engines=not args.native_only,
        engine_timeout=args.timeout,
    )


def _unescape_delim(s: str) -> str:
    return {"\\t": "\t", "tab": "\t", "\\;": ";"}.get(s, s)


def _parse_pages(s: str) -> Optional[tuple[int, int]]:
    """'1-10' / '5' / '3-' 를 0-base [start, end) 로."""
    s = s.strip()
    try:
        if "-" in s:
            a, _, b = s.partition("-")
            start = int(a) - 1 if a.strip() else 0
            end = int(b) if b.strip() else 10**9
            return (max(0, start), end)
        n = int(s)
        return (n - 1, n)
    except ValueError:
        print(f"경고: 쪽 범위를 이해하지 못해 무시합니다: {s}", file=sys.stderr)
        return None


def _collect(inputs: Iterable[str], recursive: bool) -> Iterable[Path]:
    """인자를 실제 파일 목록으로. 폴더와 glob 패턴을 지원한다."""
    seen: set[Path] = set()
    for raw in inputs:
        p = Path(raw)
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            for f in sorted(it):
                if f.is_file() and _known(f) and f not in seen:
                    seen.add(f)
                    yield f
        elif p.is_file():
            if p not in seen:
                seen.add(p)
                yield p
        else:
            # 셸이 확장하지 않은 glob 패턴을 직접 처리한다(Windows cmd 등).
            parent = p.parent if str(p.parent) else Path(".")
            try:
                matches = sorted(parent.glob(p.name))
            except (OSError, ValueError):
                matches = []
            if not matches:
                print(f"경고: 찾을 수 없습니다: {raw}", file=sys.stderr)
            for f in matches:
                if f.is_file() and f not in seen:
                    seen.add(f)
                    yield f


def _known(p: Path) -> bool:
    return formats.format_of_ext(p.suffix) is not None


def _report(results: list, *, quiet: bool) -> int:
    ok = [r for r in results if r.ok]
    bad = [r for r in results if not r.ok]
    warned = [r for r in ok if r.warnings]

    if not quiet:
        print()
        print("-" * 60)
    print(f"성공 {len(ok)}건 / 실패 {len(bad)}건", end="")
    print(f" / 경고 {len(warned)}건" if warned else "")

    for r in warned:
        for w in r.warnings:
            print(f"  경고  {r.source.name}: {w}")
    for r in bad:
        print(f"  실패  {r.source.name}", file=sys.stderr)
        for line in (r.error or "").split("\n"):
            print(f"        {line}", file=sys.stderr)
    return 0 if not bad else 1


def _show_targets(path: Path) -> int:
    if not path.is_file():
        print(f"오류: 파일이 없습니다: {path}", file=sys.stderr)
        return 1
    try:
        src = formats.detect(path)
    except Exception as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1
    targets = pipeline.supported_targets(src)
    print(f"{path.name} ({formats.spec(src).label})")
    print("  변환 가능:", ", ".join(targets) if targets else "(없음)")
    return 0


def format_table() -> str:
    rows = ["형식      확장자          읽기  쓰기  비고", "-" * 74]
    for key in list(formats.CORE_FORMATS) + [
        k for k in formats.FORMATS if k not in formats.CORE_FORMATS
    ]:
        s = formats.FORMATS[key]
        rows.append(
            f"{s.key:<9s} {', '.join(s.ext):<15s} "
            f"{'O' if s.can_read else '-':^5s} {'O' if s.can_write else '-':^5s} {s.note}"
        )
    return "\n".join(rows)


def _launch_gui() -> int:
    try:
        from .gui.app import run
    except ImportError as e:
        print(f"GUI를 열 수 없습니다: {e}", file=sys.stderr)
        print("tkinter가 설치되어 있는지 확인하세요.", file=sys.stderr)
        return 1
    return run()


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
