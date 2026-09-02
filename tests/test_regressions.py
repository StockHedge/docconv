"""회귀 시험 — 한 번 겪은 결함이 다시 나오지 않는지 지킨다.

여기 있는 항목은 전부 **실제 문서에서 터졌던 것**이다. 각 시험은 그 결함의
근본 원인을 직접 겨냥한다. 새 결함을 고칠 때마다 여기에 한 줄 늘리는 것이
"같은 실수를 반복하지 않는다"의 실질적인 내용이다.

실행::

    python -m tests.test_regressions
"""

from __future__ import annotations

import io
import re
import shutil
import struct
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docconv import pipeline  # noqa: E402
from docconv.ir import (  # noqa: E402
    Border,
    BorderStyle,
    Cell,
    Document,
    Image,
    Paragraph,
    Table,
)
from docconv.options import ConvertOptions  # noqa: E402
from docconv.writers.hwpx_writer import write_hwpx  # noqa: E402

WORK = ROOT / "output" / "regressions"


class Failure(Exception):
    pass


def check(cond: bool, msg: str) -> None:
    if not cond:
        raise Failure(msg)


# --------------------------------------------------------------------------
# 1. 알 수 없는 서식 값이 서식을 켜면 안 된다
# --------------------------------------------------------------------------


def test_unknown_format_value_does_not_enable() -> str:
    """`shape="3D"` 같은 미지의 값을 취소선으로 오판하지 않는지.

    실제 사고: 어떤 한/글 저장본이 취소선 없음을 `shape="NONE"` 이 아니라
    `shape="3D"` 로 적어, "NONE이 아니면 적용" 판정 탓에 **본문 180개 조각
    전부에 취소선**이 그어졌다.
    """
    from docconv.readers import hwpx_reader

    src = WORK / "unknown_shape.hwpx"
    doc = Document(blocks=[Paragraph.of("취소선이 없어야 하는 문장입니다.")])
    write_hwpx(doc, src, ConvertOptions())

    # 저장된 HWPX의 strikeout/underline 값을 미지의 값으로 바꿔치기한다.
    patched = WORK / "unknown_shape_patched.hwpx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(patched, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "Contents/header.xml":
                text = data.decode("utf-8")
                text = re.sub(r'(<hh:strikeout shape=")[^"]*(")', r"\g<1>3D\g<2>", text)
                text = re.sub(
                    r'(<hh:underline type=")[^"]*(")', r"\g<1>WAVE\g<2>", text
                )
                data = text.encode("utf-8")
            zout.writestr(item, data)

    got = hwpx_reader.read(patched)
    runs = [r for p in got.paragraphs() for r in p.runs if r.text.strip()]
    check(bool(runs), "문단을 읽지 못했다")
    check(
        not any(r.strike for r in runs),
        f'shape="3D" 를 취소선으로 오판했다 ({sum(r.strike for r in runs)}개)',
    )
    check(
        not any(r.underline for r in runs),
        'type="WAVE" 를 밑줄로 오판했다',
    )
    return f"미지의 서식 값 무시 확인 (run {len(runs)}개)"


def test_unknown_border_type_draws_nothing() -> str:
    """알 수 없는 테두리 종류가 실선으로 둔갑하지 않는지."""
    from docconv.readers.hwpx_reader import _parse_border_fill
    from docconv.util.safety import parse_xml

    xml = (
        '<hh:borderFill xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" id="9">'
        '<hh:leftBorder type="완전히-모르는-값" width="0.4 mm" color="#000000"/>'
        '<hh:rightBorder type="NONE" width="0.1 mm" color="none"/>'
        '<hh:topBorder type="SOLID" width="0.4 mm" color="#000000"/>'
        '<hh:bottomBorder type="NONE" width="0.1 mm" color="none"/>'
        "</hh:borderFill>"
    )
    fill = _parse_border_fill(parse_xml(xml.encode("utf-8")))
    check(fill.borders is not None, "테두리를 읽지 못했다")
    left, top, right, bottom = fill.borders
    check(not left.visible, "모르는 종류를 실선으로 그렸다")
    check(top.visible, "SOLID 테두리를 놓쳤다")
    return "미지의 테두리 종류는 긋지 않음"


# --------------------------------------------------------------------------
# 2. 표가 쪽을 넘겨도 유령 배경이 남으면 안 된다
# --------------------------------------------------------------------------


def test_no_ghost_cell_background_on_later_pages() -> str:
    """여러 쪽에 걸친 표에서 셀 배경이 다음 쪽에 복사되지 않는지.

    실제 사고: 조판 엔진이 첫 쪽 셀 배경을 이후 모든 쪽에 같은 좌표로 다시
    그렸다. 높이 6pt짜리 띠가 본문 글줄을 가로질러 **취소선처럼 보였다**.
    """
    import pymupdf

    # 배경 있는 머리글 + 쪽을 넘길 만큼 긴 본문
    t = Table(header_row=True, col_widths_pt=[120.0, 320.0])
    t.rows.append(
        [
            Cell.of("항목", background=(217, 217, 217)),
            Cell.of("내용", background=(217, 217, 217)),
        ]
    )
    for i in range(40):
        t.rows.append(
            [
                Cell.of(f"{i + 1}행", background=(240, 240, 240)),
                Cell.of("쪽을 넘기도록 채우는 긴 본문입니다. " * 3),
            ]
        )
    doc = Document(blocks=[Paragraph.of("유령 배경 회귀 시험"), t])

    out = WORK / "ghost.pdf"
    from docconv.writers import pdf_writer

    # 저수준 폴백으로 떨어졌는지 감시한다. 폴백은 서식을 통째로 버리므로
    # "유령이 없다"는 결과가 나와도 사실은 표가 사라진 것일 수 있다.
    # 실제로 이 시험이 처음 통과했을 때가 그런 경우였다.
    fell_back = []
    real_fallback = pdf_writer._write_fallback

    def watch(*a, **k):
        fell_back.append(True)
        return real_fallback(*a, **k)

    pdf_writer._write_fallback = watch
    try:
        pdf_writer.write_pdf(doc, out, ConvertOptions())
    finally:
        pdf_writer._write_fallback = real_fallback

    check(not fell_back, "조판이 실패해 폴백으로 떨어졌다(서식이 버려졌다)")

    pdf = pymupdf.open(str(out))
    shapes = sum(len(pdf[i].get_drawings()) for i in range(pdf.page_count))
    check(shapes > 0, "표 테두리가 하나도 그려지지 않았다")
    check(pdf.page_count >= 2, f"표본이 여러 쪽이어야 한다 (현재 {pdf.page_count}쪽)")
    ghosts = 0
    for pno in range(1, pdf.page_count):
        for dr in pdf[pno].get_drawings():
            fill = dr.get("fill")
            if not fill:
                continue
            rgb = tuple(round(v, 2) for v in fill)
            if rgb in ((1.0, 1.0, 1.0), (0.0, 0.0, 0.0)):
                continue
            # 높이 2pt 미만은 표 테두리 선이다. 유령 배경만 센다.
            if 2.0 <= dr["rect"].height <= 8.0 and dr["rect"].width > 20:
                ghosts += 1
    pages = pdf.page_count
    pdf.close()
    check(ghosts == 0, f"유령 배경 띠 {ghosts}개가 남았다")
    return f"유령 배경 없음 ({pages}쪽 표)"


# --------------------------------------------------------------------------
# 3. HWP 표 셀 좌표 파싱
# --------------------------------------------------------------------------


def test_hwp_cell_header_offsets() -> str:
    """셀 LIST_HEADER의 문단 수가 INT32라는 전제를 지키는지.

    실제 사고: INT16으로 읽어 이후 필드가 2바이트씩 밀렸고, col과 row가
    뒤바뀌어 셀이 엉뚱한 자리에 놓였다. 7x4 표의 첫 행이 통째로 비었다.
    """
    from docconv.readers.hwp_reader import _parse_cell_header

    # 실제 파일에서 뜬 바이트: 문단수=1, 속성=0x20, col=1, row=0, span=1x1
    payload = bytes.fromhex("010000002000000001000000010001009b2d0000") + b"\x00" * 27
    col, row, cspan, rspan = _parse_cell_header(payload)
    check(col == 1, f"col 이 1이어야 하는데 {col} 이다")
    check(row == 0, f"row 가 0이어야 하는데 {row} 이다")
    check((cspan, rspan) == (1, 1), f"span 이 1x1이어야 하는데 {cspan}x{rspan} 이다")
    return "HWP 셀 좌표 오프셋 유지"


def test_hwp_picture_tag_constant() -> str:
    """그림 레코드 태그가 85인지. 83은 CURVE라 그림을 못 찾는다."""
    from docconv.readers.hwp_reader import HWPTAG_SHAPE_COMPONENT_PICTURE

    check(
        HWPTAG_SHAPE_COMPONENT_PICTURE == 85,
        f"그림 태그가 85여야 하는데 {HWPTAG_SHAPE_COMPONENT_PICTURE} 이다",
    )
    return "그림 레코드 태그 85 유지"


# --------------------------------------------------------------------------
# 4. IR 순회가 중첩을 빠뜨리면 안 된다
# --------------------------------------------------------------------------


def test_nested_content_is_counted() -> str:
    """표 셀 안의 이미지·문단이 통계와 순회에 잡히는지.

    실제 사고: 표 안의 그림이 `images()` 에 안 잡혀 "이미지 0개"로 보고됐다.
    한/글 양식 문서는 도장·서명을 표 칸에 넣는 경우가 흔하다.
    """
    inner = Table(
        rows=[[Cell(blocks=[Paragraph.of("셀 안 문단"), Image(data=b"x", fmt="png")])]]
    )
    doc = Document(blocks=[Paragraph.of("바깥 문단"), inner])
    st = doc.stats()
    check(st["images"] == 1, f"중첩 이미지를 놓쳤다 (images={st['images']})")
    check(st["paragraphs"] == 2, f"중첩 문단을 놓쳤다 (paragraphs={st['paragraphs']})")
    check(len(list(doc.images())) == 1, "images() 가 중첩을 훑지 않는다")
    return "중첩 내용 집계 확인"


# --------------------------------------------------------------------------
# 5. PDF 글꼴 서브셋
# --------------------------------------------------------------------------


def test_pdf_font_is_subset() -> str:
    """한글 PDF에 글꼴 전체가 임베딩되지 않는지.

    실제 사고: 서브셋을 빠뜨려 26MB, 서브셋 뒤에 쪽 번호를 그려 7.6MB가 됐다.
    """
    import pymupdf

    from docconv.writers.pdf_writer import write_pdf

    doc = Document(
        blocks=[Paragraph.of("한글 글꼴 서브셋 확인 문장입니다. 가나다라마바사.")]
    )
    from docconv.ir import PageNumberSpec

    doc.page_number = PageNumberSpec()  # 쪽 번호도 켜서 순서 문제까지 잡는다
    out = WORK / "subset.pdf"
    write_pdf(doc, out, ConvertOptions())

    size_mb = out.stat().st_size / 1024**2
    check(size_mb < 2.0, f"PDF가 {size_mb:.1f}MB — 글꼴이 통째로 들어갔다")

    pdf = pymupdf.open(str(out))
    names = {
        f[3] for pno in range(pdf.page_count) for f in pdf[pno].get_fonts(full=True)
    }
    pdf.close()
    # 서브셋된 글꼴은 "ABCDEF+이름" 형태의 접두사를 갖는다.
    check(
        all("+" in n for n in names) if names else True,
        f"서브셋되지 않은 글꼴이 있다: {names}",
    )
    return f"글꼴 서브셋 확인 ({out.stat().st_size:,}B)"


# --------------------------------------------------------------------------


TESTS = (
    test_unknown_format_value_does_not_enable,
    test_unknown_border_type_draws_nothing,
    test_no_ghost_cell_background_on_later_pages,
    test_hwp_cell_header_offsets,
    test_hwp_picture_tag_constant,
    test_nested_content_is_counted,
    test_pdf_font_is_subset,
)


def main() -> int:
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    print("회귀 시험 — 한 번 겪은 결함이 다시 나오는지 확인")
    print("=" * 66)
    failed = 0
    for fn in TESTS:
        name = fn.__name__.replace("test_", "")
        try:
            note = fn()
        except Failure as e:
            failed += 1
            print(f"  실패  {name}\n          {e}")
        except Exception as e:  # noqa: BLE001
            failed += 1
            print(f"  오류  {name}\n          {type(e).__name__}: {e}")
        else:
            print(f"  통과  {name:<44s} {note}")
    print("=" * 66)
    print(
        f"{len(TESTS) - failed}/{len(TESTS)} 통과"
        + ("" if not failed else f" — 실패 {failed}건")
    )
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
