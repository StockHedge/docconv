"""시험용 표본 문서 생성.

테스트가 특정 개인의 파일에 의존하면 저장소를 내려받은 사람이 시험을 돌릴 수
없다. 그래서 필요한 표본을 **프로그램 자신이 만들어 쓴다.** 마침 HWPX Writer가
있으므로, 검증하고 싶은 요소(제목/본문/서식/표/병합/배경/테두리)를 골고루 담은
문서를 IR로 조립해 내보내면 된다.

표본이 갖춰야 할 것

* 여러 수준의 제목과 본문
* 굵게/기울임/밑줄/취소선/글자색/크기 변화
* 표 - 머리글, 가로 병합, 세로 병합, 셀 배경, 열 너비
* 문단 정렬(왼쪽/가운데/오른쪽/양쪽)
* 한글·한자·영문·숫자·특수문자가 섞인 텍스트
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docconv.ir import (  # noqa: E402
    Align,
    Border,
    BorderStyle,
    Cell,
    DocMeta,
    Document,
    ListKind,
    PageNumberSpec,
    Paragraph,
    Run,
    Table,
)
from docconv.options import ConvertOptions  # noqa: E402
from docconv.writers.hwpx_writer import write_hwpx  # noqa: E402

#: 변환 후에도 살아남아야 하는 표본 문자열.
#:
#: 한글/영문/숫자를 고루 섞되, **한글과 라틴·숫자의 경계를 걸치지 않게** 고른다.
#: 일부 변환기(특히 Word가 만드는 PDF)는 CJK와 라틴 사이에 자동 자간을 넣어
#: 별도 조각으로 분리하는데, 그러면 텍스트 추출 시 그 틈이 공백이 된다.
#: 예를 들어 "1,234,567원" 은 "1,234,567 원" 으로 나와 멀쩡한 변환이 실패로 잡힌다.
PROBES: tuple[str, ...] = (
    "문서 변환기",
    "DocumentConverter",
    "가나다라마바사",
    "1,234,567",
)


def _solid(color=(120, 120, 120), w=0.5) -> Border:
    return Border(style=BorderStyle.SOLID, width_pt=w, color=color)


def _no() -> Border:
    return Border()


def _grid() -> tuple[Border, Border, Border, Border]:
    return (_solid(), _solid(), _solid(), _solid())


def build_document() -> Document:
    """검증 요소를 고루 담은 표본 문서를 IR로 조립한다."""
    doc = Document(
        meta=DocMeta(
            title="문서 변환기 시험 표본",
            author="docconv",
            subject="변환 무결성 확인용",
            keywords="변환, 시험, 표본",
        )
    )
    doc.page_number = PageNumberSpec(position="BOTTOM_CENTER", side_char="-")

    B = doc.blocks
    B.append(Paragraph.of("문서 변환기 시험 표본", heading=1, align=Align.CENTER))
    B.append(
        Paragraph.of(
            "이 문서는 DocumentConverter 의 변환 결과를 확인하려고 만든 표본입니다. "
            "글자와 서식이 형식을 오가면서 어떻게 살아남는지 봅니다.",
        )
    )

    B.append(Paragraph.of("1. 글자 서식", heading=2))
    B.append(
        Paragraph(
            runs=[
                Run(text="보통 글자, "),
                Run(text="굵게", bold=True),
                Run(text=", "),
                Run(text="기울임", italic=True),
                Run(text=", "),
                Run(text="밑줄", underline=True),
                Run(text=", "),
                Run(text="취소선", strike=True),
                Run(text=", "),
                Run(text="붉은 글자", color=(200, 30, 30)),
                Run(text=", "),
                Run(text="큰 글자", size_pt=15.0),
                Run(text=" 가 한 문단에 섞여 있습니다."),
            ]
        )
    )
    B.append(
        Paragraph.of(
            "한글 가나다라마바사, 한자 漢字, 영문 DocumentConverter, "
            "숫자 1,234,567원, 기호 ·—「」※ 를 함께 씁니다."
        )
    )

    B.append(Paragraph.of("2. 문단 정렬", heading=2))
    for align, label in (
        (Align.LEFT, "왼쪽 정렬 문단입니다."),
        (Align.CENTER, "가운데 정렬 문단입니다."),
        (Align.RIGHT, "오른쪽 정렬 문단입니다."),
        (
            Align.JUSTIFY,
            "양쪽 정렬 문단입니다. 줄이 길어지면 좌우 끝이 가지런히 맞습니다.",
        ),
    ):
        B.append(Paragraph.of(label, align=align))

    B.append(Paragraph.of("3. 목록", heading=2))
    for i, t in enumerate(("첫째 항목", "둘째 항목", "셋째 항목"), start=1):
        p = Paragraph.of(t)
        p.list_kind = ListKind.BULLET
        B.append(p)

    B.append(Paragraph.of("4. 표", heading=2))
    B.append(_build_table())

    B.append(Paragraph.of("5. 긴 본문", heading=2))
    for i in range(1, 7):
        B.append(
            Paragraph.of(
                f"{i}번째 단락입니다. 쪽을 넘길 때 내용이 잘리지 않는지, "
                "표와 문단의 순서가 뒤바뀌지 않는지 확인하려고 여러 줄을 둡니다. "
                "가나다라마바사아자차카타파하 1,234,567원."
            )
        )
    return doc


def _build_table() -> Table:
    """머리글 + 가로 병합 + 세로 병합 + 배경 + 테두리를 모두 쓰는 표."""
    head_bg = (222, 234, 255)
    tint = (240, 240, 240)

    def cell(text: str, *, bg=None, span=1, rspan=1, borders=None) -> Cell:
        c = Cell.of(text)
        c.background = bg
        c.col_span = span
        c.row_span = rspan
        c.borders = borders if borders is not None else _grid()
        return c

    t = Table(header_row=True, col_widths_pt=[90.0, 150.0, 210.0])
    t.rows.append(
        [
            cell("항목", bg=head_bg),
            cell("값", bg=head_bg),
            cell("설명", bg=head_bg),
        ]
    )
    t.rows.append(
        [
            cell("가로 병합", span=2),
            Cell(merged_placeholder=True),
            cell("두 칸을 하나로 합친 행"),
        ]
    )
    t.rows.append(
        [
            cell("세로 병합", rspan=2, bg=tint),
            cell("첫째 줄"),
            cell("아래 행과 이어집니다"),
        ]
    )
    t.rows.append(
        [
            Cell(merged_placeholder=True),
            cell("둘째 줄"),
            cell("왼쪽 칸이 위와 합쳐져 있습니다"),
        ]
    )
    t.rows.append(
        [
            cell("테두리 없음", borders=(_no(), _no(), _no(), _no())),
            cell("좌우만", borders=(_solid(), _no(), _solid(), _no())),
            cell("칸마다 테두리가 다릅니다"),
        ]
    )
    t.rows.append([cell("숫자"), cell("1,234,567원"), cell("자릿점과 단위가 섞인 값")])
    return t


def write_sample_hwpx(path: Path) -> Path:
    """표본을 HWPX로 저장한다. 다른 형식은 여기서 변환해 만든다."""
    path.parent.mkdir(parents=True, exist_ok=True)
    write_hwpx(build_document(), path, ConvertOptions())
    return path


def main() -> int:
    out = ROOT / "output" / "시험표본.hwpx"
    write_sample_hwpx(out)
    print(f"생성: {out}  ({out.stat().st_size:,}B)")

    from docconv.readers import hwpx_reader

    doc = hwpx_reader.read(out)
    text = doc.to_text()
    print(f"되읽기: {len(text):,}자, {doc.stats()}")
    missing = [p for p in PROBES if p not in text]
    print("표본 문자열:", "전부 확인" if not missing else f"누락 {missing}")
    return 0 if not missing else 1


if __name__ == "__main__":
    sys.exit(main())
