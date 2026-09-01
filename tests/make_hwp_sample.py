"""HWP 5.0 표본 파일 생성기 (테스트 전용).

한컴오피스가 없는 환경에서 `.hwp` **읽기** 경로를 검증하려면 표본이 필요한데,
이 프로그램은 HWP 쓰기를 지원하지 않는다(그럴 만한 이유가 있다 - README 참고).
그래서 파서를 검증할 최소한의 파일을 명세대로 직접 조립한다.

    FileHeader          256바이트 평문. 시그니처 + 버전 + 플래그
    DocInfo             (선택) 서식 테이블
    BodyText/Section0   zlib(raw deflate) 압축된 레코드 스트림

OLE 복합 문서 컨테이너는 Windows COM의 `StgCreateDocfile` 로 만든다.
olefile은 읽기 전용이라 쓸 수 없다.

**한계**: 여기서 만든 파일은 명세를 따르지만 한/글이 실제로 저장한 파일과
같지는 않다. 레코드 순회·압축 해제·텍스트 디코딩 같은 파서의 뼈대는 검증되지만,
실제 문서의 다양한 컨트롤과 서식까지 확인해 주지는 못한다.
"""

from __future__ import annotations

import struct
import sys
import zlib
from pathlib import Path

# HWP 레코드 태그
HWPTAG_BEGIN = 0x010
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51
HWPTAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52

SIGNATURE = b"HWP Document File"


def record(tag: int, level: int, payload: bytes) -> bytes:
    """32비트 헤더(tag 10 | level 10 | size 12) + 페이로드.

    size가 12비트에 안 들어가면 0xFFF 표식을 두고 다음 4바이트에 실제 크기를
    쓴다. 이 확장 규칙을 빠뜨리면 스트림 전체가 어긋난다.
    """
    size = len(payload)
    if size < 0xFFF:
        head = (tag & 0x3FF) | ((level & 0x3FF) << 10) | (size << 20)
        return struct.pack("<I", head) + payload
    head = (tag & 0x3FF) | ((level & 0x3FF) << 10) | (0xFFF << 20)
    return struct.pack("<I", head) + struct.pack("<I", size) + payload


def para_header(n_chars: int) -> bytes:
    return struct.pack(
        "<IIHBBHHHI",
        n_chars,  # 글자 수
        0,  # control mask
        0,  # para shape id
        0,  # style id
        0,  # column type
        1,  # char shape 개수
        0,  # range tag 개수
        1,  # line seg 개수
        0,  # instance id
    )


def para_text(text: str) -> bytes:
    return text.encode("utf-16-le")


def para_char_shape() -> bytes:
    # (문자위치 0, char_shape_id 0)
    return struct.pack("<II", 0, 0)


def build_section(paragraphs: list[str]) -> bytes:
    out = bytearray()
    for p in paragraphs:
        out += record(HWPTAG_PARA_HEADER, 0, para_header(len(p)))
        out += record(HWPTAG_PARA_TEXT, 1, para_text(p))
        out += record(HWPTAG_PARA_CHAR_SHAPE, 1, para_char_shape())
    return bytes(out)


def build_file_header(compressed: bool) -> bytes:
    head = bytearray(256)
    head[0 : len(SIGNATURE)] = SIGNATURE
    # 버전 5.0.3.0 (리틀엔디언으로 build, revision, minor, major)
    struct.pack_into("<BBBB", head, 32, 0, 3, 0, 5)
    props = 0x01 if compressed else 0x00
    struct.pack_into("<I", head, 36, props)
    return bytes(head)


def write_hwp(path: Path, paragraphs: list[str], *, compressed: bool = True) -> Path:
    """OLE 복합 문서로 .hwp 를 만든다. Windows 전용."""
    try:
        import pythoncom
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("이 표본 생성기는 Windows + pywin32 가 필요합니다.") from e

    from win32com.storagecon import (
        STGM_CREATE,
        STGM_READWRITE,
        STGM_SHARE_EXCLUSIVE,
    )

    if path.exists():
        path.unlink()

    mode = STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE
    stg = pythoncom.StgCreateDocfile(str(path), mode)

    def put(name: str, data: bytes, parent=None) -> None:
        target = parent if parent is not None else stg
        stream = target.CreateStream(name, mode, 0, 0)
        stream.Write(data)

    section = build_section(paragraphs)
    if compressed:
        # HWP는 zlib 헤더 없는 raw deflate를 쓴다.
        comp = zlib.compressobj(9, zlib.DEFLATED, -15)
        section = comp.compress(section) + comp.flush()

    put("FileHeader", build_file_header(compressed))
    # 서식 테이블은 비워 두되, 압축 플래그가 켜져 있으면 빈 스트림도 압축
    # 형식이어야 한다. 그렇지 않으면 압축 해제 단계에서 걸린다.
    empty_docinfo = b""
    if compressed:
        c = zlib.compressobj(9, zlib.DEFLATED, -15)
        empty_docinfo = c.compress(b"") + c.flush()
    put("DocInfo", empty_docinfo)

    body = stg.CreateStorage("BodyText", mode, 0, 0)
    put("Section0", section, parent=body)

    body.Commit(0)
    stg.Commit(0)
    del body, stg
    return path


SAMPLE_PARAGRAPHS = [
    "문서 변환기 시험 표본",
    "형식: HWP 5.0 (합성)",
    "용도: 읽기 경로 검증",
    "DocumentConverter 가 만든 표본입니다.",
    "한글 가나다라마바사, 한자 漢字, 숫자 1,234,567원을 함께 담습니다.",
    "레코드 순회와 압축 해제, 텍스트 디코딩이 제대로 도는지 봅니다.",
]


def main() -> int:
    out = Path(__file__).resolve().parent.parent / "output" / "hwp표본.hwp"
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        write_hwp(out, SAMPLE_PARAGRAPHS)
    except Exception as e:
        print(f"표본 생성 실패: {type(e).__name__}: {e}")
        return 1
    print(f"생성: {out}  ({out.stat().st_size:,}B)")

    # 바로 읽어 본다.
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from docconv.readers import hwp_reader

    doc = hwp_reader.read(out)
    text = doc.to_text()
    print(f"읽기: {len(text)}자, 문단 {doc.stats()['paragraphs']}개")
    for line in text.splitlines()[:8]:
        print(f"   {line!r}")

    from tests.sample import PROBES

    ok = all(p in text for p in PROBES)
    print("표본 검증:", "통과" if ok else "실패")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
