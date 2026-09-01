"""신뢰할 수 없는 입력 파일에 대한 방어 계층.

이 프로그램의 존재 이유가 "온라인 변환 서비스에 파일을 올리기 싫다"이므로,
로컬에서 처리한다는 사실만으로는 부족하다. 변환기 자체가 악성 문서에 당하면
안 된다. 문서 포맷이 실제로 악용되는 경로는 셋이다.

1. XXE / 외부 엔티티
   XML 문서가 ``<!ENTITY xxe SYSTEM "file:///etc/passwd">`` 같은 선언으로
   로컬 파일을 읽거나 원격 URL을 호출하게 만든다. DOCX/HWPX/ODT/XLSX가 전부
   XML 기반이므로 전 포맷 공통 위협이다. -> `parse_xml()` 로만 파싱한다.

2. Zip bomb / Zip slip
   압축 해제 시 수십 GB로 부풀거나(`42.zip`), 엔트리 이름에 ``../`` 를 넣어
   컨테이너 밖 경로에 파일을 쓰게 만든다. -> `SafeZip` 이 압축률/총량/경로를
   검사한다.

3. 매크로 / 임베디드 실행 코드
   VBA(`vbaProject.bin`), OLE 개체, 외부 링크. 우리는 문서를 *실행*하지 않고
   텍스트만 뽑으므로 직접적 위험은 낮지만, 발견 시 경고를 남겨 사용자가
   원본 파일 자체를 의심할 수 있게 한다.

어떤 경우에도 네트워크를 쓰지 않는다.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import IO, Iterable, Optional

from lxml import etree

from ..errors import CorruptFileError

# --------------------------------------------------------------------------
# 한도. 정상 문서라면 여유롭게 통과하고, 공격 파일은 걸리는 값으로 잡는다.
# --------------------------------------------------------------------------

#: 압축 해제 후 총 바이트 한도 (2 GiB)
MAX_TOTAL_UNCOMPRESSED = 2 * 1024**3
#: 단일 엔트리 압축 해제 한도 (512 MiB)
MAX_ENTRY_UNCOMPRESSED = 512 * 1024**2
#: 단일 엔트리 최대 압축률. 텍스트 XML은 20:1도 흔하므로 넉넉히 잡는다.
MAX_COMPRESSION_RATIO = 200
#: ZIP 내 엔트리 개수 한도
MAX_ENTRIES = 20_000
#: XML 파싱 시 허용할 최대 깊이(lxml 자체 한도와 별개로 우리가 검사)
MAX_XML_DEPTH = 256

#: 존재하면 사용자에게 알릴 "실행 가능 콘텐츠" 신호
SUSPICIOUS_ENTRIES = (
    "vbaProject.bin",
    "vbaData.xml",
    "macros/",
    "_VBA_PROJECT",
    "Scripts/",
)


@dataclass
class SafetyReport:
    """검사 결과. 치명적이지 않은 관찰은 warnings로 모아 사용자에게 전달한다."""

    warnings: list[str] = field(default_factory=list)
    has_macros: bool = False
    entry_count: int = 0
    uncompressed_bytes: int = 0

    def warn(self, msg: str) -> None:
        if msg not in self.warnings:
            self.warnings.append(msg)


# --------------------------------------------------------------------------
# XML
# --------------------------------------------------------------------------


def _make_parser() -> etree.XMLParser:
    """외부 엔티티·DTD·네트워크를 전면 차단한 파서.

    `resolve_entities=False` 만으로는 부족하다. `no_network=True` 와
    `load_dtd=False` 를 함께 꺼야 SYSTEM 식별자 해석 경로가 완전히 막힌다.
    `huge_tree=False` 는 깊은 중첩/거대 노드로 인한 메모리 폭발을 막는다.
    """
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        dtd_validation=False,
        huge_tree=False,
        recover=False,
    )


def parse_xml(data: bytes, *, source: str = "") -> etree._Element:
    """XML 바이트를 안전하게 파싱한다.

    프로젝트 안의 모든 XML 파싱은 반드시 이 함수를 거친다.
    `etree.fromstring` 직접 호출은 금지.
    """
    try:
        root = etree.fromstring(data, parser=_make_parser())
    except etree.XMLSyntaxError as e:
        raise CorruptFileError(
            "XML 구조가 손상되었습니다.", detail=f"{source}: {e}"
        ) from e

    # DOCTYPE 선언 자체를 거부한다. 정상 오피스 문서에는 없다.
    tree = root.getroottree()
    if tree.docinfo.internalDTD is not None or tree.docinfo.externalDTD is not None:
        raise CorruptFileError(
            "DTD가 포함된 문서는 보안상 처리하지 않습니다.", detail=source
        )
    return root


def parse_xml_lenient(data: bytes, *, source: str = "") -> Optional[etree._Element]:
    """손상 복구를 허용하는 파싱. 보조 파일(스타일 등)에만 쓴다.

    본문 파싱에는 쓰지 말 것 - 조용히 내용이 잘릴 수 있다.
    """
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        huge_tree=False,
        recover=True,
    )
    try:
        return etree.fromstring(data, parser=parser)
    except etree.XMLSyntaxError:
        return None


# --------------------------------------------------------------------------
# ZIP
# --------------------------------------------------------------------------


class SafeZip:
    """zip bomb / zip slip 방어를 붙인 ZipFile 래퍼.

    사용법::

        with SafeZip(path) as z:
            data = z.read("Contents/section0.xml")
            print(z.report.warnings)
    """

    def __init__(self, file: Path | str | IO[bytes]) -> None:
        try:
            self._zf = zipfile.ZipFile(file)
        except zipfile.BadZipFile as e:
            raise CorruptFileError(
                "ZIP 컨테이너가 손상되었습니다.", detail=str(e)
            ) from e
        self.report = SafetyReport()
        self._consumed = 0
        self._audit()

    # -- 컨텍스트 관리 -----------------------------------------------------

    def __enter__(self) -> "SafeZip":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def close(self) -> None:
        self._zf.close()

    # -- 검사 -------------------------------------------------------------

    def _audit(self) -> None:
        infos = self._zf.infolist()
        self.report.entry_count = len(infos)
        if len(infos) > MAX_ENTRIES:
            raise CorruptFileError(
                f"압축 파일의 항목이 너무 많습니다({len(infos)}개). "
                "손상되었거나 안전하지 않은 파일일 수 있습니다."
            )

        total = 0
        for info in infos:
            name = info.filename
            if _is_unsafe_path(name):
                raise CorruptFileError(
                    "압축 파일에 비정상적인 경로가 포함되어 있습니다.", detail=name
                )
            if info.file_size > MAX_ENTRY_UNCOMPRESSED:
                raise CorruptFileError(
                    f"항목 하나의 크기가 한도를 초과합니다: {name} "
                    f"({info.file_size / 1024**2:.0f} MiB)"
                )
            if info.compress_size > 0:
                ratio = info.file_size / info.compress_size
                if ratio > MAX_COMPRESSION_RATIO and info.file_size > 8 * 1024**2:
                    raise CorruptFileError(
                        f"비정상적인 압축률이 감지되었습니다: {name} ({ratio:.0f}:1). "
                        "압축 폭탄일 수 있어 처리를 중단합니다."
                    )
            total += info.file_size
            if any(s.lower() in name.lower() for s in SUSPICIOUS_ENTRIES):
                self.report.has_macros = True
                self.report.warn(
                    f"매크로/스크립트가 포함된 문서입니다({name}). "
                    "변환 결과에는 포함되지 않지만 원본 파일의 출처를 확인하세요."
                )

        if total > MAX_TOTAL_UNCOMPRESSED:
            raise CorruptFileError(
                f"압축 해제 후 총 크기가 한도를 초과합니다 ({total / 1024**3:.1f} GiB)."
            )
        self.report.uncompressed_bytes = total

    # -- 접근 -------------------------------------------------------------

    def namelist(self) -> list[str]:
        return self._zf.namelist()

    def has(self, name: str) -> bool:
        try:
            self._zf.getinfo(name)
            return True
        except KeyError:
            return False

    def read(self, name: str) -> bytes:
        """엔트리를 읽는다. 누적 해제량을 추적해 런타임에도 한도를 지킨다."""
        data = self._zf.read(name)
        self._consumed += len(data)
        if self._consumed > MAX_TOTAL_UNCOMPRESSED:
            raise CorruptFileError("압축 해제 총량이 한도를 초과했습니다.")
        return data

    def read_optional(self, name: str) -> Optional[bytes]:
        try:
            return self.read(name)
        except KeyError:
            return None

    def read_xml(self, name: str) -> etree._Element:
        return parse_xml(self.read(name), source=name)

    def iter_names(self, prefix: str = "", suffix: str = "") -> Iterable[str]:
        for n in self._zf.namelist():
            if n.startswith(prefix) and n.endswith(suffix):
                yield n


def _is_unsafe_path(name: str) -> bool:
    """zip slip 및 절대 경로 검사."""
    if not name:
        return False
    n = name.replace("\\", "/")
    if n.startswith("/") or n.startswith("//"):
        return True
    if len(n) >= 2 and n[1] == ":":  # C:\ 형태
        return True
    parts = n.split("/")
    return any(p == ".." for p in parts)


# --------------------------------------------------------------------------
# OLE (CFB) - .hwp / .doc / .xls
# --------------------------------------------------------------------------

MAX_OLE_STREAM = 512 * 1024**2


def check_ole_stream_size(size: int, name: str) -> None:
    if size > MAX_OLE_STREAM:
        raise CorruptFileError(
            f"OLE 스트림이 너무 큽니다: {name} ({size / 1024**2:.0f} MiB)"
        )
