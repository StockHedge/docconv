# docconv — 로컬 문서 변환기

PDF · DOCX · DOC · HWP · HWPX · XLSX · XLS · CSV 를 서로 변환한다.
파일을 외부 서비스에 올리지 않고 **전부 내 컴퓨터에서** 처리한다.

한국어 문서(HWP/HWPX)를 한컴오피스 없이 다루는 것이 이 프로그램의 중심이다.
HWPX는 순수 파이썬으로 읽고 쓰며, 구형 HWP도 읽는다.

Windows 우선, macOS/Linux 에서도 동작한다. Python 3.10 이상.

| | |
|---|---|
| 전 조합 시험 | 8종 사이 가능한 49개 경로 전부 통과 (실패 0) |
| 한글 깨짐 | PDF 글꼴 임베딩·서브셋, CSV BOM, DOCX eastAsia 처리 |
| 의존성 | 순수 파이썬 6개. 외부 프로그램 없이 대부분 동작 |
| 라이선스 | MIT |

---

## 빠른 시작

### Windows

```
최초 설치.bat      한 번만 실행 - 필요한 패키지를 설치하고 환경을 점검한다
변환기 실행.bat    창을 띄운다. 파일을 끌어다 놓고 형식을 고른 뒤 [변환 시작]
```

### macOS / Linux

```bash
pip install -r requirements.txt
./docconv.sh              # 창 열기
./docconv.sh 파일.hwpx --to pdf
```

### 명령줄

```bash
python -m docconv 기획서.hwpx --to pdf
python -m docconv *.hwp --to docx --out-dir 결과
python -m docconv 보고서.pdf --to xlsx --tables-only
python -m docconv --doctor          # 환경 진단
```

---

## 지원 범위

| 형식 | 읽기 | 쓰기 | 비고 |
|---|:--:|:--:|---|
| PDF | O | O | 읽기는 텍스트·표·이미지 추출. 스캔본은 OCR이 없어 글자를 못 뽑는다 |
| DOCX | O | O | 네이티브 |
| DOC | O | O | 외부 엔진 필요 (LibreOffice 또는 MS Word) |
| HWPX | O | O | 네이티브. 한컴오피스 없이 읽고 쓴다 |
| HWP | O | **X** | **읽기 전용** — 아래 설명 참고 |
| XLSX | O | O | 네이티브 |
| XLS | O | O | 읽기는 네이티브(xlrd), 쓰기는 외부 엔진 필요 |
| CSV / TSV | O | O | 인코딩 자동 판별, 출력은 Excel 호환 UTF-8 BOM |
| TXT · MD · HTML · ODT · RTF · JSON | O | 일부 | 덤으로 지원 |

### `.hwp` 로 저장할 수 없는 이유

HWP 5.0 은 공개된 명세로 **읽을** 수는 있지만, 한컴오피스 없이 유효한 파일을
**만드는** 것은 현실적으로 불가능하다. 대신 **HWPX 로 저장**하면 한/글에서 그대로
열리고, 거기서 `다른 이름으로 저장 → .hwp` 하면 된다. HWPX 는 한/글 2010 이후
버전의 표준 형식이다.

---

## 동작 방식

### 중간표현(IR)을 거친다

8개 형식을 직접 연결하면 56가지 변환기가 필요하다. 대신 모든 입력을 공통
중간표현으로 바꾸고, 거기서 각 형식으로 내보낸다. 필요한 모듈이 16개로 준다.

```
[입력] → Reader → Document IR → Writer → [출력]
                   문단·표·이미지
```

### 서식은 어디까지 옮기나

한/글 문서는 셀 서식을 셀에 직접 쓰지 않고 `borderFillIDRef` 로 참조 테이블을
가리킨다. 그래서 이 테이블을 읽어야 **배경 음영과 테두리**가 살아난다. 표마다
칸마다 테두리 유무가 다른 양식 문서(제목 칸만 상자, 본문은 좌우선만)를 그대로
재현하는 데 필요하다. 쪽 번호도 본문이 아니라 구역 설정에 들어 있어 따로 옮긴다.

옮기는 것: 글꼴·크기·굵기·기울임·밑줄·취소선·글자색·음영, 문단 정렬·들여쓰기·
줄간격, 표의 셀 병합·열 너비·배경·4방향 테두리, 이미지, 용지 크기·여백, 쪽 번호.

옮기지 못하는 것: 수식, 차트, 도형, 각주/미주 위치, 머리말/꼬리말 본문,
글상자, 필드(자동 날짜 등)는 텍스트로만 남거나 생략된다.

### 엔진 3계층 — 자동 선택, 자동 폴백

| 순위 | 엔진 | 플랫폼 | 담당 |
|---|---|---|---|
| 1 | **native** (순수 파이썬) | 전부 | pdf, docx, hwpx, hwp(읽기), xlsx, xls(읽기), csv |
| 2 | **libreoffice** (headless) | 전부 | doc, xls 쓰기, 폭넓은 폴백 |
| 3 | **msoffice** (COM) | Windows | doc, xls 최종 폴백 |

앞의 엔진이 실패하면 자동으로 다음을 시도한다. 어느 엔진도 직접 처리하지 못하면
**중간 형식을 경유**한다. 경로는 너비 우선 탐색으로 찾으므로 항상 가장 짧은 길을
고른다(경유가 늘수록 서식 손실이 쌓이기 때문).

```
doc → [MS Word] → docx → [native] → hwpx          2단계
doc → [MS Word] → docx → [native] → xlsx → [Excel] → xls    3단계
```

3단계까지 찾는 이유가 있다. `doc → xls` 는 어떤 2단계 조합으로도 불가능하다 —
`doc→docx` 는 되지만 `docx→xls` 가 안 되고(Excel은 docx를 못 읽는다),
`csv→xls` 는 되지만 `doc→csv` 가 안 된다(Word는 csv를 못 만든다).

`--doctor` 로 지금 어떤 엔진이 살아 있는지 볼 수 있다.

---

## 한글이 깨지지 않게 하는 장치

이런 변환기가 가장 자주 실패하는 지점이라 따로 다뤘다.

**PDF 글꼴 임베딩과 서브셋.** PDF 표준 글꼴에는 한글 글리프가 없어서 글꼴 파일을
문서에 넣어야 한다. 그런데 한글 글꼴은 글리프가 2만 자를 넘어 파일 자체가 10MB를
넘는다. 맑은 고딕 Regular(13.4MB) + Bold(12.6MB)를 그대로 넣으면 **26MB짜리 PDF**가
나온다. 그래서 실제 쓰인 글자만 남기는 서브셋을 적용한다 — **26.4MB → 236KB**.

**CSV 의 UTF-8 BOM.** Excel 은 BOM 없는 UTF-8 CSV 를 시스템 코드페이지(한국어
Windows 는 CP949)로 해석해 한글을 깨뜨린다. 그래서 출력 기본값이 `utf-8-sig` 다.

**CSV 읽을 때 인코딩 판별.** CP949 저장본과 UTF-8 이 섞여 돌아다닌다. BOM →
디코딩 성공 여부 → 한글/제어문자 비율 점수 순으로 고른다. `latin-1` 은 어떤
바이트열이든 성공하므로 감점 처리한다.

**DOCX 의 eastAsia 글꼴.** Word 는 문자 계열별로 글꼴을 따로 고른다. `w:ascii` 만
지정하면 한글은 엉뚱한 기본 글꼴로 나온다. `w:eastAsia` 를 함께 지정한다.

**Word COM 의 텍스트 저장 형식.** `wdFormatUnicodeText = 7` 은 평문이 아니라
UTF-16 이다. 이름만 보고 쓰면 결과가 UTF-16LE 로 저장되어 이후 단계에서 전부
깨진다. 저장 뒤 인코딩을 확인해 UTF-8 로 되돌린다.

**PDF 표의 열 너비.** 조판 엔진이 `width:83%` 같은 백분율을 존중하지 않아 제목이
좁은 칸에 밀려 두 줄로 깨졌다. pt 절대값으로 환산해 넘긴다. 열 너비를 아예 주지
않으면 대부분 빈 열이 극단적으로 좁아져 그 칸의 긴 값이 잘린다.

---

## 옵션

### 표 · 스프레드시트

| 옵션 | 값 | 설명 |
|---|---|---|
| `--merged` | `placeholder` | 병합 셀의 좌상단에만 값, 나머지는 빈 칸 (기본) |
| | `duplicate` | 병합 영역 전체에 같은 값 복제 — 필터·피벗에 유리 |
| | `keep` | 실제 병합 유지 (XLSX 만) |
| `--tables-only` | | 문단은 버리고 표만 추출 |
| `--single-sheet` | | 여러 표를 시트 한 장으로 합침 |
| `--csv-encoding` | `utf-8-sig` | 출력 인코딩 (기본값이 Excel 호환) |

### 문서

| 옵션 | 설명 |
|---|---|
| `--font 경로` | PDF 에 쓸 글꼴 파일 지정 (기본: 시스템에서 자동 탐색) |
| `--pages 1-10` | PDF 에서 읽을 쪽 범위 |
| `--password` | 암호가 걸린 입력 파일 |
| `--no-format` | 서식을 버리고 텍스트만 옮김 |

### 엔진

| 옵션 | 설명 |
|---|---|
| `--engine native` | 특정 엔진 강제 |
| `--native-only` | 외부 프로그램을 전혀 쓰지 않음 |
| `--timeout 300` | 외부 엔진 제한 시간(초) |

---

## 지원 범위를 넓히려면

`.doc` 과 `.xls` 쓰기는 외부 엔진이 있어야 한다.

**LibreOffice** (권장, 모든 OS)

```
Windows : winget install TheDocumentFoundation.LibreOffice
macOS   : brew install --cask libreoffice
Linux   : sudo apt install libreoffice
```

**MS Office** (Windows, 이미 설치돼 있다면)

```
pip install pywin32
```

---

## 구조

```
docconv/
  ir.py            중간표현 - 모든 변환의 중심
  formats.py       형식 정의와 판별(확장자 + 매직 바이트)
  options.py       변환 옵션
  pipeline.py      엔진 선택, 경유 경로 탐색, 실행
  cli.py           명령줄
  readers/         각 형식 → IR
    hwpx_reader.py   OWPML(ZIP+XML) 파싱
    hwp_reader.py    HWP 5.0 바이너리 레코드 파싱
    pdf_reader.py    PyMuPDF, 표 영역 인식 후 문단 재구성
    docx_reader.py   순서 보존 + XML 레벨 병합 셀 처리
    sheet_reader.py  XLSX / XLS / CSV
    text_reader.py   TXT / MD / HTML / RTF / ODT
  writers/         IR → 각 형식
    hwpx_writer.py   OWPML 생성 (서식 테이블 2패스)
    pdf_writer.py    Story 조판 + 글꼴 서브셋
    docx_writer.py   eastAsia 글꼴, 병합 후처리
    sheet_writer.py  XLSX / CSV
    text_writer.py   TXT / MD / HTML / JSON
    flatten.py       문서 → 격자 평탄화
  engines/         native / libreoffice / msoffice
  util/            단위 환산, 글꼴 탐색, 안전 검사
  gui/app.py       tkinter 화면
tests/
  sample.py           시험용 표본 문서 생성 (IR -> HWPX)
  test_all_pairs.py   8종 전 조합(56가지) 전수 시험
  test_matrix.py      보너스 형식 포함 왕복 시험
  make_hwp_sample.py  HWP 5.0 표본 생성기(읽기 검증용)
docs/
  구현노트.md          실제로 부딪힌 함정과 해결책
```

---

## 검증

```bash
python -m tests.test_all_pairs   # 핵심 8종 전 조합(56가지) 전수 시험
python -m tests.test_matrix      # 보너스 형식 포함 왕복 시험
python -m docconv --doctor       # 환경 진단
```

`test_all_pairs` 는 8종 사이의 가능한 모든 방향을 하나씩 실제로 변환하고,
결과를 다시 읽어 표본 문자열이 살아남았는지 센다. MS Office 가 설치된 환경에서:

```
OK 49 / 부분 0 / 원리적 불가 7 / 실패 0
```

**시험용 문서는 프로그램이 직접 만든다**(`tests/sample.py`). 제목·서식·표·병합·
배경·테두리를 고루 담은 문서를 IR로 조립해 HWPX로 내보낸 뒤 그걸 원본으로 쓴다.
특정 파일을 준비할 필요가 없다.

손에 있는 실제 문서로 확인하고 싶으면 환경 변수로 지정한다.

```bash
# Windows
set DOCCONV_TEST_SOURCE=C:\경로\내문서.hwpx
# macOS / Linux
export DOCCONV_TEST_SOURCE=~/문서/내문서.hwpx
```

`.hwp` 를 입력으로 쓰는 행은 명세대로 조립한 합성 표본을 쓴다
(`tests/make_hwp_sample.py`). 한컴오피스가 없으면 진짜 `.hwp` 를 만들 수 없기
때문인데, 레코드 순회·압축 해제·텍스트 디코딩은 이것으로 검증되지만 실제 한/글이
저장한 파일의 다양한 컨트롤까지 확인해 주지는 못한다.

---

## 알려진 한계

- **스캔 PDF** 는 텍스트 레이어가 없어 글자를 뽑을 수 없다. OCR 은 범위 밖이다.
- **PDF 읽기는 재구성이다.** PDF 에는 문단 개념이 없어 좌표에서 역산한다.
  복잡한 다단 편집물은 읽기 순서가 어긋날 수 있다.
- **`.hwp` 쓰기 불가.** 위 설명 참고.
- **수식·차트·도형** 은 텍스트로만 옮겨지거나 생략된다.
- HWP 바이너리의 **그림 추출**은 버전에 따라 실패할 수 있다(구조 오프셋이 유동적).
- **표 행이 페이지 경계에서 통째로 다음 장으로 밀린다.** 한/글은 셀 내부를 쪼개
  나누지만 PDF 조판 엔진은 행 단위로만 나눈다. 그래서 긴 행 앞에서 페이지 하단이
  빌 수 있다. 표를 문단으로 풀면 해결되지만 테두리와 배경을 잃으므로 원본 충실도를
  택했다.
- **CSV 는 한 파일에 시트 하나만 담긴다.** 문서에서 표가 여러 장 나오면
  `이름.csv` 와 `이름_표1.csv` 처럼 갈라진다. 내용이 사라지는 것은 아니다.

---

## 더 읽을거리

만들면서 실측으로 확인한 함정들을 [`docs/구현노트.md`](docs/구현노트.md) 에 정리했다.
PDF 글꼴 서브셋, HWPX의 `borderFill` 참조 구조, HWP 레코드 헤더의 12비트 크기 필드,
python-docx 병합 셀 판별의 함정, Word COM 상수 오독 같은 것들이다. 비슷한 걸 만드는
사람에게는 시간을 꽤 아껴 줄 것이다.

## 라이선스

MIT. 자세한 내용은 [LICENSE](LICENSE) 참고.
