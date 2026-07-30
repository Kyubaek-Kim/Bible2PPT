# Bible2PPT

성경 구절을 골라 배경 이미지가 깔린 PowerPoint 슬라이드를 만들어 주는 **경량 데스크톱 앱**입니다.
개발 지식이 없어도 실행파일을 더블클릭해 사용할 수 있고, DB 등 외부 인프라 없이 로컬 JSON + 설정
파일만 사용합니다.

> 이번 릴리스의 실제 동작 타깃은 **Windows 전용**입니다. 다만 OS 의존 코드(경로/폴더 열기/폰트
> 등록 등)를 전용 모듈로 격리해 두어, 향후 macOS 이식이 최소 수정으로 가능하도록 설계했습니다.

## 목차

- [주요 기능](#주요-기능)
- [사용자용 (Windows)](#사용자용-windows)
  - [내 성경 파일 올리기 (업로드/등록)](#내-성경-파일-올리기-업로드등록)
  - [내 설정·데이터는 어디에 저장되나요?](#내-설정데이터는-어디에-저장되나요)
- [포함된 번역본과 출처/라이선스](#포함된-번역본과-출처라이선스)
- [개발자용](#개발자용)
- [아키텍처](#아키텍처)

---

## 주요 기능

- **다중 번역본 교차 배치** — 여러 번역본을 동시에 골라 같은 절을 나란히 한 슬라이드에 배치.
- **유연한 구절 입력** — 드롭다운(책·장·절, 장 넘김 범위 지원) 또는 자유 형식 직접 입력.
- **여러 구절 일괄 처리** — 구절+제목을 목록에 담아 개별 PPT 또는 통합 PPT로 한 번에 생성.
- **자동 페이지네이션** — 절 묶음이 슬라이드에 넘치지 않도록 줄 수를 계산하고 필요 시 글자 크기를
  자동 축소. 절 번호는 내어쓰기로 왼쪽에 정렬.
- **타이포/레이아웃 커스터마이징** — 화면 비율, 글꼴·크기·굵기, 그리고 제목·구절·본문 상자 위치를
  드래그로 조정해 저장.
- **배경 이미지** — 기본 배경 또는 사용자 이미지(비율 자동 크롭, 잘릴 양 사전 안내).
- **내 성경 파일 등록** — `.txt`/`.json` 업로드 → 검토 → 등록(원문 보존, 다양한 인코딩 자동 인식).
- **UI 다국어(한/영)** 및 설정 영속화(다음 실행 시 복원).
- **본문 원문 무결성 보장** — 어떤 단계에서도 성경 본문을 수정하지 않음(아래 참고).

---

## 사용자용 (Windows)

1. 배포된 `Bible2PPT.exe`를 내려받아 더블클릭하면 실행됩니다. (별도 설치 불필요 — 번역본/폰트/
   기본 배경이 실행파일에 포함되어 있습니다.)
2. 화면은 세로형(휴대폰 크기) 레이아웃입니다. 위에서부터:
   - **화면 언어**: UI 표시 언어(한국어/영어). *본문 번역본 선택과는 별개입니다.*
   - **번역본**: 여러 개를 동시에(체크박스) 선택할 수 있습니다. 각 항목에는 언어가 괄호로 표기됩니다
     (예: `개역한글 (한국어)`, `King James Version (영어)`). 목록이 길면 `더보기`로 접고, `자주 사용`
     체크한 번역본은 체크한 순서대로 맨 위에 고정됩니다. 각 행의 `기준` 라디오로 고른 **책·장·절 기준
     성경**은 (1) 위 장·절 드롭다운의 절 개수 기준이자 (2) 교차 배치 순서의 맨 앞이 됩니다.
   - **구절 선택**: 성경/장/시작 절/끝 절 드롭다운으로 선택하거나, "직접 입력"에 자유 형식으로
     입력합니다. `창세기 15:1-15`, `창 15 1 15`, `창15:1~15`, `창 1:23-2:5`(장 넘김) 모두 인식합니다.
     끝이 시작보다 앞서는 잘못된 범위나 0 이하의 장·절은 거부되고 안내가 표시됩니다.
   - **제목**: 비워 두면 구간 정보(예: `창세기 1:1-5`)가 대신 표시됩니다.
   - **담기**: 입력한 구절을 목록에 추가합니다. 여러 구절을 담아 한 번에 생성할 수 있습니다.
   - **생성 방식**: (구절 목록 바로 아래) "구절별 개별 PPT" 또는 "1개 PPT 통합".
   - **화면 설정**: 화면 비율(16:9 / 4:3 / A4), 글자체(미리보기 제공), 본문 글자크기, 본문 굵게.
     하단 **화면 구성 커스터마이징**에서 제목·구절·본문 상자를 드래그로 재배치하고 제목·구절의
     글자체/크기/굵기를 지정한 뒤 저장할 수 있습니다(초기화로 기본 배치 복원).
   - **배경 이미지**: 기본 배경 또는 사용자 이미지 첨부. 비율이 다르면 잘릴 양을 픽셀·cm로 알려
     주고 확인을 받은 뒤 비율에 맞춰 크롭합니다. 과거 배경 히스토리에서 다시 고를 수 있습니다.
   - **저장 위치**: 기본은 문서 폴더이며 변경 가능합니다. 생성 후 폴더를 바로 열 수 있습니다.
   - **생성**: 화면 하단에 항상 고정된 버튼으로, 번역본·구절을 고르지 않았으면 안내창이 뜹니다.

### 내 성경 파일 올리기 (업로드/등록)

`번역본 파일 업로드`에서 `.txt` 또는 `.json`을 선택하면 형식을 자동 감지해 파싱합니다.

- 지원 예: txt `창 1:1 본문` / `Genesis 1:1 text`(탭·다중 공백 구분), json 중첩 구조
  (`{"창": {"1": {"1": "본문"}}}`), 평면 구조(`{"창1:1": "본문"}`), 행 배열 등.
- **검토 단계**에서 책/장/절 통계, 누락·중복·형식오류 라인, 정경과 절 개수가 다른 책을 보여 줍니다.
  검토를 통과해야 등록할 수 있습니다.
- 등록 시 이름/언어/약자를 지정하면 번역본 목록(드롭다운)에 바로 나타납니다. **본문 텍스트는 원문
  그대로 저장**되며, 원본 파일도 함께 보관됩니다.

### 내 설정·데이터는 어디에 저장되나요?

앱은 DB 없이 로컬 파일만 사용합니다. 설정·등록한 성경·배경 히스토리는 **OS 표준 사용자 데이터
폴더**에 저장되어 다음 실행 때 자동 복원됩니다(프로그램 폴더는 건드리지 않음).

| 항목 | Windows | macOS(이식 예정) | Linux(개발) |
|------|---------|------------------|-------------|
| 설정/데이터 | `%APPDATA%\Bible2PPT` | `~/Library/Application Support/Bible2PPT` | `$XDG_DATA_HOME/Bible2PPT` |
| 기본 출력 폴더 | 사용자 문서 폴더 | 〃 | 〃 |

등록한 성경은 위 폴더의 `bibles/`에, 첨부한 배경 원본은 배경 히스토리 폴더에 보관됩니다. 출력 폴더는
UI에서 언제든 바꿀 수 있습니다.

---

## 포함된 번역본과 출처/라이선스

`scripts/fetch_bibles.py`로 공개 저장소에서 내려받아 정경 `(책ID, 장, 절)` 스키마
`data/bibles/<코드>.json`로 저장합니다. **저작권 판본(개역개정 등)은 제외**했습니다.

| 코드 | 이름 | 언어 | 출처 | 라이선스 |
|------|------|------|------|----------|
| KRV | 개역한글 | 한국어 | getbible.net (`korean`) | Public domain (1961) |
| KJV | King James Version | 영어 | getbible.net | Public domain |
| ASV | American Standard Version | 영어 | getbible.net | Public domain |
| WEB | World English Bible | 영어 | getbible.net | Public domain |
| YLT | Young's Literal Translation | 영어 | getbible.net | Public domain |
| TR | Textus Receptus (NT) | 헬라어 | getbible.net | Public domain |
| WH | Westcott-Hort (NT) | 헬라어 | getbible.net | Public domain |
| LXX | Septuagint (OT) | 헬라어 | getbible.net | Public domain |
| WLC | Westminster Leningrad Codex (OT) | 히브리어 | getbible.net | Public domain / CC |
| ALEPPO | Aleppo Codex (OT) | 히브리어 | getbible.net | Public domain |
| VULGATE | Clementine Vulgate | 라틴어 | getbible.net | Public domain |

번들 폰트: **나눔스퀘어 Bold(NanumSquare Bold, 기본)**, **나눔고딕/나눔고딕 Bold(NanumGothic)** — 모두 SIL Open Font License (`data/fonts/OFL.txt`). 글꼴 드롭다운에는 이 번들 글꼴과 함께 Windows 기본 글꼴(**맑은 고딕·굴림·돋움·바탕·궁서**)이 표시됩니다. 본문 굵기는 체크박스로 켜고 끌 수 있습니다.

> 요청 목록 중 SBLGNT, Nestle1904는 getbible.net에 없어 별도 어댑터가 필요하며, 스크립트에
> 출처와 함께 TODO로 표시해 두었습니다.

---

## 개발자용

### 요구사항(의존성) 파일 — `requirements.txt` vs `requirements-dev.txt`

두 파일을 나눈 이유는 **앱을 실행하는 데 꼭 필요한 것**과 **개발·빌드에만 필요한 것**을
분리하기 위해서입니다(파이썬 프로젝트의 일반적 관례).

| 파일 | 용도 | 포함 패키지 |
|------|------|-----------|
| `requirements.txt` | **런타임** 의존성만 | `python-pptx`, `Pillow`, `requests` (Tkinter는 CPython 내장) |
| `requirements-dev.txt` | 런타임 + **개발/빌드 도구** | 첫 줄의 `-r requirements.txt`로 위를 포함 + `pytest`(테스트)·`ruff`(린트)·`pyinstaller`(.exe 빌드)·`fonttools`(폰트 처리) |

따라서 배포물에 테스트/빌드 도구 같은 불필요한 패키지가 섞이지 않으며, **개발자·CI는
`requirements-dev.txt` 하나만** 설치하면 런타임+도구가 한 번에 깔립니다.

### 설치 & 실행

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python main.py
```

Tkinter는 CPython에 기본 포함됩니다. (리눅스에서 개발 시 `sudo apt install python3-tk` 필요할 수 있음.)

### 번역본 다시 받기 / 정경 데이터 재생성

```bash
python scripts/build_canon.py            # data/canon.json
python scripts/fetch_bibles.py           # data/bibles/*.json + data/versification/kjv.json
python scripts/fetch_bibles.py KRV KJV   # 일부만
python scripts/make_icon.py              # run_icon.ico / run_icon.png
```

### 테스트 & 린트

```bash
pytest -q          # core 로직 단위/통합 테스트
ruff check .       # 린트
```

GUI는 헤드리스(Xvfb)로 스모크 검증할 수 있습니다(리눅스):

```bash
PYTHONPATH=. xvfb-run -a python -c "import ui.app as a; app=a.App(); app.update(); app.destroy()"
```

CI(GitHub Actions)는 Python 3.10/3.12에서 `ruff`와 `pytest`를 실행합니다.

### Windows 실행파일 빌드

```bash
pip install -r requirements-dev.txt
pyinstaller Bible2PPT.spec        # dist/Bible2PPT.exe
```

`Bible2PPT.spec`은 데이터 수집/Analysis를 OS 독립적으로 두고, 아이콘·번들 방식만 `sys.platform`
분기로 격리해 두어 추후 macOS `.app` 브랜치를 쉽게 켤 수 있습니다. (GitHub Actions 크로스 빌드/
서명·공증은 이번 범위 밖입니다.)

---

## 아키텍처

로직은 `core/`, 표현은 `ui/`에 두어 분리합니다. `core`는 Tkinter를 import하지 않으므로 단위 테스트/
재사용이 쉽습니다. 전체 설계·데이터 흐름·모듈 책임은 [`docs/DESIGN.md`](docs/DESIGN.md)에 정리했습니다.

```
main.py                엔트리포인트 (얇음)
core/
  paths.py             OS 독립 리소스/사용자 데이터 경로 (Windows 구현 + macOS 스텁)
  platform_util.py     폴더 열기 · 폰트 등록 등 OS 의존 동작 격리
  parser.py            구절 파싱·정규화 (직접 입력/드롭다운/장 넘김)
  bible.py             정경 책ID · 번역본 레지스트리 · 구절 범위 확장
  alignment.py         번역본 간 절 정렬 (표시 레이어 매핑, 본문 무수정)
  importer.py          사용자 성경 업로드 파싱/검증/등록
  ppt.py               슬라이드 엔진 (비율·폰트·페이지네이션·배경·교차 배치)
  image_util.py        배경 크롭/알림 (Pillow)
  fonts.py             번들 폰트 + 미리보기 지원
  i18n.py              UI 언어 테이블
  settings.py          설정 영속화
  generator.py         상위 오케스트레이션 (UI가 호출)
ui/app.py              Tkinter GUI
data/                  canon / bibles / i18n / versification / fonts / 기본 배경
scripts/               canon·번역본·아이콘 생성 스크립트
tests/                 core 단위 테스트
```

### 핵심 원칙 — 본문 원문 무결성

정렬·정규화·파싱·전처리 등 어떤 단계에서도 **성경 본문 텍스트 원문을 수정/훼손하지 않습니다.**
저장은 항상 원문 그대로, 절 정렬은 표시(display) 레이어의 매핑으로만 처리합니다
(밀림=매핑 정렬, 합본=`N-M절` 라벨, 누락=`(해당 절 없음)`).
