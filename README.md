# Bible2PPT

성경 구절을 골라 배경 이미지가 깔린 PowerPoint 슬라이드를 만들어 주는 **경량 데스크톱 앱**입니다.
개발 지식이 없어도 실행파일을 더블클릭해 사용할 수 있고, DB 등 외부 인프라 없이 로컬 JSON + 설정
파일만 사용합니다.

> 이번 릴리스의 실제 동작 타깃은 **Windows 전용**입니다. 다만 OS 의존 코드(경로/폴더 열기/폰트
> 등록 등)를 전용 모듈로 격리해 두어, 향후 macOS 이식이 최소 수정으로 가능하도록 설계했습니다.

---

## 사용자용 (Windows)

1. 배포된 `Bible2PPT.exe`를 내려받아 더블클릭하면 실행됩니다. (별도 설치 불필요 — 번역본/폰트/
   기본 배경이 실행파일에 포함되어 있습니다.)
2. 화면은 세로형(휴대폰 크기) 레이아웃입니다. 위에서부터:
   - **화면 언어**: UI 표시 언어(한국어/영어). *본문 번역본 선택과는 별개입니다.*
   - **번역본**: 여러 개를 동시에 선택할 수 있습니다. 각 항목에는 언어가 괄호로 표기됩니다
     (예: `개역한글 (한국어)`, `King James Version (영어)`). 기본 번역본은 개역한글입니다.
   - **구절 선택**: 성경/장/시작 절/끝 절 드롭다운으로 선택하거나, "직접 입력"에 자유 형식으로
     입력합니다. `창세기 15:1-15`, `창 15 1 15`, `창15:1~15`, `창 1:23-2:5`(장 넘김) 모두 인식합니다.
   - **제목**: 비워 두면 구간 정보(예: `창세기 1:1-5`)가 대신 표시됩니다.
   - **담기**: 입력한 구절을 목록에 추가합니다. 여러 구절을 담아 한 번에 생성할 수 있습니다.
   - **옵션**: 화면 비율(16:9 / 4:3 / A4), 글자체(미리보기 제공), 본문 글자크기.
   - **배경 이미지**: 기본 배경 또는 사용자 이미지 첨부. 비율이 다르면 잘릴 양을 픽셀·cm로 알려
     주고 확인을 받은 뒤 비율에 맞춰 크롭합니다. 과거 배경 히스토리에서 다시 고를 수 있습니다.
   - **저장 위치**: 기본은 문서 폴더이며 변경 가능합니다. 생성 후 폴더를 바로 열 수 있습니다.
   - **생성 방식**: "구절별 개별 PPT" 또는 "1개 PPT 통합".
   - **생성**: 최종 PPT를 만듭니다.

### 내 성경 파일 올리기 (업로드/등록)

`번역본 파일 업로드`에서 `.txt` 또는 `.json`을 선택하면 형식을 자동 감지해 파싱합니다.

- 지원 예: txt `창 1:1 본문` / `Genesis 1:1 text`(탭·다중 공백 구분), json 중첩 구조
  (`{"창": {"1": {"1": "본문"}}}`), 평면 구조(`{"창1:1": "본문"}`), 행 배열 등.
- **검토 단계**에서 책/장/절 통계, 누락·중복·형식오류 라인, 정경과 절 개수가 다른 책을 보여 줍니다.
  검토를 통과해야 등록할 수 있습니다.
- 등록 시 이름/언어/약자를 지정하면 번역본 목록(드롭다운)에 바로 나타납니다. **본문 텍스트는 원문
  그대로 저장**되며, 원본 파일도 함께 보관됩니다.

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

번들 폰트: **나눔스퀘어 볼드(NanumSquare Bold, 기본)**, **나눔고딕(NanumGothic)** — 모두 SIL Open Font License (`data/fonts/OFL.txt`). 글꼴 드롭다운에는 이 둘과 Windows 기본 글꼴 **맑은 고딕**이 표시됩니다.

> 요청 목록 중 SBLGNT, Nestle1904는 getbible.net에 없어 별도 어댑터가 필요하며, 스크립트에
> 출처와 함께 TODO로 표시해 두었습니다.

---

## 개발자용

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
pytest -q
ruff check .
```

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
재사용이 쉽습니다.

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
