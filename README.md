# md-to-pptx-skill-mng

마크다운 문서를 `#` 헤딩 기준으로 슬라이드로 분리하고, URL 스크린샷을 포함한 전문적인 PPTX 파일을 생성하는 프로젝트입니다.

## 개요

이 프로젝트는 **md-to-pptx** 스킬을 관리하며, 마크다운 소스 문서를 읽어 다음 워크플로우로 PPTX를 생성합니다:

```
마크다운 파일 → 파싱(# 기준) → 콘텐츠 분석 → URL 캡처 → JSON spec → 프리뷰 HTML → 사용자 확인 → PPTX 생성
```

- **슬라이드 구분**: `#` (h1) 헤딩을 기준으로 슬라이드 분리
- **URL 자동 캡처**: 본문의 `http://`, `https://` 링크를 Playwright로 스크린샷
- **프리뷰 우선**: 테마·레이아웃 선택 후 사용자 확인을 거쳐 PPTX 생성
- **디자인 원칙**: PPT_Design_Guide_2026 기반 (타이포그래피, 색상, 여백, 접근성)

## 프로젝트 구조

```
md-to-pptx-skill-mng/
├── README.md                 # 이 문서
├── GEMINI.md                 # Gemini 실행 시 프로젝트 가이드 (필독)
├── .ai/
│   └── skills/
│       └── md-to-pptx/       # md-to-pptx 스킬
│           ├── SKILL.md      # 스킬 정의 및 워크플로우 (핵심 참조)
│           ├── scripts/
│           │   ├── capture.mjs      # URL → 스크린샷 캡처 (Playwright)
│           │   ├── preview.py      # JSON → 프리뷰 HTML 생성
│           │   └── generate_pptx.py # JSON spec → PPTX 생성
│           ├── package.json
│           ├── references/
│           │   └── design-rules.md  # 디자인 규칙 상세
│           └── node_modules/       # Playwright 등 (캡처용)
├── sources/                  # 마크다운 소스 문서
│   ├── react-official-site.md
│   └── spring-official-site.md
├── output/                   # 생성 결과물 (소스별 하위 폴더)
│   └── <소스명>/
│       ├── slides.json       # 파싱된 슬라이드 JSON
│       ├── preview.html      # 프리뷰 HTML
│       ├── preview.choices.json # 사용자 선택 (테마, 레이아웃)
│       ├── *.pptx            # 최종 PPTX 파일
│       └── _captures/        # URL 스크린샷 이미지
└── .vscode/
    └── settings.json
```

## 사전 요구사항


| 도구           | 용도                              |
| ------------ | ------------------------------- |
| **Node.js**  | Playwright 기반 URL 캡처            |
| **Python 3** | preview.py, generate_pptx.py 실행 |
| **pip**      | python-pptx, Pillow 설치          |


### 의존성 설치

```bash
# 스킬 디렉토리로 이동
cd .ai/skills/md-to-pptx

# Playwright (URL 캡처용)
npm install
npx playwright install chromium

# Python 패키지 (PPTX 생성용)
pip install python-pptx Pillow
```

## 사용 방법

### 1. 마크다운 소스 작성

`sources/` 폴더에 마크다운 파일을 작성합니다. `#` (h1)가 슬라이드 구분자입니다.

```markdown
# 1
https://example.com
페이지 제목
페이지 설명 텍스트

# 2
https://example.com/page2
다른 슬라이드
내용...
```

### 2. AI 에이전트로 변환 요청

Cursor, Claude, Gemini 등에서 다음처럼 요청합니다:

- "이 마크다운을 PPT로 변환해줘"
- "sources/react-official-site.md로 발표 자료 만들어줘"
- "md를 pptx로 변환"

### 3. 워크플로우 (AI 에이전트 실행)

1. `sources/*.md` 읽기 → `#` 기준 슬라이드 분리
2. URL 발견 → `capture.mjs`로 스크린샷 → `output/<소스명>/_captures/`
3. JSON spec 생성 + URL 완전성 검증
4. `preview.py` → 프리뷰 HTML → 브라우저에서 확인
5. 사용자 OK → `choices.json` 반영 → `generate_pptx.py` → PPTX 생성

## 지원 레이아웃


| 레이아웃            | 설명                  |
| --------------- | ------------------- |
| `title`         | 타이틀 슬라이드            |
| `content`       | 텍스트 중심              |
| `content-image` | 좌측 텍스트 + 우측 이미지     |
| `image-full`    | 풀블리드 이미지            |
| `two-images`    | 좌우 분할               |
| `grid-images`   | 그리드 배치 (2×2, 3×1 등) |
| `table`         | 표                   |
| `code`          | 코드 블록               |
| `closing`       | 클로징 슬라이드            |


## 테마

6가지 빌트인 테마: `dark`, `light`, `minimal`, `consulting`, `pitch`, `education`

프리뷰 HTML에서 테마를 선택하고 "확인"을 누르면 `preview.choices.json`에 저장됩니다.

## 핵심 문서


| 문서                                                                 | 용도                                   |
| ------------------------------------------------------------------ | ------------------------------------ |
| [GEMINI.md](./GEMINI.md)                                           | Gemini 실행 시 프로젝트 구조·체크리스트            |
| [.ai/skills/md-to-pptx/SKILL.md](./.ai/skills/md-to-pptx/SKILL.md) | 전체 워크플로우, 파싱 규칙, URL 완전성 검증, 레이아웃 타입 |


**AI 에이전트 실행 전**: `GEMINI.md`와 `SKILL.md`를 반드시 숙지한 후 시행합니다.

## 스크립트 직접 실행

```bash
# URL 스크린샷 캡처
node .ai/skills/md-to-pptx/scripts/capture.mjs <URL> <저장경로.png> [--full-page]

# 프리뷰 HTML 생성 (로컬 서버 모드)
python .ai/skills/md-to-pptx/scripts/preview.py <slides.json> <output.html> --images-dir <경로>

# PPTX 생성
python .ai/skills/md-to-pptx/scripts/generate_pptx.py <spec.json> <output.pptx>
```

