---
name: md-to-pptx
description: 마크다운 문서를 읽어 `#` 헤딩 기준으로 슬라이드를 분리하고, 이미지·URL 스크린샷을 포함한 전문적인 PPTX 파일을 생성하는 스킬. 사용자가 "마크다운을 PPT로", "md를 pptx로 변환", "이 문서로 발표자료 만들어줘", "슬라이드 만들어줘", "PPT 생성", "프레젠테이션 파일 만들어줘", "/md-to-pptx" 등을 요청하면 이 스킬을 사용한다. .pptx 파일 출력이 필요한 모든 프레젠테이션 요청에 트리거된다. 단, 웹 기반 HTML 슬라이드를 원하는 경우에는 web-presentation 스킬을 사용한다.
---

# md-to-pptx — 마크다운 → PPTX 변환 스킬

---

## ⚠️ 필수 준수 사항 (반드시 숙지 후 시행)

**이 스킬(SKILL.md)과 GEMINI.md의 내용을 반드시 전부 숙지한 후에만 작업을 시작한다.**

- **숙지 의무**: 본 문서의 모든 Step(0~5), 규칙, 주의사항을 하나도 빠짐없이 읽고 이해한다.
- **시행 의무**: 숙지한 내용을 하나도 누락 없이 그대로 실행한다. 단계 생략, 규칙 무시, URL 캡처 누락은 금지된다.
- **검증 의무**: Step 0(진행 모드), Step 0.5(출력 경로), Step 4.5(URL 완전성 검증), Step 3(프리뷰 확인) 등 필수 단계를 건너뛰지 않는다.

**위 사항을 준수하지 않으면 작업 결과가 무효이다.**

---

마크다운 문서를 `#` 헤딩 기준으로 슬라이드를 분리하고, PPT_Design_Guide_2026 원칙에 따라 전문적인 PPTX 파일을 생성한다.

## 워크플로우 개요

```
마크다운 파일 → 파싱(# 기준) → 콘텐츠 분석 → URL 캡처 → JSON spec → 프리뷰 HTML → 사용자 확인 → PPTX 생성
```

**중요**: 프리뷰를 먼저 생성하여 사용자가 테마와 구성을 확인한 뒤, 사용자가 OK하면 PPTX를 생성한다. **본 스킬의 모든 Step과 규칙을 하나도 누락 없이 시행한다.**

---

## Step 0: 진행 모드 선택

스킬 실행 시작 시 사용자에게 **진행 모드**를 질문한다:

> **진행 모드를 선택해주세요:**
>
> 1. **자동 진행** — 프리뷰 확인만 받고, 나머지 단계는 자동으로 진행합니다. **(이 모드에서는 모든 도구 실행 권한(Action Required)을 '이번 세션 동안 허용(Allow for this session)'으로 처리한 것으로 간주하여 중단 없이 연속 실행합니다.)**
> 2. **단계별 확인** — 각 주요 단계마다 진행 여부를 확인합니다.

사용자가 "1", "자동", "auto" 등으로 응답하면 → **자동 모드**: 프리뷰 확인(Step 3)에서만 사용자 응답을 기다리고, 나머지 단계는 중간 확인 없이 연속 실행한다. **CLI에서 권한 요청(Action Required) 발생 시 사용자가 'Allow for this session'을 선택하면 이후 모든 도구 실행이 자동 승인되어 끊김 없는 자동화가 가능함을 안내한다.**

사용자가 "2", "단계별", "step" 등으로 응답하면 → **단계별 모드**: 각 Step 완료 후 "계속 진행할까요?" 확인을 받는다.

사용자가 모드를 선택하지 않고 바로 작업 지시를 하거나, 모호한 응답을 하면 → **자동 모드**를 기본값으로 사용한다.

---

## Step 0.5: 출력 경로 결정 (필수)

이 스킬은 **다양한 프로젝트에서 호출**될 수 있다. 출력 파일(프리뷰 HTML, choices.json, PPTX, 캡처 이미지)의 저장 경로를 **스킬이 호출된 프로젝트의 작업 디렉토리**를 기준으로 결정해야 한다.

### 경로 결정 규칙

1. **현재 작업 디렉토리(CWD) 확인** — `pwd` 또는 환경 정보에서 현재 프로젝트 루트를 파악한다.
2. **프로젝트 구조 탐색** — 현재 프로젝트에 이미 PPTX 관련 출력 디렉토리가 있는지 확인한다 (`ls`로 탐색).
3. **출력 디렉토리 결정**:
   - 사용자가 출력 경로를 명시적으로 지정한 경우 → 그 경로를 사용한다.
   - 입력 마크다운 파일이 특정 디렉토리에 있으면 → 해당 마크다운 파일과 같은 디렉토리에 출력한다.
   - 위 조건이 없으면 → 현재 작업 디렉토리 아래에 `md-pptx-convert/` 디렉토리를 생성하여 사용한다.

**절대 이전 프로젝트의 경로를 재사용하지 않는다.** 매 실행 시 현재 프로젝트의 CWD를 기준으로 경로를 새로 결정한다.

```bash
# 예시: 현재 프로젝트 확인
pwd
# → /home/user/my-project

# 출력 디렉토리
# → /home/user/my-project/md-pptx-convert/
```

---

## Step 1: 마크다운 파싱

사용자가 제공한 마크다운 파일을 읽고 `#`(h1) 기준으로 슬라이드를 분리한다.

### 파싱 규칙

- `#` (h1)만 슬라이드 구분자로 사용한다. `##`, `###` 등은 슬라이드 내부의 구조적 요소로 처리한다.
- 첫 번째 `#` 이전의 내용은 무시하거나, 내용이 있으면 첫 슬라이드의 부제목/메타 정보로 활용한다.
- 각 슬라이드 섹션에서 다음을 감지한다:
  - **로컬 이미지**: `![alt](로컬경로)` 패턴 → 로컬 파일 경로
  - **이미지 URL**: `![alt](https://...)` 패턴 또는 이미지 확장자(`.png`, `.jpg`, `.gif`, `.webp`, `.svg` 등)로 끝나는 URL → **직접 다운로드**하여 슬라이드에 원본 이미지로 삽입
  - **웹사이트 URL**: `http://` 또는 `https://`로 시작하는 링크 중 이미지가 아닌 것 → **스크린샷 캡처**하여 슬라이드에 삽입
  - **구조적 텍스트 요소**: 아래 마크다운 문법에 따라 `body_elements` 배열로 구조화

### 마크다운 문법 → body_elements 매핑 (필수)

슬라이드 내부(`#` 사이)의 마크다운 요소를 분석하여, 단순 문자열 `body` 대신 **`body_elements` 배열**로 구조화한다. 각 요소는 대응하는 슬라이드 디자인으로 렌더링된다.

| 마크다운 문법 | type | 슬라이드 디자인 | 설명 |
|---|---|---|---|
| `## 제목` | `heading` (level:2) | 24pt Bold + accent 하단 라인 | 슬라이드 내 섹션 구분 |
| `### 소제목` | `heading` (level:3) | 20pt Bold | 소항목 제목 |
| 일반 텍스트 | `paragraph` | 18pt 본문 텍스트 | 줄바꿈 단위로 분리 |
| `- 항목` / `* 항목` | `bullet_list` | accent 색상 불릿 + 들여쓰기 | 다단계 지원 (level 0~3) |
| `1. 항목` | `numbered_list` | accent 색상 번호 + 들여쓰기 | 순서가 있는 목록 |
| `> 인용구` | `blockquote` | secondary 배경 + 좌측 accent 바 + 이탤릭 | 강조/인용 |
| `---` | `divider` | 수평 구분선 | 슬라이드 내 시각적 분리 |
| ` ```lang ``` ` | `code_block` | 어두운 배경 + Consolas 모노스페이스 + 언어 라벨 | 코드 표시 |
| `\| 헤더 \| ... \|` | `inline_table` | 미니멀 테이블 (accent 헤더, 줄무늬) | 본문 내 소형 테이블 |

**body_elements 구조 예시:**

```json
{
  "layout": "content",
  "title": "기술 스택",
  "body_elements": [
    { "type": "heading", "level": 2, "text": "백엔드" },
    { "type": "bullet_list", "items": ["Spring Boot 4.0", "Java 21", "MySQL 8.0"] },
    { "type": "divider" },
    { "type": "heading", "level": 2, "text": "프론트엔드" },
    { "type": "paragraph", "text": "Vanilla JS 기반의 경량 프론트엔드" },
    { "type": "blockquote", "text": "SPA 프레임워크 없이 순수 JS로 구현하여 번들 사이즈를 최소화" },
    { "type": "code_block", "language": "java", "code": "@RestController\npublic class ApiController {\n    ...\n}" },
    { "type": "inline_table", "headers": ["기술", "버전"], "rows": [["Spring Boot", "4.0"], ["Java", "21"]] }
  ]
}
```

**중첩 리스트 표현:**

리스트 항목에 레벨 정보를 포함하여 다단계를 표현한다:

```json
{
  "type": "bullet_list",
  "items": [
    "최상위 항목",
    { "text": "하위 항목", "level": 1 },
    { "text": "더 깊은 항목", "level": 2 },
    "다시 최상위"
  ]
}
```

**호환성**: `body_elements`가 없으면 기존 `body` 문자열 방식으로 렌더링한다. 두 필드가 모두 있으면 `body_elements`를 우선 사용한다.

### URL 완전성 규칙 (필수)

**슬라이드 내 모든 URL은 반드시 캡처하여 이미지로 포함해야 한다. 단 하나의 누락도 허용하지 않는다.**

- 각 `#` 섹션에서 발견된 **모든** `http://`, `https://` URL을 추출한다.
- URL이 1개이면 `content-image`, 2개이면 `two-images`, 3개 이상이면 `grid-images` 레이아웃을 사용한다.
- 텍스트 + URL 2개 이상이면 `two-images` 또는 `grid-images`에 `body` 필드도 **반드시** 함께 포함한다. 복수 이미지 슬라이드에서 텍스트가 누락되면 안 된다.
- JSON spec 생성 후, **검증 단계**를 실행한다:
  1. 원본 마크다운의 각 섹션에서 URL 목록을 재추출한다.
  2. JSON spec의 각 슬라이드에서 `image`, `images` 필드의 이미지 파일 수를 센다.
  3. URL 수 ≠ 이미지 수인 슬라이드가 있으면 **즉시 수정**한다.
  4. 검증 결과를 사용자에게 표로 보고한다.

### 슬라이드 타입 자동 결정

각 섹션의 콘텐츠를 분석하여 최적의 레이아웃을 자동 선택한다:

| 콘텐츠 구성                | 레이아웃        | 설명                                   |
| -------------------------- | --------------- | -------------------------------------- |
| 제목만 (첫 슬라이드)       | `title`         | 타이틀 슬라이드                        |
| 텍스트만                   | `content`       | 텍스트 중심 슬라이드                   |
| 텍스트 + 이미지 1개        | `content-image` | 좌측 텍스트 + 우측 이미지              |
| 이미지 1개만               | `image-full`    | 풀블리드 이미지                        |
| 이미지 2개                 | `two-images`    | 좌우 분할                              |
| 이미지 3개+                | `grid-images`   | 그리드 배치 (프리뷰에서 레이아웃 선택) |
| 표 포함                    | `table`         | 미니멀 표 디자인                       |
| 코드블록 포함              | `code`          | 코드 표시 슬라이드                     |
| 마지막 슬라이드 (감사/Q&A) | `closing`       | 클로징 슬라이드                        |

Claude는 콘텐츠를 분석하여 위 규칙을 기반으로 판단하되, 콘텐츠의 의미와 맥락도 함께 고려하여 최적의 레이아웃을 선택한다. 예를 들어 비교 내용이면 `comparison`, 타임라인이면 `timeline`, KPI 수치가 있으면 `kpi` 등으로 판단할 수 있다.

### 콘텐츠 오버플로우 자동 분할

한 `#` 섹션의 내용이 슬라이드 한 장에 담기엔 너무 많으면, generate_pptx.py가 **자동으로 여러 슬라이드로 분할**한다. 이 기능 덕분에 텍스트가 슬라이드 밖으로 벗어나거나 요소끼리 겹치는 문제가 방지된다.

**동작 원리 (3단계 분할):**

1. **오버플로우 감지** — 슬라이드의 `body_elements` 총 높이가 가용 높이(content 레이아웃 기준 5.8")를 **실제로 초과할 때만** 분할을 수행한다. 초과하지 않으면 h2가 여러 개여도 분할하지 않는다.
2. **h2 기준 분할** — 오버플로우가 감지되면, `## 서브타이틀`(heading level 2) 경계에서 우선 분할한다. 분할된 슬라이드의 제목은 해당 h2 텍스트가 된다.
3. **높이 기반 추가 분할** — h2 섹션 하나가 여전히 가용 높이를 초과하면, 요소 높이를 추정하여 적절한 지점에서 추가 분할한다.
4. **제목 자동 생성** — 분할된 슬라이드에 h2 제목이 없으면, 해당 슬라이드의 첫 번째 콘텐츠(h3 heading, 문장 첫 구절, 리스트 키워드 등)에서 자동으로 제목을 생성한다. `(계속)` 같은 임의 접미사는 사용하지 않는다.

**추가 규칙:**
- `content-image` 레이아웃에서 분할 시, 첫 슬라이드만 이미지를 유지하고 연속 슬라이드는 `content` 레이아웃으로 변경된다.
- 요소 간에는 `0.2"` 수직 여백이 자동 적용된다.
- 타이틀 슬라이드(`layout_title`, `layout_closing`)는 제목 줄 수에 따라 accent bar와 부제목 위치가 동적으로 조정된다.

**JSON spec 작성 시 참고:**
- `body_elements`에 콘텐츠를 최대한 충실히 담되, 분량 초과를 걱정하지 않아도 된다.
- 생성 엔진이 자동으로 적절한 지점에서 분할한다.
- 다만 의미적으로 분리가 자연스러운 경우(예: 하위 주제가 2개 이상), 파싱 단계에서 미리 2개 슬라이드로 나누는 것이 더 보기 좋을 수 있다.

---

## Step 2: URL 처리 (스크린샷 캡처 / 이미지 다운로드)

마크다운 본문에서 발견된 URL을 **유형에 따라 자동 분류**하여 처리한다:

- **YouTube URL** (`youtube.com/watch?v=`, `youtu.be/`) → **프리뷰에서 iframe 임베드**, **PPTX에서 온라인 비디오 임베드 (PowerPoint 2013+에서 인라인 재생 가능)** + 썸네일 포스터 프레임 + ▶ 재생 버튼 + 폴백 하이퍼링크. 스크린샷도 함께 캡처하여 PPTX 썸네일로 사용한다. JSON spec에 `video_url` 필드를 추가한다.
- **이미지 URL** (`.png`, `.jpg`, `.jpeg`, `.gif`, `.webp`, `.svg`, `.bmp`, `.ico`, `.tiff`, `.avif` 확장자 또는 Content-Type이 `image/*`) → **직접 다운로드**하여 원본 이미지를 슬라이드에 삽입
- **웹사이트 URL** (위에 해당하지 않는 일반 페이지) → **Playwright 스크린샷 캡처**하여 슬라이드에 삽입

capture.mjs가 URL 유형을 자동 판별하므로, 호출 방식은 동일하다. 이미지 URL이면 다운로드, 웹사이트면 스크린샷을 수행한다. YouTube URL도 웹사이트로 캡처하되, JSON spec에 `video_url` 필드를 별도 추가한다.

### 사전 준비

```bash
cd <skill-path>
npm ls playwright 2>/dev/null || npm install playwright
npx playwright install chromium
```

### 실행

```bash
node <skill-path>/scripts/capture.mjs <URL> <저장경로.png> [옵션]
```

| 옵션          | 설명                                    | 기본값     |
| ------------- | --------------------------------------- | ---------- |
| `--full-page` | 전체 페이지 스크롤 캡처 (웹사이트만)   | viewport만 |
| `--width`     | 뷰포트 너비 px (웹사이트만)            | 1920       |
| `--height`    | 뷰포트 높이 px (웹사이트만)            | 1080       |
| `--wait`      | 로드 후 대기 시간 초 (웹사이트만)      | 2          |
| `--device`    | 모바일 디바이스 에뮬레이션 (웹사이트만) | 없음       |

### URL 유형 판별 로직

1. URL 경로의 확장자가 이미지 확장자 목록에 포함되면 → 이미지로 판별
2. 확장자로 판별이 안 되면 → HTTP HEAD 요청으로 `Content-Type` 확인
3. `Content-Type`이 `image/*`이면 → 이미지 다운로드
4. 그 외 → 웹사이트 스크린샷 캡처

캡처/다운로드한 이미지는 출력 PPTX와 같은 디렉토리의 `_captures/` 폴더에 저장한다. 파일명은 URL의 도메인 기반으로 자동 생성한다.

---

## Step 3: 프리뷰 HTML 생성 및 사용자 선택

PPTX를 생성하기 전에 **프리뷰 HTML 페이지**를 생성하여 사용자가 다음을 한 화면에서 선택할 수 있게 한다:

1. **테마 선택** — 6가지 빌트인 테마 중 선택
2. **다중 이미지 레이아웃 선택** — 이미지가 2개 이상인 슬라이드의 배치 방식 선택

### 프리뷰 실행

```bash
python <skill-path>/scripts/preview.py <slides-json> <output-html> [--images-dir <경로>]
```

- `slides-json`: Step 1에서 파싱한 슬라이드 정보 JSON
- `output-html`: 생성할 HTML 파일 경로
- `--images-dir`: 이미지 파일이 위치한 디렉토리
- 기본적으로 **로컬 서버 모드**로 실행된다 (자동 저장 지원). `--no-serve` 옵션으로 비활성화 가능.

**중요 (필수 준수 - 절대 경로 배너 의무)**: 생성된 프리뷰 HTML의 **최상단(Header 바로 아래)**에는 반드시 **Step 0.5에서 결정된 결과물 저장소의 절대 경로**를 **강조된 배너(Banner)** 형태로 명시해야 한다.
- 안내 내용: `출력 디렉토리의 절대 경로`, `choices.json이 저장될 절대 파일 경로`
- 시각적 요구사항: 배경색 대비를 통한 강조, "저장 경로 안내" 명칭 사용
- 목적: 사용자가 `choices.json`을 저장해야 할 위치를 0.1초 만에 파악하게 함.
**이 단계를 누락하거나 모호하게 표시하는 경우, 해당 작업 결과는 무효이며 즉시 재수행해야 한다.**

**실행 방식**: preview.py를 **백그라운드로** 실행한다. 서버가 시작되면 자동으로 브라우저가 열린다.

```bash
python <skill-path>/scripts/preview.py <slides-json> <output-html> --images-dir <경로> &
```

사용자가 "확인" 버튼을 클릭하면 choices.json이 **다이얼로그 없이 자동 저장**된다 (서버 POST `/save-choices`).

생성된 HTML을 브라우저에서 열면:

- 상단: 테마 프리뷰 카드 6개 (클릭하여 선택)
- 하단: 프레젠테이션 구성 뷰어 (인라인 슬라이드 뷰어 + 전체 화면 모드)
  - 도트 네비게이션으로 슬라이드 탐색
  - 다중 이미지 슬라이드: 레이아웃 선택 UI 표시
  - "전체 화면" 버튼 또는 F5로 프레젠테이션 모드 진입
- 확인 바: "확인" 버튼 → choices.json 저장 (프리뷰/PPT와 같은 경로)

프리뷰 HTML 생성 시 **기본 choices.json이 자동 생성**된다 (slides.json의 meta.theme 값 사용). 사용자가 프리뷰에서 테마를 변경하고 "확인"을 누르면 choices.json이 덮어쓰여진다.

**워크플로우**: 프리뷰 생성(기본 choices.json 자동 생성) → 사용자에게 프리뷰 안내 → 사용자가 OK/진행 응답 → choices.json을 읽어 JSON spec의 meta.theme에 반영 → PPTX 생성.

사용자가 OK하면 **choices.json을 읽고 바로 PPTX를 생성**한다. choices.json에는 사용자가 프리뷰에서 선택한 테마가 이미 반영되어 있으므로, 별도로 테마를 물어보거나 "기본 테마로 생성합니다" 같은 안내를 하지 않는다. 이런 안내는 사용자가 자신의 선택이 무시되었다고 오해할 수 있기 때문이다.

### 프리뷰 결과 JSON 형식

```json
{
  "theme": "dark",
  "slide_overrides": {
    "3": { "layout": "grid-2x2" },
    "5": { "layout": "main-sub" }
  }
}
```

---

## Step 4: JSON Spec 생성

파싱 결과 + 사용자 선택을 합쳐 PPTX 생성용 JSON spec을 만든다.

```json
{
  "meta": {
    "title": "프레젠테이션 제목",
    "author": "",
    "date": "2026-03-14",
    "theme": "dark"
  },
  "slides": [
    {
      "layout": "title",
      "title": "제목 텍스트",
      "subtitle": "부제목",
      "notes": "발표자 노트"
    },
    {
      "layout": "content-image",
      "title": "슬라이드 제목",
      "body_elements": [
        { "type": "bullet_list", "items": ["항목 1", "항목 2", "**강조** 텍스트"] },
        { "type": "blockquote", "text": "핵심 메시지를 강조" }
      ],
      "image": "./images/photo.png",
      "notes": ""
    },
    {
      "layout": "grid-images",
      "title": "화면 캡처 모음",
      "images": ["./cap1.png", "./cap2.png", "./cap3.png"],
      "grid": "2x2",
      "notes": ""
    }
  ]
}
```

### 지원 레이아웃 전체 목록

| 레이아웃        | 필수 필드                   | 선택 필드                                 |
| --------------- | --------------------------- | ----------------------------------------- |
| `title`         | `title`                     | `subtitle`, `date`, `author`              |
| `section`       | `title`                     | `subtitle`                                |
| `content`       | `title`, `body` 또는 `body_elements` | `notes`                           |
| `content-image` | `title`, (`body` 또는 `body_elements`), `image` | `image_position`(left/right), `video_url`, `notes` |
| `image-full`    | `image`                     | `title`, `caption`, `overlay_text`        |
| `two-images`    | `images`(2개)               | `title`, (`body` 또는 `body_elements`), `captions`, `notes` |
| `grid-images`   | `images`(3개+)              | `title`, `grid`(2x2/3x1/1x3/2x1), `notes` |
| `comparison`    | `title`, `left`, `right`    | `notes` (left/right 내에 `body_elements` 사용 가능) |
| `table`         | `title`, `headers`, `rows`  | `highlight_rows`, `notes`                 |
| `code`          | `title`, `code`, `language` | `notes`                                   |
| `kpi`           | `title`, `metrics`          | `notes`                                   |
| `timeline`      | `title`, `events`           | `notes`                                   |
| `closing`       | `title`                     | `subtitle`, `contact`                     |

> **참고**: `body_elements`와 `body`가 동시에 있으면 `body_elements`를 우선 사용한다. 마크다운 문서 파싱 시에는 구조적 요소(제목, 인용구, 코드블록, 테이블, 구분선 등)가 포함된 경우 반드시 `body_elements`를 사용한다.

---

## Step 5: PPTX 생성

JSON spec을 읽어 python-pptx로 PPTX 파일을 생성한다.

### 사전 준비

```bash
pip install python-pptx Pillow
```

### 실행

```bash
python <skill-path>/scripts/generate_pptx.py <spec.json> <output.pptx>
```

### 테마 시스템

6가지 빌트인 테마를 제공한다:

| 테마         | 배경      | 텍스트    | 강조색    | 적합 상황       |
| ------------ | --------- | --------- | --------- | --------------- |
| `dark`       | `#1A1A2E` | `#FFFFFF` | `#E94560` | 테크, 스타트업  |
| `light`      | `#FFFFFF` | `#1A1A1A` | `#2563EB` | 범용, 비즈니스  |
| `minimal`    | `#FAFAFA` | `#333333` | `#6366F1` | 미니멀, 교육    |
| `consulting` | `#FFFFFF` | `#002F6C` | `#C41230` | 컨설팅, 보고서  |
| `pitch`      | `#0F0A1A` | `#FFFFFF` | `#7B2FBE` | 투자 IR, 피치덱 |
| `education`  | `#F8F7FF` | `#1E1B4B` | `#F97316` | 교육, 강의      |

### 디자인 원칙 (PPT_Design_Guide_2026 기반)

생성 시 다음 규칙을 자동 적용한다:

- **타이포그래피**: 제목 48pt+ Bold, 본문 20pt+ Regular, 최대 2종 폰트
- **색상**: 60-30-10 법칙 (주색 60%, 보조색 30%, 강조색 10%)
- **여백**: 슬라이드 면적의 30%+ 여백 확보, 요소 간 최소 16px
- **표**: 세로선 제거, 헤더 아래 가로선 1개만, 숫자 우측정렬
- **차트**: 3D 효과/그림자 금지, Direct Labeling, Data-Ink Ratio 극대화
- **접근성**: WCAG AA 대비율 4.5:1 이상, 색맹 지원 (Blue-Orange 계열)
- **슬라이드 크기**: 16:9 와이드스크린 (13.333" × 7.5")

디자인 규칙의 상세 내용은 `references/design-rules.md`를 참조한다.

---

## 전체 워크플로우 요약

```
0. 진행 모드 선택 (자동 진행 / 단계별 확인) — 기본값: 자동 진행
1. 마크다운 파일 읽기
2. # 기준으로 슬라이드 분리 + 콘텐츠 분석
3. URL 발견 시 → capture.mjs로 스크린샷 캡처 (섹션 내 모든 URL 캡처 필수)
4. JSON spec 생성
4.5. ★ URL 완전성 검증 — 원본 마크다운 재확인하여 URL 수 ≠ 이미지 수 슬라이드 즉시 수정
5. preview.py로 프리뷰 HTML 생성 (기본 choices.json 자동 생성) → 브라우저에서 열기
6. 사용자에게 프리뷰 안내 (슬라이드 구성 요약 + URL 검증 결과 표시)
7. 사용자가 OK/진행 응답 → choices.json 읽기 → JSON spec의 theme 반영
8. generate_pptx.py로 PPTX 생성
9. 결과 파일 경로를 사용자에게 안내
```

**진행 모드에 따른 동작:**

- **자동 모드**: Step 6(프리뷰 확인)에서만 사용자 응답을 기다리고, 나머지는 연속 실행한다.
- **단계별 모드**: 각 Step 완료 후 "계속 진행할까요?" 확인을 받는다.

**핵심 순서**: 프리뷰 먼저 → 사용자 OK → choices.json 기반 PPTX 생성.
프리뷰에서 사용자가 선택한 테마가 choices.json에 이미 저장되어 있으므로, 사용자 응답 후 choices.json을 읽고 바로 PPTX를 생성한다. 테마를 다시 확인하거나 "기본 테마" 언급을 하지 않는다.

**⚠️ 본 스킬의 모든 Step(0~5)과 규칙을 하나도 누락 없이 시행한다. 단계 생략은 금지된다.**

### 주의사항

- **이미지 경로 기준**: slides.json의 이미지 경로는 **slides.json 파일 위치 기준** 상대경로로 작성한다. 캡처 이미지는 slides.json과 같은 디렉토리의 `_captures/` 하위에 저장하고, 경로는 `./_captures/파일명.png` 형식을 사용한다. 프로젝트 루트 기준 경로(`./output/...`)로 작성하면 generate_pptx.py가 경로를 찾지 못한다.
- URL 캡처 실패 시 해당 이미지를 건너뛰고 텍스트로 URL을 표시한다
- 로그인 필요 페이지는 캡처할 수 없다 (공개 페이지만 지원)
- 이미지가 1개인 슬라이드는 레이아웃 선택 없이 자동 배치한다
- 이미지가 2개 이상인 슬라이드만 프리뷰에서 레이아웃 선택 UI를 표시한다
- **프리뷰 브라우저 오픈**: preview.py를 서버 모드(기본값)로 실행하면 자동으로 브라우저가 열린다. 별도로 `start`/`open` 명령을 실행하면 2번 열리므로, 서버 모드에서는 수동 오픈하지 않는다. 비서버 모드(`--no-serve`)에서만 수동으로 HTML을 연다.
- **다이어그램 지원**: ERD, 플로우차트 등 다이어그램이 필요한 경우 mermaid CLI로 이미지를 생성하여 `image-full` 또는 `content-image` 레이아웃으로 삽입한다. 코드블록 텍스트보다 시각적 다이어그램이 훨씬 효과적이다.
  ```bash
  npx -y @mermaid-js/mermaid-cli -i diagram.mmd -o diagram.png -w 1920 -H 1080 --backgroundColor transparent
  ```

### 트러블슈팅

문제 발생 시 `references/troubleshooting.md`를 참조한다. 이미지 경로, 프리뷰 중복 오픈, 불릿 정렬, YouTube 재생, 코드블록 오버플로우 등 주요 이슈의 원인과 해결 방법이 정리되어 있다.

### 마크다운 본문 포맷 지원

슬라이드 body 텍스트에서 다음 마크다운 문법을 지원한다:

- `**굵게**` → Bold
- `*기울임*` → Italic
- `~~취소선~~` → Strikethrough (oxml `sngStrike` 속성으로 렌더링)
- `` `인라인코드` `` → Consolas 모노스페이스
- `` `코드` `` → 고정폭 폰트
- `- 항목` → 불릿 리스트
- `1. 항목` → 번호 리스트
