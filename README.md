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

`sources/` 폴더에 마크다운 파일을 작성합니다. `#` (h1)가 슬라이드 구분자이며, 그 아래의 마크다운 요소들이 슬라이드 콘텐츠가 됩니다.

> **구성은 자유입니다.** 아래는 다양한 마크다운 요소가 프리뷰/PPTX에 어떻게 반영되는지 보여주는 예시입니다. 반드시 이 구성을 따를 필요는 없으며, 프로젝트 성격에 맞게 자유롭게 작성하면 됩니다.

#### 기본 구조: `#`이 슬라이드를 나눈다

```markdown
# 첫 번째 슬라이드 제목
이 영역의 모든 내용이 하나의 슬라이드가 됩니다.

# 두 번째 슬라이드 제목
여기부터 새 슬라이드입니다.
```

- `#` (h1) = 슬라이드 구분자. **첫 번째 `#`은 타이틀 슬라이드**, 마지막은 클로징 슬라이드로 자동 판별됩니다.
- `#` 앞의 텍스트는 타이틀 슬라이드의 부제목/메타로 활용됩니다.
- 콘텐츠가 한 슬라이드에 담기지 않으면 **자동으로 분할**됩니다 (걱정 없이 내용을 충실히 작성하세요).

---

#### 마크다운 요소별 렌더링 안내

아래는 `sources/studymate-project.md`에서 사용된 실제 패턴입니다.

##### 타이틀 슬라이드 — `#` + 부제목 텍스트

```markdown
# StudyMate — AI 기반 스터디 매칭 플랫폼

**팀명: 코드브릿지 (CodeBridge)**
2026년 1학기 캡스톤 디자인 최종 발표
```

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| `# 제목` | 슬라이드 메인 타이틀 (54pt Bold, 중앙 정렬) |
| `#` 아래 텍스트 | 부제목 (24pt, accent bar 아래 배치) |

##### 텍스트 + 구조화 요소 — `##`, 리스트, 인용구, 구분선

```markdown
# 비즈니스 개요

## 문제 정의

대학생의 **78%**가 스터디 그룹을 원하지만...

- 에브리타임, 카카오톡 등 기존 채널은 **정보 비대칭**이 심함
- 실력 수준, 학습 목표가 맞지 않아 **중도 이탈률 62%**

---

## 솔루션: StudyMate

> AI가 학습 성향을 분석하여 최적의 스터디 그룹을 자동 매칭하는 플랫폼

### 타겟 사용자

1. 전공 시험 대비 스터디를 찾는 대학생
2. 코딩 테스트 준비 그룹을 원하는 취준생
```

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| `## 소제목` | 24pt Bold + accent 하단 라인 (슬라이드 내 섹션 구분) |
| `### 소소제목` | 20pt Bold (소항목 제목) |
| `- 항목` | accent 색상 네이티브 불릿 목록 |
| `1. 항목` | accent 색상 번호 목록 (다단계 지원) |
| `> 인용구` | secondary 배경 + 좌측 accent 바 + 이탤릭 |
| `---` | 수평 구분선 (시각적 분리) |
| `**굵게**` | Bold |
| 내용이 길면 | `##` 경계에서 **자동 분할** (여러 슬라이드로) |

##### 표 — 파이프 테이블

```markdown
| 구분          | 내용                          | 예상 단가    |
| ------------- | ----------------------------- | ----------- |
| 프리미엄 구독  | AI 심층 매칭, 학습 분석 리포트  | ₩4,900/월   |
| 기업 제휴     | 채용 연계 스터디 스폰서십       | ₩500,000/건 |
```

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| 파이프 테이블 | `table` 레이아웃 자동 선택. accent 헤더, 줄무늬 행, 세로선 제거 |
| 본문 내 소형 테이블 | `inline_table`로 body_elements 내부에 렌더링 |

##### URL 스크린샷 — 웹사이트 자동 캡처

```markdown
# 화면 시나리오

## 메인 페이지

<https://www.wanted.co.kr>

사용자가 처음 접속하면 보이는 랜딩 페이지입니다.

## 매칭 결과 화면

<https://www.rocketpunch.com>

AI가 분석한 호환성 점수와 함께 추천 스터디 목록을 표시합니다.
```

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| `https://...` URL | Playwright로 **자동 스크린샷 캡처** → 이미지로 삽입 |
| URL 1개 + 텍스트 | `content-image` (좌측 텍스트 + 우측 이미지) |
| URL 2개 | `two-images` (좌우 분할) |
| URL 3개+ | `grid-images` (그리드 배치) |
| `##`별 URL | 각 `##`가 별도 슬라이드로 분할되어 각각 스크린샷 포함 |

##### YouTube — 영상 임베드

```markdown
# 시연 영상

프로젝트 시연 영상입니다.

<https://www.youtube.com/watch?v=dQw4w9WgXcQ>
```

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| YouTube URL | 프리뷰: iframe 임베드 재생 |
| | PPTX: **온라인 비디오 임베드** (PowerPoint 2013+에서 인라인 재생) + 썸네일 + ▶ 버튼 |

##### 코드블록 — 구문 표시

````markdown
```java
@RestController
public class ApiController {
    @GetMapping("/hello")
    public String hello() { return "Hello!"; }
}
```
````

| 요소 | 프리뷰/PPTX 반영 |
|------|------------------|
| ` ```lang ``` ` | 어두운 배경 + Consolas 모노스페이스 + 언어 라벨 |
| 12줄 초과 시 | 자동 truncate + `...` 표시 |

##### 다이어그램 — mermaid 이미지 변환

ERD, 플로우차트 등은 **코드블록보다 이미지**가 효과적입니다. 마크다운에 텍스트 ERD를 작성해도 되지만, 별도로 mermaid 파일을 만들어 이미지로 변환하면 훨씬 보기 좋습니다.

```bash
# mermaid 다이어그램 → PNG 이미지 변환
npx -y @mermaid-js/mermaid-cli -i erd.mmd -o erd.png -w 1920 -H 1080
```

변환된 이미지는 `_captures/` 폴더에 넣고, slides.json에서 `image-full` 레이아웃으로 삽입됩니다.

---

#### 인라인 텍스트 서식

| 마크다운 | 렌더링 결과 |
|----------|-------------|
| `**굵게**` | **Bold** |
| `*기울임*` | *Italic* |
| `~~취소선~~` | ~~Strikethrough~~ |
| `` `인라인코드` `` | `Consolas 모노스페이스` |

---

#### 요약: 마크다운 요소 → 슬라이드 매핑 한눈에 보기

| 마크다운 패턴 | 자동 선택 레이아웃 | 비고 |
|---------------|-------------------|------|
| `#` 첫 번째 (+ 부제목) | `title` | 타이틀 슬라이드 |
| `#` + 텍스트만 | `content` | `body_elements`로 구조화 |
| `#` + 텍스트 + URL 1개 | `content-image` | 좌 텍스트 + 우 스크린샷 |
| `#` + URL만 1개 | `image-full` | 풀블리드 이미지 |
| `#` + URL 2개 | `two-images` | 좌우 분할 |
| `#` + URL 3개+ | `grid-images` | 그리드 배치 |
| `#` + 파이프 테이블 | `table` | 미니멀 표 디자인 |
| `#` + 코드블록 | `code` | 코드 표시 슬라이드 |
| `#` + YouTube URL | `content-image` + 비디오 | 인라인 재생 가능 |
| `#` 마지막 (감사/Q&A) | `closing` | 클로징 슬라이드 |
| 콘텐츠 초과 시 | 자동 분할 | `##` 경계 우선, 높이 기반 추가 분할 |

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
| [references/troubleshooting.md](./.ai/skills/md-to-pptx/references/troubleshooting.md) | 트러블슈팅 가이드 (이미지 경로, 불릿, YouTube 등 10개 항목) |


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

