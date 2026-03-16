# Gemini 실행 시 프로젝트 구조 가이드

---

## ⚠️ 필수 준수 사항 (반드시 숙지 후 시행)

**이 문서(GEMINI.md)와 `.ai/skills/md-to-pptx/SKILL.md`의 내용을 반드시 전부 숙지한 후에만 작업을 시작한다.**

- **숙지 의무**: 두 문서의 모든 섹션, 규칙, 워크플로우를 하나도 빠짐없이 읽고 이해한다.
- **시행 의무**: 숙지한 내용을 하나도 누락 없이 그대로 실행한다. 단계 생략, 규칙 무시, 경로 임의 변경은 금지된다.
- **검증 의무**: URL 완전성 검증, 출력 경로 결정, 프리뷰 확인 등 필수 단계를 건너뛰지 않는다.

**위 사항을 준수하지 않으면 작업 결과가 무효이다.**

---

## 프로젝트 개요

**md-to-pptx-skill-mng** — 마크다운 문서를 `#` 헤딩 기준으로 슬라이드로 분리하고, URL 스크린샷을 포함한 전문적인 PPTX 파일을 생성하는 프로젝트입니다.

---

## 전체 프로젝트 구조

```
md-to-pptx-skill-mng/
├── GEMINI.md                    # 이 문서 (Gemini용 프로젝트 가이드)
├── .ai/
│   └── skills/
│       └── md-to-pptx/          # md-to-pptx 스킬
│           ├── SKILL.md        # 스킬 정의 및 워크플로우 (핵심 참조)
│           ├── scripts/
│           │   ├── capture.mjs      # URL → 스크린샷 캡처 (Playwright)
│           │   ├── preview.py       # JSON → 프리뷰 HTML 생성
│           │   └── generate_pptx.py # JSON spec → PPTX 생성
│           ├── package.json
│           └── node_modules/        # Playwright 등 (캡처용)
├── sources/                    # 마크다운 소스 문서
│   ├── spring-official-site.md
│   └── react-official-site.md
├── output/                     # 생성 결과물
│   ├── slides.json             # 파싱된 슬라이드 JSON
│   ├── preview.html            # 프리뷰 HTML
│   ├── preview.choices.json    # 사용자 선택 (테마, 레이아웃)
│   ├── *.pptx                  # 최종 PPTX 파일
│   └── _captures/              # URL 스크린샷 이미지
│       ├── spring-main.png
│       ├── react-learn.png
│       └── ...
└── .vscode/
    └── settings.json           # 에디터 설정
```

---

## 디렉터리별 설명

| 경로                     | 역할                                                                 |
| ------------------------ | -------------------------------------------------------------------- |
| `sources/`               | 마크다운 소스. `#`(h1)가 슬라이드 구분자. URL 포함 시 자동 캡처 대상 |
| `output/`                | slides.json, preview.html, choices.json, PPTX, \_captures/ 저장      |
| `.ai/skills/md-to-pptx/` | 스킬 정의(SKILL.md) 및 실행 스크립트                                 |
| `output/_captures/`      | URL 스크린샷 PNG 파일 (도메인 기반 파일명)                           |

---

## 핵심 파일

| 파일                             | 용도                                                                        |
| -------------------------------- | --------------------------------------------------------------------------- |
| `.ai/skills/md-to-pptx/SKILL.md` | **필독** — 전체 워크플로우, 파싱 규칙, URL 완전성 검증, 레이아웃 타입, 테마 |
| `sources/*.md`                   | 입력 마크다운. `#` 기준 슬라이드 분리                                       |
| `output/slides.json`             | 파싱 결과 JSON (PPTX 생성용 spec 기반)                                      |
| `output/preview.choices.json`    | 사용자 선택(테마, 다중 이미지 레이아웃)                                     |

---

## 워크플로우 요약

```
1. sources/*.md 읽기 → # 기준 슬라이드 분리
2. {sources.md} 폴더를 생성 후 폴더에 이후 작업파일 저장
3. URL 발견 → capture.mjs로 스크린샷 → output/{sources.md}/_captures/
4. JSON spec 생성 + URL 완전성 검증
5. preview.py → 프리뷰 HTML → 사용자 확인
6. choices.json 반영 → generate_pptx.py → PPTX 생성
```

---

## Gemini 실행 시 체크리스트 (반드시 모두 수행)

**아래 항목을 하나도 누락 없이 수행한다. 체크리스트 미수행 시 작업 결과는 무효이다.**

1. **프로젝트 구조 확인**: 이 문서의 구조도와 실제 디렉터리 일치 여부
2. **SKILL.md 필독**: 작업 전 `.ai/skills/md-to-pptx/SKILL.md` 전체를 반드시 읽고 숙지
3. **경로 기준**: 마크다운은 `sources/`, 출력은 `output/` 사용
4. **스킬 경로**: `<skill-path>` = `.ai/skills/md-to-pptx`
5. **의존성**: Playwright(캡처), python-pptx, Pillow(PPTX), Python(preview)

---

## 마크다운 슬라이드 형식

```
# 1                    ← 슬라이드 번호 (구분자)
https://example.com    ← URL (자동 캡처)
페이지 설명
내부 설명

# 2
내용 구성만
```

---

*이 문서는 Gemini가 프로젝트를 이해하고 올바르게 실행할 수 있도록 작성되었습니다. **GEMINI.md와 SKILL.md를 반드시 숙지한 후, 하나도 누락 없이 시행한다.***
