"""
통합 프리뷰 HTML 생성기

슬라이드 정보 JSON을 읽어 테마 선택 + 프레젠테이션 뷰어를 하나의 HTML 페이지로 생성한다.
사용자가 선택을 완료하면 choices.json 파일을 프리뷰/PPT와 같은 경로에 저장한다.

Usage:
    python preview.py <slides.json> <output.html> [--images-dir <경로>] [--serve]
"""

import json
import sys
import os
import base64
import http.server
import webbrowser
import threading
from pathlib import Path

THEMES = {
    "dark": {
        "name": "Dark",
        "bg": "#1A1A2E", "text": "#FFFFFF", "subtitle": "#A0A0B8",
        "accent": "#E94560", "secondary": "#16213E", "card_bg": "#16213E",
        "desc": "테크 · 스타트업"
    },
    "light": {
        "name": "Light",
        "bg": "#FFFFFF", "text": "#1A1A1A", "subtitle": "#6B7280",
        "accent": "#2563EB", "secondary": "#F3F4F6", "card_bg": "#F3F4F6",
        "desc": "범용 · 비즈니스"
    },
    "minimal": {
        "name": "Minimal",
        "bg": "#FAFAFA", "text": "#333333", "subtitle": "#888888",
        "accent": "#6366F1", "secondary": "#F0F0F0", "card_bg": "#FFFFFF",
        "desc": "미니멀 · 교육"
    },
    "consulting": {
        "name": "Consulting",
        "bg": "#FFFFFF", "text": "#002F6C", "subtitle": "#4A5568",
        "accent": "#C41230", "secondary": "#F7F8FA", "card_bg": "#F7F8FA",
        "desc": "컨설팅 · 보고서"
    },
    "pitch": {
        "name": "Pitch",
        "bg": "#0F0A1A", "text": "#FFFFFF", "subtitle": "#C4B5D0",
        "accent": "#7B2FBE", "secondary": "#1A1030", "card_bg": "#1A1030",
        "desc": "투자 IR · 피치덱"
    },
    "education": {
        "name": "Education",
        "bg": "#F8F7FF", "text": "#1E1B4B", "subtitle": "#6B6B8D",
        "accent": "#F97316", "secondary": "#EDE9FE", "card_bg": "#FFFFFF",
        "desc": "교육 · 강의"
    },
}

IMAGE_LAYOUTS = {
    "grid-2x2": {"name": "2×2 그리드", "icon": "⊞", "group": "images"},
    "grid-3x1": {"name": "3열 가로", "icon": "⫼", "group": "images"},
    "grid-1x3": {"name": "3행 세로", "icon": "≡", "group": "images"},
    "main-sub": {"name": "메인 + 서브", "icon": "◧", "group": "images"},
    "side-by-side": {"name": "좌우 분할", "icon": "◫", "group": "images"},
    "text-left-img-right": {"name": "텍스트↔이미지", "icon": "◨", "group": "text-image"},
    "text-top-img-bottom": {"name": "텍스트↕이미지", "icon": "⬒", "group": "text-image"},
    "img-left-text-right": {"name": "이미지↔텍스트", "icon": "◧", "group": "text-image"},
    "text-left-imgs-stack": {"name": "텍스트+이미지스택", "icon": "▣", "group": "text-image"},
    "text-img-alternating": {"name": "텍스트·이미지교차", "icon": "▤", "group": "text-image"},
    "text-img-grid-mixed": {"name": "텍스트+이미지그리드", "icon": "▦", "group": "text-image"},
}


def image_to_data_uri(path: str) -> str:
    """이미지를 base64 data URI로 변환한다."""
    if not os.path.exists(path):
        return ""
    ext = Path(path).suffix.lower()
    mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
            "gif": "image/gif", "webp": "image/webp", "svg": "image/svg+xml"}
    mime_type = mime.get(ext.lstrip("."), "image/png")
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    return f"data:{mime_type};base64,{data}"


def generate_preview_html(slides_json_path: str, output_html: str, images_dir: str = ""):
    with open(slides_json_path, "r", encoding="utf-8") as f:
        slides_data = json.load(f)

    slides = slides_data if isinstance(slides_data, list) else slides_data.get("slides", [])
    meta = slides_data.get("meta", {}) if isinstance(slides_data, dict) else {}
    default_theme = meta.get("theme", "light")

    # 경로 정규화 (절대 경로로 변환)
    if images_dir:
        images_dir = os.path.abspath(images_dir)

    # 이미지가 있는 모든 슬라이드의 이미지 수집
    multi_image_slides = {}
    for i, s in enumerate(slides):
        images = s.get("images", [])
        single_image = s.get("image", "")
        all_imgs = images if images else ([single_image] if single_image else [])
        if all_imgs:
            resolved = []
            for img in all_imgs:
                if not os.path.isabs(img) and images_dir:
                    img = os.path.abspath(os.path.join(images_dir, img))
                else:
                    img = os.path.abspath(img)
                resolved.append(img)
            multi_image_slides[i] = resolved

    # 테마 JSON for JS
    themes_js = json.dumps(THEMES, ensure_ascii=False)
    layouts_js = json.dumps(IMAGE_LAYOUTS, ensure_ascii=False)
    slides_js = json.dumps(slides, ensure_ascii=False)

    # 다중 이미지 슬라이드의 data URI 맵
    image_data = {}
    for idx, paths in multi_image_slides.items():
        image_data[str(idx)] = [image_to_data_uri(p) for p in paths]
    image_data_js = json.dumps(image_data, ensure_ascii=False)

    # choices.json 저장 경로 — 프리뷰 HTML과 같은 디렉토리 (절대 경로)
    output_dir = str(Path(output_html).parent.resolve()).replace("\\", "/")
    choices_filename = Path(output_html).stem + ".choices.json"
    choices_full_path = str((Path(output_html).parent / choices_filename).resolve()).replace("\\", "/")

    # 기본 choices.json을 미리 생성 — 사용자가 저장 안 해도 기본값으로 동작
    default_choices = {
        "theme": default_theme,
        "slide_overrides": {}
    }
    with open(choices_full_path, "w", encoding="utf-8") as cf:
        json.dump(default_choices, cf, ensure_ascii=False, indent=2)
    print(f"기본 choices.json 생성: {choices_full_path}")

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PPTX 프리뷰 — 테마 & 구성 확인</title>
<style>
  :root {{
    --bg: #0f0f14;
    --surface: #1a1a24;
    --surface2: #24243a;
    --text: #e8e8f0;
    --text2: #9898b0;
    --accent: #6366f1;
    --accent-hover: #818cf8;
    --border: #2a2a40;
    --success: #10b981;
    --radius: 12px;
  }}
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{
    font-family: 'Pretendard', -apple-system, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.6;
    padding: 2rem;
    max-width: 1200px;
    margin: 0 auto;
  }}
  h1 {{
    font-size: 1.8rem;
    font-weight: 700;
    margin-bottom: 0.5rem;
  }}
  .subtitle {{ color: var(--text2); margin-bottom: 2rem; font-size: 0.95rem; }}

  /* 섹션 제목 */
  .section-title {{
    font-size: 1.2rem;
    font-weight: 600;
    margin: 2rem 0 1rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
  }}
  .section-title .badge {{
    background: var(--accent);
    color: white;
    font-size: 0.7rem;
    padding: 2px 8px;
    border-radius: 20px;
  }}

  /* ─── 테마 선택 ─── */
  .theme-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
    gap: 1rem;
    margin-bottom: 1rem;
  }}
  .theme-card {{
    background: var(--surface);
    border: 2px solid var(--border);
    border-radius: var(--radius);
    padding: 0.8rem;
    cursor: pointer;
    transition: all 0.2s;
  }}
  .theme-card:hover {{ border-color: var(--accent); transform: translateY(-2px); }}
  .theme-card.selected {{ border-color: var(--accent); box-shadow: 0 0 0 3px rgba(99,102,241,0.3); }}
  .theme-preview {{
    width: 100%;
    aspect-ratio: 16/9;
    border-radius: 6px;
    margin-bottom: 0.6rem;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.85rem;
    font-weight: 600;
    position: relative;
    overflow: hidden;
  }}
  .theme-preview .accent-bar {{
    position: absolute;
    bottom: 0; left: 0;
    width: 100%; height: 4px;
  }}
  .theme-name {{ font-weight: 600; font-size: 0.9rem; }}
  .theme-desc {{ color: var(--text2); font-size: 0.75rem; margin-top: 2px; }}
  .theme-check {{
    display: none;
    position: absolute;
    top: 6px; right: 6px;
    background: var(--accent);
    color: white;
    width: 22px; height: 22px;
    border-radius: 50%;
    align-items: center;
    justify-content: center;
    font-size: 13px;
  }}
  .theme-card.selected .theme-check {{ display: flex; }}

  /* ─── 프레젠테이션 구성 (인라인 뷰어) ─── */
  .viewer-container {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
  }}
  .viewer-toolbar {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.6rem 1rem;
    border-top: 1px solid var(--border);
    background: var(--surface2);
  }}
  .viewer-slide_info {{
    font-size: 0.85rem;
    color: var(--text2);
    display: flex;
    align-items: center;
    gap: 0.8rem;
  }}
  .viewer-slide_info .layout-tag {{
    background: rgba(99,102,241,0.15);
    color: var(--accent);
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
  }}
  .viewer-actions {{
    display: flex;
    gap: 0.5rem;
  }}
  .viewer-btn {{
    background: none;
    border: 1px solid var(--border);
    color: var(--text2);
    padding: 0.3rem 0.7rem;
    border-radius: 6px;
    font-size: 0.8rem;
    cursor: pointer;
    transition: all 0.15s;
    display: inline-flex;
    align-items: center;
    gap: 0.3rem;
  }}
  .viewer-btn:hover {{ background: var(--accent); color: white; border-color: var(--accent); }}

  .viewer-stage {{
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 1.5rem;
    background: #000;
    min-height: 300px;
  }}
  .viewer-slide {{
    width: 100%;
    max-width: 960px;
    aspect-ratio: 16 / 9;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    gap: 0.6rem;
    padding: 4%;
    position: relative;
    overflow: hidden;
    box-shadow: 0 4px 30px rgba(0,0,0,0.4);
    transition: background 0.3s;
  }}
  .viewer-slide .s-title {{
    font-weight: 700;
    font-size: clamp(1rem, 2.5vw, 1.8rem);
    text-align: center;
    line-height: 1.3;
  }}
  .viewer-slide .s-subtitle {{
    font-size: clamp(0.7rem, 1.2vw, 1rem);
    text-align: center;
    opacity: 0.7;
  }}
  .viewer-slide .s-body {{
    font-size: clamp(0.65rem, 1vw, 0.85rem);
    text-align: left;
    width: 100%;
    max-height: 70%;
    overflow: auto;
    white-space: pre-wrap;
    line-height: 1.6;
    opacity: 0.9;
  }}
  .viewer-slide .s-code {{
    font-family: 'Consolas', 'Monaco', monospace;
    font-size: clamp(0.55rem, 0.9vw, 0.75rem);
    background: rgba(0,0,0,0.4);
    border-radius: 8px;
    padding: 0.8rem 1rem;
    width: 100%;
    max-height: 70%;
    overflow: auto;
    white-space: pre;
    line-height: 1.5;
    color: #ddd;
  }}

  .viewer-nav {{
    position: absolute;
    top: 50%;
    transform: translateY(-50%);
    background: rgba(0,0,0,0.55);
    backdrop-filter: blur(4px);
    border: 1.5px solid rgba(255,255,255,0.2);
    color: white;
    width: 44px; height: 64px;
    border-radius: 12px;
    font-size: 1.5rem;
    font-weight: 700;
    cursor: pointer;
    opacity: 0.85;
    transition: all 0.2s ease;
    z-index: 10;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
  }}
  .viewer-nav:hover {{ opacity: 1; background: var(--accent); border-color: var(--accent); transform: translateY(-50%) scale(1.08); box-shadow: 0 4px 16px rgba(0,0,0,0.4); }}
  .viewer-nav:active {{ transform: translateY(-50%) scale(0.95); }}
  .viewer-nav.prev {{ left: 0.6rem; }}
  .viewer-nav.next {{ right: 0.6rem; }}

  /* 슬라이드 도트 네비게이션 */
  .viewer-dots {{
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 6px;
    padding: 0.8rem 1rem;
    border-top: 1px solid var(--border);
    flex-wrap: wrap;
  }}
  .viewer-dot {{
    width: 8px; height: 8px;
    border-radius: 50%;
    background: var(--border);
    cursor: pointer;
    transition: all 0.2s;
    border: none;
    padding: 0;
  }}
  .viewer-dot:hover {{ background: var(--text2); }}
  .viewer-dot.active {{ background: var(--accent); width: 20px; border-radius: 4px; }}

  /* 이미지 레이아웃 선택 (다중 이미지 슬라이드) */
  .layout-options {{
    padding: 0.8rem 1rem;
    border-bottom: 1px solid var(--border);
    background: var(--surface2);
  }}
  .layout-label {{
    font-size: 0.8rem;
    color: var(--text2);
    margin-bottom: 0.5rem;
  }}
  .layout-btns {{
    display: flex;
    gap: 0.5rem;
    flex-wrap: wrap;
  }}
  .layout-btn {{
    background: var(--surface);
    border: 2px solid var(--border);
    border-radius: 8px;
    padding: 0.4rem 0.8rem;
    color: var(--text);
    cursor: pointer;
    font-size: 0.85rem;
    transition: all 0.15s;
    display: flex;
    align-items: center;
    gap: 0.4rem;
  }}
  .layout-btn:hover {{ border-color: var(--accent); }}
  .layout-btn.selected {{ border-color: var(--accent); background: rgba(99,102,241,0.15); }}
  .layout-icon {{ font-size: 1.1rem; }}

  /* ─── 확인 버튼 바 ─── */
  .confirm-bar {{
    position: sticky;
    bottom: 1rem;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.5rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
    backdrop-filter: blur(10px);
    box-shadow: 0 -4px 20px rgba(0,0,0,0.3);
    margin-top: 2rem;
  }}
  .confirm-info {{ color: var(--text2); font-size: 0.9rem; }}
  .confirm-path {{ color: var(--text2); font-size: 0.75rem; margin-top: 2px; opacity: 0.6; }}
  .confirm-btn {{
    background: var(--accent);
    color: white;
    border: none;
    padding: 0.7rem 2rem;
    border-radius: 8px;
    font-size: 1rem;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.2s;
  }}
  .confirm-btn:hover {{ background: var(--accent-hover); }}
  .confirm-btn.done {{ background: var(--success); }}

  /* ─── 경로 설정 팝업 (모달) ─── */
  .modal-overlay {{
    display: none;
    position: fixed;
    inset: 0;
    z-index: 100000;
    background: rgba(0,0,0,0.8);
    backdrop-filter: blur(8px);
    align-items: center;
    justify-content: center;
  }}
  .modal-overlay.active {{ display: flex; }}
  .path-modal {{
    background: #1a1a2e;
    border: 2px solid var(--accent);
    border-radius: 16px;
    padding: 2rem;
    max-width: 600px;
    width: 90%;
    box-shadow: 0 20px 50px rgba(0,0,0,0.5);
    animation: modalPop 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275);
  }}
  @keyframes modalPop {{ from {{ transform: scale(0.9); opacity: 0; }} to {{ transform: scale(1); opacity: 1; }} }}
  
  .modal-title {{ font-size: 1.3rem; font-weight: 700; margin-bottom: 1.5rem; color: #fff; display: flex; align-items: center; gap: 10px; }}
  .modal-field {{ margin-bottom: 1.2rem; }}
  .modal-label {{ display: block; font-size: 0.8rem; color: var(--text2); margin-bottom: 0.5rem; text-transform: uppercase; font-weight: 700; }}
  .modal-input {{
    width: 100%;
    background: #020617;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    color: #38bdf8;
    font-family: 'Consolas', monospace;
    font-size: 0.9rem;
    outline: none;
  }}
  .modal-input:focus {{ border-color: var(--accent); box-shadow: 0 0 0 2px rgba(99,102,241,0.2); }}
  
  .modal-actions {{ display: flex; gap: 1rem; margin-top: 2rem; }}
  .modal-btn {{ flex: 1; padding: 0.8rem; border-radius: 8px; font-weight: 700; cursor: pointer; transition: all 0.2s; border: none; }}
  .modal-btn.cancel {{ background: #334155; color: #fff; }}
  .modal-btn.save {{ background: var(--accent); color: white; }}
  .modal-btn:hover {{ opacity: 0.9; transform: translateY(-1px); }}

  /* ─── 풀스크린 프레젠테이션 모드 ─── */
  .pres-overlay {{
    display: none;
    position: fixed;
    inset: 0;
    z-index: 9999;
    background: #000;
    flex-direction: column;
  }}
  .pres-overlay.active {{ display: flex; }}
  .pres-viewport {{
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    position: relative;
  }}
  .pres-slide {{
    width: 90vw;
    max-width: calc(90vh * 16 / 9);
    aspect-ratio: 16 / 9;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    gap: 0.8rem;
    padding: 4%;
    position: relative;
    overflow: hidden;
    box-shadow: 0 4px 40px rgba(0,0,0,0.5);
    transition: background 0.3s;
  }}
  .pres-slide .s-title {{
    font-weight: 700;
    font-size: clamp(1.2rem, 3vw, 2.4rem);
    text-align: center;
    line-height: 1.3;
  }}
  .pres-slide .s-subtitle {{
    font-size: clamp(0.8rem, 1.5vw, 1.2rem);
    text-align: center;
    opacity: 0.7;
  }}
  .pres-slide .s-body {{
    font-size: clamp(0.7rem, 1.2vw, 1rem);
    text-align: left;
    width: 100%;
    max-height: 70%;
    overflow: auto;
    white-space: pre-wrap;
    line-height: 1.6;
    opacity: 0.9;
  }}
  .pres-slide .s-code {{
    font-family: 'Consolas', 'Monaco', monospace;
    font-size: clamp(0.6rem, 1vw, 0.85rem);
    background: rgba(0,0,0,0.4);
    border-radius: 8px;
    padding: 1rem 1.2rem;
    width: 100%;
    max-height: 70%;
    overflow: auto;
    white-space: pre;
    line-height: 1.5;
    color: #ddd;
  }}
  .pres-nav {{
    position: absolute;
    top: 50%;
    transform: translateY(-50%);
    background: rgba(0,0,0,0.5);
    backdrop-filter: blur(4px);
    border: 1.5px solid rgba(255,255,255,0.25);
    color: white;
    width: 52px; height: 76px;
    border-radius: 14px;
    font-size: 1.8rem;
    font-weight: 700;
    cursor: pointer;
    opacity: 0.85;
    transition: all 0.25s ease;
    z-index: 10;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 2px 12px rgba(0,0,0,0.35);
  }}
  .pres-viewport:hover .pres-nav {{ opacity: 1; }}
  .pres-nav:hover {{ background: var(--accent); border-color: var(--accent); transform: translateY(-50%) scale(1.1); box-shadow: 0 4px 20px rgba(0,0,0,0.5); }}
  .pres-nav:active {{ transform: translateY(-50%) scale(0.95); }}
  .pres-nav.prev {{ left: 1.2rem; }}
  .pres-nav.next {{ right: 1.2rem; }}
  .pres-progress {{
    height: 3px;
    background: rgba(255,255,255,0.1);
    flex-shrink: 0;
  }}
  .pres-progress-bar {{
    height: 100%;
    background: var(--accent);
    transition: width 0.3s ease;
  }}
  .pres-footer {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 0.4rem 1.2rem;
    color: rgba(255,255,255,0.4);
    font-size: 0.75rem;
    flex-shrink: 0;
  }}
  .pres-close {{
    background: none;
    border: 1px solid rgba(255,255,255,0.2);
    color: rgba(255,255,255,0.6);
    padding: 0.2rem 0.8rem;
    border-radius: 4px;
    cursor: pointer;
    font-size: 0.75rem;
  }}
  .pres-close:hover {{ background: rgba(255,255,255,0.1); color: white; }}
</style>
</head>
<body>

<!-- 절대 경로 배너 (필수 준수 사항) -->
<div style="background: #2563EB; color: #FFFFFF; padding: 1.5rem; border-radius: var(--radius); margin-bottom: 2rem; border-left: 8px solid #1D4ED8; box-shadow: 0 4px 12px rgba(0,0,0,0.2);">
  <h2 style="margin: 0 0 0.5rem 0; font-size: 1.1rem; display: flex; align-items: center; gap: 0.5rem;">
    <span style="font-size: 1.4rem;">📂</span> [필독] 결과물 저장 경로 안내
  </h2>
  <div style="font-family: 'Consolas', monospace; font-size: 0.95rem; background: rgba(0,0,0,0.2); padding: 0.8rem; border-radius: 6px; word-break: break-all;">
    <div style="margin-bottom: 0.4rem;"><span style="color: #93C5FD; font-weight: bold;">출력 디렉토리:</span> {os.path.abspath(output_dir)}</div>
    <div><span style="color: #93C5FD; font-weight: bold;">설정 파일(choices.json):</span> {os.path.abspath(choices_full_path)}</div>
  </div>
  <p style="margin-top: 0.8rem; font-size: 0.85rem; opacity: 0.9;">
    ※ 위 경로는 현재 프로젝트의 작업 환경을 기준으로 결정되었습니다. <strong>choices.json</strong>이 해당 경로에 존재해야 PPTX가 올바르게 생성됩니다.
  </p>
</div>

<h1>PPTX 프리뷰</h1>
<p class="subtitle">테마를 선택하고 슬라이드를 확인한 뒤 "확인"을 눌러주세요.</p>

<!-- 테마 선택 -->
<div class="section-title">테마 선택 <span class="badge">필수</span></div>
<div class="theme-grid" id="themeGrid"></div>

<!-- 프레젠테이션 구성 -->
<div class="section-title">
  프레젠테이션 구성 <span class="badge" id="slideCount"></span>
</div>
<div class="viewer-container">
  <div id="layoutOptions" style="display:none"></div>
  <div class="viewer-stage" id="viewerStage">
    <button class="viewer-nav prev" onclick="viewerNav(-1)">&#8249;</button>
    <div class="viewer-slide" id="viewerSlide"></div>
    <button class="viewer-nav next" onclick="viewerNav(1)">&#8250;</button>
  </div>
  <div class="viewer-dots" id="viewerDots"></div>
  <div class="viewer-toolbar">
    <div class="viewer-slide_info">
      <span id="viewerSlideNum">1 / 1</span>
      <span class="layout-tag" id="viewerLayoutTag">title</span>
    </div>
    <div class="viewer-actions">
      <button class="viewer-btn" onclick="openPresentation()" title="전체 화면 (F5)">&#9654; 전체 화면</button>
    </div>
  </div>
</div>

<!-- 확인 바 -->
<div class="confirm-bar">
  <div>
    <div class="confirm-info" id="confirmInfo">테마를 선택해주세요</div>
    <div class="confirm-path" id="confirmPath"></div>
  </div>
  <button class="confirm-btn" id="confirmBtn" onclick="openPathModal()">확인 — 선택 저장</button>
</div>

<!-- 경로 설정 팝업 -->
<div class="modal-overlay" id="pathModal">
  <div class="path-modal">
    <div class="modal-title"><span>📍</span> 결과물 출력 및 설정 저장 경로</div>
    <div class="modal-field">
      <label class="modal-label">결과물 출력 디렉토리 (CWD 기반)</label>
      <input type="text" class="modal-input" id="inputOutputDir">
    </div>
    <div class="modal-field">
      <label class="modal-label">설정 저장 파일명 (choices.json)</label>
      <input type="text" class="modal-input" id="inputChoicesFile">
    </div>
    <div style="font-size: 0.8rem; color: #94a3b8; margin-top: 1rem; border-left: 3px solid var(--accent); padding-left: 10px;">
      위 경로는 시스템 설정에서 결정된 기본값입니다. 필요 시 수정할 수 있습니다.
    </div>
    <div class="modal-actions">
      <button class="modal-btn cancel" onclick="closePathModal()">취소</button>
      <button class="modal-btn save" onclick="saveChoices()">저장 및 확인</button>
    </div>
  </div>
</div>

<!-- 풀스크린 프레젠테이션 오버레이 -->
<div class="pres-overlay" id="presOverlay">
  <div class="pres-viewport" id="presViewport">
    <button class="pres-nav prev" onclick="presNav(-1)">&#8249;</button>
    <div class="pres-slide" id="presSlide"></div>
    <button class="pres-nav next" onclick="presNav(1)">&#8250;</button>
  </div>
  <div class="pres-progress"><div class="pres-progress-bar" id="presProgressBar"></div></div>
  <div class="pres-footer">
    <span id="presInfo"></span>
    <span>&#8592;&#8594; / Space / 스와이프 | Esc 닫기</span>
    <button class="pres-close" onclick="closePresentation()">ESC 닫기</button>
  </div>
</div>

<script>
const THEMES = {themes_js};
const LAYOUTS = {layouts_js};
const SLIDES = {slides_js};
const IMAGE_DATA = {image_data_js};
const CHOICES_FILENAME = {json.dumps(choices_filename)};
const CHOICES_DIR = {json.dumps(output_dir)};

let selectedTheme = {json.dumps(default_theme)};
let slideOverrides = {{}};
let viewerIdx = 0;

// ─── 경로 설정 팝업 제어 ───
function openPathModal() {{
  document.getElementById('inputOutputDir').value = CHOICES_DIR;
  document.getElementById('inputChoicesFile').value = CHOICES_FILENAME;
  document.getElementById('pathModal').classList.add('active');
}}

function closePathModal() {{
  document.getElementById('pathModal').classList.remove('active');
}}

// ─── 테마 카드 ───
function renderThemes() {{
  const grid = document.getElementById('themeGrid');
  grid.innerHTML = '';
  for (const [key, t] of Object.entries(THEMES)) {{
    const card = document.createElement('div');
    card.className = 'theme-card' + (key === selectedTheme ? ' selected' : '');
    card.onclick = () => selectTheme(key);
    card.innerHTML = `
      <div class="theme-preview" style="background:${{t.bg}};color:${{t.text}}">
        <span class="theme-check">✓</span>
        <span>Aa 가나다</span>
        <div class="accent-bar" style="background:${{t.accent}}"></div>
      </div>
      <div class="theme-name">${{t.name}}</div>
      <div class="theme-desc">${{t.desc}}</div>
    `;
    grid.appendChild(card);
  }}
}}

function selectTheme(key) {{
  selectedTheme = key;
  renderThemes();
  renderViewerSlide();
  updateConfirmInfo();
}}

// ─── 이미지 렌더링 헬퍼 ───
function getSlideImages(slideIdx) {{
  return IMAGE_DATA[String(slideIdx)] || [];
}}

function renderImgTag(src, caption, style) {{
  if (!src) return `<div style="${{style}};background:#333;display:flex;align-items:center;justify-content:center;color:#888;font-size:0.7rem;border-radius:6px">이미지 없음</div>`;
  return `<div style="${{style}};position:relative;overflow:hidden;border-radius:6px">
    <img src="${{src}}" style="width:100%;height:100%;object-fit:contain;display:block" />
    ${{caption ? `<div style="position:absolute;bottom:0;left:0;right:0;background:linear-gradient(transparent,rgba(0,0,0,0.7));color:#fff;font-size:0.55rem;padding:0.2rem 0.4rem">${{caption}}</div>` : ''}}
  </div>`;
}}

function renderImageGrid(imgs, captions, layoutKey, t) {{
  captions = captions || [];
  const n = imgs.length;
  if (n === 0) return '';

  if (n === 1) {{
    return `<div style="width:100%;height:60%;display:flex;justify-content:center">${{renderImgTag(imgs[0], captions[0], 'width:80%;height:100%')}}</div>`;
  }}

  switch (layoutKey) {{
    case 'grid-3x1':
      return `<div style="display:grid;grid-template-columns:repeat(${{Math.min(n,3)}},1fr);gap:0.4rem;width:100%;height:55%">
        ${{imgs.slice(0,3).map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
      </div>${{n > 3 ? `<div style="display:grid;grid-template-columns:repeat(${{n-3}},1fr);gap:0.4rem;width:100%;height:25%;margin-top:0.4rem">${{imgs.slice(3).map((img,i) => renderImgTag(img, captions[i+3], 'width:100%;height:100%')).join('')}}</div>` : ''}}`;
    case 'grid-1x3':
      return `<div style="display:grid;grid-template-rows:repeat(${{Math.min(n,3)}},1fr);gap:0.4rem;width:60%;height:70%;margin:0 auto">
        ${{imgs.slice(0,3).map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
      </div>`;
    case 'main-sub':
      return `<div style="display:grid;grid-template-columns:2fr 1fr;grid-template-rows:1fr 1fr;gap:0.4rem;width:100%;height:65%">
        <div style="grid-row:1/3">${{renderImgTag(imgs[0], captions[0], 'width:100%;height:100%')}}</div>
        ${{imgs.slice(1,3).map((img,i) => renderImgTag(img, captions[i+1], 'width:100%;height:100%')).join('')}}
        ${{n > 3 ? `</div><div style="display:grid;grid-template-columns:repeat(${{n-3}},1fr);gap:0.4rem;width:100%;height:20%;margin-top:0.4rem">${{imgs.slice(3).map((img,i) => renderImgTag(img, captions[i+3], 'width:100%;height:100%')).join('')}}</div>` : '</div>'}}`;
    case 'side-by-side':
      return `<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:0.4rem;width:100%;height:65%">
        ${{imgs.slice(0,2).map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
      </div>${{n > 2 ? `<div style="display:grid;grid-template-columns:repeat(${{Math.min(n-2,3)}},1fr);gap:0.4rem;width:100%;height:25%;margin-top:0.4rem">${{imgs.slice(2).map((img,i) => renderImgTag(img, captions[i+2], 'width:100%;height:100%')).join('')}}</div>` : ''}}`;
    case 'grid-2x2':
    default:
      const cols = n <= 2 ? n : 2;
      return `<div style="display:grid;grid-template-columns:repeat(${{cols}},1fr);gap:0.4rem;width:100%;height:65%">
        ${{imgs.map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
      </div>`;
  }}
}}

function getEffectiveLayout(slideIdx) {{
  if (slideOverrides[slideIdx]?.layout) return slideOverrides[slideIdx].layout;
  const s = SLIDES[slideIdx];
  const gridMap = {{'2x2':'grid-2x2','3x1':'grid-3x1','1x3':'grid-1x3','2x1':'side-by-side'}};
  return gridMap[s.grid] || 'grid-2x2';
}}

// ─── body_elements → HTML 렌더링 ───
function renderBodyElements(elements, t, fontSize) {{
  if (!elements || !elements.length) return '';
  fontSize = fontSize || 'clamp(0.6rem,0.9vw,0.8rem)';
  let html = '';
  elements.forEach(el => {{
    switch (el.type) {{
      case 'heading':
        if (el.level === 2) {{
          html += `<div style="font-size:clamp(0.75rem,1.3vw,1rem);font-weight:700;color:${{t.text}};margin:0.5rem 0 0.3rem;padding-bottom:0.2rem;border-bottom:2px solid ${{t.accent}}">${{el.text}}</div>`;
        }} else {{
          html += `<div style="font-size:clamp(0.7rem,1.1vw,0.9rem);font-weight:700;color:${{t.text}};margin:0.4rem 0 0.2rem">${{el.text}}</div>`;
        }}
        break;
      case 'paragraph':
        html += `<div style="font-size:${{fontSize}};color:${{t.text}};line-height:1.6;margin:0.2rem 0">${{el.text}}</div>`;
        break;
      case 'bullet_list':
        html += '<div style="margin:0.2rem 0">';
        (el.items || []).forEach(item => {{
          const text = typeof item === 'string' ? item : item.text;
          const level = typeof item === 'object' ? (item.level || 0) : 0;
          const indent = level * 1.2;
          const bullets = ['●','○','■','▸'];
          html += `<div style="font-size:${{fontSize}};color:${{t.text}};line-height:1.8;padding-left:${{indent + 1}}rem;position:relative">
            <span style="position:absolute;left:${{indent}}rem;color:${{t.accent}};font-size:0.5em;top:0.35em">${{bullets[level] || '●'}}</span>${{text}}
          </div>`;
        }});
        html += '</div>';
        break;
      case 'numbered_list':
        html += '<div style="margin:0.2rem 0">';
        (el.items || []).forEach((item, i) => {{
          const text = typeof item === 'string' ? item : item.text;
          html += `<div style="font-size:${{fontSize}};color:${{t.text}};line-height:1.8;padding-left:1.5rem;position:relative">
            <span style="position:absolute;left:0;color:${{t.accent}};font-weight:700">${{i+1}}.</span>${{text}}
          </div>`;
        }});
        html += '</div>';
        break;
      case 'blockquote':
        html += `<div style="margin:0.3rem 0;padding:0.4rem 0.8rem;background:${{t.secondary}};border-left:3px solid ${{t.accent}};border-radius:0 6px 6px 0;font-style:italic;font-size:${{fontSize}};color:${{t.text}};line-height:1.6">${{el.text}}</div>`;
        break;
      case 'divider':
        html += `<hr style="border:none;border-top:1px solid ${{t.subtitle}}40;margin:0.4rem 0" />`;
        break;
      case 'code_block':
        const lang = el.language || '';
        html += `<div style="margin:0.3rem 0;background:#1e1e2e;border-radius:6px;padding:0.5rem 0.7rem;position:relative">
          ${{lang ? `<span style="position:absolute;top:0.3rem;right:0.5rem;font-size:0.55rem;color:#888">${{lang}}</span>` : ''}}
          <pre style="margin:0;font-family:Consolas,'Courier New',monospace;font-size:clamp(0.5rem,0.7vw,0.7rem);color:#e0e0e0;white-space:pre-wrap;line-height:1.5">${{(el.code||'').replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</pre>
        </div>`;
        break;
      case 'inline_table':
        html += '<div style="margin:0.3rem 0;overflow:auto">';
        html += `<table style="width:100%;border-collapse:collapse;font-size:clamp(0.5rem,0.75vw,0.7rem)">`;
        html += '<thead><tr>';
        (el.headers || []).forEach(h => {{
          html += `<th style="padding:0.3rem 0.5rem;text-align:left;border-bottom:2px solid ${{t.accent}};color:${{t.text}};font-weight:600">${{h}}</th>`;
        }});
        html += '</tr></thead><tbody>';
        (el.rows || []).forEach((row, ri) => {{
          const bg = ri % 2 === 1 ? t.secondary : 'transparent';
          html += `<tr style="background:${{bg}}">`;
          row.forEach(cell => {{
            html += `<td style="padding:0.25rem 0.5rem;border-bottom:1px solid ${{t.secondary}};color:${{t.text}}">${{cell}}</td>`;
          }});
          html += '</tr>';
        }});
        html += '</tbody></table></div>';
        break;
    }}
  }});
  return html;
}}

// ─── YouTube 헬퍼 ───
function extractYouTubeId(url) {{
  if (!url) return null;
  const m = url.match(/(?:youtube\\.com\\/watch\\?v=|youtu\\.be\\/|youtube\\.com\\/embed\\/)([^&?/]+)/);
  return m ? m[1] : null;
}}

// ─── 슬라이드 렌더링 (공통) ───
function renderSlideContent(s, t, containerClass, slideIdx) {{
  const layout = s.layout || 'content';
  const title = s.title || '';
  const subtitle = s.subtitle || '';
  const body = s.body_elements ? renderBodyElements(s.body_elements, t) : (s.body || '');
  const code = s.code || '';
  const imgs = getSlideImages(slideIdx);
  const captions = s.captions || [];

  let html = '';

  // ─── YouTube 영상 슬라이드 ───
  if (s.video_url) {{
    const videoId = extractYouTubeId(s.video_url);
    if (videoId) {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        ${{body ? `<div style="flex:1;overflow:auto"><div style="color:${{t.text}}">${{body}}</div></div>` : ''}}
        <div style="flex:${{body ? '1.2' : '2'}};min-height:0">
          <iframe src="https://www.youtube.com/embed/${{videoId}}?rel=0" style="width:100%;height:100%;border:none;border-radius:8px" allowfullscreen></iframe>
        </div>
      </div>`;
      return html;
    }}
  }}

  // content-image: 좌측 텍스트 + 우측 이미지
  if (layout === 'content-image' && imgs.length > 0) {{
    if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
    html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
      <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}}">${{body}}</div></div>
      <div style="flex:1">${{renderImgTag(imgs[0], captions[0], 'width:100%;height:100%')}}</div>
    </div>`;
    return html;
  }}

  // image-full: 풀블리드 이미지
  if (layout === 'image-full' && imgs.length > 0) {{
    html += `<img src="${{imgs[0]}}" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain" />`;
    if (title) html += `<div style="position:relative;z-index:1;background:rgba(0,0,0,0.5);padding:0.5rem 1rem;border-radius:6px"><div class="s-title" style="color:#fff">${{title}}</div></div>`;
    return html;
  }}

  // two-images: 텍스트 + 좌우 이미지 (레이아웃 오버라이드가 없을 때만)
  if (layout === 'two-images' && imgs.length >= 2 && !slideOverrides[slideIdx]?.layout) {{
    if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
    if (body) {{
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
        <div style="flex:1.2;display:grid;grid-template-columns:1fr 1fr;gap:0.4rem">
          ${{imgs.slice(0,2).map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
        </div>
      </div>`;
    }} else {{
      html += `<div style="display:grid;grid-template-columns:1fr 1fr;gap:0.6rem;width:100%;height:60%">
        ${{imgs.slice(0,2).map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
      </div>`;
    }}
    return html;
  }}

  // grid-images / two-images(오버라이드 있을 때): 다중 이미지 그리드 (레이아웃 선택 적용)
  if ((layout === 'grid-images' || layout === 'two-images' || (s.images && s.images.length >= 2)) && imgs.length >= 2) {{
    const effectiveLayout = getEffectiveLayout(slideIdx);

    // 텍스트+이미지 복합 레이아웃
    if (effectiveLayout === 'text-left-img-right') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
        <div style="flex:1;display:grid;grid-template-columns:repeat(${{Math.min(imgs.length,2)}},1fr);gap:0.3rem">
          ${{imgs.map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
        </div>
      </div>`;
      return html;
    }}

    if (effectiveLayout === 'img-left-text-right') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        <div style="flex:1;display:grid;grid-template-columns:repeat(${{Math.min(imgs.length,2)}},1fr);gap:0.3rem">
          ${{imgs.map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
        </div>
        <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
      </div>`;
      return html;
    }}

    if (effectiveLayout === 'text-top-img-bottom') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;flex-direction:column;gap:0.5rem;width:100%;flex:1;min-height:0">
        <div style="flex:0 0 auto;max-height:35%;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
        <div style="flex:1;display:grid;grid-template-columns:repeat(${{Math.min(imgs.length,3)}},1fr);gap:0.3rem;min-height:0">
          ${{imgs.map((img,i) => renderImgTag(img, captions[i], 'width:100%;height:100%')).join('')}}
        </div>
      </div>`;
      return html;
    }}

    // 레이아웃 1) 텍스트(좌) + 이미지 세로 스택(우)
    if (effectiveLayout === 'text-left-imgs-stack') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
        <div style="flex:1;display:flex;flex-direction:column;gap:0.3rem">
          ${{imgs.map((img,i) => renderImgTag(img, captions[i], 'width:100%;flex:1;min-height:0')).join('')}}
        </div>
      </div>`;
      return html;
    }}

    // 레이아웃 2) 텍스트·이미지 교차 배치 — 좌우 2열, 텍스트/이미지 교차
    if (effectiveLayout === 'text-img-alternating') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      const bodyLines = body ? body.split('\\n').filter(l => l.trim()) : [];
      const linesPerBlock = Math.max(1, Math.ceil(bodyLines.length / imgs.length));
      let altHtml = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:0.5rem;width:100%;flex:1;min-height:0;overflow:auto;align-content:start">';
      imgs.forEach((img, i) => {{
        const blockLines = bodyLines.slice(i * linesPerBlock, (i + 1) * linesPerBlock).join('\\n');
        altHtml += `<div style="overflow:auto;display:flex;align-items:center"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.55rem,0.85vw,0.75rem);white-space:pre-wrap;line-height:1.5">${{blockLines || ''}}</div></div>`;
        altHtml += `<div style="min-height:120px">${{renderImgTag(img, captions[i], 'width:100%;height:100%')}}</div>`;
      }});
      const remainLines = bodyLines.slice(imgs.length * linesPerBlock).join('\\n');
      if (remainLines) altHtml += `<div style="grid-column:1/-1"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.55rem,0.85vw,0.75rem);white-space:pre-wrap;line-height:1.5">${{remainLines}}</div></div>`;
      altHtml += '</div>';
      html += altHtml;
      return html;
    }}

    // 레이아웃 3) 텍스트+첫 이미지 상단, 나머지 이미지 그리드 하단
    if (effectiveLayout === 'text-img-grid-mixed') {{
      if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
      html += `<div style="display:flex;flex-direction:column;gap:0.4rem;width:100%;flex:1;min-height:0">
        <div style="display:flex;gap:0.6rem;flex:0 0 50%;min-height:0">
          <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.55rem,0.85vw,0.75rem);white-space:pre-wrap;line-height:1.5">${{body}}</div></div>
          <div style="flex:0 0 40%">${{renderImgTag(imgs[0], captions[0], 'width:100%;height:100%')}}</div>
        </div>
        ${{imgs.length > 1 ? `<div style="display:grid;grid-template-columns:repeat(${{Math.min(imgs.length - 1, 4)}},1fr);gap:0.3rem;flex:1;min-height:0">
          ${{imgs.slice(1).map((img,i) => renderImgTag(img, captions[i+1], 'width:100%;height:100%')).join('')}}
        </div>` : ''}}
      </div>`;
      return html;
    }}

    // 이미지 레이아웃 (기본) — body가 있으면 텍스트도 표시
    if (title) html += `<div class="s-title" style="color:${{t.text}};font-size:clamp(0.8rem,1.8vw,1.2rem);width:100%;text-align:left">${{title}}</div>`;
    if (body) {{
      html += `<div style="display:flex;gap:0.8rem;width:100%;flex:1;min-height:0">
        <div style="flex:1;overflow:auto"><div class="s-body" style="color:${{t.text}};font-size:clamp(0.6rem,0.9vw,0.8rem);white-space:pre-wrap;line-height:1.6">${{body}}</div></div>
        <div style="flex:1.2">${{renderImageGrid(imgs, captions, effectiveLayout, t)}}</div>
      </div>`;
    }} else {{
      html += renderImageGrid(imgs, captions, effectiveLayout, t);
    }}
    return html;
  }}

  // 단일 이미지가 있는 content 슬라이드
  if (imgs.length === 1 && layout !== 'title' && layout !== 'section' && layout !== 'closing') {{
    if (title) html += `<div class="s-title" style="color:${{t.text}};width:100%;text-align:left">${{title}}</div>`;
    if (body) html += `<div class="s-body" style="color:${{t.text}};margin-bottom:0.4rem">${{body}}</div>`;
    html += `<div style="width:80%;height:45%;margin:0 auto">${{renderImgTag(imgs[0], captions[0], 'width:100%;height:100%')}}</div>`;
    return html;
  }}

  // 기본 텍스트 기반 렌더링
  if (title) html += `<div class="s-title" style="color:${{t.text}}">${{title}}</div>`;
  if (subtitle) html += `<div class="s-subtitle" style="color:${{t.subtitle}}">${{subtitle}}</div>`;

  if (code) {{
    html += `<div class="s-code">${{code.replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</div>`;
  }} else if (s.metrics) {{
    html += '<div style="display:flex;gap:0.8rem;width:100%;justify-content:center;flex-wrap:wrap">';
    s.metrics.forEach(m => {{
      html += `<div style="background:${{t.card_bg}};border-radius:12px;padding:1rem 1.2rem;text-align:center;flex:1;min-width:100px">
        <div style="font-size:clamp(1.2rem,2.5vw,2rem);font-weight:700;color:${{t.accent}}">${{m.value}}</div>
        <div style="font-size:0.8rem;color:${{t.text}};margin-top:0.3rem">${{m.label}}</div>
        ${{m.change ? `<div style="font-size:0.7rem;color:${{t.subtitle}};margin-top:0.2rem">${{m.change}}</div>` : ''}}
      </div>`;
    }});
    html += '</div>';
  }} else if (s.left && s.right) {{
    const renderSide = (d) => {{
      const data = typeof d === 'string' ? {{body: d}} : d;
      return `<div style="background:${{t.card_bg}};border-radius:12px;padding:0.8rem 1rem;flex:1">
        ${{data.title ? `<div style="font-weight:700;color:${{t.accent}};margin-bottom:0.4rem;font-size:clamp(0.8rem,1.3vw,1rem)">${{data.title}}</div>` : ''}}
        <div style="font-size:clamp(0.65rem,0.9vw,0.8rem);color:${{t.text}};white-space:pre-wrap;line-height:1.5">${{data.body || ''}}</div>
      </div>`;
    }};
    html += `<div style="display:flex;gap:0.8rem;width:100%">${{renderSide(s.left)}}${{renderSide(s.right)}}</div>`;
  }} else if (s.headers && s.rows) {{
    html += '<div style="width:100%;overflow:auto;max-height:65%">';
    html += `<table style="width:100%;border-collapse:collapse;font-size:clamp(0.6rem,0.9vw,0.8rem)">`;
    html += '<thead><tr>';
    s.headers.forEach(h => {{
      html += `<th style="padding:0.4rem 0.6rem;text-align:left;border-bottom:2px solid ${{t.accent}};color:${{t.text}};font-weight:600">${{h}}</th>`;
    }});
    html += '</tr></thead><tbody>';
    s.rows.forEach(row => {{
      html += '<tr>';
      row.forEach(cell => {{
        html += `<td style="padding:0.3rem 0.6rem;border-bottom:1px solid ${{t.secondary}};color:${{t.text}}">${{cell}}</td>`;
      }});
      html += '</tr>';
    }});
    html += '</tbody></table></div>';
  }} else if (s.events) {{
    html += '<div style="display:flex;gap:0.4rem;width:100%;align-items:flex-start;flex-wrap:wrap;justify-content:center">';
    s.events.forEach((e, i) => {{
      html += `<div style="text-align:center;flex:1;min-width:80px;max-width:180px">
        <div style="width:10px;height:10px;border-radius:50%;background:${{t.accent}};margin:0 auto 0.3rem"></div>
        <div style="font-weight:700;color:${{t.accent}};font-size:clamp(0.65rem,0.9vw,0.8rem)">${{e.date}}</div>
        <div style="font-size:clamp(0.55rem,0.8vw,0.7rem);color:${{t.text}};margin-top:0.2rem">${{e.description || e.title || ''}}</div>
      </div>`;
      if (i < s.events.length - 1) html += `<div style="flex:0 0 auto;margin-top:4px;color:${{t.accent}}">—</div>`;
    }});
    html += '</div>';
  }} else if (body) {{
    html += `<div class="s-body" style="color:${{t.text}}">${{body}}</div>`;
  }}

  return html;
}}

// ─── 인라인 뷰어 ───
function renderViewerSlide() {{
  const s = SLIDES[viewerIdx];
  const t = THEMES[selectedTheme];
  const slide = document.getElementById('viewerSlide');
  const layout = s.layout || 'content';

  slide.style.background = t.bg;
  slide.innerHTML = renderSlideContent(s, t, null, viewerIdx);

  // 슬라이드 정보
  document.getElementById('viewerSlideNum').textContent = `${{viewerIdx + 1}} / ${{SLIDES.length}}`;
  document.getElementById('viewerLayoutTag').textContent = layout;
  document.getElementById('slideCount').textContent = SLIDES.length + '장';

  // 도트 네비게이션
  renderDots();

  // 다중 이미지 레이아웃 옵션
  renderLayoutOptions();
}}

function renderDots() {{
  const container = document.getElementById('viewerDots');
  container.innerHTML = '';
  SLIDES.forEach((_, i) => {{
    const dot = document.createElement('button');
    dot.className = 'viewer-dot' + (i === viewerIdx ? ' active' : '');
    dot.onclick = () => {{ viewerIdx = i; renderViewerSlide(); }};
    container.appendChild(dot);
  }});
}}

function renderLayoutOptions() {{
  const container = document.getElementById('layoutOptions');
  const s = SLIDES[viewerIdx];
  const images = s?.images || [];
  const hasImages = images.length >= 2 || (getSlideImages(viewerIdx).length >= 2);
  if (!hasImages) {{
    container.style.display = 'none';
    return;
  }}

  container.style.display = 'block';
  const currentLayout = slideOverrides[viewerIdx]?.layout || getEffectiveLayout(viewerIdx);
  const hasText = !!(s.body || (s.captions && s.captions.some(c => c)));
  const imgCount = getSlideImages(viewerIdx).length;

  if (hasText) {{
    // 텍스트가 있는 경우: 복합 레이아웃만 표시
    let textImgBtns = '';
    for (const [lk, lv] of Object.entries(LAYOUTS)) {{
      if (lv.group !== 'text-image') continue;
      const sel = lk === currentLayout ? ' selected' : '';
      textImgBtns += `<button class="layout-btn${{sel}}" onclick="selectLayout(${{viewerIdx}}, '${{lk}}')">
        <span class="layout-icon">${{lv.icon}}</span> ${{lv.name}}
      </button>`;
    }}
    container.innerHTML = `
      <div class="layout-options">
        <div class="layout-label">텍스트 + 이미지 (${{imgCount}}개) 복합 레이아웃 선택:</div>
        <div class="layout-btns">${{textImgBtns}}</div>
      </div>
    `;
  }} else {{
    // 텍스트가 없는 경우: 이미지 전용 레이아웃만 표시
    let imgBtns = '';
    for (const [lk, lv] of Object.entries(LAYOUTS)) {{
      if (lv.group !== 'images') continue;
      const sel = lk === currentLayout ? ' selected' : '';
      imgBtns += `<button class="layout-btn${{sel}}" onclick="selectLayout(${{viewerIdx}}, '${{lk}}')">
        <span class="layout-icon">${{lv.icon}}</span> ${{lv.name}}
      </button>`;
    }}
    container.innerHTML = `
      <div class="layout-options">
        <div class="layout-label">이미지 전용 레이아웃 선택 (${{imgCount}}개):</div>
        <div class="layout-btns">${{imgBtns}}</div>
      </div>
    `;
  }}
}}

function viewerNav(dir) {{
  viewerIdx = Math.max(0, Math.min(SLIDES.length - 1, viewerIdx + dir));
  renderViewerSlide();
}}

function selectLayout(slideIdx, layoutKey) {{
  if (!slideOverrides[slideIdx]) slideOverrides[slideIdx] = {{}};
  slideOverrides[slideIdx].layout = layoutKey;
  renderViewerSlide();
  updateConfirmInfo();
}}

function updateConfirmInfo() {{
  const info = document.getElementById('confirmInfo');
  const pathEl = document.getElementById('confirmPath');
  const overrideCount = Object.keys(slideOverrides).length;
  const fullPath = CHOICES_DIR + '/' + CHOICES_FILENAME;

  info.textContent = `테마: ${{THEMES[selectedTheme].name}} | 슬라이드: ${{SLIDES.length}}장` +
    (overrideCount > 0 ? ` | 레이아웃 커스텀: ${{overrideCount}}개` : '');

  pathEl.innerHTML = `<span style="color:#F97316;font-weight:700">저장 예정 경로:</span> <span style="font-family:monospace;background:rgba(255,255,255,0.05);padding:2px 4px;border-radius:4px">${{fullPath}}</span>`;
  pathEl.style.fontSize = '0.85rem';
  pathEl.style.marginTop = '4px';
}}

// ─── 저장 ───
async function saveChoices() {{
  const choices = {{
    theme: selectedTheme,
    slide_overrides: {{}}
  }};
  for (const [k, v] of Object.entries(slideOverrides)) {{
    choices.slide_overrides[k] = v;
  }}
  
  const customDir = document.getElementById('inputOutputDir').value;
  const customFile = document.getElementById('inputChoicesFile').value;
  const fullPath = customDir + '/' + customFile;
  
  const payload = {{
    choices: choices,
    path: fullPath
  }};
  
  const json = JSON.stringify(payload, null, 2);

  // 1) 로컬 서버 POST (--serve 모드)
  if (location.protocol !== 'file:') {{
    try {{
      const res = await fetch('/save-choices', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: json
      }});
      if (res.ok) {{
        closePathModal();
        showSaveSuccess('지정된 경로에 저장 완료!');
        return;
      }}
    }} catch (e) {{
      // 서버 없으면 fallback
    }}
  }}

  // 2) File System Access API (file:// 모드에서 직접 파일 저장)
  if (window.showSaveFilePicker) {{
    try {{
      const handle = await window.showSaveFilePicker({{
        suggestedName: customFile,
        types: [{{
          description: 'JSON',
          accept: {{ 'application/json': ['.json'] }}
        }}]
      }});
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(choices, null, 2));
      await writable.close();
      closePathModal();
      showSaveSuccess('파일 저장 완료!');
      return;
    }} catch (e) {{
      if (e.name === 'AbortError') return; // 사용자가 취소
      // API 실패 시 fallback
    }}
  }}

  // 3) Fallback: Blob 다운로드 + 경로 안내
  const blob = new Blob([JSON.stringify(choices, null, 2)], {{ type: 'application/json' }});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = customFile;
  a.click();
  URL.revokeObjectURL(url);

  closePathModal();
  showSavePathWarning(fullPath);
}}

function showSavePathWarning(targetPath) {{
  // 기존 제거
  const existing = document.getElementById('savePathWarning');
  if (existing) existing.remove();

  // 풀스크린 모달 오버레이 — 사용자가 반드시 확인해야 진행 가능
  const overlay = document.createElement('div');
  overlay.id = 'savePathWarning';
  overlay.style.cssText = 'position:fixed;inset:0;z-index:999999;background:rgba(0,0,0,0.85);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px)';
  overlay.innerHTML = `
    <div style="background:#1a1a2e;border:2px solid #F97316;border-radius:16px;padding:2.5rem;max-width:640px;width:90%;color:#fff;box-shadow:0 0 60px rgba(249,115,22,0.3);animation:pulseGlow 2s ease-in-out infinite">
      <div style="text-align:center;margin-bottom:1.5rem">
        <div style="font-size:3rem;margin-bottom:0.5rem">⚠️</div>
        <div style="font-size:1.4rem;font-weight:800;color:#F97316;letter-spacing:0.5px">자동 저장 불가 (직접 이동 필요)</div>
      </div>
      <div style="background:rgba(249,115,22,0.1);border:1px solid rgba(249,115,22,0.3);border-radius:10px;padding:1.2rem;margin-bottom:1.2rem">
        <div style="font-size:0.85rem;color:#FDBA74;margin-bottom:0.6rem;font-weight:600">다운로드된 파일을 반드시 아래 경로로 이동시키세요:</div>
        <div style="background:#000;border-radius:8px;padding:0.8rem 1rem;font-family:Consolas,Monaco,monospace;font-size:0.9rem;word-break:break-all;user-select:all;cursor:text;color:#4ADE80;border:1px solid #333">
          ${{targetPath}}
        </div>
      </div>
      <div style="background:rgba(220,38,38,0.15);border:1px solid rgba(220,38,38,0.3);border-radius:8px;padding:0.8rem 1rem;margin-bottom:1.5rem">
        <div style="font-size:0.85rem;color:#FCA5A5;line-height:1.6">
          <strong style="color:#F87171">⛔ 경로 불일치 시 선택 사항이 무시됩니다.</strong><br>
          브라우저 보안 정책으로 인해 직접 저장이 차단되었습니다. 파일을 다운로드 후 위 경로에 덮어씌워주세요.
        </div>
      </div>
      <button onclick="closeSaveWarning()" style="display:block;width:100%;padding:0.9rem;background:#F97316;color:#fff;border:none;border-radius:10px;font-size:1rem;font-weight:700;cursor:pointer;transition:background 0.2s" onmouseover="this.style.background='#EA580C'" onmouseout="this.style.background='#F97316'">
        확인했습니다 — 파일을 수동으로 이동하겠습니다
      </button>
    </div>
  `;
  document.body.appendChild(overlay);

  // 스타일 주입 (glow 애니메이션)
  if (!document.getElementById('saveWarningStyles')) {{
    const style = document.createElement('style');
    style.id = 'saveWarningStyles';
    style.textContent = '@keyframes pulseGlow {{ 0%,100% {{ box-shadow: 0 0 30px rgba(249,115,22,0.2); }} 50% {{ box-shadow: 0 0 60px rgba(249,115,22,0.5); }} }}';
    document.head.appendChild(style);
  }}

  // 하단 확인 바에 영구 경고 표시
  const pathEl = document.getElementById('confirmPath');
  pathEl.innerHTML = '<span style="color:#F97316;font-weight:700">⚠️ 수동 저장 필요:</span> ' + targetPath;
  pathEl.style.fontSize = '0.85rem';
  pathEl.style.padding = '0.3rem 0';
}}

function closeSaveWarning() {{
  const el = document.getElementById('savePathWarning');
  if (el) el.remove();
}}

function showSaveSuccess(msg) {{
  const btn = document.getElementById('confirmBtn');
  btn.textContent = msg || '저장 완료!';
  btn.className = 'confirm-btn done';
  setTimeout(() => {{
    btn.textContent = '확인 — 선택 저장';
    btn.className = 'confirm-btn';
  }}, 3000);
}}

// ─── 풀스크린 프레젠테이션 모드 ───
let presIdx = 0;

function openPresentation() {{
  presIdx = viewerIdx;
  document.getElementById('presOverlay').classList.add('active');
  document.body.style.overflow = 'hidden';
  renderPresSlide();
  updateHash();
}}

function closePresentation() {{
  document.getElementById('presOverlay').classList.remove('active');
  document.body.style.overflow = '';
  viewerIdx = presIdx;
  renderViewerSlide();
  history.replaceState(null, '', location.pathname + location.search);
}}

function presNav(dir) {{
  presIdx = Math.max(0, Math.min(SLIDES.length - 1, presIdx + dir));
  renderPresSlide();
  updateHash();
}}

function renderPresSlide() {{
  const s = SLIDES[presIdx];
  const t = THEMES[selectedTheme];
  const slide = document.getElementById('presSlide');

  slide.style.background = t.bg;
  slide.innerHTML = renderSlideContent(s, t, null, presIdx);

  const pct = SLIDES.length > 1 ? ((presIdx) / (SLIDES.length - 1)) * 100 : 100;
  document.getElementById('presProgressBar').style.width = pct + '%';
  document.getElementById('presInfo').textContent = `${{presIdx + 1}} / ${{SLIDES.length}} — ${{THEMES[selectedTheme].name}}`;
}}

function updateHash() {{
  history.replaceState(null, '', '#slide-' + (presIdx + 1));
}}

// ─── 키보드 ───
document.addEventListener('keydown', (e) => {{
  const overlay = document.getElementById('presOverlay');
  const isFullscreen = overlay.classList.contains('active');

  if (!isFullscreen) {{
    if (e.key === 'F5') {{ e.preventDefault(); openPresentation(); }}
    // 인라인 뷰어: 좌우 화살표
    if (e.key === 'ArrowRight') {{ e.preventDefault(); viewerNav(1); }}
    if (e.key === 'ArrowLeft') {{ e.preventDefault(); viewerNav(-1); }}
    return;
  }}

  switch(e.key) {{
    case 'ArrowRight': case 'ArrowDown': case ' ': case 'PageDown':
      e.preventDefault(); presNav(1); break;
    case 'ArrowLeft': case 'ArrowUp': case 'PageUp':
      e.preventDefault(); presNav(-1); break;
    case 'Escape':
      e.preventDefault(); closePresentation(); break;
    case 'Home':
      e.preventDefault(); presIdx = 0; renderPresSlide(); updateHash(); break;
    case 'End':
      e.preventDefault(); presIdx = SLIDES.length - 1; renderPresSlide(); updateHash(); break;
  }}
}});

// ─── 터치/스와이프 ───
(function() {{
  let startX = 0, startY = 0;
  function addSwipe(el, navFn) {{
    el.addEventListener('touchstart', (e) => {{
      startX = e.touches[0].clientX;
      startY = e.touches[0].clientY;
    }}, {{passive: true}});
    el.addEventListener('touchend', (e) => {{
      const dx = e.changedTouches[0].clientX - startX;
      const dy = e.changedTouches[0].clientY - startY;
      if (Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 50) {{
        navFn(dx < 0 ? 1 : -1);
      }}
    }}, {{passive: true}});
  }}
  addSwipe(document.getElementById('presViewport'), presNav);
  addSwipe(document.getElementById('viewerStage'), viewerNav);
}})();

// ─── URL 해시 복원 ───
if (location.hash.startsWith('#slide-')) {{
  const n = parseInt(location.hash.replace('#slide-', ''));
  if (n >= 1 && n <= SLIDES.length) {{
    presIdx = n - 1;
    viewerIdx = n - 1;
    setTimeout(openPresentation, 100);
  }}
}}

// ─── 초기화 ───
renderThemes();
renderViewerSlide();
updateConfirmInfo();

// file:// 프로토콜일 때 서버 모드 안내
if (location.protocol === 'file:') {{
  const pathEl = document.getElementById('confirmPath');
  if (pathEl) {{
    pathEl.textContent = '💡 로컬 서버 모드로 실행하면 다이얼로그 없이 자동 저장됩니다.';
    pathEl.style.color = '#F97316';
    pathEl.style.fontSize = '0.8rem';
  }}
}}
</script>
</body>
</html>"""

    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"프리뷰 HTML 생성 완료: {output_html}")
    print(f"  슬라이드: {len(slides)}장")
    print(f"  다중 이미지 슬라이드: {len(multi_image_slides)}개")
    print(f"  choices 저장 경로: {choices_full_path}")


import random

def serve_preview(html_path: str, choices_path: str, port: int = 0):
    """로컬 HTTP 서버로 프리뷰를 제공하고, POST /save-choices로 JSON을 저장한다.
    port=0이면 10000-60000 사이의 랜덤 포트를 시도한다."""
    html_dir = str(Path(html_path).parent.resolve())
    html_name = Path(html_path).name

    class Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=html_dir, **kwargs)

        def do_POST(self):
            if self.path == "/save-choices":
                length = int(self.headers.get("Content-Length", 0))
                data = json.loads(self.rfile.read(length).decode("utf-8"))
                
                # 클라이언트가 보낸 경로 또는 기본 경로 사용
                target_path = data.get("path", choices_path)
                choices_data = data.get("choices", {})
                
                # 디렉토리 생성
                target_dir = os.path.dirname(target_path)
                if target_dir and not os.path.exists(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                
                with open(target_path, "w", encoding="utf-8") as f:
                    json.dump(choices_data, f, ensure_ascii=False, indent=2)
                
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
                print(f"\nchoices.json 저장 완료: {target_path}")
            else:
                self.send_response(404)
                self.end_headers()

        def log_message(self, format, *args):
            pass  # 로그 억제

    server = None
    # 포트 충돌 방지를 위해 랜덤 시도 (최대 10회)
    attempts = 0
    while attempts < 10:
        try:
            current_port = port if port != 0 else random.randint(10000, 60000)
            server = http.server.HTTPServer(("localhost", current_port), Handler)
            break
        except OSError:
            if port != 0: raise # 지정된 포트 실패 시 중단
            attempts += 1
            continue
    
    if not server:
        print("사용 가능한 포트를 찾을 수 없습니다.")
        return

    actual_port = server.server_address[1]
    url = f"http://localhost:{actual_port}/{html_name}"
    print(f"프리뷰 서버: {url}")
    print("Ctrl+C로 종료")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버 종료")
        server.shutdown()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python preview.py <slides.json> <output.html> [--images-dir <경로>] [--serve]")
        sys.exit(1)

    slides_path = sys.argv[1]
    output_path = sys.argv[2]
    img_dir = ""
    do_serve = "--no-serve" not in sys.argv  # 기본값: 서버 모드 (자동 저장 지원)

    if "--images-dir" in sys.argv:
        idx = sys.argv.index("--images-dir")
        if idx + 1 < len(sys.argv):
            img_dir = sys.argv[idx + 1]

    generate_preview_html(slides_path, output_path, img_dir)

    if do_serve:
        choices_path = str(Path(output_path).parent / (Path(output_path).stem + ".choices.json"))
        serve_preview(output_path, choices_path)
