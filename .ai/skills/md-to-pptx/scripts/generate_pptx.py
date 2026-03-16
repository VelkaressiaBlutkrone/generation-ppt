"""
PPTX 생성 엔진

JSON spec 파일을 읽어 python-pptx로 프레젠테이션을 생성한다.
PPT_Design_Guide_2026 원칙을 자동 적용한다.

Usage:
    python generate_pptx.py <spec.json> <output.pptx> [choices.json]
"""

import json
import sys
import os
import re
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsmap
from PIL import Image
from copy import deepcopy
from lxml import etree

# ─── 슬라이드 크기 (16:9) ─────────────────────────────────────────
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# ─── 여백 ─────────────────────────────────────────────────────────
MARGIN_LEFT = Inches(0.8)
MARGIN_RIGHT = Inches(0.8)
MARGIN_TOP = Inches(0.6)
MARGIN_BOTTOM = Inches(0.5)
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_HEIGHT = SLIDE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM

# ─── 폰트 ─────────────────────────────────────────────────────────
FONT_NAME = "Pretendard"
FONT_NAME_EA = "Pretendard"  # East Asian fallback
FONT_FALLBACKS = ["Malgun Gothic", "맑은 고딕", "Apple SD Gothic Neo", "sans-serif"]


def set_font_with_ea(run, font_name=FONT_NAME, ea_name=FONT_NAME_EA):
    """폰트에 East Asian typeface를 함께 설정한다 (한글 렌더링 안정화)."""
    run.font.name = font_name
    # oxml로 <a:ea> East Asian typeface 추가
    rPr = run._r.get_or_add_rPr()
    # 기존 <a:ea> 제거
    for ea in rPr.findall(qn('a:ea')):
        rPr.remove(ea)
    ea_elem = etree.SubElement(rPr, qn('a:ea'))
    ea_elem.set('typeface', ea_name)


# ─── 테이블 oxml 헬퍼 ─────────────────────────────────────────────

def _set_cell_border(cell, side, width_pt=0, color_hex=None):
    """셀의 특정 변(side)에 테두리를 설정한다. width_pt=0이면 선 제거."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    border_tag = {
        'top': 'a:lnT', 'bottom': 'a:lnB',
        'left': 'a:lnL', 'right': 'a:lnR'
    }[side]

    # 기존 제거
    for old in tcPr.findall(qn(border_tag)):
        tcPr.remove(old)

    ln = etree.SubElement(tcPr, qn(border_tag))
    if width_pt == 0:
        ln.set('w', '0')
        ln.set('cmpd', 'sng')
        no_fill = etree.SubElement(ln, qn('a:noFill'))
    else:
        ln.set('w', str(int(width_pt * 12700)))  # pt → EMU
        ln.set('cmpd', 'sng')
        sf = etree.SubElement(ln, qn('a:solidFill'))
        srgb = etree.SubElement(sf, qn('a:srgbClr'))
        srgb.set('val', color_hex.lstrip('#') if color_hex else '000000')


def _set_cell_margins(cell, top=Pt(8), bottom=Pt(8), left=Pt(12), right=Pt(12)):
    """셀 내부 여백(padding)을 설정한다."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set('marT', str(int(top)))
    tcPr.set('marB', str(int(bottom)))
    tcPr.set('marL', str(int(left)))
    tcPr.set('marR', str(int(right)))

# ─── 테마 정의 ─────────────────────────────────────────────────────
THEMES = {
    "dark": {
        "bg": "#1A1A2E",
        "text": "#FFFFFF",
        "subtitle": "#A0A0B8",
        "accent": "#E94560",
        "secondary": "#16213E",
        "card_bg": "#16213E",
    },
    "light": {
        "bg": "#FFFFFF",
        "text": "#1A1A1A",
        "subtitle": "#6B7280",
        "accent": "#2563EB",
        "secondary": "#F3F4F6",
        "card_bg": "#F3F4F6",
    },
    "minimal": {
        "bg": "#FAFAFA",
        "text": "#333333",
        "subtitle": "#888888",
        "accent": "#6366F1",
        "secondary": "#F0F0F0",
        "card_bg": "#FFFFFF",
    },
    "consulting": {
        "bg": "#FFFFFF",
        "text": "#002F6C",
        "subtitle": "#4A5568",
        "accent": "#C41230",
        "secondary": "#F7F8FA",
        "card_bg": "#F7F8FA",
    },
    "pitch": {
        "bg": "#0F0A1A",
        "text": "#FFFFFF",
        "subtitle": "#C4B5D0",
        "accent": "#7B2FBE",
        "secondary": "#1A1030",
        "card_bg": "#1A1030",
    },
    "education": {
        "bg": "#F8F7FF",
        "text": "#1E1B4B",
        "subtitle": "#6B6B8D",
        "accent": "#F97316",
        "secondary": "#EDE9FE",
        "card_bg": "#FFFFFF",
    },
}


def hex_to_rgb(hex_color: str) -> RGBColor:
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_slide_bg(slide, hex_color):
    """슬라이드 배경색을 설정한다."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(hex_color)


def add_text_with_markdown(paragraph, text, theme, font_size=Pt(18), is_body=True):
    """마크다운 스타일(**굵게**, *기울임*)을 해석하여 텍스트를 추가한다."""
    if not text:
        return

    # 간단한 마크다운 파싱 (정규표현식)
    # **bold**, *italic*, ~~strike~~, `code`
    parts = re.split(r'(\*\*.*?\*\*|\*.*?\*|~~.*?~~|`.*?`)', text)
    
    for part in parts:
        if not part: continue
        
        run = paragraph.add_run()
        run.font.size = font_size
        run.font.color.rgb = hex_to_rgb(theme["text"])
        set_font_with_ea(run)
        
        if part.startswith('**') and part.endswith('**'):
            run.text = part[2:-2]
            run.font.bold = True
        elif part.startswith('*') and part.endswith('*'):
            run.text = part[1:-1]
            run.font.italic = True
        elif part.startswith('~~') and part.endswith('~~'):
            run.text = part[2:-2]
            # python-pptx doesn't have native strikethrough in common API, skipping for simplicity or use oxml
        elif part.startswith('`') and part.endswith('`'):
            run.text = part[1:-1]
            run.font.name = "Consolas"
        else:
            run.text = part


# ─── 레이아웃 핸들러 ──────────────────────────────────────────────

def layout_title(prs, slide_data, theme):
    """제목 슬라이드"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    set_slide_bg(slide, theme["bg"])

    # 메인 제목
    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, Inches(2.5), CONTENT_WIDTH, Inches(1.5)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(54)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 부제목
    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        sub_box = slide.shapes.add_textbox(
            MARGIN_LEFT, Inches(4.2), CONTENT_WIDTH, Inches(0.8)
        )
        tf2 = sub_box.text_frame
        para2 = tf2.paragraphs[0]
        para2.alignment = PP_ALIGN.CENTER
        run2 = para2.add_run()
        run2.text = subtitle
        run2.font.size = Pt(24)
        run2.font.color.rgb = hex_to_rgb(theme["subtitle"])
        set_font_with_ea(run2)

    # 강조 데코레이션
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(5.5), Inches(4.0), Inches(2.3), Pt(4)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    line.line.fill.background()

    return slide


def layout_section(prs, slide_data, theme):
    """섹션 구분 슬라이드"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["secondary"])

    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, Inches(3.0), CONTENT_WIDTH, Inches(1.5)
    )
    tf = title_box.text_frame
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 사이드 바 데코
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Pt(0), Inches(2.8), Inches(0.4), Inches(1.8)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    bar.line.fill.background()

    return slide


def layout_content(prs, slide_data, theme):
    """기본 콘텐츠 슬라이드 (불릿 리스트)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    # 제목
    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8)
    )
    tf = title_box.text_frame
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 본문
    body = slide_data.get("body", "")
    if body:
        content_box = slide.shapes.add_textbox(
            MARGIN_LEFT, MARGIN_TOP + Inches(1.0), CONTENT_WIDTH, CONTENT_HEIGHT - Inches(1.0)
        )
        tf = content_box.text_frame
        tf.word_wrap = True
        
        lines = body.split('\n')
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            # 불릿 포인트 처리
            clean_line = line.strip()
            if clean_line.startswith('- ') or clean_line.startswith('* '):
                p.level = 0
                add_text_with_markdown(p, clean_line[2:], theme)
            elif clean_line.startswith('  - ') or clean_line.startswith('  * '):
                p.level = 1
                add_text_with_markdown(p, clean_line[4:], theme)
            else:
                add_text_with_markdown(p, line, theme)

    return slide


def layout_content_image(prs, slide_data, theme):
    """좌측 텍스트, 우측 이미지 레이아웃"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    # 제목
    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8)
    )
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 텍스트 영역 (좌측 45%)
    body_width = CONTENT_WIDTH * 0.45
    content_box = slide.shapes.add_textbox(
        MARGIN_LEFT, MARGIN_TOP + Inches(1.0), body_width, CONTENT_HEIGHT - Inches(1.0)
    )
    tf = content_box.text_frame
    tf.word_wrap = True
    body = slide_data.get("body", "")
    for i, line in enumerate(body.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        add_text_with_markdown(p, line, theme, font_size=Pt(18))

    # 이미지 영역 (우측 50%)
    img_path = slide_data.get("image")
    if img_path and os.path.exists(img_path):
        img_left = MARGIN_LEFT + body_width + Inches(0.4)
        img_top = MARGIN_TOP + Inches(1.0)
        img_width = CONTENT_WIDTH * 0.5
        img_height = CONTENT_HEIGHT - Inches(1.0)
        
        # 가로세로 비율 유지하며 채우기
        slide.shapes.add_picture(img_path, img_left, img_top, width=img_width)

    return slide


def layout_image_full(prs, slide_data, theme):
    """전체 화면 이미지 레이아웃"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, "#000000")

    img_path = slide_data.get("image")
    if img_path and os.path.exists(img_path):
        # 꽉 채우기
        slide.shapes.add_picture(img_path, 0, 0, width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

    # 하단 캡션/제목
    title = slide_data.get("title", "")
    if title:
        overlay = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, SLIDE_HEIGHT - Inches(1.2), SLIDE_WIDTH, Inches(1.2)
        )
        overlay.fill.solid()
        overlay.fill.fore_color.rgb = RGBColor(0, 0, 0)
        overlay.fill.transparency = 0.4
        overlay.line.fill.background()

        txt = slide.shapes.add_textbox(Inches(0.5), SLIDE_HEIGHT - Inches(1.0), SLIDE_WIDTH - Inches(1.0), Inches(0.8))
        para = txt.text_frame.paragraphs[0]
        run = para.add_run()
        run.text = title
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_font_with_ea(run)

    return slide


def layout_two_images(prs, slide_data, theme):
    """텍스트 + 2개 이미지 병렬 배치"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    # 제목
    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 텍스트 (상단 30%)
    body = slide_data.get("body", "")
    if body:
        txt_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP + Inches(0.9), CONTENT_WIDTH, Inches(1.5))
        tf = txt_box.text_frame; tf.word_wrap = True
        for i, line in enumerate(body.split('\n')):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            add_text_with_markdown(p, line, theme, font_size=Pt(16))

    # 이미지 2개 (하단 60%)
    imgs = slide_data.get("images", [])
    if len(imgs) >= 2:
        w = (CONTENT_WIDTH - Inches(0.4)) / 2
        h = Inches(3.5)
        top = SLIDE_HEIGHT - h - MARGIN_BOTTOM
        if os.path.exists(imgs[0]):
            slide.shapes.add_picture(imgs[0], MARGIN_LEFT, top, width=w)
        if os.path.exists(imgs[1]):
            slide.shapes.add_picture(imgs[1], MARGIN_LEFT + w + Inches(0.4), top, width=w)

    return slide


def layout_grid_images(prs, slide_data, theme):
    """이미지 그리드 레이아웃 (2x2 등)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(28); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    imgs = slide_data.get("images", [])
    if not imgs: return slide

    n = len(imgs)
    cols = 2 if n > 1 else 1
    rows = (n + 1) // 2
    
    gap = Inches(0.2)
    w = (CONTENT_WIDTH - (cols-1)*gap) / cols
    h = (CONTENT_HEIGHT - Inches(1.0) - (rows-1)*gap) / rows
    
    for i, img_path in enumerate(imgs):
        if not os.path.exists(img_path): continue
        r, c = i // cols, i % cols
        slide.shapes.add_picture(
            img_path, 
            MARGIN_LEFT + c*(w+gap), 
            MARGIN_TOP + Inches(1.0) + r*(h+gap),
            width=w
        )
    return slide

def layout_text_left_img_right(prs, slide_data, theme):
    """텍스트(좌) + 이미지 그리드(우)"""
    return layout_content_image(prs, slide_data, theme) # 공용으로 사용하거나 커스텀 구현

def layout_img_left_text_right(prs, slide_data, theme):
    """이미지 그리드(좌) + 텍스트(우)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])
    
    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", ""); run.font.size = Pt(32); run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"]); set_font_with_ea(run)

    img_width = CONTENT_WIDTH * 0.5
    imgs = slide_data.get("images", [])
    if imgs and os.path.exists(imgs[0]):
        slide.shapes.add_picture(imgs[0], MARGIN_LEFT, MARGIN_TOP + Inches(1.0), width=img_width)

    body_left = MARGIN_LEFT + img_width + Inches(0.4)
    content_box = slide.shapes.add_textbox(body_left, MARGIN_TOP + Inches(1.0), CONTENT_WIDTH - img_width - Inches(0.4), CONTENT_HEIGHT - Inches(1.0))
    tf = content_box.text_frame; tf.word_wrap = True
    body = slide_data.get("body", "")
    for i, line in enumerate(body.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        add_text_with_markdown(p, line, theme, font_size=Pt(18))
    return slide

def layout_text_top_img_bottom(prs, slide_data, theme):
    """텍스트(상) + 이미지 가로 나열(하)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])
    # 제목
    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", ""); run.font.size = Pt(28); run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"]); set_font_with_ea(run)
    # 본문
    body = slide_data.get("body", "")
    txt_h = Inches(1.5)
    txt_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP + Inches(0.9), CONTENT_WIDTH, txt_h)
    tf = txt_box.text_frame; tf.word_wrap = True
    for i, line in enumerate(body.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        add_text_with_markdown(p, line, theme, font_size=Pt(16))
    # 이미지들
    imgs = slide_data.get("images", [])
    if imgs:
        n = len(imgs)
        gap = Inches(0.2)
        w = (CONTENT_WIDTH - (n-1)*gap) / n
        top = MARGIN_TOP + Inches(0.9) + txt_h + Inches(0.3)
        for i, img in enumerate(imgs):
            if os.path.exists(img):
                slide.shapes.add_picture(img, MARGIN_LEFT + i*(w+gap), top, width=w)
    return slide


def layout_comparison(prs, slide_data, theme):
    """좌우 비교 레이아웃"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    col_w = (CONTENT_WIDTH - Inches(0.6)) / 2
    
    def render_col(data, x):
        box = slide.shapes.add_textbox(x, MARGIN_TOP + Inches(1.2), col_w, CONTENT_HEIGHT - Inches(1.2))
        tf = box.text_frame; tf.word_wrap = True
        
        # 컬럼 제목
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = data.get("title", "")
        r0.font.size = Pt(22); r0.font.bold = True; r0.font.color.rgb = hex_to_rgb(theme["accent"])
        set_font_with_ea(r0)
        
        # 컬럼 내용
        for line in data.get("body", "").split('\n'):
            p = tf.add_paragraph()
            add_text_with_markdown(p, line, theme, font_size=Pt(16))

    render_col(slide_data.get("left", {}), MARGIN_LEFT)
    render_col(slide_data.get("right", {}), MARGIN_LEFT + col_w + Inches(0.6))

    return slide


def layout_table(prs, slide_data, theme):
    """데이터 테이블 슬라이드"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    headers = slide_data.get("headers", [])
    rows = slide_data.get("rows", [])
    if not headers: return slide

    rows_count = len(rows) + 1
    cols_count = len(headers)
    
    table_shape = slide.shapes.add_table(
        rows_count, cols_count, 
        MARGIN_LEFT, MARGIN_TOP + Inches(1.2), 
        CONTENT_WIDTH, Inches(0.5) * rows_count
    )
    table = table_shape.table

    # 헤더 스타일
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
        _set_cell_border(cell, 'bottom', width_pt=2, color_hex="#FFFFFF")
        para = cell.text_frame.paragraphs[0]
        run = para.add_run()
        run.text = str(h)
        run.font.size = Pt(18); run.font.bold = True; run.font.color.rgb = RGBColor(255, 255, 255)
        set_font_with_ea(run)

    # 데이터 행
    for r, row_data in enumerate(rows):
        for c, val in enumerate(row_data):
            cell = table.cell(r + 1, c)
            cell.text = str(val)
            para = cell.text_frame.paragraphs[0]
            para.font.size = Pt(16)
            para.font.color.rgb = hex_to_rgb(theme["text"])
            # 줄무늬 배경 (Zebra striping)
            if r % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = hex_to_rgb(theme["secondary"])

    return slide


def layout_code(prs, slide_data, theme):
    """코드 블록 레이아웃"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(28); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    # 코드 배경 박스
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        MARGIN_LEFT, MARGIN_TOP + Inches(1.0),
        CONTENT_WIDTH, CONTENT_HEIGHT - Inches(1.0)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(30, 30, 30) # 항상 어두운 배경
    bg.line.color.rgb = hex_to_rgb(theme["accent"])

    code_box = slide.shapes.add_textbox(
        MARGIN_LEFT + Inches(0.2), MARGIN_TOP + Inches(1.1),
        CONTENT_WIDTH - Inches(0.4), CONTENT_HEIGHT - Inches(1.2)
    )
    tf = code_box.text_frame
    tf.word_wrap = False
    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = slide_data.get("code", "")
    run.font.name = "Consolas"
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(220, 220, 220)

    return slide


def layout_kpi(prs, slide_data, theme):
    """KPI / 지표 레이아웃 (카드형)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    title_box.text_frame.paragraphs[0].text = slide_data.get("title", "")

    metrics = slide_data.get("metrics", [])
    if not metrics: return slide

    card_count = len(metrics)
    card_w = (CONTENT_WIDTH - (card_count-1)*Inches(0.3)) / card_count
    card_h = Inches(2.5)
    
    for i, m in enumerate(metrics):
        left = MARGIN_LEFT + i * (card_w + Inches(0.3))
        top = Inches(3.0)
        
        # 카드 배경
        rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h)
        rect.fill.solid(); rect.fill.fore_color.rgb = hex_to_rgb(theme["card_bg"])
        rect.line.color.rgb = hex_to_rgb(theme["secondary"])
        
        # 값
        v_box = slide.shapes.add_textbox(left, top + Inches(0.5), card_w, Inches(0.8))
        p_v = v_box.text_frame.paragraphs[0]; p_v.alignment = PP_ALIGN.CENTER
        r_v = p_v.add_run(); r_v.text = str(m.get("value", ""))
        r_v.font.size = Pt(44); r_v.font.bold = True; r_v.font.color.rgb = hex_to_rgb(theme["accent"])
        
        # 라벨
        l_box = slide.shapes.add_textbox(left, top + Inches(1.4), card_w, Inches(0.5))
        p_l = l_box.text_frame.paragraphs[0]; p_l.alignment = PP_ALIGN.CENTER
        r_l = p_l.add_run(); r_l.text = str(m.get("label", ""))
        r_l.font.size = Pt(18); r_l.font.color.rgb = hex_to_rgb(theme["text"])

    return slide


def layout_timeline(prs, slide_data, theme):
    """타임라인 레이아웃"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(MARGIN_LEFT, MARGIN_TOP, CONTENT_WIDTH, Inches(0.8))
    run = title_box.text_frame.paragraphs[0].add_run()
    run.text = slide_data.get("title", "")
    run.font.size = Pt(32); run.font.bold = True; run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    events = slide_data.get("events", [])
    if not events: return slide

    # 가로 중앙 선
    line_y = SLIDE_HEIGHT / 2 + Inches(0.5)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, MARGIN_LEFT, line_y, CONTENT_WIDTH, Pt(2)
    )
    line.fill.solid(); line.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    line.line.fill.background()

    count = len(events)
    step = CONTENT_WIDTH / count
    
    for i, ev in enumerate(events):
        x = MARGIN_LEFT + i * step + (step / 2)
        
        # 원 점
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x - Pt(6), line_y - Pt(6), Pt(12), Pt(12))
        dot.fill.solid(); dot.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
        dot.line.fill.background()
        
        # 날짜 (선 위)
        d_box = slide.shapes.add_textbox(x - Inches(0.8), line_y - Inches(0.6), Inches(1.6), Inches(0.4))
        p_d = d_box.text_frame.paragraphs[0]; p_d.alignment = PP_ALIGN.CENTER
        r_d = p_d.add_run(); r_d.text = ev.get("date", "")
        r_d.font.size = Pt(18); r_d.font.bold = True; r_d.font.color.rgb = hex_to_rgb(theme["accent"])
        
        # 설명 (선 아래)
        desc_box = slide.shapes.add_textbox(x - Inches(1.0), line_y + Inches(0.3), Inches(2.0), Inches(1.0))
        tf = desc_box.text_frame; tf.word_wrap = True
        p_desc = tf.paragraphs[0]; p_desc.alignment = PP_ALIGN.CENTER
        r_desc = p_desc.add_run(); r_desc.text = ev.get("description", ev.get("title", ""))
        r_desc.font.size = Pt(14); r_desc.font.color.rgb = hex_to_rgb(theme["text"])

    return slide


def layout_closing(prs, slide_data, theme):
    """클로징 슬라이드"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])

    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, Inches(2.5), CONTENT_WIDTH, Inches(1.5)
    )
    tf = title_box.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = slide_data.get("title", "감사합니다")
    run.font.size = Pt(54)
    run.font.bold = True
    run.font.color.rgb = hex_to_rgb(theme["text"])
    set_font_with_ea(run)

    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        sub_box = slide.shapes.add_textbox(
            MARGIN_LEFT, Inches(4.2), CONTENT_WIDTH, Inches(0.8)
        )
        tf2 = sub_box.text_frame
        para2 = tf2.paragraphs[0]
        para2.alignment = PP_ALIGN.CENTER
        run2 = para2.add_run()
        run2.text = subtitle
        run2.font.size = Pt(22)
        run2.font.color.rgb = hex_to_rgb(theme["subtitle"])
        set_font_with_ea(run2)

    contact = slide_data.get("contact", "")
    if contact:
        c_box = slide.shapes.add_textbox(
            MARGIN_LEFT, Inches(5.2), CONTENT_WIDTH, Inches(0.5)
        )
        tf3 = c_box.text_frame
        para3 = tf3.paragraphs[0]
        para3.alignment = PP_ALIGN.CENTER
        run3 = para3.add_run()
        run3.text = contact
        run3.font.size = Pt(16)
        run3.font.color.rgb = hex_to_rgb(theme["subtitle"])
        set_font_with_ea(run3)

    # 강조 라인
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(5.5), Inches(4.0), Inches(2.3), Pt(4)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = hex_to_rgb(theme["accent"])
    line.line.fill.background()

    return slide


# ─── 레이아웃 디스패처 ──────────────────────────────────────────────

LAYOUT_HANDLERS = {
    "title": layout_title,
    "section": layout_section,
    "content": layout_content,
    "content-image": layout_content_image,
    "image-full": layout_image_full,
    "two-images": layout_two_images,
    "grid-images": layout_grid_images,
    "text-left-img-right": layout_text_left_img_right,
    "img-left-text-right": layout_img_left_text_right,
    "text-top-img-bottom": layout_text_top_img_bottom,
    "comparison": layout_comparison,
    "table": layout_table,
    "code": layout_code,
    "kpi": layout_kpi,
    "timeline": layout_timeline,
    "closing": layout_closing,
}


def add_speaker_notes(slide, notes: str):
    """슬라이드에 발표자 노트를 추가한다."""
    if notes:
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = notes


def resolve_image_paths(spec, base_dir):
    """슬라이드의 이미지 경로를 spec.json 기준 절대경로로 변환한다."""
    for slide_data in spec.get("slides", []):
        # 단일 이미지
        if "image" in slide_data and slide_data["image"]:
            img = slide_data["image"]
            if not os.path.isabs(img):
                slide_data["image"] = os.path.normpath(os.path.join(base_dir, img))
        # 다중 이미지
        if "images" in slide_data and slide_data["images"]:
            resolved = []
            for img in slide_data["images"]:
                if not os.path.isabs(img):
                    resolved.append(os.path.normpath(os.path.join(base_dir, img)))
                else:
                    resolved.append(img)
            slide_data["images"] = resolved


def generate(spec_path: str, output_path: str, choices_path: str = None):
    """JSON spec으로부터 PPTX를 생성한다."""
    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    choices = {}
    if choices_path and os.path.exists(choices_path):
        with open(choices_path, "r", encoding="utf-8") as f:
            choices = json.load(f)
        print(f"choices.json 로드 완료: {choices_path}")

    # 이미지 경로를 spec.json 기준으로 해석
    spec_dir = os.path.dirname(os.path.abspath(spec_path))
    resolve_image_paths(spec, spec_dir)

    meta = spec.get("meta", {})
    # choices.json의 테마가 우선
    theme_name = choices.get("theme", meta.get("theme", "light"))
    theme = THEMES.get(theme_name, THEMES["light"])

    # 슬라이드별 레이아웃 오버라이드 적용
    slide_overrides = choices.get("slide_overrides", {})
    
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    for i, slide_data in enumerate(spec.get("slides", [])):
        # choices.json에 해당 슬라이드 오버라이드가 있으면 반영
        if str(i) in slide_overrides:
            override = slide_overrides[str(i)]
            if "layout" in override:
                slide_data["layout"] = override["layout"]

        layout = slide_data.get("layout", "content")
        handler = LAYOUT_HANDLERS.get(layout, layout_content)
        slide = handler(prs, slide_data, theme)

        notes = slide_data.get("notes", "")
        if notes:
            add_speaker_notes(slide, notes)

    prs.save(output_path)
    print(f"PPTX 생성 완료: {output_path}")
    print(f"  테마: {theme_name}")
    print(f"  슬라이드 수: {len(spec.get('slides', []))}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_pptx.py <spec.json> <output.pptx> [choices.json]")
        sys.exit(1)

    spec_arg = sys.argv[1]
    output_arg = sys.argv[2]
    choices_arg = sys.argv[3] if len(sys.argv) > 3 else None

    generate(spec_arg, output_arg, choices_arg)
