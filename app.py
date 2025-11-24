#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pattern Worksheet Generator - Vercel Final Fix
- Removed 'outputs' folder creation (Fixes Read-only file system error)
- Fully in-memory processing
"""

try:
    from flask import Flask, render_template, request, send_file, jsonify
    import openpyxl
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    import os
    import glob
    import random
    from datetime import datetime
    from io import BytesIO
except ImportError as e:
    print(f"CRITICAL ERROR: {e}")
    raise e

app = Flask(__name__)

# --- 경로 설정 ---
# 현재 파일의 위치를 기준으로 잡음
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FOLDER = os.path.join(BASE_DIR, 'databases')
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')
FONT_DIR = os.path.join(BASE_DIR, 'fonts')

# [중요] OUTPUT_FOLDER 생성 코드 삭제함 (Vercel은 쓰기 금지)

# 템플릿 폴더 위치 강제 지정
app.template_folder = TEMPLATE_DIR

# --- 폰트 설정 ---
def setup_korean_font():
    try:
        local_font = os.path.join(FONT_DIR, 'NanumGothic.ttf')
        if os.path.exists(local_font):
            pdfmetrics.registerFont(TTFont('KoreanFont', local_font))
            return 'KoreanFont'
        return 'Helvetica' 
    except:
        return 'Helvetica'

KOREAN_FONT = setup_korean_font()

def load_patterns_from_excel(filename):
    file_path = os.path.join(DB_FOLDER, filename)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"DB 파일을 찾을 수 없습니다: {filename}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    ws_overview = wb["Pattern Overview"]
    pattern_info = {}
    for row in ws_overview.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            pattern_info[int(row[0])] = {
                'number': int(row[0]),
                'name': str(row[1]),
                'unit': str(row[3]) if len(row) > 3 and row[3] else 'Level A'
            }
            
    ws_detail = wb["Pattern Details"]
    patterns = {}
    for row in ws_detail.iter_rows(min_row=2, values_only=True):
        try:
            p_num = int(row[0])
            section = row[2]
            content = row[4]
            answer = row[5] if len(row) > 5 and row[5] else ""
            
            if p_num not in patterns:
                patterns[p_num] = {
                    'pattern_num': p_num,
                    'pattern_name': pattern_info.get(p_num, {}).get('name', ''),
                    'unit': pattern_info.get(p_num, {}).get('unit', 'Level A'),
                    'speaking1': [], 'speaking2': [], 'unscramble': []
                }
            
            if section == 'Speaking I':
                patterns[p_num]['speaking1'].append(content)
            elif section == 'Speaking II':
                patterns[p_num]['speaking2'].append((content, answer))
            elif section == 'Unscramble':
                scrambled = row[6].strip('()') if row[6] else ""
                patterns[p_num]['unscramble'].append((content, scrambled, answer))
        except:
            continue
            
    return patterns

def distribute_questions(selected_patterns, target_count=5):
    result = {'speaking1': [], 'speaking2': [], 'unscramble': []}
    if not selected_patterns: return result
    
    pattern_count = len(selected_patterns)
    items_per = target_count // pattern_count
    remainder = target_count % pattern_count
    
    for section in ['speaking1', 'speaking2', 'unscramble']:
        for i, p in enumerate(selected_patterns):
            count = items_per + (1 if i < remainder else 0)
            pool = p[section][:]
            random.shuffle(pool)
            result[section].extend(pool[:count])
            
    return result

def create_worksheet_in_memory(pattern_data, selected_patterns, book_title, student_name="", student_date=""):
    buffer = BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=10*mm,
        bottomMargin=10*mm,
        leftMargin=15*mm,
        rightMargin=15*mm
    )
    
    story = []
    p_nums = ", ".join([str(p['pattern_num']) for p in selected_patterns])
    clean_book_title = book_title.replace('.xlsx', '')
    
    title_style = ParagraphStyle('Title', fontSize=12, fontName='Helvetica-Bold', alignment=TA_CENTER, spaceAfter=5)
    section_style = ParagraphStyle('Section', fontSize=11, fontName='Helvetica-Bold', spaceBefore=0, spaceAfter=0)
    item_style = ParagraphStyle('Item', fontSize=10, fontName='Helvetica', leftIndent=0, spaceBefore=2, spaceAfter=2)
    item_kr_style = ParagraphStyle('ItemKr', fontSize=10, fontName=KOREAN_FONT, leftIndent=0, spaceBefore=2, spaceAfter=2)
    line_style = ParagraphStyle('Line', fontSize=10, font