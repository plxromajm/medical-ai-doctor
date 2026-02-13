# ==========================================
# 라이브러리 불러오기 (import)
# ==========================================
import streamlit as st
from google import genai
import json
import os
from datetime import datetime, timedelta
import docx
from pypdf import PdfReader
import random
from pptx import Presentation as PptxPresentation
import pandas as pd
from io import BytesIO
from docx import Document as DocxDocument
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ==========================================
# 1. 프로그램 기본 설정
# ==========================================
GOOGLE_API_KEY = "AIzaSyA4xWRH8HnIWmAWAOnU1D9w8eNoOGYJsMM"  # 선생님 키 확인!
client = genai.Client(api_key=GOOGLE_API_KEY)
MODEL = 'gemini-2.5-flash'
DB_FILE = "medical_flashcards.json"

st.set_page_config(page_title="MEDI-Quiz", page_icon="🩺", layout="wide")

# ==========================================
# CSS 스타일 설정
# ==========================================
st.markdown("""
<style>
    .question-box {
        background-color: #f8f9fa; padding: 25px; border-radius: 12px; 
        border: 1px solid #e9ecef; margin-bottom: 25px; font-size: 1.1rem; line-height: 1.6;
    }
    .option-row {
        display: flex; align-items: center; margin-bottom: 10px; padding: 10px;
        border-radius: 8px; transition: background-color 0.2s;
    }
    .option-row:hover { background-color: #f1f3f5; }
    .option-text { flex-grow: 1; margin-left: 15px; font-size: 1rem; }
    .eliminated { text-decoration: line-through; color: #adb5bd; }
    .stButton button { width: 100%; }
    .options-box {
        background-color: #f8f9fa; padding: 20px; border-radius: 12px; 
        border: 1px solid #e9ecef; margin-bottom: 25px;
    }
    .option-item {
        display: flex; align-items: center; padding: 12px 15px; 
        margin-bottom: 10px; border-radius: 8px; transition: background-color 0.2s;
    }
    .option-item:hover { background-color: #e9ecef; }
    .option-number {
        font-size: 1.1rem; font-weight: bold; margin-right: 15px; min-width: 30px;
    }
    
    /* 정리본 표 스타일 */
    .summary-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 20px;
        font-size: 0.95rem;
    }
    .summary-table th {
        background-color: #495057;
        color: white;
        padding: 10px;
        text-align: center;
        border: 1px solid #dee2e6;
        font-size: 1.1rem;
    }
    .summary-table td {
        border: 1px solid #dee2e6;
        padding: 10px;
        vertical-align: top;
    }
    .summary-header {
        background-color: #e9ecef;
        font-weight: bold;
        width: 20%;
        text-align: center;
        vertical-align: middle !important;
    }
    
    /* 하이라이트 스타일 */
    .hl-yellow { background-color: #fff3bf; padding: 2px 4px; border-radius: 3px; }
    .hl-blue { color: #1971c2; font-weight: bold; }
    .hl-gray { color: #adb5bd; }

    /* 파일 업로더 드래그앤드롭 스타일 */
    [data-testid="stFileUploader"] {
        background-color: #ffffff;
        border: 2px dashed #dee2e6;
        border-radius: 12px;
        padding: 25px 20px 15px 20px;
        transition: border-color 0.3s, background-color 0.3s;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #FF6B35;
        background-color: #fff8f5;
    }
    /* 라벨을 가운데 정렬, 굵게 */
    [data-testid="stFileUploader"] label {
        width: 100% !important;
        text-align: center !important;
    }
    [data-testid="stFileUploader"] label p {
        text-align: center !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: #212529 !important;
    }
    /* 드롭존 자체 테두리 제거 */
    [data-testid="stFileUploaderDropzone"] {
        border: none !important;
        background: transparent !important;
        padding: 15px 10px !important;
    }
    /* Browse 버튼 색상 */
    [data-testid="stFileUploaderDropzone"] button {
        color: #FF6B35 !important;
        border-color: #FF6B35 !important;
    }
    [data-testid="stFileUploaderDropzone"] button:hover {
        background-color: #FF6B35 !important;
        color: white !important;
    }
    /* 드롭존 안내 텍스트 */
    [data-testid="stFileUploaderDropzone"] span {
        color: #868e96 !important;
    }
    [data-testid="stFileUploaderDropzone"] small {
        color: #adb5bd !important;
    }

    /* 탭 글씨 크기 1.6배 */
    [data-testid="stTabs"] button[role="tab"] p {
        font-size: 1.6rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 백엔드 함수들
# ==========================================

# 워드 표 셀 배경색 설정을 위한 함수 (XML 조작)
def set_cell_background(cell, color_hex):
    cell_properties = cell._element.get_or_add_tcPr()
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color_hex)
    cell_properties.append(shading_elm)

def load_cards():
    if not os.path.exists(DB_FILE): return []
    with open(DB_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            return [card for card in data if 'options' in card and isinstance(card['options'], list)]
        except: return []

def save_all_cards(cards):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(cards, f, ensure_ascii=False, indent=4)

def save_card_to_file(question, options, correct_index, explanation):
    cards = load_cards()
    cards.append({
        "question": question, "options": options, "correct_index": correct_index,
        "explanation": explanation, "next_review": datetime.now().strftime("%Y-%m-%d"), "interval": 1
    })
    save_all_cards(cards)

def delete_card(index):
    cards = load_cards()
    if 0 <= index < len(cards): del cards[index]; save_all_cards(cards)

def update_card_schedule(card_index, is_correct):
    cards = load_cards()
    if card_index < len(cards):
        card = cards[card_index]
        if is_correct:
            card['interval'] = card['interval'] * 2 + 1
            st.toast(f"🎉 정답! {card['interval']}일 뒤에 봅니다.")
        else:
            card['interval'] = 1
            st.toast("🥲 오답... 내일 다시 복습!")
        card['next_review'] = (datetime.now() + timedelta(days=card['interval'])).strftime("%Y-%m-%d")
        save_all_cards(cards)

def read_file(file):
    try:
        if file.name.endswith('.pdf'):
            reader = PdfReader(file)
            return "\n".join([page.extract_text() for page in reader.pages])
        elif file.name.endswith('.docx'):
            doc = docx.Document(file)
            return "\n".join([para.text for para in doc.paragraphs])
    except: return ""
    return ""

# ==========================================
# 3. 화면 구성
# ==========================================
st.markdown("<h1 style='text-align: center; color: #FF6B35; font-size: 3.2rem;'>MEDI-Quiz</h1>", unsafe_allow_html=True)

if 'generated_quiz' not in st.session_state: st.session_state['generated_quiz'] = None
if 'show_explanation' not in st.session_state: st.session_state['show_explanation'] = False
if 'summary_data' not in st.session_state: st.session_state['summary_data'] = None
if 'user_style' not in st.session_state: st.session_state['user_style'] = ""

tab1, tab2, tab3, tab4 = st.tabs(["📝 문제 생성", "🧠 실전 모의고사", "🗂️ 문제 관리", "📋 정리본 형성"])

# ==========================================
# [탭 1] 문제 생성
# ==========================================
with tab1:
    uploaded_file = st.file_uploader("📄  학습 자료 업로드  ·  PDF / PPT / DOCX", type=['docx', 'pdf', 'pptx'], key="tab1_uploader")
    study_content = read_file(uploaded_file) if uploaded_file else ""
    if uploaded_file and study_content:
        st.success(f"파일 읽기 성공! ({len(study_content)}자)")

    if st.button("⚡ 5문제 출제하기", type="primary", use_container_width=True, disabled=not bool(study_content)):
        with st.spinner("출제위원이 5개 문제를 만들고 있습니다..."):
            try:
                medical_categories = ["순환기내과", "호흡기내과", "소화기내과", "신장내과", "내분비내과", "감염내과", "류마티스내과", "신경과", "일반외과", "산부인과", "소아청소년과", "응급의학과", "예방의학", "피부과", "정신건강의학과"]
                selected_categories = random.sample(medical_categories, 5)
                categories_str = ", ".join(selected_categories)

                prompt = f"""
                당신은 의사 국가고시 출제위원입니다. 다음 내용을 바탕으로 5지선다형 객관식 문제 5개를 만드세요.
                [필수 출제 계통] {categories_str} (순서대로)
                [내용] {study_content[:15000]}
                [출력] 반드시 JSON 배열 형식:
                [
                    {{"question": "질문", "options": ["보기1", "보기2", "보기3", "보기4", "보기5"], "correct_index": 0, "explanation": "해설"}}, ...
                ]
                """
                response = client.models.generate_content(model=MODEL, contents=prompt)
                quizzes = json.loads(response.text.replace("```json", "").replace("```", ""))

                if isinstance(quizzes, list):
                    for quiz in quizzes:
                        save_card_to_file(quiz['question'], quiz['options'], quiz['correct_index'], quiz['explanation'])
                    st.success(f"✅ {len(quizzes)}개 문제가 생성되어 저장되었습니다!")
                else: st.error("형식 오류")
            except Exception as e: st.error(f"오류: {e}")

# ==========================================
# [탭 2] 실전 모의고사
# ==========================================
with tab2:
    cards = load_cards()
    today = datetime.now().strftime("%Y-%m-%d")
    due_cards = [(i, c) for i, c in enumerate(cards) if c['next_review'] <= today]

    if not due_cards:
        st.info("🎉 오늘 풀 문제가 없습니다!")
    else:
        idx, card = due_cards[0]
        if 'current_quiz_idx' not in st.session_state or st.session_state.current_quiz_idx != idx:
            st.session_state.current_quiz_idx = idx
            st.session_state.selected_opt = None
            st.session_state.eliminated_opts = set()
            st.session_state.show_explanation = False

        st.write(f"남은 문제: **{len(due_cards)}개**")
        st.markdown(f"""<div class="question-box"><b>Q.</b> {card['question']}</div>""", unsafe_allow_html=True)
        st.write("---")

        circle_numbers = ["①", "②", "③", "④", "⑤"]
        st.markdown('<div class="options-box">', unsafe_allow_html=True)

        for i, opt_text in enumerate(card['options']):
            col_num, col_text, col_sel, col_elim = st.columns([1, 10, 1.5, 2])
            
            if i in st.session_state.eliminated_opts:
                col_num.markdown(f'<span class="option-number eliminated">{circle_numbers[i]}</span>', unsafe_allow_html=True)
            else:
                col_num.markdown(f'<span class="option-number">{circle_numbers[i]}</span>', unsafe_allow_html=True)

            text_style = "eliminated" if i in st.session_state.eliminated_opts else ""
            if st.session_state.selected_opt == i:
                col_text.markdown(f'<div class="{text_style}" style="font-weight: bold; color: #1971c2;">{opt_text}</div>', unsafe_allow_html=True)
            else:
                col_text.markdown(f'<div class="{text_style}">{opt_text}</div>', unsafe_allow_html=True)

            btn_label = "●" if st.session_state.selected_opt == i else "○"
            if col_sel.button(btn_label, key=f"sel_{idx}_{i}"):
                st.session_state.selected_opt = i
                st.rerun()

            elim_label = "해제" if i in st.session_state.eliminated_opts else "오답"
            btn_type = "secondary" if i in st.session_state.eliminated_opts else "primary"
            if col_elim.button(elim_label, key=f"elim_{idx}_{i}", type=btn_type):
                if i in st.session_state.eliminated_opts: st.session_state.eliminated_opts.remove(i)
                else: st.session_state.eliminated_opts.add(i)
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        st.write("---")

        if st.button("🚀 정답 확인", type="primary", use_container_width=True, disabled=st.session_state.show_explanation):
            if st.session_state.selected_opt is None: st.warning("답을 선택해주세요!")
            else:
                st.session_state.show_explanation = True
                if st.session_state.selected_opt == card['correct_index']:
                    st.balloons(); st.success("✅ 정답입니다!"); update_card_schedule(idx, True)
                else:
                    st.error(f"❌ 오답입니다. 정답은 {circle_numbers[card['correct_index']]}번 입니다."); update_card_schedule(idx, False)
                st.rerun()

        if st.session_state.show_explanation:
            with st.expander("💡 해설 보기", expanded=True): st.info(card['explanation'])
            if st.button("➡️ 다음 문제 풀기", type="primary", use_container_width=True):
                st.session_state.show_explanation = False; st.rerun()

# ==========================================
# [탭 3] 문제 관리
# ==========================================
with tab3:
    st.header("🗂️ 문제 리스트")
    cards = load_cards()
    if not cards: st.write("저장된 문제가 없습니다.")
    else:
        circle_numbers = ["①", "②", "③", "④", "⑤"]
        for i, card in enumerate(cards):
            with st.expander(f"#{i+1}. {card['question'][:40]}..."):
                st.markdown(f'<div class="question-box">**Q.** {card["question"]}</div>', unsafe_allow_html=True)
                st.markdown('<div class="options-box">', unsafe_allow_html=True)
                for opt_i, opt_text in enumerate(card['options']):
                    if opt_i == card['correct_index']:
                        st.markdown(f'<div class="option-item" style="background-color: #e7f5ff;"><span class="option-number" style="color: #1971c2;">{circle_numbers[opt_i]}</span><span style="color: #1971c2; font-weight: bold;">{opt_text}</span></div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="option-item"><span class="option-number">{circle_numbers[opt_i]}</span><span>{opt_text}</span></div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                st.caption(f"💡 해설: {card['explanation']}")
                if st.button("🗑️ 삭제", key=f"del_{i}", type="secondary"): delete_card(i); st.rerun()

# ==========================================
# [탭 4] 정리본 형성 (테이블 형식 업데이트)
# ==========================================
with tab4:
    st.info("강의자료와 족보를 업로드하면 주제별 표 형식의 정리본을 만듭니다.")
    
    st.markdown("""
    <div style="background-color:#f8f9fa; padding:12px; border-radius:8px; margin-bottom:15px;">
        <b>색상 범례:</b>&nbsp;&nbsp;
        <span class="hl-yellow">■ 정답 선지</span>&nbsp;&nbsp;
        <span class="hl-blue">■ 족보 출제(강의 관련)</span>&nbsp;&nbsp;
        <span class="hl-gray">■ 족보 출제(강의 무관)</span>
    </div>
    """, unsafe_allow_html=True)

    lecture_content = ""
    jokbo_content = ""
    col_upload1, col_upload2 = st.columns(2)

    with col_upload1:
        uploaded_summaries = st.file_uploader("📚  강의자료 업로드  ·  PDF / PPT", type=['pdf', 'pptx'], key="summary_uploader", accept_multiple_files=True)
        if uploaded_summaries:
            all_texts = []
            for f in uploaded_summaries:
                if f.name.endswith('.pdf'):
                    try:
                        reader = PdfReader(f)
                        text = "\n".join([p.extract_text() or "" for p in reader.pages])
                        if text.strip(): all_texts.append(text)
                    except: pass
                elif f.name.endswith('.pptx'):
                    try:
                        prs = PptxPresentation(f)
                        txt = []
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame: txt.append(shape.text_frame.text)
                        all_texts.append("\n".join(txt))
                    except: pass
            lecture_content = "\n\n".join(all_texts)
            if lecture_content: st.success(f"강의자료 읽기 성공! ({len(lecture_content)}자)")

    with col_upload2:
        uploaded_jokbo = st.file_uploader("📝  족보 업로드  ·  PDF / DOCX", type=['pdf', 'docx'], key="jokbo_uploader")
        if uploaded_jokbo:
            if uploaded_jokbo.name.endswith('.pdf'):
                try:
                    reader = PdfReader(uploaded_jokbo)
                    jokbo_content = "\n".join([p.extract_text() or "" for p in reader.pages])
                except: pass
            elif uploaded_jokbo.name.endswith('.docx'):
                try:
                    doc = docx.Document(uploaded_jokbo)
                    jokbo_content = "\n".join([p.text for p in doc.paragraphs])
                except: pass
            if jokbo_content: st.success(f"족보 읽기 성공! ({len(jokbo_content)}자)")

    st.divider()

    if st.button("📋 통합 표 정리본 생성", type="primary", use_container_width=True, disabled=not bool(lecture_content)):
        with st.spinner("AI가 강의와 족보를 분석하여 표를 만들고 있습니다..."):
            try:
                # 프롬프트: JSON 구조를 "Topic" -> "Subsections" 형태로 변경
                prompt = f"""
                당신은 의대 학습 정리 전문가입니다.
                강의자료를 메인 주제(질환 등)별로 나누고, 각 주제 하위에 소주제(임상양상, 진단, 치료 등)를 포함한 표 형태로 정리하세요.
                
                [색상 태그 규칙]
                - 족보 정답 선지 내용: <yellow>내용</yellow>
                - 족보 오답 선지(강의 관련): <blue>내용</blue>
                - 족보 오답 선지(강의 무관): <gray>내용</gray>
                
                [입력 자료]
                강의: {lecture_content[:30000]}
                족보: {jokbo_content[:20000]}

                [출력 형식 - JSON 배열]
                반드시 아래 구조를 지키세요.
                [
                  {{
                    "main_topic": "메인 주제명 (예: 급성 A형 간염)",
                    "sub_sections": [
                      {{ "key": "개요/정의", "value": "내용..." }},
                      {{ "key": "임상양상", "value": "발열, 황달..." }},
                      {{ "key": "진단", "value": "IgM anti-HAV <yellow>양성</yellow>..." }},
                      {{ "key": "치료", "value": "보존적 치료..." }}
                    ]
                  }},
                  ...
                ]
                """
                response = client.models.generate_content(model=MODEL, contents=prompt)
                st.session_state['summary_data'] = json.loads(response.text.replace("```json", "").replace("```", "").strip())
                st.rerun()
            except Exception as e: st.error(f"오류: {e}")

    # ── 결과 표시 및 워드 다운로드 ──
    if st.session_state['summary_data']:
        st.divider()
        st.subheader("📋 통합 정리본")
        
        # 1. 화면 표시 (HTML Table)
        for item in st.session_state['summary_data']:
            main_topic = item.get('main_topic', '주제 없음')
            
            # HTML Table 시작
            html_code = f"""
            <table class="summary-table">
                <thead>
                    <tr><th colspan="2">{main_topic}</th></tr>
                </thead>
                <tbody>
            """
            
            for sub in item.get('sub_sections', []):
                key = sub.get('key', '')
                value = sub.get('value', '')
                
                # 태그 변환 (HTML 표시용)
                value = value.replace('\n', '<br>')
                value = re.sub(r'<(yellow)>(.*?)</\1>', r'<span class="hl-yellow">\2</span>', value)
                value = re.sub(r'<(blue)>(.*?)</\1>', r'<span class="hl-blue">\2</span>', value)
                value = re.sub(r'<(gray)>(.*?)</\1>', r'<span class="hl-gray">\2</span>', value)
                
                html_code += f"""
                <tr>
                    <td class="summary-header">{key}</td>
                    <td>{value}</td>
                </tr>
                """
            
            html_code += "</tbody></table>"
            st.markdown(html_code, unsafe_allow_html=True)

        # 2. 워드 파일 생성 (표 스타일 적용)
        try:
            doc_out = DocxDocument()
            
            # 제목
            title = doc_out.add_heading('의대 강의/족보 통합 정리본', level=0)
            title.alignment = 1 # 가운데 정렬
            
            # 범례
            legend = doc_out.add_paragraph()
            legend.alignment = 1
            run_y = legend.add_run('■ 정답  ')
            run_y.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run_b = legend.add_run('■ 관련 오답  ')
            run_b.font.color.rgb = RGBColor(0x19, 0x71, 0xC2)
            run_g = legend.add_run('■ 무관 오답')
            run_g.font.color.rgb = RGBColor(0xAD, 0xB5, 0xBD)
            doc_out.add_paragraph() # 빈 줄

            for item in st.session_state['summary_data']:
                main_topic = item.get('main_topic', '')
                sub_sections = item.get('sub_sections', [])
                
                if not sub_sections: continue

                # 표 생성 (행 수: 소주제 개수 + 1(제목행), 열 수: 2)
                table = doc_out.add_table(rows=0, cols=2)
                table.style = 'Table Grid' # 격자 스타일
                
                # 1행: 메인 주제 (병합)
                row_main = table.add_row()
                cell_main = row_main.cells[0]
                cell_main.merge(row_main.cells[1])
                cell_main.text = main_topic
                
                # 메인 주제 스타일 (진한 회색 배경, 흰 글씨, 가운데 정렬)
                set_cell_background(cell_main, "495057") # Hex color
                run_main = cell_main.paragraphs[0].runs[0]
                run_main.font.color.rgb = RGBColor(255, 255, 255)
                run_main.bold = True
                run_main.font.size = Pt(12)
                cell_main.paragraphs[0].alignment = 1

                # 소주제 행들 추가
                for sub in sub_sections:
                    key = sub.get('key', '')
                    content = sub.get('value', '')
                    
                    row = table.add_row()
                    
                    # 왼쪽 셀 (소주제): 회색 배경
                    cell_key = row.cells[0]
                    cell_key.text = key
                    cell_key.width = Cm(3.5) # 너비 고정
                    set_cell_background(cell_key, "E9ECEF") # 연한 회색
                    cell_key.paragraphs[0].runs[0].bold = True
                    cell_key.vertical_alignment = 0 # Top 정렬

                    # 오른쪽 셀 (내용): 태그 파싱하여 스타일 적용
                    cell_val = row.cells[1]
                    cell_val.vertical_alignment = 0
                    p = cell_val.paragraphs[0]
                    
                    # 정규식으로 태그 분리해서 순서대로 넣기
                    parts = re.split(r'(<(?:yellow|blue|gray)>.*?</(?:yellow|blue|gray)>)', content)
                    for part in parts:
                        if not part: continue
                        
                        # 태그 확인
                        tag_match = re.match(r'<(yellow|blue|gray)>(.*?)</\1>', part)
                        if tag_match:
                            tag_type = tag_match.group(1)
                            text_body = tag_match.group(2)
                            run = p.add_run(text_body)
                            
                            if tag_type == 'yellow':
                                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            elif tag_type == 'blue':
                                run.font.color.rgb = RGBColor(0x19, 0x71, 0xC2)
                                run.bold = True
                            elif tag_type == 'gray':
                                run.font.color.rgb = RGBColor(0xAD, 0xB5, 0xBD)
                        else:
                            # 일반 텍스트
                            p.add_run(part)

                doc_out.add_paragraph() # 표 사이 간격

            bio = BytesIO()
            doc_out.save(bio)
            bio.seek(0)
            
            st.download_button("💾 표 정리본 다운로드 (Word)", data=bio, file_name="통합_표_정리본.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        except Exception as e:
            st.error(f"워드 생성 오류: {e}")