# ==========================================
# DELTA | PREMIUM DUE DILIGENCE PIPELINE
# ==========================================

import streamlit as st
import io
import re
import math
import hashlib
from datetime import datetime
import numpy as np

# Core Geospatial, Document, & Vision Engine Binaries
import pymupdf as fitz  
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import matplotlib.pyplot as plt
import shutil
import os

# Explicit dynamic resolution of Linux system path bindings for Tesseract
if not shutil.which("tesseract"):
    fallback_paths = ["/usr/bin/tesseract", "/usr/local/bin/tesseract"]
    for path in fallback_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            break
else:
    pytesseract.pytesseract.tesseract_cmd = shutil.which("tesseract")

# ==========================================
# BLOCK 1: STATE HYDRATION
# ==========================================
if 'initialized' not in st.session_state:
    st.session_state.initialized = True
    st.session_state.uploaded_titles = {}
    st.session_state.title_hashes = {}
    st.session_state.title_order = []           
    st.session_state.title_roles = {}           
    st.session_state.base_title_idx = 0      
    st.session_state.counter_title_idx = 1       
    st.session_state.extracted_tech_desc = ""
    st.session_state.parsed_polygon = None

# ==========================================
# BLOCK 2: SYSTEM DESIGN SYSTEM (CSS)
# ==========================================
def inject_luxury_system_css():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500&family=Inter:wght@300;400;500;600&display=swap');
            
            .stApp {
                background-color: #F8F9FA !important;
                color: #202124 !important;
                font-family: 'Inter', sans-serif !important;
            }
            
            [data-testid="stHeader"] { background-color: rgba(0,0,0,0) !important; border-bottom: none !important; }
            footer {visibility: hidden !important;}
            .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 95% !important; }
            
            .brand-title {
                font-family: 'Inter', sans-serif !important;
                color: #1A1A1A !important;
                font-weight: 600;
                font-size: 22px;
                letter-spacing: -0.5px;
                margin-bottom: 0.2rem;
            }
            .brand-subtitle { font-family: 'Inter', sans-serif; color: #5F6368; font-size: 11px; letter-spacing: 0.5px; margin-bottom: 1.5rem; }
            
            .doc-cell {
                background-color: #FFFFFF;
                padding: 16px 20px;
                border: 1px solid #DADCE0;
                border-radius: 6px;
                height: 100%;
                margin-bottom: 12px;
            }
            .empty-cell { background-color: transparent !important; border: none !important; }
            
            .stFileUploader { border: 1px dashed #DACCE0 !important; background-color: #FFFFFF !important; padding: 1rem !important; border-radius: 6px; }
            .stButton > button {
                background-color: #1A73E8 !important; color: #FFFFFF !important; border: none !important;
                border-radius: 4px !important; font-family: 'Inter', sans-serif; font-size: 12px !important; font-weight: 500;
                padding: 0.5rem 1.5rem !important; transition: all 0.2s ease; width: auto;
            }
            .stButton > button:hover { background-color: #1557B0 !important; }
            
            .stream-paragraph { color: #3C4043; font-size: 13px; line-height: 1.6; word-wrap: break-word; margin-bottom: 0px; }
            .add-token { background-color: #E6F4EA !important; color: #137333 !important; padding: 2px 4px; border-radius: 2px; }
            .del-token { background-color: #FCE8E6 !important; color: #C5221F !important; text-decoration: line-through; padding: 2px 4px; border-radius: 2px; }
            .trace-flag { font-family: 'Inter', sans-serif; color: #5F6368; font-size: 11px; font-weight: 600; text-transform: uppercase; margin-bottom: 0.5rem; display: block; padding-bottom: 4px; }
            
            .advisory-panel { background-color: #FFFFFF; border: 1px solid #DADCE0; padding: 12px; margin-bottom: 8px; border-radius: 6px; }
            .advisory-header { font-family: 'Inter', sans-serif; font-weight: 600; color: #202124; font-size: 12px; margin-bottom: 0.25rem; }
            
            .minimap-row { width: 100%; height: 6px; border-radius: 1px; margin-top: 24px; }
            .minimap-equal { background-color: #E8EAED; }
            .minimap-replace { background-color: #FCE8E6; }
            .minimap-delete { background-color: #D93025; }
            .minimap-insert { background-color: #1E8E3E; }
            
            div[data-baseweb="select"] { background-color: #FFFFFF !important; }
            input { border-radius: 4px !important; background-color: #FFFFFF !important; color: #202124 !important; border: 1px solid #DADCE0 !important; font-size: 12px !important; }
            textarea { border-radius: 4px !important; background-color: #FFFFFF !important; color: #202124 !important; border: 1px solid #DADCE0 !important; font-size: 12px !important; }
        </style>
    """, unsafe_allow_html=True)

# ==========================================
# BLOCK 3: ENGINE PARSING, OCR & HASHING
# ==========================================
def preprocess_image_for_ocr(image_bytes):
    image = Image.open(io.BytesIO(image_bytes)).convert('L')
    image = image.filter(ImageFilter.SHARPEN)
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    return image

def extract_text_via_ocr(file_bytes):
    try:
        img = preprocess_image_for_ocr(file_bytes)
        text = pytesseract.image_to_string(img)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        return lines
    except Exception as e:
        return [f"[OCR Error: {str(e)}]"]

def parse_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    paragraphs = []
    for page in doc:
        text_blocks = page.get_text("blocks")
        if len(text_blocks) > 0:
            for b in text_blocks:
                text = b[4].strip()
                if text: paragraphs.append(" ".join(text.split()))
        else:
            pix = page.get_pixmap(dpi=200)
            img_data = pix.tobytes("png")
            paragraphs.extend(extract_text_via_ocr(img_data))
    return paragraphs

def parse_docx(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip(): paragraphs.append(" ".join(p.text.split()))
    return paragraphs

def load_due_diligence_matrices(uploaded_files):
    for file in uploaded_files:
        if file.name not in st.session_state.uploaded_titles:
            bytes_data = file.read()
            file_hash = hashlib.sha256(bytes_data).hexdigest()
            st.session_state.title_hashes[file.name] = file_hash
            
            if file.name.endswith('.pdf'):
                parsed_text = parse_pdf(bytes_data)
            elif file.name.endswith('.docx'):
                parsed_text = parse_docx(bytes_data)
            elif file.name.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp')):
                parsed_text = extract_text_via_ocr(bytes_data)
            else:
                continue
            
            st.session_state.uploaded_titles[file.name] = parsed_text
            if file.name not in st.session_state.title_order:
                st.session_state.title_order.append(file.name)
                st.session_state.title_roles[file.name] = "Baseline Core"

# ==========================================
# BLOCK 4: FUZZY ALIGNMENT ENGINE
# ==========================================
def compute_fuzzy_alignment_matrix(left_paras, right_paras, threshold=0.35):
    import difflib
    matched_right_indices = set()
    alignment_opcodes = []
    
    for i, lp in enumerate(left_paras):
        best_ratio = 0.0
        best_j = None
        lp_tokens = sorted(lp.lower().split())
        
        for j, rp in enumerate(right_paras):
            if j in matched_right_indices: continue
            rp_tokens = sorted(rp.lower().split())
            ratio = difflib.SequenceMatcher(None, lp_tokens, rp_tokens).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_j = j
                
        if best_ratio >= threshold and best_j is not None:
            matched_right_indices.add(best_j)
            if best_ratio > 0.95 and lp == right_paras[best_j]: 
                alignment_opcodes.append(('equal', i, i, best_j, best_j))
            else: 
                alignment_opcodes.append(('replace', i, i, best_j, best_j))
        else:
            alignment_opcodes.append(('delete', i, i, None, None))
            
    for j in range(len(right_paras)):
        if j not in matched_right_indices:
            alignment_opcodes.append(('insert', None, None, j, j))
            
    alignment_opcodes.sort(key=lambda x: (x[1] if x[1] is not None else float('inf'), x[3] if x[3] is not None else 0))
    return alignment_opcodes

def compute_token_diff_html(text1, text2):
    import difflib
    matcher = difflib.SequenceMatcher(None, text1.split(), text2.split())
    out1, out2 = [], []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        w1 = " ".join(text1.split()[i1:i2]); w2 = " ".join(text2.split()[j1:j2])
        if tag == 'equal': out1.append(w1); out2.append(w2)
        elif tag == 'replace': 
            out1.append(f'<span class="del-token">{w1}</span>')
            out2.append(f'<span class="add-token">{w2}</span>')
        elif tag == 'delete': out1.append(f'<span class="del-token">{w1}</span>')
        elif tag == 'insert': out2.append(f'<span class="add-token">{w2}</span>')
    return " ".join(out1), " ".join(out2)

# ==========================================
# BLOCK 5: TECHNICAL DESCRIPTION GEOMETRY ENGINE
# ==========================================
def parse_technical_description(text_content):
    cleaned = text_content.replace('\n', ' ')
    pattern = r'THENCE\s+([N|S])\s*[\.]?\s*(\d+)\s*DEG\.\s*(\d+)\'?\s*([E|W])\s*[\.]?[\, ]?\s*(\d+(?:\.\d+)?)\s*M\.'
    matches = re.findall(pattern, cleaned, re.IGNORECASE)
    
    if not matches:
        pattern_fallback = r'(N|S)\s*(\d+)\s*DEG\s*(\d+)\s*(E|W)\s*(\d+(?:\.\d+)?)'
        matches = re.findall(pattern_fallback, cleaned, re.IGNORECASE)
    
    current_x, current_y = 0.0, 0.0
    coordinates = [(current_x, current_y)]
    
    for i, (ns, deg, mins, ew, dist) in enumerate(matches):
        try:
            d_val = float(dist)
            deg_val = float(deg)
            min_val = float(mins)
            
            angle_rad = math.radians(deg_val + (min_val / 60.0))
            ns_sign = 1.0 if ns.upper() == 'N' else -1.0
            ew_sign = 1.0 if ew.upper() == 'E' else -1.0
            
            dy = d_val * math.cos(angle_rad) * ns_sign
            dx = d_val * math.sin(angle_rad) * ew_sign
            
            current_x += dx
            current_y += dy
            coordinates.append((current_x, current_y))
        except ValueError:
            continue
            
    if len(coordinates) > 1:
        coordinates[-1] = (0.0, 0.0)
        
    return coordinates

def generate_white_blueprint_pdf(coordinates):
    doc = fitz.open()
    page_w, page_h = 612, 792 
    page = doc.new_page(width=page_w, height=page_h)
    
    page.draw_rect(fitz.Rect(0, 0, page_w, page_h), color=None, fill=(1, 1, 1))
    page.insert_text(fitz.Point(54, 54), "AUTOMATED LOT SURVEY PLAN", fontsize=12, fontname="Helvetica-Bold", color=(0.1, 0.1, 0.1))
    
    if len(coordinates) > 1:
        x_vals = [c[0] for c in coordinates]
        y_vals = [c[1] for c in coordinates]
        
        min_x, max_x = min(x_vals), max(x_vals)
        min_y, max_y = min(y_vals), max(y_vals)
        
        span_x = max_x - min_x if max_x != min_x else 1.0
        span_y = max_y - min_y if max_y != min_y else 1.0
        
        scale = min(400.0 / span_x, 400.0 / span_y)
        center_x, center_y = 306, 430
        
        map_cx = min_x + (span_x / 2.0)
        map_cy = min_y + (span_y / 2.0)
        
        transformed_pts = []
        for x, y in coordinates:
            px = center_x + (x - map_cx) * scale
            py = center_y - (y - map_cy) * scale 
            transformed_pts.append(fitz.Point(px, py))
            
        page.draw_line(fitz.Point(center_x-20, center_y), fitz.Point(center_x+20, center_y), color=(0.8, 0.8, 0.8), width=0.5)
        page.draw_line(fitz.Point(center_x, center_y-20), fitz.Point(center_x, center_y+20), color=(0.8, 0.8, 0.8), width=0.5)
        
        for idx in range(len(transformed_pts) - 1):
            p1 = transformed_pts[idx]
            p2 = transformed_pts[idx+1]
            page.draw_line(p1, p2, color=(0.1, 0.1, 0.1), width=1.5)
            page.draw_circle(p1, 2.5, color=(0.1, 0.1, 0.1), fill=(1, 1, 1))
            page.insert_text(fitz.Point(p1.x + 4, p1.y - 4), f"P{idx+1}", fontsize=7, color=(0.2, 0.2, 0.2))
            
    page.draw_line(fitz.Point(54, 720), fitz.Point(558, 720), color=(0.9, 0.9, 0.9), width=1)
    page.insert_text(fitz.Point(54, 740), "Generated directly from document technical transcripts for reference use.", fontsize=7, color=(0.5, 0.5, 0.5))
    
    return doc.write()

# ==========================================
# BLOCK 6: HIGH-FIDELITY EXPORTERS
# ==========================================
def export_due_diligence_docx(left_paras, right_paras, title_left, title_right, alignment_opcodes):
    doc = Document()
    
    # Target Section Setup
    section = doc.sections[-1]
    
    # Enforce Landscape Orientation
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Inches(11.69), Inches(8.27) # Exact A4 Dimensions
    
    # Inject Narrow Margins via Low-Level OpenXML Section Elements (0.5 Inches / 36pt Margin)
    sectPr = section._sectPr
    pgMar = OxmlElement('w:pgMar')
    pgMar.set(qn('w:top'), '720')    # 0.5 inch
    pgMar.set(qn('w:bottom'), '720') # 0.5 inch
    pgMar.set(qn('w:left'), '720')   # 0.5 inch
    pgMar.set(qn('w:right'), '720')  # 0.5 inch
    pgMar.set(qn('w:header'), '360')
    pgMar.set(qn('w:footer'), '360')
    sectPr.append(pgMar)
    
    # Heading Layout Elements
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("DUE DILIGENCE DISCREPANCY REPORT")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.name = 'Arial'
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = f"Baseline Source ({title_left})"
    hdr_cells[1].text = f"Target Comparison ({title_right})"
    hdr_cells[2].text = "Variance Summary"
    
    for tag, i1, _, j1, _ in alignment_opcodes:
        row = table.add_row()
        if tag == 'equal':
            row.cells[0].text = left_paras[i1] if i1 is not None else ""
            row.cells[1].text = right_paras[j1] if j1 is not None else ""
            row.cells[2].text = "No variance found."
        elif tag == 'replace':
            row.cells[0].text = left_paras[i1] if i1 is not None else ""
            row.cells[1].text = right_paras[j1] if j1 is not None else ""
            row.cells[2].text = "Mismatched Parameter Text. Modification identified."
        elif tag == 'delete':
            row.cells[0].text = left_paras[i1] if i1 is not None else ""
            row.cells[1].text = "[Omitted Field]"
            row.cells[2].text = "Baseline data point absent from target evaluation document."
        elif tag == 'insert':
            row.cells[0].text = "[Absent Frame]"
            row.cells[1].text = right_paras[j1] if j1 is not None else ""
            row.cells[2].text = "Injected target clause verified."
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# BLOCK 8: UI PHASE 2 - SYSTEM DASHBOARD REVIEW
# ==========================================
def render_due_diligence_workspace():
    st.markdown('<div class="brand-title">DELTA DUE DILIGENCE WORKSPACE</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">Automated Land Title Ingestion & Boundary Analysis Matrix</div>', unsafe_allow_html=True)
    
    tab_review, tab_geospatial = st.tabs(["❖ Core Title Discrepancy Matrix", "⚡ Boundary Technical Description Plotter"])
    
    with tab_review:
        st.markdown("### Document Ingestion")
        uploaded_files = st.file_uploader("Upload titles / files", 
                                          type=['pdf', 'docx', 'png', 'jpg', 'jpeg', 'tiff'], 
                                          accept_multiple_files=True, label_visibility="collapsed")
        
        if uploaded_files:
            load_due_diligence_matrices(uploaded_files)
            
            st.markdown("<div style='margin: 1rem 0;'>", unsafe_allow_html=True)
            for filename in st.session_state.title_order:
                col1, col2 = st.columns([7, 3])
                with col1: 
                    st.markdown(f'<p style="font-size:13px; color:#1A1A1A; padding-top:4px;">📁 {filename}</p>', unsafe_allow_html=True)
                with col2:
                    roles = ["Baseline Core", "Target Comparison", "Secondary Copy", "Tax Declaration Reference"]
                    st.session_state.title_roles[filename] = st.selectbox(f"Type_{filename}", roles, index=0, label_visibility="collapsed", key=f"dd_role_{filename}")
            st.markdown("</div>", unsafe_allow_html=True)

        ordered_files = st.session_state.title_order
        roles = [st.session_state.title_roles[f] for f in ordered_files]
        
        if len(ordered_files) >= 2:
            st.markdown("---")
            t_col1, t_col2 = st.columns(2)
            with t_col1: 
                st.session_state.base_title_idx = st.selectbox("Baseline Core", range(len(ordered_files)), format_func=lambda x: f"BASE // {ordered_files[x]} ({roles[x]})")
            with t_col2: 
                st.session_state.counter_title_idx = st.selectbox("Target Frame", range(len(ordered_files)), format_func=lambda x: f"EVAL // {ordered_files[x]} ({roles[x]})", index=min(1, len(ordered_files)-1))
                
            base_file = ordered_files[st.session_state.base_title_idx]
            counter_file = ordered_files[st.session_state.counter_title_idx]
            
            left_paras = st.session_state.uploaded_titles[base_file]
            right_paras = st.session_state.uploaded_titles[counter_file]
            
            alignment_opcodes = compute_fuzzy_alignment_matrix(left_paras, right_paras)
            
            st.markdown(f"""
                <div class="crypto-banner">
                    [BASE_SHA256]: {st.session_state.title_hashes.get(base_file, "N/A")}<br/>
                    [EVAL_SHA256]: {st.session_state.title_hashes.get(counter_file, "N/A")}
                </div>
            """, unsafe_allow_html=True)
            
            h_col1, _, h_col2, _ = st.columns([5, 0.2, 5, 4.5])
            with h_col1: st.markdown("<p style='font-size:11px; font-weight:600; color:#5F6368;'>BASELINE SOURCE</p>", unsafe_allow_html=True)
            with h_col2: st.markdown("<p style='font-size:11px; font-weight:600; color:#5F6368;'>COMPARISON TARGET</p>", unsafe_allow_html=True)
            
            change_index = 1
            for idx, (tag, i1, _, j1, _) in enumerate(alignment_opcodes):
                r_col1, r_col2, r_col3, r_col4 = st.columns([5, 0.2, 5, 4.5])
                
                with r_col1:
                    if tag == 'equal':
                        st.markdown(f'<div class="doc-cell"><p class="stream-paragraph">{left_paras[i1]}</p></div>', unsafe_allow_html=True)
                    elif tag == 'replace':
                        h1, _ = compute_token_diff_html(left_paras[i1], right_paras[j1])
                        st.markdown(f'<div class="doc-cell"><span class="trace-flag">▲ Variance</span><p class="stream-paragraph">{h1}</p></div>', unsafe_allow_html=True)
                    elif tag == 'delete':
                        st.markdown(f'<div class="doc-cell"><span class="trace-flag" style="color:#C5221F;">◼ Missing Item</span><p class="stream-paragraph"><span class="del-token">{left_paras[i1]}</span></p></div>', unsafe_allow_html=True)
                    elif tag == 'insert':
                        st.markdown('<div class="doc-cell empty-cell"></div>', unsafe_allow_html=True)
                        
                with r_col2:
                    if tag == 'equal': st.markdown('<div class="minimap-row minimap-equal"></div>', unsafe_allow_html=True)
                    elif tag == 'replace': st.markdown('<div class="minimap-row minimap-replace"></div>', unsafe_allow_html=True)
                    elif tag == 'delete': st.markdown('<div class="minimap-row minimap-delete"></div>', unsafe_allow_html=True)
                    elif tag == 'insert': st.markdown('<div class="minimap-row minimap-insert"></div>', unsafe_allow_html=True)
                    
                with r_col3:
                    if tag == 'equal':
                        st.markdown(f'<div class="doc-cell"><p class="stream-paragraph">{right_paras[j1]}</p></div>', unsafe_allow_html=True)
                    elif tag == 'replace':
                        _, h2 = compute_token_diff_html(left_paras[i1], right_paras[j1])
                        st.markdown(f'<div class="doc-cell"><span class="trace-flag">▲ Variance</span><p class="stream-paragraph">{h2}</p></div>', unsafe_allow_html=True)
                    elif tag == 'delete':
                        st.markdown('<div class="doc-cell empty-cell"></div>', unsafe_allow_html=True)
                    elif tag == 'insert':
                        st.markdown(f'<div class="doc-cell"><span class="trace-flag" style="color:#137333;">◆ External Injection</span><p class="stream-paragraph"><span class="add-token">{right_paras[j1]}</span></p></div>', unsafe_allow_html=True)
                        
                with r_col4:
                    if tag != 'equal':
                        unique_id = f"dd_row_{i1}_{j1}_{idx}"
                        st.markdown(f'<div class="advisory-panel"><div class="advisory-header"><span style="color:#D93025;">Alert #{change_index}</span></div></div>', unsafe_allow_html=True)
                        st.radio("Status", ["Clear", "Variance", "Risk", "Hold"], key=f"risk_{unique_id}", horizontal=True, label_visibility="collapsed")
                        st.text_input("Findings", key=f"note_dd_{unique_id}", placeholder="Notes...", label_visibility="collapsed")
                        change_index += 1
                    else:
                        st.write("")
                        
            st.markdown("<br/>", unsafe_allow_html=True)
            export_bytes = export_due_diligence_docx(left_paras, right_paras, base_file, counter_file, alignment_opcodes)
            st.download_button("📥 Export Matrix (.docx)", data=export_bytes, file_name=f"DD_Matrix_{datetime.now().strftime('%Y%m%d')}.docx")
        elif len(ordered_files) == 1:
            st.info("Upload at least two records to run comparison analysis structures.")

    with tab_geospatial:
        st.markdown("### Technical Description Plotter")
        
        # Dual Ingestion Architecture: Manual Text String Entry Block vs Photo/OCR Extraction Frame
        input_mode = st.radio("Select Input Mode", ["Manual Text Entry", "Photo / Scan Document Upload"], horizontal=True)
        
        extracted_text_target = ""
        
        if input_mode == "Manual Text Entry":
            default_tech_desc = """THENCE S. 32 DEG. 00'E., 18.58 M. TO POINT 2; THENCE S. 69 DEG. 49'W., 50.64 M. TO POINT 3; THENCE N. 5 DEG. 19'W., 10.19 M. TO POINT 4; THENCE N. 35 DEG. 56'W., 7.77 M. TO POINT 5; THENCE N. 35 DEG. 56'W., 1.96 M. TO POINT 6; THENCE N. 28 DEG. 59'W., 9.72 M. TO POINT 7; THENCE N. 82 DEG. 12'E., 49.55 M. TO THE POINT OF BEGINNING;"""
            extracted_text_target = st.text_area("Paste Technical Text Block", value=default_tech_desc, height=150)
        else:
            uploaded_scan = st.file_uploader("Upload boundary image / photo scan", type=['png', 'jpg', 'jpeg', 'tiff', 'bmp'])
            if uploaded_scan:
                scan_bytes = uploaded_scan.read()
                with st.spinner("Processing OCR Extract..."):
                    ocr_lines = extract_text_via_ocr(scan_bytes)
                    extracted_text_target = " ".join(ocr_lines)
                st.success("Text extracted completely.")
                with st.expander("View Extracted OCR Transcript"):
                    extracted_text_target = st.text_area("Adjust Transcribed Text Pattern", value=extracted_text_target, height=150)
        
        if extracted_text_target:
            coords = parse_technical_description(extracted_text_target)
            if len(coords) > 1:
                st.success(f"Tracked [{len(coords)-1}] vertex coordinates.")
                
                geo_col1, geo_col2 = st.columns([6, 4])
                
                with geo_col1:
                    fig, ax = plt.subplots(figsize=(5, 5), facecolor='#FFFFFF')
                    x_pts = [c[0] for c in coords]
                    y_pts = [c[1] for c in coords]
                    
                    ax.plot(x_pts, y_pts, color='#1A73E8', linestyle='-', linewidth=1.5, marker='o', markerfacecolor='#FFFFFF', markeredgecolor='#1A73E8', markersize=4)
                    for i, (xp, yp) in enumerate(coords[:-1]):
                        ax.text(xp + 1, yp + 1, f"P{i+1}", fontsize=8)
                        
                    ax.set_aspect('equal', 'box')
                    ax.grid(True, color='#E8EAED', linestyle='--', linewidth=0.5)
                    ax.set_facecolor('#FFFFFF')
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)
                    ax.spines['left'].set_color('#CCCCCC')
                    ax.spines['bottom'].set_color('#CCCCCC')
                    st.pyplot(fig)
                    plt.close(fig)
                    
                with geo_col2:
                    st.markdown("#### Export Asset Formats Matrix")
                    
                    pdf_bytes = generate_white_blueprint_pdf(coords)
                    st.download_button(label="📥 Export Lot Plan PDF",
                                       data=pdf_bytes,
                                       file_name=f"LOT_PLAN_{datetime.now().strftime('%Y%m%d')}.pdf",
                                       mime="application/pdf")
                    
                    st.markdown("---")
                    
                    kml_bytes = generate_kml_payload(coords)
                    st.download_button(label="📥 Export Vector KML",
                                       data=kml_bytes,
                                       file_name=f"LOT_GEOSPATIAL_{datetime.now().strftime('%Y%m%d')}.kml",
                                       mime="application/vnd.google-earth.kml+xml")
                    
                    st.markdown("---")
                    
                    st.dataframe([{"Point": f"Point {idx+1}", "X (Eastings)": f"{pt[0]:.2f} m", "Y (Northings)": f"{pt[1]:.2f} m"} for idx, pt in enumerate(coords[:-1])], use_container_width=True)
            else:
                st.error("Failed to parse bearing patterns from text inputs.")

    st.markdown("<br/><br/><hr style='border-color:#DADCE0;'/>", unsafe_allow_html=True)
    if st.button("Reset Session"):
        st.session_state.uploaded_titles = {}
        st.session_state.title_hashes = {}
        st.session_state.title_order = []
        st.session_state.title_roles = {}
        st.rerun()
        
# ==========================================
# BLOCK 9: ENTRYPOINT ORCHESTRATION
# ==========================================
def main():
    st.set_page_config(page_title="DELTA DUE DILIGENCE", layout="wide", initial_sidebar_state="collapsed")
    inject_luxury_system_css()
    render_due_diligence_workspace()

if __name__ == "__main__": 
    main()
