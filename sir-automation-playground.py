import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries
import re
import io
import requests
from copy import copy
import os
import hashlib
from openpyxl import load_workbook
import base64
import streamlit.components.v1 as components

# ReportLab libraries needed to draw a highly precise PDF grid layout from raw spreadsheet cells
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="trs.sitesourcing.viewer",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- PROGRAMMATIC LIGHT MODE LOCK ---
_config_dir = ".streamlit"
_config_file = os.path.join(_config_dir, "config.toml")
if not os.path.exists(_config_file):
    os.makedirs(_config_dir, exist_ok=True)
    with open(_config_file, "w", encoding="utf-8") as f:
        f.write("[theme]\nbase=\"light\"\n")

# --- CUSTOM CSS FOR CLEAN UI CONTROLS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    * { font-family: 'Roboto', 'Segoe UI', sans-serif !important; }
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    button[title="View source"] {display: none !important;}
    .stAppDeployButton {display: none !important;}
    div[data-testid="stStatusWidget"] {display: none !important;}
    
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        max-width: 100% !important;
    }
    .stButton > button, .stDownloadButton > button {
        background-color: #e8e8e8 !important;
        color: #333333 !important;
        border: 1px solid #d0d0d0 !important;
        border-radius: 2px !important;
        padding: 0.1rem 0.3rem !important;
        font-size: 0.7rem !important;
        min-height: 24px !important;
        height: 24px !important;
        width: 100%;
    }
    .stSelectbox label { display: none !important; }
    .stSelectbox > div > div {
        background-color: #fafafa !important;
        border-color: #d0d0d0 !important;
        min-height: 24px !important;
        height: 24px !important;
    }
    .stSelectbox > div > div > div { padding-top: 0 !important; padding-bottom: 0 !important; font-size: 0.7rem !important; }
    div[data-testid="stHorizontalBlock"] { gap: 0.3rem !important; align-items: center !important; }
    .info-text { font-size: 0.7rem; color: #333; text-align: right; margin: 0; padding: 0; line-height: 24px; font-weight: 500; }
    
    .pdf-frame-container {
        border: 1px solid #d0d0d0;
        margin-top: 0.5rem;
        background-color: #525659;
        padding: 4px;
        border-radius: 2px;
    }
</style>
""", unsafe_allow_html=True)

# --- LOGIN VERIFICATION LOGIC ---
TARGET_HASH = "6e7dfba0b39da481db37c3263c61cac6"
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

def check_password(password):
    return hashlib.md5(password.encode('utf-8')).hexdigest() == TARGET_HASH

if not st.session_state.authenticated:
    r1_col1, r1_col2, r1_col3 = st.columns([1, 1.2, 1])
    with r1_col2:
        st.markdown("<h3 style='text-align: center; margin-top:50px;'>Access Required</h3>", unsafe_allow_html=True)
        password_input = st.text_input("Enter password:", type="password", label_visibility="collapsed")
        if st.button("Login", use_container_width=True) or password_input:
            if check_password(password_input):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid token string provided.")
    st.stop()

# --- CONFIGURATION ---
SOURCE_URL = "https://docs.google.com/spreadsheets/d/14nhO9u7zJRcOoux8I7l2IzwU7iQZNW9fRX6TCip47CE/export?format=xlsx"
TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1uS3xmnPi0o4c_EayQtURYDSMMPRDRGSb/export?format=xlsx"

# --- HELPER FUNCTIONS ---
@st.cache_data(ttl=3600)
def download_file(url):
    try:
        response = requests.get(url, timeout=30)
        return io.BytesIO(response.content) if response.status_code == 200 else None
    except:
        return None

def get_placeholders(sheet):
    placeholders = set()
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                for match in matches:
                    name = match.split(":")[0].strip() if ":" in match else match.strip()
                    placeholders.add(name)
    return sorted(list(placeholders))

def clone_cell_styles(source_cell, target_cell):
    if source_cell.font:
        target_cell.font = Font(
            name=source_cell.font.name, size=source_cell.font.size, bold=source_cell.font.bold,
            italic=source_cell.font.italic, color=copy(source_cell.font.color)
        )
    if source_cell.alignment:
        target_cell.alignment = Alignment(horizontal=source_cell.alignment.horizontal, vertical=source_cell.alignment.vertical, wrap_text=True)
    if source_cell.fill:
        target_cell.fill = copy(source_cell.fill)

def copy_and_merge_aware_injection(template_ws, target_ws, coord, data_value):
    if not target_ws: return
    target_cell = target_ws[coord]
    template_cell = template_ws[coord]
    target_cell.value = data_value
    clone_cell_styles(template_cell, target_cell)

def sanitize_tab_name(name, existing_names):
    clean_name = re.sub(r'[\\/*?\[\]:]', '', str(name))[:31]
    if clean_name not in existing_names:
        existing_names.add(clean_name)
        return clean_name
    counter = 1
    while True:
        new_name = f"{clean_name[:27]} ({counter})"
        if new_name not in existing_names:
            existing_names.add(new_name)
            return new_name
        counter += 1

# --- REPORTLAB CONVERSION ENGINE (A1:P67 Range Target Matrix) ---
def generate_pdf_from_range(workbook, sheet_name=None, range_string="A1:P67"):
    ws = workbook[sheet_name] if sheet_name else workbook.active
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    
    page_width = 792  # 11 inches landscape
    page_height = 612 # 8.5 inches landscape
    margin = 18
    usable_width = page_width - (margin * 2)
    
    excel_widths = [ws.column_dimensions[get_column_letter(c)].width for c in range(min_col, max_col + 1)]
    excel_widths = [w if w else 10 for w in excel_widths]
    sum_widths = sum(excel_widths)
    pdf_col_widths = [(w / sum_widths) * usable_width for w in excel_widths]
    
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer, pagesize=landscape(letter),
        leftMargin=margin, rightMargin=margin, topMargin=margin, bottomMargin=margin
    )
    
    table_data = []
    t_styles = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#B0B0B0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ]
    
    for r_idx, r in enumerate(range(min_row, max_row + 1)):
        row_cells = []
        for c_idx, c in enumerate(range(min_col, max_col + 1)):
            cell = ws.cell(row=r, column=c)
            val = cell.value if cell.value is not None else ''
            
            if isinstance(val, float) and val.is_integer():
                val = int(val)
            elif hasattr(val, 'strftime'):
                val = val.strftime('%B %d, %Y')
            val_str = str(val).strip()
            if re.match(r'^\d+\.0$', val_str):
                val_str = val_str.split('.')[0]
                
            bg_hex = '#FFFFFF'
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                rgb_str = cell.fill.fgColor.rgb
                if len(rgb_str) == 8: bg_hex = '#' + rgb_str[2:]
                elif len(rgb_str) == 6: bg_hex = '#' + rgb_str
            
            font_hex = '#333333'
            is_bold = False
            if cell.font:
                is_bold = bool(cell.font.bold)
                if cell.font.color and cell.font.color.rgb:
                    rgb_f = cell.font.color.rgb
                    if len(rgb_f) == 8: font_hex = '#' + rgb_f[2:]
                    elif len(rgb_f) == 6: font_hex = '#' + rgb_f
            
            if bg_hex.lower() in ['#800000', '#8c0000', '#7a0000'] or "SITE INFORMATION REPORT" in val_str.upper():
                font_hex = '#FFFFFF'
                is_bold = True
                align_p = 'center'
            else:
                align_p = 'left'
                if cell.alignment and cell.alignment.horizontal:
                    align_p = cell.alignment.horizontal
            
            t_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(bg_hex)))
            
            p_style = ParagraphStyle(
                name=f'CellStyle_{r}_{c}', fontName='Helvetica-Bold' if is_bold else 'Helvetica',
                fontSize=6.5, leading=8, textColor=colors.HexColor(font_hex), alignment=0 if align_p == 'left' else (1 if align_p == 'center' else 2)
            )
            row_cells.append(Paragraph(val_str, p_style))
        table_data.append(row_cells)
        
    for merged_range in ws.merged_cells.ranges:
        m_min_c, m_min_r, m_max_c, m_max_r = range_boundaries(str(merged_range))
        if m_min_r >= min_row and m_max_r <= max_row and m_min_c >= min_col and m_max_c <= max_col:
            sc_idx = m_min_c - min_col
            sr_idx = m_min_r - min_row
            ec_idx = m_max_c - min_col
            er_idx = m_max_r - min_row
            t_styles.append(('SPAN', (sc_idx, sr_idx), (ec_idx, er_idx)))
            
    final_table = Table(table_data, colWidths=pdf_col_widths)
    final_table.setStyle(TableStyle(t_styles))
    
    doc.build([final_table])
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()

# --- LOAD DATA ---
@st.cache_data(ttl=3600)
def load_data():
    source_data = download_file(SOURCE_URL)
    template_data = download_file(TEMPLATE_URL)
    if source_data is None or template_data is None: return None, None, None
    
    df = pd.read_excel(source_data)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    
    temp_wb = load_workbook(template_data)
    placeholders = get_placeholders(temp_wb.active)
    template_data.seek(0)
    return df, placeholders, template_data.getvalue()

with st.spinner("Loading Data Assets..."):
    df, placeholders, template_bytes_raw = load_data()

if df is None or template_bytes_raw is None:
    st.error("Failed to load data. Please check connection profiles.")
    st.stop()

template_data = io.BytesIO(template_bytes_raw)

# --- CONTROLS ROW ---
trade_areas = sorted(df["TRADE AREA"].dropna().unique())
col1, col2, col3, col4, col5, col6 = st.columns([1.5, 1.5, 0.6, 0.7, 0.7, 0.6])

with col1:
    selected_ta = st.selectbox("Trade Area", options=trade_areas, index=0 if trade_areas else None, key="ta_select")
with col2:
    sites_in_ta = sorted(df[df["TRADE AREA"] == selected_ta]["SITE NAME"].dropna().unique()) if selected_ta else []
    selected_site = st.selectbox("Site Name", options=sites_in_ta, index=0 if sites_in_ta else None, key="site_select")
with col3:
    if st.button("Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

site_excel_bytes = None
if selected_ta and selected_site:
    site_data = df[(df["TRADE AREA"] == selected_ta) & (df["SITE NAME"] == selected_site)]
    if not site_data.empty:
        row = site_data.iloc[0]
        template_data.seek(0)
        wb = load_workbook(template_data)
        base_sheet = wb.active
        for row_cells in base_sheet.iter_rows():
            for cell in row_cells:
                if isinstance(cell.value, str) and ("{{" in cell.value):
                    new_val = cell.value
                    for ph in placeholders:
                        target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                        if re.search(target_regex, new_val):
                            raw_data_val = row.get(ph.upper(), "")
                            if pd.isna(raw_data_val) or raw_data_val is None: raw_data_val = ""
                            if isinstance(raw_data_val, float) and raw_data_val.is_integer(): val_str = str(int(raw_data_val))
                            elif hasattr(raw_data_val, 'strftime'): val_str = raw_data_val.strftime('%B %d, %Y')
                            else: val_str = str(raw_data_val)
                            new_val = re.sub(target_regex, val_str, new_val)
                    cell.value = new_val.strip() if new_val else ""
        
        ex_buf = io.BytesIO()
        wb.save(ex_buf)
        site_excel_bytes = ex_buf.getvalue()

with col4:
    if site_excel_bytes:
        safe_filename = f"{selected_site}_{selected_ta}".replace("/", "-").replace("\\", "-")
        st.download_button("Excel Report", data=site_excel_bytes, file_name=f"{safe_filename}.xlsx", use_container_width=True)

with col5:
    if selected_ta:
        if st.button("Trade Report", use_container_width=True):
            with st.spinner("Generating..."):
                ta_data = df[df["TRADE AREA"] == selected_ta]
                template_data.seek(0)
                wb = load_workbook(template_data)
                base_sheet = wb.active
                base_sheet.title = "TEMPLATE_TO_DELETE"
                existing_tabs = set()
                for _, r_row in ta_data.iterrows():
                    s_name = r_row.get("SITE NAME", "Unknown")
                    safe_tab_name = sanitize_tab_name(s_name, existing_tabs)
                    new_sheet = wb.copy_worksheet(base_sheet)
                    new_sheet.title = safe_tab_name
                    for row_cells in new_sheet.iter_rows():
                        for cell in row_cells:
                            if isinstance(cell.value, str) and "{{" in cell.value:
                                new_val = cell.value
                                for ph in placeholders:
                                    target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                                    if re.search(target_regex, new_val):
                                        raw_data_val = r_row.get(ph.upper(), "")
                                        if pd.isna(raw_data_val) or raw_data_val is None: raw_data_val = ""
                                        if isinstance(raw_data_val, float) and raw_data_val.is_integer(): val_str = str(int(raw_data_val))
                                        elif hasattr(raw_data_val, 'strftime'): val_str = raw_data_val.strftime('%B %d, %Y')
                                        else: val_str = str(raw_data_val)
                                        new_val = re.sub(target_regex, val_str, new_val)
                                cell.value = new_val.strip() if new_val else ""
                wb.remove(base_sheet)
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                st.download_button("Download Bulk", data=wb_buffer.getvalue(), file_name="Trade_Report.xlsx", use_container_width=True)

with col6:
    st.markdown(f"<p class='info-text'>Sites: {len(df['SITE NAME'].dropna().unique())}</p>", unsafe_allow_html=True)

# --- INLINE REPORT PDF VIEWER ENGAGEMENT LAYER ---
if site_excel_bytes:
    try:
        active_wb = load_workbook(io.BytesIO(site_excel_bytes))
        pdf_data_bytes = generate_pdf_from_range(active_wb, range_string="A1:P67")
        
        # Convert to an fully isolated sandbox data string
        base64_pdf = base64.b64encode(pdf_data_bytes).decode('utf-8')
        
        # Wrapping it directly into an explicit iframe string document template
        html_string = f'''
            <iframe src="data:application/pdf;base64,{base64_pdf}#toolbar=0&navpanes=0&scrollbar=1" 
                    width="100%" 
                    height="800px" 
                    style="border:none;">
            </iframe>
        '''
        
        # components.html mounts the element onto a clean decoupled sub-domain. 
        # This completely tricks Brave Shields, allowing the PDF to bypass any blocking layers.
        with st.container():
            st.markdown('<div class="pdf-frame-container">', unsafe_allow_html=True)
            components.html(html_string, height=810, scrolling=False)
            st.markdown('</div>', unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Error compiling canvas range mapping stream: {str(e)}")
else:
    st.info("Select a Trade Area and Site to view the report.")
