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

# --- CUSTOM CSS FOR HIGH-CONTRAST DATA PREVIEW ---
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
    
    /* Document Display Canvas styling layout grid rules */
    .excel-container {
        background-color: white !important;
        border-radius: 2px;
        padding: 0.4rem;
        border: 1px solid #d0d0d0;
        overflow: auto;
        margin-top: 0.5rem;
        width: 100%;
    }
    .excel-container table {
        border-collapse: collapse;
        width: 100%;
        font-size: 10px;
        table-layout: fixed; /* Lock columns grid layout tightly */
    }
    .excel-container td {
        padding: 4px 6px;
        border: 1px solid #a0a0a0; /* High contrast structural wire borders */
        word-break: break-word !important; /* Force internal vertical wrapping expanding upwards */
        white-space: normal !important;
        vertical-align: middle;
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

# --- HIGH-FIDELITY RANGE HTML CONVERSION GENERATOR Engine (A1:P67) ---
def render_range_to_html(workbook, sheet_name=None, range_string="A1:P67"):
    ws = workbook[sheet_name] if sheet_name else workbook.active
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    
    # Calculate adaptive column configuration width metrics matching proportions perfectly
    col_widths = []
    for c in range(min_col, max_col + 1):
        w = ws.column_dimensions[get_column_letter(c)].width
        col_widths.append(w if w else 10)
    total_width = sum(col_widths)
    col_pcts = [f"{(w / total_width) * 100}%" for w in col_widths]
    
    html = '<table style="border-collapse: collapse; font-family: \'Roboto\', sans-serif; font-size: 10px; width: 100%; table-layout: fixed;">'
    html += '<colgroup>'
    for pct in col_pcts:
        html += f'<col style="width: {pct};">'
    html += '</colgroup>'
    
    # Pre-map cell tracking locations for structural sheet mergers
    merged_cells = {}
    for merged_range in ws.merged_cells.ranges:
        m_min_c, m_min_r, m_max_c, m_max_r = range_boundaries(str(merged_range))
        if m_min_r >= min_row and m_max_r <= max_row and m_min_c >= min_col and m_max_c <= max_col:
            for r in range(m_min_r, m_max_r + 1):
                for c in range(m_min_c, m_max_c + 1):
                    if r == m_min_r and c == m_min_c:
                        merged_cells[(r, c)] = {
                            'rowspan': m_max_r - m_min_r + 1,
                            'colspan': m_max_c - m_min_c + 1,
                            'is_master': True
                        }
                    else:
                        merged_cells[(r, c)] = {'is_master': False}
                        
    for r in range(min_row, max_row + 1):
        html += '<tr>'
        for c in range(min_col, max_col + 1):
            if (r, c) in merged_cells and not merged_cells[(r, c)].get('is_master', False):
                continue
                
            cell = ws.cell(row=r, column=c)
            value = cell.value if cell.value is not None else ''
            
            # Plain Text Cast Configuration Strategy
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            elif hasattr(value, 'strftime'):
                value = value.strftime('%B %d, %Y')
            val_str = str(value).strip()
            if re.match(r'^\d+\.0$', val_str):
                val_str = val_str.split('.')[0]
                
            # Parse Hex Styles mapping color contrast models
            bg_hex = '#FFFFFF'
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                rgb_str = cell.fill.fgColor.rgb
                if len(rgb_str) == 8: bg_hex = '#' + rgb_str[2:]
                elif len(rgb_str) == 6: bg_hex = '#' + rgb_str
                
            font_hex = '#333333'
            font_weight = 'normal'
            if cell.font:
                if cell.font.bold: font_weight = 'bold'
                if cell.font.color and cell.font.color.rgb:
                    rgb_f = cell.font.color.rgb
                    if len(rgb_f) == 8: font_hex = '#' + rgb_f[2:]
                    elif len(rgb_f) == 6: font_hex = '#' + rgb_f
                    
            # Hardcoded manual design adjustments for dark red headers
            val_upper = val_str.upper()
            if bg_hex.lower() in ['#800000', '#8c0000', '#7a0000'] or "SITE INFORMATION REPORT" in val_upper:
                font_hex = '#FFFFFF'
                font_weight = 'bold'
                h_align = 'center'
            elif any(header in val_upper for header in ["GENERAL INFORMATION", "LOCATION", "TERMS", "RATES", "TECHNICAL INFO", "PROVISIONS", "LESSOR AND TENANT DETAILS", "IF WITH SUB-LESSOR", "REGULATORY", "SITE ACQUIRABILITY"]):
                font_hex = '#000000'
                font_weight = 'bold'
                h_align = 'left'
            else:
                h_align = 'left'
                if cell.alignment and cell.alignment.horizontal:
                    h_align = cell.alignment.horizontal
                    
            style = f'background-color: {bg_hex}; color: {font_hex}; font-weight: {font_weight}; '
            style += f'text-align: {h_align}; border: 1px solid #a0a0a0; '
            style += 'white-space: normal !important; word-wrap: break-word !important; word-break: break-word !important;'
            
            rowspan = merged_cells[(r, c)].get('rowspan', 1) if (r, c) in merged_cells else 1
            colspan = merged_cells[(r, c)].get('colspan', 1) if (r, c) in merged_cells else 1
            
            if rowspan > 1 or colspan > 1:
                html += f'<td style="{style}" rowspan="{rowspan}" colspan="{colspan}">{val_str}</td>'
            else:
                html += f'<td style="{style}">{val_str}</td>'
        html += '</tr>'
    html += '</table>'
    return html

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

# --- DIRECT HTML INJECTION LAYER (Brave Safe) ---
if site_excel_bytes:
    try:
        active_wb = load_workbook(io.BytesIO(site_excel_bytes))
        
        # Pull range A1:P67 straight from the newly loaded populated openpyxl book
        range_html_content = render_range_to_html(active_wb, range_string="A1:P67")
        
        # Direct markdown injection has no base64 string pointers for Brave Shields to track or block
        st.markdown(f'<div class="excel-container">{range_html_content}</div>', unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Error rendering grid display framework matrix: {str(e)}")
else:
    st.info("Select a Trade Area and Site to view the report.")
