import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import io
import requests
from copy import copy
import os
import hashlib
from openpyxl import load_workbook
import tempfile
from openpyxl.drawing.image import Image
import base64
from io import BytesIO

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

# --- CUSTOM CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    
    * {
        font-family: 'Roboto', 'Segoe UI', sans-serif !important;
    }
    
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    button[title="View source"] {display: none !important;}
    .stAppDeployButton {display: none !important;}
    div[data-testid="stStatusWidget"] {display: none !important;}
    
    .block-container {
        padding-top: 0.1rem !important;
        padding-bottom: 0.1rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        max-width: 100% !important;
    }
    
    .stButton > button {
        background-color: #e8e8e8 !important;
        color: #333333 !important;
        border: 1px solid #d0d0d0 !important;
        border-radius: 3px !important;
        padding: 0.2rem 0.4rem !important;
        font-weight: 400 !important;
        font-size: 0.7rem !important;
        min-height: 28px !important;
        height: 28px !important;
        line-height: 1 !important;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #f5f5f5 !important;
        border-color: #b0b0b0 !important;
    }
    
    .stSelectbox > div > div {
        background-color: #fafafa !important;
        border-color: #d0d0d0 !important;
        color: #333333 !important;
        min-height: 28px !important;
        height: 28px !important;
    }
    .stSelectbox > div > div > div {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }
    .stSelectbox label {
        color: #333333 !important;
        font-weight: 400 !important;
        font-size: 0.7rem !important;
        margin-bottom: 0.1rem !important;
    }
    
    .stMarkdown, .stMarkdown * {
        color: #333333 !important;
    }
    
    .control-section {
        background-color: #f8f8f8 !important;
        border-radius: 2px;
        padding: 0.2rem 0.4rem !important;
        border: 1px solid #e8e8e8;
        margin-bottom: 0.2rem;
    }
    
    .excel-container {
        background-color: white !important;
        border-radius: 2px;
        padding: 0.3rem;
        border: 1px solid #e8e8e8;
        overflow: auto;
        height: calc(100vh - 160px);
    }
    
    .excel-container table {
        border-collapse: collapse;
        width: 100%;
        font-size: 10px;
    }
    
    .excel-container td {
        padding: 2px 4px;
        border: 1px solid #d0d0d0;
    }
    
    .footer-overlay {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        right: 0 !important;
        height: 25px !important;
        background-color: #fafafa !important;
        z-index: 9999999 !important;
        border-top: 1px solid #e8e8e8 !important;
        pointer-events: none !important;
    }
    
    .stAlert {
        padding: 0.2rem 0.4rem !important;
        font-size: 0.7rem !important;
    }
    
    div[data-testid="stVerticalBlock"] {
        gap: 0.1rem !important;
    }
    
    div[data-testid="stHorizontalBlock"] {
        gap: 0.3rem !important;
        align-items: center !important;
    }
    
    .stDownloadButton > button {
        background-color: #e8e8e8 !important;
        color: #333333 !important;
        border: 1px solid #d0d0d0 !important;
        border-radius: 3px !important;
        padding: 0.2rem 0.4rem !important;
        font-weight: 400 !important;
        font-size: 0.7rem !important;
        min-height: 28px !important;
        height: 28px !important;
        line-height: 1 !important;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background-color: #f5f5f5 !important;
        border-color: #b0b0b0 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- PERSISTENT FOOTER OVERLAY ---
st.markdown("""
<div class="footer-overlay"></div>
""", unsafe_allow_html=True)

# --- LOGIN VERIFICATION LOGIC ---
TARGET_HASH = "6e7dfba0b39da481db37c3263c61cac6"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

def check_password(password):
    hashed = hashlib.md5(password.encode('utf-8')).hexdigest()
    return hashed == TARGET_HASH

# --- 3x3 LOGIN GRID LAYOUT ---
if not st.session_state.authenticated:
    st.write("##")
    st.write("##")
    
    r1_col1, r1_col2, r1_col3 = st.columns([1, 1.2, 1])
    
    with r1_col2:
        with st.container():
            st.markdown("<h3 style='text-align: center; margin-bottom: 20px;'>Access Required</h3>", unsafe_allow_html=True)
            password_input = st.text_input("Enter password:", type="password", label_visibility="collapsed")
            login_btn = st.button("Login", use_container_width=True)
            
            if login_btn or password_input:
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
        if response.status_code == 200:
            return io.BytesIO(response.content)
        return None
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
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            color=copy(source_cell.font.color),
            underline=source_cell.font.underline,
            strike=source_cell.font.strike,
            vertAlign=source_cell.font.vertAlign,
            scheme=source_cell.font.scheme
        )
    if source_cell.alignment:
        target_cell.alignment = Alignment(
            horizontal=source_cell.alignment.horizontal,
            vertical=source_cell.alignment.vertical,
            text_rotation=source_cell.alignment.text_rotation,
            wrap_text=source_cell.alignment.wrap_text,
            shrink_to_fit=source_cell.alignment.shrink_to_fit,
            indent=source_cell.alignment.indent
        )
    if source_cell.border:
        target_cell.border = Border(
            left=copy(source_cell.border.left),
            right=copy(source_cell.border.right),
            top=copy(source_cell.border.top),
            bottom=copy(source_cell.border.bottom),
            diagonal=copy(source_cell.border.diagonal),
            diagonal_direction=source_cell.border.diagonal_direction,
            outline=copy(source_cell.border.outline),
            vertical=copy(source_cell.border.vertical),
            horizontal=copy(source_cell.border.horizontal)
        )
    if source_cell.fill:
        target_cell.fill = copy(source_cell.fill)

def copy_and_merge_aware_injection(template_ws, target_ws, coord, data_value):
    if not target_ws:
        return
    target_cell = target_ws[coord]
    template_cell = template_ws[coord]
    target_cell.value = data_value
    clone_cell_styles(template_cell, target_cell)
    
    merged_range_string = None
    for merged_range in template_ws.merged_cells.ranges:
        if template_cell.coordinate in merged_range:
            merged_range_string = str(merged_range)
            break
    
    if merged_range_string:
        min_col, min_row, max_col, max_row = range_boundaries(merged_range_string)
        overlapping_target_ranges = []
        for target_range in target_ws.merged_cells.ranges:
            t_min_c, t_min_r, t_max_c, t_max_r = range_boundaries(str(target_range))
            if not (max_col < t_min_c or min_col > t_max_c or max_row < t_min_r or min_row > t_max_r):
                overlapping_target_ranges.append(target_range)
        for target_conflict in overlapping_target_ranges:
            try:
                target_ws.unmerge_cells(str(target_conflict))
            except:
                pass
        try:
            target_ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        except:
            pass
        
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                sub_coord = f"{get_column_letter(c)}{r}"
                if sub_coord != coord:
                    clone_cell_styles(template_ws[sub_coord], target_ws[sub_coord])

def render_excel_to_html(workbook, sheet_name=None):
    if sheet_name is None:
        ws = workbook.active
    else:
        ws = workbook[sheet_name]
    
    max_row = ws.max_row
    max_col = ws.max_column
    
    html = '<table style="border-collapse: collapse; font-family: \'Roboto\', sans-serif; font-size: 10px; width: 100%;">'
    
    merged_cells = {}
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if row == min_row and col == min_col:
                    merged_cells[(row, col)] = {
                        'rowspan': max_row - min_row + 1,
                        'colspan': max_col - min_col + 1,
                        'is_master': True
                    }
                else:
                    merged_cells[(row, col)] = {'is_master': False}
    
    for row in range(1, max_row + 1):
        html += '<tr>'
        for col in range(1, max_col + 1):
            if (row, col) in merged_cells and not merged_cells[(row, col)].get('is_master', False):
                continue
            
            cell = ws.cell(row, col)
            value = cell.value if cell.value is not None else ''
            
            bg_color = 'white'
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                bg_color = cell.fill.fgColor.rgb
                if bg_color.startswith('FF'):
                    bg_color = '#' + bg_color[2:]
            
            font_color = '#333333'
            font_weight = 'normal'
            font_size = '10px'
            if cell.font:
                if cell.font.color and cell.font.color.rgb:
                    font_color = cell.font.color.rgb
                    if font_color.startswith('FF'):
                        font_color = '#' + font_color[2:]
                if cell.font.bold:
                    font_weight = 'bold'
                if cell.font.size:
                    font_size = f'{cell.font.size}px'
            
            h_align = 'left'
            v_align = 'middle'
            wrap_text = False
            if cell.alignment:
                h_align = cell.alignment.horizontal or 'left'
                v_align = cell.alignment.vertical or 'middle'
                wrap_text = cell.alignment.wrap_text or False
            
            border_style = '1px solid #d0d0d0'
            
            style = f'background-color: {bg_color}; color: {font_color}; font-weight: {font_weight}; font-size: {font_size}; '
            style += f'text-align: {h_align}; vertical-align: {v_align}; padding: 2px 4px; border: {border_style}; '
            
            if wrap_text:
                style += 'white-space: normal; word-wrap: break-word; max-width: 300px; '
            else:
                style += 'white-space: nowrap; '
            
            rowspan = 1
            colspan = 1
            if (row, col) in merged_cells and merged_cells[(row, col)].get('is_master', False):
                rowspan = merged_cells[(row, col)].get('rowspan', 1)
                colspan = merged_cells[(row, col)].get('colspan', 1)
            
            if rowspan > 1 or colspan > 1:
                html += f'<td style="{style}" rowspan="{rowspan}" colspan="{colspan}">{str(value)}</td>'
            else:
                html += f'<td style="{style}">{str(value)}</td>'
        
        html += '</tr>'
    
    html += '</table>'
    return html

@st.cache_data(ttl=3600)
def load_data():
    source_data = download_file(SOURCE_URL)
    template_data = download_file(TEMPLATE_URL)
    
    if source_data is None or template_data is None:
        return None, None, None, None, None
    
    df = pd.read_excel(source_data)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    
    template_wb = load_workbook(template_data)
    template_sheet = template_wb.active
    placeholders = get_placeholders(template_sheet)
    
    return df, template_wb, template_sheet, placeholders, template_data

# --- LOAD DATA ---
with st.spinner("Loading..."):
    df, template_wb, template_sheet, placeholders, template_data = load_data()

if df is None or template_wb is None:
    st.error("Failed to load data. Please check your internet connection.")
    st.stop()

# --- CONTROLS ROW ---
trade_areas = sorted(df["TRADE AREA"].dropna().unique())

# Create columns for controls
col1, col2, col3, col4, col5, col6 = st.columns([1.2, 1.2, 0.8, 0.8, 0.8, 0.8])

with col1:
    selected_ta = st.selectbox(
        "Trade Area",
        options=trade_areas,
        index=0 if trade_areas else None,
        key="ta_select",
        label_visibility="collapsed"
    )

with col2:
    if selected_ta:
        sites_in_ta = df[df["TRADE AREA"] == selected_ta]["SITE NAME"].dropna().unique()
        sites_in_ta = sorted(sites_in_ta)
        selected_site = st.selectbox(
            "Site Name",
            options=sites_in_ta,
            index=0 if len(sites_in_ta) > 0 else None,
            key="site_select",
            label_visibility="collapsed"
        )
    else:
        selected_site = st.selectbox(
            "Site Name",
            options=[],
            key="site_select",
            label_visibility="collapsed"
        )

with col3:
    if st.button("Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with col4:
    if selected_ta and selected_site:
        site_data = df[(df["TRADE AREA"] == selected_ta) & (df["SITE NAME"] == selected_site)]
        if not site_data.empty:
            row = site_data.iloc[0]
            
            # Create single site download
            template_data.seek(0)
            wb = load_workbook(template_data)
            base_sheet = wb.active
            
            for row_cells in base_sheet.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, str) and "{{" in cell.value:
                        new_val = cell.value
                        for ph in placeholders:
                            target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                            if re.search(target_regex, new_val):
                                raw_data_val = row.get(ph.upper(), "")
                                if pd.isna(raw_data_val) or raw_data_val is None:
                                    raw_data_val = ""
                                val_str = str(raw_data_val)
                                new_val = re.sub(target_regex, val_str, new_val)
                        cell.value = new_val.strip() if new_val else ""
            
            wb_buffer = io.BytesIO()
            wb.save(wb_buffer)
            safe_filename = f"{selected_site}_{selected_ta}".replace("/", "-").replace("\\", "-")
            
            st.download_button(
                "Site Report",
                data=wb_buffer.getvalue(),
                file_name=f"{safe_filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

with col5:
    if selected_ta:
        # Create trade area download (all sites in trade area)
        ta_data = df[df["TRADE AREA"] == selected_ta]
        
        if st.button("Trade Report", use_container_width=True):
            with st.spinner("Generating..."):
                template_data.seek(0)
                wb = load_workbook(template_data)
                base_sheet = wb.active
                base_sheet.title = "TEMPLATE_TO_DELETE"
                existing_tabs = set()
                
                for _, row in ta_data.iterrows():
                    site_name = row.get("SITE NAME", "Unknown")
                    safe_tab_name = sanitize_tab_name(site_name, existing_tabs)
                    new_sheet = wb.copy_worksheet(base_sheet)
                    new_sheet.title = safe_tab_name
                    
                    for row_cells in new_sheet.iter_rows():
                        for cell in row_cells:
                            if isinstance(cell.value, str) and "{{" in cell.value:
                                new_val = cell.value
                                for ph in placeholders:
                                    target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                                    if re.search(target_regex, new_val):
                                        raw_data_val = row.get(ph.upper(), "")
                                        if pd.isna(raw_data_val) or raw_data_val is None:
                                            raw_data_val = ""
                                        val_str = str(raw_data_val)
                                        new_val = re.sub(target_regex, val_str, new_val)
                                cell.value = new_val.strip() if new_val else ""
                
                wb.remove(base_sheet)
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                safe_filename = str(selected_ta).replace("/", "-").replace("\\", "-")
                
                st.download_button(
                    "Download",
                    data=wb_buffer.getvalue(),
                    file_name=f"{safe_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="ta_download"
                )

def sanitize_tab_name(name, existing_names):
    illegal_chars = r'[\\/*?\[\]:]'
    clean_name = re.sub(illegal_chars, '', str(name))
    base_name = clean_name[:31]
    if base_name not in existing_names:
        existing_names.add(base_name)
        return base_name
    counter = 1
    while True:
        suffix = f" ({counter})"
        max_len = 31 - len(suffix)
        new_name = f"{clean_name[:max_len]}{suffix}"
        if new_name not in existing_names:
            existing_names.add(new_name)
            return new_name
        counter += 1

with col6:
    total_sites = len(df["SITE NAME"].dropna().unique())
    st.markdown(f"<p style='text-align: center; font-size: 0.7rem; color: #666; margin: 0; padding-top: 8px;'>Sites: {total_sites}</p>", unsafe_allow_html=True)

# --- REPORT VIEWER ---
if selected_ta and selected_site:
    try:
        site_data = df[(df["TRADE AREA"] == selected_ta) & (df["SITE NAME"] == selected_site)]
        
        if site_data.empty:
            st.warning("No data found for the selected site.")
        else:
            template_data.seek(0)
            wb = load_workbook(template_data)
            base_sheet = wb.active
            
            row = site_data.iloc[0]
            
            for row_cells in base_sheet.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, str) and "{{" in cell.value:
                        new_val = cell.value
                        has_injected = False
                        for ph in placeholders:
                            target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                            if re.search(target_regex, new_val):
                                raw_data_val = row.get(ph.upper(), "")
                                if pd.isna(raw_data_val) or raw_data_val is None:
                                    raw_data_val = ""
                                val_str = str(raw_data_val)
                                new_val = re.sub(target_regex, val_str, new_val)
                                if val_str.strip() != "":
                                    has_injected = True
                        if has_injected and new_val != "":
                            copy_and_merge_aware_injection(base_sheet, base_sheet, cell.coordinate, new_val.strip())
                        elif new_val == "":
                            cell.value = ""
                        else:
                            cell.value = new_val.strip() if new_val else ""
            
            html_content = render_excel_to_html(wb)
            st.markdown('<div class="excel-container">', unsafe_allow_html=True)
            st.markdown(html_content, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
    
    except Exception as e:
        st.error(f"Error: {str(e)}")
else:
    st.info("Select a Trade Area and Site to view the report.")
