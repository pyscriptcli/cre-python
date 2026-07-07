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
from PIL import Image, ImageDraw, ImageFont

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
    
    .photo-container {
        border: 1px solid #d0d0d0;
        margin-top: 0.5rem;
        background-color: #ffffff;
        padding: 4px;
        border-radius: 2px;
        width: 100%;
        text-align: center;
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

# Dynamically download true vector fonts into memory to ensure anti-aliased clean rendering in headless deployments
@st.cache_data(ttl=86400)
def download_vector_fonts():
    reg_url = "https://github.com/googlefonts/roboto/raw/main/src/hinted/Roboto-Regular.ttf"
    bold_url = "https://github.com/googlefonts/roboto/raw/main/src/hinted/Roboto-Bold.ttf"
    try:
        reg_font = requests.get(reg_url, timeout=10).content
        bold_font = requests.get(bold_url, timeout=10).content
        return io.BytesIO(reg_font), io.BytesIO(bold_font)
    except:
        return None, None

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
    if source_cell.border:
        target_cell.border = copy(source_cell.border)

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
        new_name = f".{clean_name[:27]} ({counter})"
        if new_name not in existing_names:
            existing_names.add(new_name)
            return new_name
        counter += 1

def parse_excel_color(openpyxl_color, default_rgb=(255, 255, 255)):
    if not openpyxl_color or openpyxl_color.type != 'rgb' or not openpyxl_color.rgb:
        return default_rgb
    rgb_str = str(openpyxl_color.rgb)
    if len(rgb_str) == 8: 
        rgb_str = rgb_str[2:]
    if len(rgb_str) == 6:
        return tuple(int(rgb_str[i:i+2], 16) for i in (0, 2, 4))
    return default_rgb

# --- PILLOW CRYSTAL-CLEAR RENDERING ENGINE (A1:P67) ---
def render_range_to_image(workbook, font_reg_bytes, font_bold_bytes, sheet_name=None, range_string="A1:P67"):
    ws = workbook[sheet_name] if sheet_name else workbook.active
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    
    # Scale width coordinates explicitly to high resolution pixel bounds
    col_widths = []
    for c in range(min_col, max_col + 1):
        w = ws.column_dimensions[get_column_letter(c)].width
        col_widths.append(int((w if w else 10) * 12)) # Heightened scaling factor for crisp details
        
    row_heights = []
    for r in range(min_row, max_row + 1):
        h = ws.row_dimensions[r].height
        row_heights.append(int((h if h else 20) * 1.8)) 
        
    img_width = sum(col_widths)
    img_height = sum(row_heights)
    
    img = Image.new('RGB', (img_width, img_height), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    col_positions = [0]
    for w in col_widths:
        col_positions.append(col_positions[-1] + w)
        
    row_positions = [0]
    for h in row_heights:
        row_positions.append(row_positions[-1] + h)
        
    merged_blocks = {}
    for merged_range in ws.merged_cells.ranges:
        m_min_c, m_min_r, m_max_c, m_max_r = range_boundaries(str(merged_range))
        if m_min_r >= min_row and m_max_r <= max_row and m_min_c >= min_col and m_max_c <= max_col:
            for r in range(m_min_r, m_max_r + 1):
                for c in range(m_min_c, m_max_c + 1):
                    if r == m_min_r and c == m_min_c:
                        merged_blocks[(r, c)] = (m_min_r, m_min_c, m_max_r, m_max_c)
                    else:
                        merged_blocks[(r, c)] = False 

    # Pass 1: Render Background Color Blocks
    for r_idx, r in enumerate(range(min_row, max_row + 1)):
        for c_idx, c in enumerate(range(min_col, max_col + 1)):
            if (r, c) in merged_blocks and merged_blocks[(r, c)] is False:
                continue
                
            if (r, c) in merged_blocks and merged_blocks[(r, c)]:
                mr_min, mc_min, mr_max, mc_max = merged_blocks[(r, c)]
                x1 = col_positions[mc_min - min_col]
                y1 = row_positions[mr_min - min_row]
                x2 = col_positions[mc_max - min_col + 1]
                y2 = row_positions[mr_max - min_row + 1]
            else:
                x1 = col_positions[c_idx]
                y1 = row_positions[r_idx]
                x2 = col_positions[c_idx + 1]
                y2 = row_positions[r_idx + 1]
                
            cell = ws.cell(row=r, column=c)
            bg_color = parse_excel_color(cell.fill.fgColor, (255, 255, 255)) if cell.fill and cell.fill.fill_type else (255, 255, 255)
            
            if bg_color == (0, 0, 0) and (cell.value is None or str(cell.value).strip() == ""):
                bg_color = (255, 255, 255)
                
            draw.rectangle([x1, y1, x2, y2], fill=bg_color)

    # Pass 2: Render Anti-Aliased Clean Text Layers & Outlines
    for r_idx, r in enumerate(range(min_row, max_row + 1)):
        for c_idx, c in enumerate(range(min_col, max_col + 1)):
            if (r, c) in merged_blocks and merged_blocks[(r, c)] is False:
                continue
                
            if (r, c) in merged_blocks and merged_blocks[(r, c)]:
                mr_min, mc_min, mr_max, mc_max = merged_blocks[(r, c)]
                x1 = col_positions[mc_min - min_col]
                y1 = row_positions[mr_min - min_row]
                x2 = col_positions[mc_max - min_col + 1]
                y2 = row_positions[mr_max - min_row + 1]
            else:
                x1 = col_positions[c_idx]
                y1 = row_positions[r_idx]
                x2 = col_positions[c_idx + 1]
                y2 = row_positions[r_idx + 1]

            cell = ws.cell(row=r, column=c)
            value = cell.value if cell.value is not None else ''
            
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            elif hasattr(value, 'strftime'):
                value = value.strftime('%B %d, %Y')
            val_str = str(value).strip()
            if re.match(r'^\d+\.0$', val_str):
                val_str = val_str.split('.')[0]
                
            draw.rectangle([x1, y1, x2, y2], outline=(190, 190, 190), width=1)
            
            if val_str:
                bg_color_check = parse_excel_color(cell.fill.fgColor, (255, 255, 255)) if cell.fill and cell.fill.fill_type else (255, 255, 255)
                
                font_color = (51, 51, 51)
                is_bold = False
                excel_font_size = 10
                
                if cell.font:
                    if cell.font.bold: is_bold = True
                    if cell.font.size: excel_font_size = cell.font.size
                    font_color = parse_excel_color(cell.font.color, (51, 51, 51))
                
                # Dynamic calculated text scaling matching resolution bounds
                scaled_font_size = max(int(excel_font_size * 1.35), 11)
                
                if bg_color_check in [(128, 0, 0), (140, 0, 0), (122, 0, 0)] or "SITE INFORMATION REPORT" in val_str.upper():
                    font_color = (255, 255, 255)
                    is_bold = True
                    h_align = "center"
                elif any(h in val_str.upper() for h in ["GENERAL INFORMATION", "LOCATION", "TERMS", "RATES", "TECHNICAL INFO", "PROVISIONS", "LESSOR AND TENANT DETAILS", "IF WITH SUB-LESSOR", "REGULATORY", "SITE ACQUIRABILITY"]):
                    font_color = (0, 0, 0)
                    is_bold = True
                    h_align = "left"
                else:
                    h_align = "left"
                    if cell.alignment and cell.alignment.horizontal:
                        h_align = cell.alignment.horizontal
                
                # Initialize standard vector font streams on the active canvas
                try:
                    font_source = font_bold_bytes if is_bold else font_reg_bytes
                    font_source.seek(0)
                    active_font = ImageFont.truetype(font_source, scaled_font_size)
                except:
                    active_font = ImageFont.load_default()
                
                # Fetch text bounding metrics to adjust structural text alignments
                bbox = draw.textbbox((0, 0), val_str, font=active_font)
                text_w = bbox[2] - bbox[0]
                text_h = bbox[3] - bbox[1]
                
                if h_align == "center":
                    tx = x1 + ((x2 - x1) - text_w) // 2
                elif h_align == "right":
                    tx = x2 - text_w - 8
                else:
                    tx = x1 + 8
                    
                ty = y1 + ((y2 - y1) - text_h) // 2 - 2
                draw.text((tx, ty), val_str, fill=font_color, font=active_font)
                
    return img

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
    # Download vector configurations
    font_reg_b, font_bold_b = download_vector_fonts()

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
                                        elif hasattr(raw_data_val, 'strftime'): val_str = r_row.get(ph.upper(), "").strftime('%B %d, %Y')
                                        else: val_str = str(raw_data_val)
                                        new_val = re.sub(target_regex, val_str, new_val)
                                cell.value = new_val.strip() if new_val else ""
                wb.remove(base_sheet)
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                st.download_button("Download Bulk", data=wb_buffer.getvalue(), file_name="Trade_Report.xlsx", use_container_width=True)

with col6:
    st.markdown(f"<p class='info-text'>Sites: {len(df['SITE NAME'].dropna().unique())}</p>", unsafe_allow_html=True)

# --- DIRECT PHOTO PREVIEW RENDER LAYER ---
if site_excel_bytes:
    try:
        active_wb = load_workbook(io.BytesIO(site_excel_bytes))
        
        # Pass vector components stream parameters straight into renderer
        cloned_image_canvas = render_range_to_image(
            active_wb, 
            io.BytesIO(font_reg_b) if font_reg_b else None, 
            io.BytesIO(font_bold_b) if font_bold_b else None, 
            range_string="A1:P67"
        )
        
        st.markdown('<div class="photo-container">', unsafe_allow_html=True)
        st.image(cloned_image_canvas, use_container_width=True, output_format="PNG")
        st.markdown('</div>', unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Error rendering image frame: {str(e)}")
else:
    st.info("Select a Trade Area and Site to view the report.")
