import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries
import re
import io
import zipfile
import requests
from copy import copy
import os

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="trs.sitesourcing report",
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
    /* Import Roboto from Google */
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    
    * {
        font-family: 'Roboto', 'Segoe UI', sans-serif !important;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Smaller header */
    .main-header {
        background-color: #003366 !important;
        padding: 0.3rem 1.2rem !important;
        border-radius: 4px;
        margin-bottom: 0.4rem !important;
    }
    .main-header h1 {
        color: white !important;
        font-weight: 500;
        margin: 0;
        font-size: 0.95rem !important;
        letter-spacing: 0.3px;
    }
    .main-header p {
        color: #cce0f5 !important;
        margin: 0;
        font-size: 0.65rem !important;
        font-weight: 300;
    }
    
    /* Button styling with better contrast */
    .stButton > button {
        background-color: #003366 !important;
        color: white !important;
        border: none;
        border-radius: 3px;
        padding: 0.25rem 0.6rem;
        font-weight: 400;
        font-size: 0.75rem;
        transition: all 0.2s ease;
        width: 100%;
        min-height: 32px;
    }
    .stButton > button:hover {
        background-color: #002244 !important;
        box-shadow: 0 2px 6px rgba(0, 51, 102, 0.3);
    }
    .stButton > button:active {
        background-color: #001a33 !important;
    }
    
    .stDownloadButton > button {
        background-color: #28a745 !important;
        color: white !important;
        border: none;
        border-radius: 3px;
        padding: 0.25rem 0.6rem;
        font-weight: 400;
        font-size: 0.75rem;
        width: 100%;
        min-height: 32px;
    }
    .stDownloadButton > button:hover {
        background-color: #218838 !important;
        box-shadow: 0 2px 6px rgba(40, 167, 69, 0.3);
    }
    
    /* Containers with better contrast */
    div[data-testid="stContainer"] {
        background-color: white !important;
        border-radius: 4px;
        padding: 0.6rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border: 1px solid #e0e0e0;
    }
    
    /* Headers with better contrast */
    h1, h2, h3, h4 {
        color: #003366 !important;
        font-weight: 500;
        font-size: 0.85rem;
        margin: 0 0 0.3rem 0;
    }
    
    /* Checkboxes with better contrast */
    .stCheckbox label {
        font-weight: 400;
        color: #1a1a1a !important;
        font-size: 0.8rem;
    }
    .stCheckbox label:hover {
        color: #003366 !important;
    }
    .stCheckbox {
        margin-bottom: 0.1rem;
    }
    .stCheckbox > div {
        background-color: white !important;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background-color: #003366 !important;
    }
    
    /* Divider */
    hr {
        border-color: #003366 !important;
        opacity: 0.15;
        margin: 0.4rem 0;
    }
    
    /* Metric cards with better contrast */
    .metric-card {
        background-color: white !important;
        border-radius: 4px;
        padding: 0.4rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        border-left: 3px solid #003366;
        text-align: center;
        margin-bottom: 0.3rem;
    }
    .metric-value {
        font-size: 1.1rem;
        font-weight: 500;
        color: #003366 !important;
        font-family: 'Roboto', sans-serif;
    }
    .metric-label {
        color: #555 !important;
        font-size: 0.6rem;
        font-weight: 400;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-family: 'Roboto', sans-serif;
    }
    
    /* Scrollable checkbox container */
    .checkbox-container {
        max-height: 280px;
        overflow-y: auto;
        padding-right: 4px;
        background-color: white !important;
    }
    .checkbox-container::-webkit-scrollbar {
        width: 4px;
    }
    .checkbox-container::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 2px;
    }
    .checkbox-container::-webkit-scrollbar-thumb {
        background: #003366;
        border-radius: 2px;
    }
    
    /* Alert boxes with better contrast */
    .stAlert {
        border-radius: 4px;
        padding: 0.4rem;
    }
    .stAlert[data-baseweb="notification"] {
        border-left-color: #003366 !important;
    }
    
    .stSuccess {
        background-color: #d4edda !important;
        color: #155724 !important;
    }
    .stSuccess * {
        color: #155724 !important;
    }
    
    .stWarning {
        background-color: #fff3cd !important;
        color: #856404 !important;
    }
    .stWarning * {
        color: #856404 !important;
    }
    
    .stError {
        background-color: #f8d7da !important;
        color: #721c24 !important;
    }
    .stError * {
        color: #721c24 !important;
    }
    
    .stInfo {
        background-color: #e8f0fe !important;
        color: #004085 !important;
    }
    .stInfo * {
        color: #004085 !important;
    }
    
    /* Spinner */
    .stSpinner > div {
        border-color: #003366 !important;
    }
    
    /* All text with good contrast */
    .stMarkdown, .stMarkdown * {
        color: #1a1a1a !important;
    }
    
    /* Select boxes */
    .stSelectbox > div > div {
        background-color: white !important;
    }
    .stSelectbox label {
        color: #003366 !important;
        font-weight: 400;
    }
    
    /* Force white backgrounds for all inputs */
    .stSelectbox > div > div > div {
        background-color: white !important;
    }
    .stMultiSelect > div > div {
        background-color: white !important;
    }
    .stTextInput > div > div > input {
        background-color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# --- CONFIGURATION ---
SOURCE_URL = "https://docs.google.com/spreadsheets/d/14nhO9u7zJRcOoux8I7l2IzwU7iQZNW9fRX6TCip47CE/export?format=xlsx"
TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1uS3xmnPi0o4c_EayQtURYDSMMPRDRGSb/export?format=xlsx"

# --- HARDCODED MAPPING ---
PLACEHOLDER_MAPPING = {
    "SITE NAME": {"column": "SITE NAME", "mask": "TEXT"},
    "TRADE AREA": {"column": "TRADE AREA", "mask": "TEXT"},
    "LEASE TYPE": {"column": "LEASE TYPE", "mask": "TEXT"},
    "SITE ID": {"column": "SITE ID", "mask": "TEXT"},
    "LOCATION/ADDRESS": {"column": "LOCATION/ADDRESS", "mask": "TEXT"},
    "CITY": {"column": "CITY", "mask": "TEXT"},
    "REGION": {"column": "REGION", "mask": "TEXT"},
    "POSTAL": {"column": "POSTAL", "mask": "TEXT"},
    "PARCEL AREA (SQM)": {"column": "PARCEL AREA (SQM)", "mask": "NUMBER"},
    "FLOOR AREA (SQM)": {"column": "FLOOR AREA (SQM)", "mask": "NUMBER"},
    "LEASE START": {"column": "LEASE START", "mask": "DATE_SHORT"},
    "LEASE END": {"column": "LEASE END", "mask": "DATE_SHORT"},
}

# --- HELPER FUNCTIONS ---
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

def format_with_mask(val, mask_pattern):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    try:
        if mask_pattern == "NUMBER":
            return f"{float(val):,.2f}"
        elif mask_pattern == "DATE_SHORT":
            return pd.to_datetime(val).strftime("%m/%d/%Y")
        elif mask_pattern == "DATE_TIME_FULL":
            return pd.to_datetime(val).strftime("%m/%d/%Y %H:%M:%S")
        elif mask_pattern == "PERCENT":
            return f"{float(val):.2f}%"
        elif mask_pattern == "CURRENCY_USD":
            return f"${float(val):,.2f}"
        elif mask_pattern == "CURRENCY_PHP":
            return f"₱{float(val):,.2f}"
        else:
            return str(val)
    except:
        return str(val)

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

# --- LOAD FILES ---
# Remove cache to always get latest files
def load_files():
    source_data = download_file(SOURCE_URL)
    template_data = download_file(TEMPLATE_URL)
    return source_data, template_data

# --- HEADER ---
st.markdown("""
<div class="main-header">
    <h1>trs.sitesourcing report</h1>
    <p>Generate trade area reports</p>
</div>
""", unsafe_allow_html=True)

# --- AUTO-REFRESH BUTTON ---
col_refresh, col_spacer = st.columns([1, 5])
with col_refresh:
    if st.button("Refresh Data", use_container_width=True):
        # Clear the cache and reload
        st.cache_data.clear()
        st.rerun()

# --- LOAD DATA ---
with st.spinner("Loading..."):
    source_data, template_data = load_files()

if source_data is None or template_data is None:
    st.error("Failed to load files. Make sure they are publicly accessible.")
    st.stop()

# --- PROCESS DATA ---
df = pd.read_excel(source_data)
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df.columns = df.columns.str.strip().str.upper()

if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
    st.error("Data must contain 'TRADE AREA' and 'SITE NAME' columns.")
    st.stop()

template_wb = openpyxl.load_workbook(template_data)
template_sheet = template_wb.active
placeholders = get_placeholders(template_sheet)

if not placeholders:
    st.warning("No placeholders found in template.")
    st.stop()

# --- SESSION STATE ---
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None
if 'selected_tas' not in st.session_state:
    st.session_state.selected_tas = []

# --- 3-COLUMN LAYOUT ---
col1, col2, col3 = st.columns([0.8, 1.4, 0.8])

# --- COLUMN 1: METRICS ---
with col1:
    st.markdown("### Stats")
    
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{len(df)}</div>
        <div class="metric-label">Records</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{len(df['TRADE AREA'].unique())}</div>
        <div class="metric-label">Trade Areas</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{len(placeholders)}</div>
        <div class="metric-label">Placeholders</div>
    </div>
    """, unsafe_allow_html=True)

# --- COLUMN 2: TRADE AREA SELECTION ---
with col2:
    st.markdown("### Select Trade Areas")
    
    unique_tas = sorted([str(ta) for ta in df["TRADE AREA"].dropna().unique()])
    
    # Select All / Clear All buttons
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("Select All", use_container_width=True):
            for ta in unique_tas:
                st.session_state[f"ta_{ta}"] = True
            st.session_state.selected_tas = unique_tas.copy()
            st.rerun()
    with btn_col2:
        if st.button("Clear All", use_container_width=True):
            for ta in unique_tas:
                st.session_state[f"ta_{ta}"] = False
            st.session_state.selected_tas = []
            st.rerun()
    
    # Initialize session state for checkboxes
    for ta in unique_tas:
        if f"ta_{ta}" not in st.session_state:
            st.session_state[f"ta_{ta}"] = True
    
    # Checkboxes in scrollable container
    st.markdown('<div class="checkbox-container">', unsafe_allow_html=True)
    selected_tas = []
    for ta in unique_tas:
        if st.checkbox(ta, key=f"ta_{ta}"):
            selected_tas.append(ta)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.session_state.selected_tas = selected_tas

# --- COLUMN 3: ACTIONS ---
with col3:
    st.markdown("### Actions")
    
    if st.session_state.zip_data is None:
        if st.button("Generate Reports", use_container_width=True, type="primary"):
            selected = st.session_state.selected_tas
            if not selected:
                st.warning("Select at least one Trade Area.")
            else:
                with st.spinner("Generating..."):
                    progress_bar = st.progress(0)
                    
                    filtered_df = df[df["TRADE AREA"].astype(str).isin(selected)]
                    groups = filtered_df.groupby("TRADE AREA")
                    total_groups = len(groups)
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for i, (trade_area, group) in enumerate(groups):
                            template_data.seek(0)
                            wb = openpyxl.load_workbook(template_data)
                            base_sheet = wb.active
                            base_sheet.title = "TEMPLATE_TO_DELETE"
                            existing_tabs = set()
                            
                            for _, row in group.iterrows():
                                site_name = row.get("SITE NAME", "Unknown")
                                safe_tab_name = sanitize_tab_name(site_name, existing_tabs)
                                new_sheet = wb.copy_worksheet(base_sheet)
                                new_sheet.title = safe_tab_name
                                
                                for row_cells in new_sheet.iter_rows():
                                    for cell in row_cells:
                                        if isinstance(cell.value, str) and "{{" in cell.value:
                                            new_val = cell.value
                                            has_injected = False
                                            for ph in placeholders:
                                                target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                                                if re.search(target_regex, new_val):
                                                    mapping = PLACEHOLDER_MAPPING.get(ph)
                                                    if mapping:
                                                        header = mapping["column"]
                                                        mask_patt = mapping["mask"]
                                                        raw_data_val = row.get(header) if header in row else None
                                                        val_str = format_with_mask(raw_data_val, mask_patt)
                                                        new_val = re.sub(target_regex, val_str, new_val)
                                                        if val_str.strip() != "":
                                                            has_injected = True
                                            if has_injected and new_val != "":
                                                copy_and_merge_aware_injection(base_sheet, new_sheet, cell.coordinate, new_val.strip())
                                            elif new_val == "":
                                                cell.value = ""
                                            else:
                                                cell.value = new_val.strip() if new_val else ""
                            
                            wb.remove(base_sheet)
                            wb_buffer = io.BytesIO()
                            wb.save(wb_buffer)
                            safe_filename = str(trade_area).replace("/", "-").replace("\\", "-")
                            zip_file.writestr(f"{safe_filename}.xlsx", wb_buffer.getvalue())
                            progress_bar.progress((i + 1) / total_groups)
                    
                    st.session_state.zip_data = zip_buffer.getvalue()
                    st.rerun()
    
    if st.session_state.zip_data is not None:
        st.success("Ready")
        st.download_button(
            "Download (.zip)",
            data=st.session_state.zip_data,
            file_name="Reports.zip",
            mime="application/zip",
            use_container_width=True
        )
        if st.button("Reset", use_container_width=True):
            st.session_state.zip_data = None
            st.rerun()
