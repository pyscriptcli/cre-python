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

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- HIDE STREAMLIT HEADER ---
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 0rem !important;
    }
    .stApp {
        background-color: #f5f5f5;
    }
    .main-header {
        background-color: #003366;
        padding: 1.5rem;
        border-radius: 0px;
        margin-bottom: 2rem;
        color: white;
        margin-top: -1rem;
    }
    .main-header h1 {
        color: white;
        font-weight: 600;
        margin: 0;
        font-size: 2rem;
    }
    .main-header p {
        color: #e6e6e6;
        margin: 0.5rem 0 0 0;
        font-size: 1rem;
    }
    .stButton > button {
        background-color: #003366;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 0.5rem 1rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #002244;
        color: white;
        box-shadow: 0 2px 8px rgba(0, 51, 102, 0.3);
    }
    .stButton > button:active {
        background-color: #001a33;
    }
    div[data-testid="stContainer"] {
        background-color: white;
        border-radius: 8px;
        padding: 1.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #e0e0e0;
    }
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid #003366;
    }
    .stAlert[data-baseweb="notification"] {
        background-color: #e8f0fe;
        border-left-color: #003366;
    }
    h1, h2, h3, h4 {
        color: #003366;
        font-weight: 600;
    }
    .stCheckbox label {
        font-weight: 500;
        color: #1a1a1a;
    }
    .stProgress > div > div {
        background-color: #003366;
    }
    hr {
        border-color: #003366;
        opacity: 0.2;
        margin: 2rem 0;
    }
    .stInfo {
        background-color: #e8f0fe;
        border-radius: 8px;
        border-left: 4px solid #003366;
    }
    .stSuccess {
        background-color: #d4edda;
        border-radius: 8px;
        border-left: 4px solid #28a745;
    }
    .stError {
        background-color: #f8d7da;
        border-radius: 8px;
        border-left: 4px solid #dc3545;
    }
    .stDownloadButton > button {
        background-color: #28a745;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    .stDownloadButton > button:hover {
        background-color: #218838;
        box-shadow: 0 2px 8px rgba(40, 167, 69, 0.3);
    }
    .metric-card {
        background-color: white;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #003366;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #003366;
    }
    .metric-label {
        color: #666;
        font-size: 0.9rem;
    }
    /* Hide the "Manage app" link */
    .stApp > header + div {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# --- CONFIGURATION ---
SOURCE_URL = "https://docs.google.com/spreadsheets/d/14nhO9u7zJRcOoux8I7l2IzwU7iQZNW9fRX6TCip47CE/export?format=xlsx"
TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1uS3xmnPi0o4c_EayQtURYDSMMPRDRGSb/export?format=xlsx"

# --- HARDCODED MAPPING ---
# This defines which template placeholders map to which data columns
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
        else:
            st.error(f"Failed to download. Status: {response.status_code}")
            return None
    except Exception as e:
        st.error(f"Download error: {str(e)}")
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
            num_val = float(val)
            return f"{num_val:,.2f}"
        elif mask_pattern == "DATE_SHORT":
            return pd.to_datetime(val).strftime("%m/%d/%Y")
        elif mask_pattern == "DATE_TIME_FULL":
            return pd.to_datetime(val).strftime("%m/%d/%Y %H:%M:%S")
        elif mask_pattern == "PERCENT":
            num_val = float(val)
            return f"{num_val:.2f}%"
        elif mask_pattern == "CURRENCY_USD":
            num_val = float(val)
            return f"${num_val:,.2f}"
        elif mask_pattern == "CURRENCY_PHP":
            num_val = float(val)
            return f"₱{num_val:,.2f}"
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
            except Exception:
                pass
                
        try:
            target_ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        except Exception:
            pass
            
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                sub_coord = f"{get_column_letter(c)}{r}"
                if sub_coord != coord:
                    clone_cell_styles(template_ws[sub_coord], target_ws[sub_coord])

# --- MAIN APP ---
st.markdown("""
<div class="main-header">
    <h1>Report Generator</h1>
    <p>Generate trade area reports from your data</p>
</div>
""", unsafe_allow_html=True)

# --- LOAD FILES ---
@st.cache_resource
def load_files():
    with st.spinner("Downloading files from Google Drive..."):
        source_data = download_file(SOURCE_URL)
        template_data = download_file(TEMPLATE_URL)
        return source_data, template_data

source_data, template_data = load_files()

if source_data is None or template_data is None:
    st.error("""
    ❌ **Failed to load files from Google Drive**
    
    **Please check:**
    1. Make sure the files are publicly accessible (Anyone with the link can view)
    2. Verify the file IDs are correct
    """)
    st.stop()

# --- Load and process data ---
df = pd.read_excel(source_data)
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df.columns = df.columns.str.strip().str.upper()
headers = list(df.columns)

if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
    st.error("ERROR: Raw data must contain 'TRADE AREA' and 'SITE NAME' columns.")
    st.stop()

template_wb = openpyxl.load_workbook(template_data)
template_sheet = template_wb.active
placeholders = get_placeholders(template_sheet)

if not placeholders:
    st.warning("No {{Placeholders}} found in the template.")
    st.stop()

# --- SESSION STATE ---
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None

# --- MAIN INTERFACE ---
# Metrics row
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{len(df)}</div>
        <div class="metric-label">Total Records</div>
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{len(df['TRADE AREA'].unique())}</div>
        <div class="metric-label">Trade Areas</div>
    </div>
    """, unsafe_allow_html=True)
with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{len(placeholders)}</div>
        <div class="metric-label">Placeholders Found</div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

st.markdown("### Select Trade Areas")

unique_tas = sorted([str(ta) for ta in df["TRADE AREA"].dropna().unique()])

# Select All / Clear All
col1, col2 = st.columns([1, 1])
with col1:
    if st.button("Select All", use_container_width=True):
        for ta in unique_tas:
            st.session_state[f"ta_{ta}"] = True
        st.rerun()
with col2:
    if st.button("Clear All", use_container_width=True):
        for ta in unique_tas:
            st.session_state[f"ta_{ta}"] = False
        st.rerun()

selected_tas = []
with st.container(height=200, border=True):
    for ta in unique_tas:
        if f"ta_{ta}" not in st.session_state:
            st.session_state[f"ta_{ta}"] = True
        if st.checkbox(ta, key=f"ta_{ta}"):
            selected_tas.append(ta)

st.divider()

# --- GENERATE REPORTS ---
if st.session_state.zip_data is None:
    if st.button("Generate Reports", use_container_width=True, type="primary"):
        if not selected_tas:
            st.warning("Please select at least one Trade Area.")
        else:
            with st.spinner("Generating reports..."):
                progress_bar = st.progress(0)
                
                filtered_df = df[df["TRADE AREA"].astype(str).isin(selected_tas)]
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
                                        
                                        # Check each placeholder in the cell
                                        for ph in placeholders:
                                            target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                                            if re.search(target_regex, new_val):
                                                # Find the mapping for this placeholder
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

# --- DOWNLOAD OPTIONS ---
if st.session_state.zip_data is not None:
    st.success("Reports generated successfully!")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            "Download Reports (.zip)",
            data=st.session_state.zip_data,
            file_name="Trade_Area_Reports.zip",
            mime="application/zip",
            use_container_width=True
        )
    
    with col2:
        st.info("""
        **Save to Google Drive:**
        1. Download the file above
        2. Go to your Google Drive folder:
           https://drive.google.com/drive/folders/1MAo_8VYditz-BV3vGx3aX31-SLzxSAD8
        3. Click 'New' -> 'File Upload' and select the downloaded zip
        """)
    
    if st.button("Start Over", use_container_width=True):
        st.session_state.zip_data = None
        st.rerun()
