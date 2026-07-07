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
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded"
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
    
    /* Strict Hiding */
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    button[title="View source"] {display: none !important;}
    .stAppDeployButton {display: none !important;}
    div[data-testid="stStatusWidget"] {display: none !important;}
    
    .main-header {
        display: none !important;
    }
    
    .block-container {
        padding-top: 0.3rem !important;
        padding-bottom: 1rem !important;
        max-width: 1400px !important;
    }
    
    .stButton > button {
        background-color: #e8e8e8 !important;
        color: #333333 !important;
        border: 1px solid #d0d0d0 !important;
        border-radius: 3px !important;
        padding: 0.4rem 0.6rem !important;
        font-weight: 500 !important;
        font-size: 0.8rem !important;
        transition: all 0.2s ease;
        width: 100%;
        min-height: 36px;
    }
    .stButton > button:hover {
        background-color: #f5f5f5 !important;
        border-color: #b0b0b0 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
    }
    
    .stSelectbox > div > div {
        background-color: #fafafa !important;
        border-color: #d0d0d0 !important;
        color: #333333 !important;
    }
    
    .stSelectbox label {
        color: #333333 !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
    }
    
    .stMarkdown, .stMarkdown * {
        color: #333333 !important;
    }
    
    .sidebar-section {
        background-color: #f5f5f5 !important;
        border-radius: 4px;
        padding: 1rem;
        margin-bottom: 1rem;
        border: 1px solid #e8e8e8;
    }
    
    .metric-card {
        background-color: #f5f5f5 !important;
        border-radius: 4px;
        padding: 0.4rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        border-left: 3px solid #999999;
        text-align: center;
        margin-bottom: 0.3rem;
    }
    .metric-value {
        font-size: 1.1rem;
        font-weight: 500;
        color: #333333 !important;
    }
    .metric-label {
        color: #666666 !important;
        font-size: 0.6rem;
        font-weight: 400;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .stAlert {
        border-radius: 4px;
        padding: 0.4rem;
    }
    
    /* Excel Viewer Container */
    .excel-container {
        background-color: white !important;
        border-radius: 4px;
        padding: 1rem;
        border: 1px solid #e8e8e8;
        overflow: auto;
        min-height: 600px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
        background-color: #f5f5f5 !important;
        border-radius: 4px;
        padding: 4px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 3px !important;
        padding: 0.4rem 1rem !important;
        background-color: transparent !important;
        color: #666666 !important;
        font-weight: 400 !important;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #333333 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08) !important;
        font-weight: 500 !important;
    }
    
    /* Footer overlay */
    .footer-overlay {
        position: fixed !important;
        bottom: 0 !important;
        left: 0 !important;
        right: 0 !important;
        height: 35px !important;
        background-color: #fafafa !important;
        z-index: 9999999 !important;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.03) !important;
        border-top: 1px solid #e8e8e8 !important;
        pointer-events: none !important;
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
    """Copy styles and handle merged cells for a single cell injection"""
    if not target_ws:
        return
    target_cell = target_ws[coord]
    template_cell = template_ws[coord]
    target_cell.value = data_value
    clone_cell_styles(template_cell, target_cell)
    
    # Handle merged cells
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
        
        # Clone styles for all cells in merged range
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                sub_coord = f"{get_column_letter(c)}{r}"
                if sub_coord != coord:
                    clone_cell_styles(template_ws[sub_coord], target_ws[sub_coord])

def render_excel_to_html(workbook, sheet_name=None):
    """Convert Excel worksheet to HTML for display"""
    if sheet_name is None:
        ws = workbook.active
    else:
        ws = workbook[sheet_name]
    
    # Get max row and column
    max_row = ws.max_row
    max_col = ws.max_column
    
    html = '<table style="border-collapse: collapse; font-family: \'Roboto\', sans-serif; font-size: 12px; width: 100%;">'
    
    # Build merged cell map
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
    
    # Generate table
    for row in range(1, max_row + 1):
        html += '<tr>'
        for col in range(1, max_col + 1):
            # Skip cells that are part of a merged range but not the master
            if (row, col) in merged_cells and not merged_cells[(row, col)].get('is_master', False):
                continue
            
            cell = ws.cell(row, col)
            value = cell.value if cell.value is not None else ''
            
            # Get cell styles
            bg_color = 'white'
            if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                bg_color = cell.fill.fgColor.rgb
                if bg_color.startswith('FF'):
                    bg_color = '#' + bg_color[2:]
            
            font_color = '#333333'
            font_weight = 'normal'
            font_size = '12px'
            if cell.font:
                if cell.font.color and cell.font.color.rgb:
                    font_color = cell.font.color.rgb
                    if font_color.startswith('FF'):
                        font_color = '#' + font_color[2:]
                if cell.font.bold:
                    font_weight = 'bold'
                if cell.font.size:
                    font_size = f'{cell.font.size}px'
            
            # Get alignment
            h_align = 'left'
            v_align = 'middle'
            wrap_text = False
            if cell.alignment:
                h_align = cell.alignment.horizontal or 'left'
                v_align = cell.alignment.vertical or 'middle'
                wrap_text = cell.alignment.wrap_text or False
            
            # Border
            border_style = '1px solid #d0d0d0'
            
            # Build style
            style = f'background-color: {bg_color}; color: {font_color}; font-weight: {font_weight}; font-size: {font_size}; '
            style += f'text-align: {h_align}; vertical-align: {v_align}; padding: 6px 8px; border: {border_style}; '
            
            if wrap_text:
                style += 'white-space: normal; word-wrap: break-word; max-width: 300px; '
            else:
                style += 'white-space: nowrap; '
            
            # Check if this cell is the master of a merged range
            rowspan = 1
            colspan = 1
            if (row, col) in merged_cells and merged_cells[(row, col)].get('is_master', False):
                rowspan = merged_cells[(row, col)].get('rowspan', 1)
                colspan = merged_cells[(row, col)].get('colspan', 1)
            
            # Build cell tag
            if rowspan > 1 or colspan > 1:
                html += f'<td style="{style}" rowspan="{rowspan}" colspan="{colspan}">{str(value)}</td>'
            else:
                html += f'<td style="{style}">{str(value)}</td>'
        
        html += '</tr>'
    
    html += '</table>'
    return html

@st.cache_data(ttl=3600)
def load_data():
    """Load and cache data from Google Sheets"""
    source_data = download_file(SOURCE_URL)
    template_data = download_file(TEMPLATE_URL)
    
    if source_data is None or template_data is None:
        return None, None, None, None, None
    
    # Load source data
    df = pd.read_excel(source_data)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    
    # Load template
    template_wb = load_workbook(template_data)
    template_sheet = template_wb.active
    placeholders = get_placeholders(template_sheet)
    
    return df, template_wb, template_sheet, placeholders, template_data

# --- LOAD DATA ---
with st.spinner("Loading data..."):
    df, template_wb, template_sheet, placeholders, template_data = load_data()

if df is None or template_wb is None:
    st.error("Failed to load data. Please check your internet connection and try again.")
    st.stop()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("### 📊 Viewer Controls")
    st.markdown("---")
    
    # Get unique trade areas
    trade_areas = sorted(df["TRADE AREA"].dropna().unique())
    
    # Check for TRADE AREA NO column
    trade_area_no_col = None
    for col in df.columns:
        if "TRADE AREA NO" in col.upper() or "TRADE AREA #" in col.upper():
            trade_area_no_col = col
            break
    
    # Trade Area dropdown
    selected_ta = st.selectbox(
        "Select Trade Area",
        options=trade_areas,
        index=0 if trade_areas else None,
        key="ta_select"
    )
    
    if selected_ta:
        # Filter sites for selected trade area
        sites_in_ta = df[df["TRADE AREA"] == selected_ta]["SITE NAME"].dropna().unique()
        sites_in_ta = sorted(sites_in_ta)
        
        # Site Name dropdown
        selected_site = st.selectbox(
            "Select Site",
            options=sites_in_ta,
            index=0 if len(sites_in_ta) > 0 else None,
            key="site_select"
        )
        
        st.markdown("---")
        
        # Display information about selected site
        if selected_site:
            site_data = df[(df["TRADE AREA"] == selected_ta) & (df["SITE NAME"] == selected_site)]
            if not site_data.empty:
                st.markdown("### 📋 Site Details")
                
                # Show key fields
                info_fields = ["SITE NAME", "ADDRESS", "CITY", "STATE", "ZIP", "PHONE"]
                for field in info_fields:
                    if field in df.columns:
                        col_upper = field.upper()
                        if col_upper in df.columns:
                            val = site_data[col_upper].iloc[0]
                            if pd.notna(val) and val != "":
                                st.markdown(f"**{field}:** {val}")
                
                # Show TRADE AREA NO if available
                if trade_area_no_col:
                    ta_no = site_data[trade_area_no_col].iloc[0]
                    if pd.notna(ta_no):
                        st.markdown(f"**Trade Area #:** {ta_no}")
    
    st.markdown("---")
    st.markdown("### 📈 Statistics")
    
    # Stats
    total_sites = len(df["SITE NAME"].dropna().unique())
    total_tas = len(trade_areas)
    total_placeholders = len(placeholders)
    
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{total_tas}</div>
        <div class="metric-label">Trade Areas</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{total_sites}</div>
        <div class="metric-label">Total Sites</div>
    </div>
    <div class="metric-card">
        <div class="metric-value">{total_placeholders}</div>
        <div class="metric-label">Placeholders</div>
    </div>
    """, unsafe_allow_html=True)

# --- MAIN CONTENT ---
st.markdown("### 👁️ Report Viewer")

if selected_ta and selected_site:
    try:
        # Get the data for the selected site
        site_data = df[(df["TRADE AREA"] == selected_ta) & (df["SITE NAME"] == selected_site)]
        
        if site_data.empty:
            st.warning("No data found for the selected site.")
        else:
            # Create a fresh workbook from template
            template_data.seek(0)
            wb = load_workbook(template_data)
            base_sheet = wb.active
            base_sheet.title = "Report"
            
            # Populate with site data
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
            
            # Display the report
            st.markdown("---")
            
            # Create tabs for different views
            tab1, tab2, tab3 = st.tabs(["📄 Report View", "📊 Data View", "ℹ️ Placeholders"])
            
            with tab1:
                # Render Excel to HTML
                html_content = render_excel_to_html(wb, "Report")
                st.markdown('<div class="excel-container">', unsafe_allow_html=True)
                st.markdown(html_content, unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Download button
                wb_buffer = io.BytesIO()
                wb.save(wb_buffer)
                safe_filename = f"{selected_site}_{selected_ta}".replace("/", "-").replace("\\", "-")
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        "⬇️ Download Excel File",
                        data=wb_buffer.getvalue(),
                        file_name=f"{safe_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            with tab2:
                # Show the raw data for the selected site
                st.markdown("### Site Data")
                
                # Display all data for this site in a clean format
                display_data = []
                for col in df.columns:
                    val = row.get(col, "")
                    if pd.isna(val):
                        val = ""
                    display_data.append({"Field": col, "Value": val})
                
                display_df = pd.DataFrame(display_data)
                st.dataframe(
                    display_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Field": st.column_config.TextColumn("Field", width="medium"),
                        "Value": st.column_config.TextColumn("Value", width="large")
                    }
                )
            
            with tab3:
                # Show all placeholders and their values
                st.markdown("### Placeholder Mapping")
                
                placeholder_data = []
                for ph in placeholders:
                    val = row.get(ph.upper(), "")
                    if pd.isna(val):
                        val = ""
                    placeholder_data.append({
                        "Placeholder": f"{{{{{ph}}}}}",
                        "Value": val,
                        "Status": "✅ Populated" if str(val).strip() != "" else "❌ Empty"
                    })
                
                ph_df = pd.DataFrame(placeholder_data)
                st.dataframe(
                    ph_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Placeholder": st.column_config.TextColumn("Placeholder", width="small"),
                        "Value": st.column_config.TextColumn("Value", width="large"),
                        "Status": st.column_config.TextColumn("Status", width="small")
                    }
                )
                
                # Summary
                populated = sum(1 for p in placeholder_data if p["Status"] == "✅ Populated")
                total = len(placeholder_data)
                st.markdown(f"**Summary:** {populated}/{total} placeholders populated")
    
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")
        st.info("Please try selecting a different site or trade area.")
else:
    st.info("👈 Select a Trade Area and Site from the sidebar to view the report.")

# Optional: Refresh button at bottom
if st.button("🔄 Refresh Data", use_container_width=True):
    st.cache_data.clear()
    st.rerun()
