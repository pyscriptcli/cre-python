import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import range_boundaries
import re
import io
import zipfile
import difflib
import math
import requests
from copy import copy

# Must be the first Streamlit command
st.set_page_config(page_title="Report Generator", layout="centered")

# --- CORE HELPER FUNCTIONS ---
def get_placeholders(sheet):
    """Extract all {{Placeholder}} variables from the Excel sheet."""
    placeholders = set()
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                placeholders.update(matches)
    return sorted(list(placeholders))

def get_placeholder_coords(sheet):
    """Map placeholders to their exact cell coordinates."""
    mapping = {}
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                for match in matches:
                    if match not in mapping:
                        mapping[match] = []
                    mapping[match].append(cell.coordinate)
    return mapping

def sanitize_tab_name(name, existing_names):
    """Strip illegal Excel characters, slice to 31 chars, and handle duplicates."""
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

def format_with_mask(val, mask_pattern, placeholder_name):
    """Applies cell-level transformation rules mirroring standard spreadsheet mask criteria."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    
    try:
        num_val = float(val) if isinstance(val, (int, float, str)) and re.match(r'^-?\d+(\.\d+)?$', str(val).strip()) else None
        
        if mask_pattern == "NUMBER" and num_val is not None:
            return f"{num_val:,.2f}"
        elif mask_pattern == "SCIENTIFIC" and num_val is not None:
            return f"{num_val:.2E}"
        elif mask_pattern == "ACCOUNTING" and num_val is not None:
            return f"$ ({num_val:,.2f})" if num_val < 0 else f"$ {num_val:,.2f}"
        elif mask_pattern == "FINANCIAL" and num_val is not None:
            return f"({num_val:,.2f})" if num_val < 0 else f"{num_val:,.2f}"
        elif mask_pattern == "CURRENCY_USD" and num_val is not None:
            return f"${num_val:,.2f}"
        elif mask_pattern == "CURRENCY_USD_ROUND" and num_val is not None:
            return f"${round(num_val):,}"
        elif mask_pattern == "CURRENCY_PHP" and num_val is not None:
            return f"₱{num_val:,.2f}"
        elif mask_pattern == "CURRENCY_PHP_ROUND" and num_val is not None:
            return f"₱{round(num_val):,}"
        elif mask_pattern == "PERCENT" and num_val is not None:
            factor = 1 if num_val > 1 or "%" in str(val) else 100
            return f"{num_val * factor:.2f}%"
            
        elif mask_pattern == "DATE_SHORT":
            return pd.to_datetime(val).strftime("%m/%d/%Y")
        elif mask_pattern == "TIME_STANDARD":
            return pd.to_datetime(val).strftime("%I:%M:%S %p")
        elif mask_pattern == "DATE_TIME_FULL":
            return pd.to_datetime(val).strftime("%m/%d/%Y %H:%M:%S")
        elif mask_pattern == "%B %d, %Y":
            return pd.to_datetime(val).strftime("%B %d, %Y")
        elif mask_pattern == "%Y-%m-%d":
            return pd.to_datetime(val).strftime("%Y-%m-%d")
        elif mask_pattern == "%d %b %Y":
            return pd.to_datetime(val).strftime("%d %b %Y")
    except Exception:
        pass
        
    if isinstance(val, str):
        if mask_pattern == "STREET_SEGMENT":
            p = [part.strip() for part in val.split(",")]
            return ", ".join(p[:max(0, len(p) - 6)]) if len(p) >= 6 else val
        elif mask_pattern == "BARANGAY_SEGMENT":
            p = [part.strip() for part in val.split(",")]
            return p[len(p) - 6] if len(p) >= 6 else ""
        elif mask_pattern == "CITY_SEGMENT":
            p = [part.strip() for part in val.split(",")]
            return p[len(p) - 5] if len(p) >= 5 else ""
        elif mask_pattern == "REGION_SEGMENT":
            p = [part.strip() for part in val.split(",")]
            return p[len(p) - 3] if len(p) >= 3 else ""
        elif mask_pattern == "POSTAL_SEGMENT":
            p = [part.strip() for part in val.split(",")]
            return p[len(p) - 2] if len(p) >= 2 else ""
            
    return str(val)

def generate_mock_value(mask_key):
    """Generates a dynamic baseline sample string to drive the real-time UI preview engine."""
    mock_registry = {
        "TEXT": "PRIME Philippines Core Workspace",
        "NUMBER": "1250.75",
        "PERCENT": "0.885",
        "SCIENTIFIC": "4520000",
        "ACCOUNTING": "7500.50",
        "FINANCIAL": "-1500.25",
        "CURRENCY_USD": "5200.50",
        "CURRENCY_USD_ROUND": "5200.50",
        "CURRENCY_PHP": "85400.65",
        "CURRENCY_PHP_ROUND": "85400.65",
        "DATE_SHORT": "2026-06-01 17:00:00",
        "TIME_STANDARD": "2026-06-01 17:00:00",
        "DATE_TIME_FULL": "2026-06-01 17:00:00",
        "%B %d, %Y": "2026-06-01 17:00:00",
        "%d %b %Y": "2026-06-01 17:00:00",
        "%Y-%m-%d": "2026-06-01 17:00:00",
        "STREET_SEGMENT": "Suite 401, Fortune Building, Pasig, Metro Manila, 1600, Philippines",
        "BARANGAY_SEGMENT": "Suite 401, Fortune Building, Pasig, Metro Manila, 1600, Philippines",
        "CITY_SEGMENT": "Suite 401, Fortune Building, Pasig, Metro Manila, 1600, Philippines"
    }
    return mock_registry.get(mask_key, "Sample String Value")

def inject_image_calibrated(target_sheet, cell_coord, file_path_str, media_dict, max_size, col_offset, row_offset):
    """Injects an image applying explicit width constraints and positional cell anchors."""
    if isinstance(file_path_str, str) and any(ext in file_path_str.lower() for ext in ['.jpg', '.jpeg', '.png']):
        filename = file_path_str.replace('\\', '/').split('/')[-1]
        
        if filename in media_dict:
            try:
                img_stream = io.BytesIO(media_dict[filename].getvalue())
                img = OpenpyxlImage(img_stream)
                
                if img.width > max_size or img.height > max_size:
                    ratio = min(max_size / img.width, max_size / img.height)
                    img.width = int(img.width * ratio)
                    img.height = int(img.height * ratio)
                
                if col_offset != 0 or row_offset != 0:
                    current_col = openpyxl.utils.column_index_from_string(re.sub(r'\d+', '', cell_coord))
                    current_row = int(re.sub(r'\D+', '', cell_coord))
                    target_coord = f"{openpyxl.utils.get_column_letter(current_col + col_offset)}{current_row + row_offset}"
                else:
                    target_coord = cell_coord
                    
                target_sheet.add_image(img, target_coord)
                return True 
            except Exception:
                pass 
    return False

def validate_public_url(url_string):
    """Executes a network validation pass to guarantee remote assets are accessible."""
    if not url_string or not isinstance(url_string, str):
        return False, "Invalid URL input parameters."
    if "drive.google.com" in url_string or "onedrive.live.com" in url_string or "sharepoint.com" in url_string:
        return True, "Cloud Storage directory endpoint detected."
    if not (url_string.startswith("http://") or url_string.startswith("https://")):
        return False, "Missing target protocol header (http/https)."
    try:
        response = requests.head(url_string, allow_redirects=True, timeout=5)
        if response.status_code == 200:
            return True, "URL confirmed public and responding cleanly (Status 200)."
        else:
            return False, f"Server responded with connection error code: {response.status_code}"
    except Exception as e:
        return False, f"Network Handshake Failure: {str(e)}"

def resolve_file_source(uploader_obj, link_str):
    """Unified engine routing data from local storage blocks or remote URL streams."""
    if uploader_obj is not None:
        return uploader_obj
        
    if link_str and link_str.strip():
        url = link_str.strip()
        
        if "drive.google.com" in url:
            file_match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
            if file_match:
                file_id = file_match.group(1)
                url = f"https://docs.google.com/uc?export=download&id={file_id}"
            else:
                st.info("📂 Mapping live cloud folder indexing vectors...")
        elif "onedrive.live.com" in url:
            url = url.replace("redir?", "download?").replace("1drv.ms", "1drv.ms/u")

        is_valid, msg = validate_public_url(url)
        if is_valid:
            try:
                res = requests.get(url, timeout=10)
                if res.status_code == 200:
                    return io.BytesIO(res.content)
            except Exception as e:
                st.error(f"Download stream pipeline crashed: {str(e)}")
    return None

def clone_cell_styles(source_cell, target_cell):
    """Deep clones style objects safely from source cell to target cell to prevent openpyxl hashing collisions."""
    if source_cell.font:
        target_cell.font = Font(
            name=source_cell.font.name,
            size=source_cell.font.size,
            bold=source_cell.font.bold,
            italic=source_cell.font.italic,
            charset=source_cell.font.charset,
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
    """
    Finds if the coordinate was part of a merged cell range in the template.
    Purges overlapping dirty target merges natively before executing clean matrix updates.
    """
    target_cell = target_ws[coord]
    template_cell = template_ws[coord]
    
    # Write value payload
    target_cell.value = data_value
    
    # Apply baseline styles to root coordinate cell
    clone_cell_styles(template_cell, target_cell)
    
    # Check if this target coordinates intersect a structural block inside template maps
    merged_range_string = None
    for merged_range in template_ws.merged_cells.ranges:
        if template_cell.coordinate in merged_range:
            merged_range_string = str(merged_range)
            break
            
    if merged_range_string:
        min_col, min_row, max_col, max_row = range_boundaries(merged_range_string)
        
        # --- NATIVE MERGE PURGE ROUTINE ---
        # Scan the target file for overlapping ranges to prevent Excel's xml tracking engine from corrupting
        overlapping_target_ranges = []
        for target_range in target_ws.merged_cells.ranges:
            t_min_c, t_min_r, t_max_c, t_max_r = range_boundaries(str(target_range))
            # If coordinates conflict or touch, clear them entirely
            if not (max_col < t_min_c or min_col > t_max_c or max_row < t_min_r or min_row > t_max_r):
                overlapping_target_ranges.append(target_range)
                
        # Unmerge any toxic ranges found inside target instances first
        for target_conflict in overlapping_target_ranges:
            try:
                target_ws.unmerge_cells(str(target_conflict))
            except Exception:
                pass
                
        # Execute clean array layout merge pass safely
        try:
            target_ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        except Exception:
            pass
            
        # Distribute style layouts down across sub-cells inside block envelope arrays
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                sub_coord = f"{openpyxl.utils.get_column_letter(c)}{r}"
                if sub_coord != coord:
                    clone_cell_styles(template_ws[sub_coord], target_ws[sub_coord])

class StreamWrapper:
    def __init__(self, data, filename):
        self.data = data
        self.name = filename
    def getvalue(self): 
        return self.data

# --- SESSION STATE INITIALIZATION ---
if "zip_data" not in st.session_state: st.session_state.zip_data = None
if "change_log" not in st.session_state: st.session_state.change_log = None

# --- UI: THE LANDING PAGE ---
st.title("Site Information Report")

mode = st.radio("Select Workflow Mode:", ["Create New Reports", "Edit / Update Existing Reports"], horizontal=True)
st.divider()

HUMAN_SPREADSHEET_MASKS = {
    "Plain text": "TEXT",
    "Number (1,000.12)": "NUMBER",
    "Percentage (10.12%)": "PERCENT",
    "Scientific (1.01E+03)": "SCIENTIFIC",
    "Accounting ($ 1,000.12)": "ACCOUNTING",
    "Financial (1,000.12)": "FINANCIAL",
    "Currency ($1,000.12)": "CURRENCY_USD",
    "Currency Rounded ($1,000)": "CURRENCY_USD_ROUND",
    "Philippine Peso (₱1,000.12)": "CURRENCY_PHP",
    "Philippine Peso Rounded (₱1,000)": "CURRENCY_PHP_ROUND",
    "Date (06/01/2026)": "DATE_SHORT",
    "Time (05:00:00 PM)": "TIME_STANDARD",
    "Date & Time Full": "DATE_TIME_FULL",
    "Date: June 01, 2026": "%B %d, %Y",
    "Date: 01 Jun 26": "%d %b %Y",
    "Date: YYYY-MM-DD": "%Y-%m-%d",
    "Address: Extract Street": "STREET_SEGMENT",
    "Address: Extract Barangay": "BARANGAY_SEGMENT",
    "Address: Extract City": "CITY_SEGMENT"
}

# ==========================================
# MODE A: CREATE NEW REPORTS
# ==========================================
if mode == "Create New Reports":
    st.markdown("### Upload Files & Data Targets")
    
    m_row1_col1, m_row1_col2 = st.columns(2)
    m_row2_col1, m_row2_col2 = st.columns(2)
    
    with m_row1_col1:
        with st.container(border=True):
            st.markdown("📊 **1. New Raw Data**")
            mode_a_type_1 = st.segmented_control("Source Type A1", ["File Upload", "Remote Link"], default="File Upload", key="mode_a_type_1", label_visibility="collapsed")
            if mode_a_type_1 == "File Upload":
                raw_file = st.file_uploader("Upload Data Sheet A", type=["xlsx", "xls"], key="new_raw", label_visibility="collapsed")
                raw_url = None
            else:
                raw_url = st.text_input("Data Link URL A", placeholder="https://example.com/data.xlsx", key="new_raw_url", label_visibility="collapsed")
                raw_file = None

    with m_row1_col2:
        with st.container(border=True):
            st.markdown("📐 **2. Excel Template**")
            mode_a_type_2 = st.segmented_control("Source Type A2", ["File Upload", "Remote Link"], default="File Upload", key="mode_a_type_2", label_visibility="collapsed")
            if mode_a_type_2 == "File Upload":
                template_file = st.file_uploader("Upload Template File A", type=["xlsx"], key="new_temp", label_visibility="collapsed")
                template_url = None
            else:
                template_url = st.text_input("Template URL A", placeholder="https://example.com/template.xlsx", key="new_temp_url", label_visibility="collapsed")
                template_file = None

    with m_row2_col1:
        with st.container(border=True):
            st.markdown("📸 **3. Photos / Docs (Folder / Link)**")
            mode_a_type_3 = st.segmented_control("Source Type A3", ["File Upload", "Remote Link"], default="File Upload", key="mode_a_type_3", label_visibility="collapsed")
            if mode_a_type_3 == "File Upload":
                media_files = st.file_uploader("Upload Images A", accept_multiple_files=True, key="new_media", type=['png', 'jpg', 'jpeg'], label_visibility="collapsed")
                media_url = None
            else:
                media_url = st.text_input("Cloud Drive Directory Link URL A", placeholder="Paste Google Drive or OneDrive Shared Folder Link", key="new_media_url", label_visibility="collapsed")
                media_files = None

    with m_row2_col2:
        st.empty()

    resolved_raw = resolve_file_source(raw_file, raw_url)
    resolved_template = resolve_file_source(template_file, template_url)

    media_dict = {}
    if media_files:
        media_dict = {f.name: f for f in media_files}
    elif media_url and media_url.strip():
        resolved_media_data = resolve_file_source(None, media_url)
        if resolved_media_data:
            try:
                with zipfile.ZipFile(resolved_media_data) as z:
                    for name in z.namelist():
                        if any(ext in name.lower() for ext in ['.png', '.jpg', '.jpeg']) and not name.split('/')[-1].startswith('.'):
                            clean_name = name.split('/')[-1]
                            media_dict[clean_name] = StreamWrapper(z.read(name), clean_name)
            except zipfile.BadZipFile:
                media_dict["photo1.jpg"] = StreamWrapper(resolved_media_data.getvalue(), "photo1.jpg")

    if resolved_raw and resolved_template:
        df = pd.read_excel(resolved_raw)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.columns = df.columns.str.strip().str.upper()
        headers = list(df.columns)
        
        if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
            st.error("ERROR: Raw data must contain exactly 'TRADE AREA' and 'SITE NAME' columns.")
            st.stop()
        
        template_wb = openpyxl.load_workbook(resolved_template)
        template_sheet = template_wb.active
        placeholders = get_placeholders(template_sheet)
        
        if not placeholders:
            st.warning("No {{Placeholders}} found in the uploaded template.")
        else:
            st.markdown("### Data Mapping & Advanced Mask Configuration")
            mapping = {}
            for ph in placeholders:
                match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
                default_index = headers.index(match[0]) if match else 0
                if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and "LOCATION/ADDRESS" in headers:
                    default_index = headers.index("LOCATION/ADDRESS")
                
                with st.container(border=True):
                    col1, col2, col3 = st.columns([1, 1.2, 1.2])
                    with col1: st.markdown(f"**{{{{{ph}}}}}**")
                    with col2: 
                        sel_col = st.selectbox("Map to column:", headers, index=default_index, key=f"map_{ph}", label_visibility="collapsed")
                    with col3:
                        sel_mask = st.selectbox("Format Mask Layout", list(HUMAN_SPREADSHEET_MASKS.keys()), index=0, key=f"mask_{ph}", label_visibility="collapsed")
                    
                    mask_id = HUMAN_SPREADSHEET_MASKS[sel_mask]
                    mock_seed = generate_mock_value(mask_id)
                    evaluated_preview = format_with_mask(mock_seed, mask_id, ph)
                    
                    pv_1, pv_2 = st.columns([1, 2])
                    with pv_1:
                        st.markdown(f"    ↳ `<small>Live Output Visual:</small>`", unsafe_allow_html=True)
                    with pv_2:
                        st.markdown(f"`{evaluated_preview}`")
                    
                    st.markdown(f"<div style='text-align: right; opacity: 0.35; font-size: 10px; font-weight: bold;'>[Source Column: {sel_col}] ──► [Injected Token Var: {{{{ {ph} }}}}]</div>", unsafe_allow_html=True)
                    
                    mapping[ph] = {"column": sel_col, "mask": mask_id}
            
            st.divider()
            st.markdown("### Select Trade Areas")
            unique_tas = sorted([str(ta) for ta in df["TRADE AREA"].dropna().unique()])
            
            col_sel, col_clr, _ = st.columns([1, 1, 3])
            if col_sel.button("Select All", key="sa1", use_container_width=True):
                for ta in unique_tas: st.session_state[f"chk_new_{ta}"] = True
                st.rerun()
            if col_clr.button("Clear All", key="ca1", use_container_width=True):
                for ta in unique_tas: st.session_state[f"chk_new_{ta}"] = False
                st.rerun()

            selected_tas = []
            with st.container(height=250, border=True):
                for ta in unique_tas:
                    if f"chk_new_{ta}" not in st.session_state: st.session_state[f"chk_new_{ta}"] = True
                    if st.checkbox(ta, key=f"chk_new_{ta}"): selected_tas.append(ta)
            
            st.divider()
            action_placeholder = st.empty()
            
            if st.session_state.zip_data is None:
                if action_placeholder.button("Generate Reports", use_container_width=True):
                    if not selected_tas:
                        st.warning("Please select at least one Trade Area.")
                    else:
                        action_placeholder.empty() 
                        progress_bar = st.progress(0)
                        
                        filtered_df = df[df["TRADE AREA"].astype(str).isin(selected_tas)]
                        groups = filtered_df.groupby("TRADE AREA")
                        total_groups = len(groups)
                        zip_buffer = io.BytesIO()
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for i, (trade_area, group) in enumerate(groups):
                                if isinstance(resolved_template, io.BytesIO):
                                    resolved_template.seek(0)
                                wb = openpyxl.load_workbook(resolved_template if isinstance(resolved_template, io.BytesIO) else template_file)
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
                                                for ph, map_conf in mapping.items():
                                                    target = f"{{{{{ph}}}}}"
                                                    if target in new_val:
                                                        header = map_conf["column"]
                                                        mask_patt = map_conf["mask"]
                                                        raw_data_val = row.get(header)
                                                        
                                                        is_image = inject_image_calibrated(new_sheet, cell.coordinate, raw_data_val, media_dict, max_size=180, col_offset=0, row_offset=0)
                                                        
                                                        if is_image:
                                                            val_str = ""
                                                        else:
                                                            val_str = format_with_mask(raw_data_val, mask_patt, ph)
                                                        
                                                        new_val = new_val.replace(target, val_str)
                                                        if val_str.strip() != "": has_injected = True
                                                
                                                if has_injected:
                                                    copy_and_merge_aware_injection(base_sheet, new_sheet, cell.coordinate, new_val.strip() if new_val else "")
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
                st.success("Reports generated successfully!")
                st.download_button("Download All Reports (.zip)", data=st.session_state.zip_data, file_name="Trade_Area_Reports.zip", mime="application/zip", use_container_width=True)
                if st.button("Start Over"):
                    st.session_state.zip_data = None
                    st.rerun()

# ==========================================
# MODE B: EDIT EXISTING REPORTS
# ==========================================
elif mode == "Edit / Update Existing Reports":
    st.markdown("### Upload Workbooks & Data Targets")
    
    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)
    
    with row1_col1:
        with st.container(border=True):
            st.markdown("📊 **1. New Raw Data**")
            src_type_1 = st.segmented_control("Source Type 1", ["File Upload", "Remote Link"], default="File Upload", key="src_type_1", label_visibility="collapsed")
            if src_type_1 == "File Upload":
                edit_raw_file = st.file_uploader("Upload Data Sheet", type=["xlsx", "xls"], key="edit_raw", label_visibility="collapsed")
                edit_raw_url = None
            else:
                edit_raw_url = st.text_input("Data Link URL", placeholder="https://example.com/new_data.xlsx", key="edit_raw_url", label_visibility="collapsed")
                edit_raw_file = None

    with row1_col2:
        with st.container(border=True):
            st.markdown("🗂️ **2. Existing Trade Area Workbooks**")
            src_type_2 = st.segmented_control("Source Type 2", ["File Upload", "Remote Link"], default="File Upload", key="src_type_2", label_visibility="collapsed")
            if src_type_2 == "File Upload":
                existing_wbs_raw = st.file_uploader("Upload Existing Sheets", type=["xlsx"], accept_multiple_files=True, key="edit_wbs", label_visibility="collapsed")
                existing_wbs_url = None
            else:
                existing_wbs_url = st.text_input("Workbook Zip/File URL", placeholder="https://example.com/workbooks.zip", key="edit_wbs_url", label_visibility="collapsed")
                existing_wbs_raw = None

    with row2_col1:
        with st.container(border=True):
            st.markdown("📐 **3. Excel Template (For Coordinates)**")
            src_type_3 = st.segmented_control("Source Type 3", ["File Upload", "Remote Link"], default="File Upload", key="src_type_3", label_visibility="collapsed")
            if src_type_3 == "File Upload":
                edit_temp_file = st.file_uploader("Upload Coordinate Template", type=["xlsx"], key="edit_temp", label_visibility="collapsed")
                edit_temp_url = None
            else:
                edit_temp_url = st.text_input("Template URL", placeholder="https://example.com/template.xlsx", key="edit_temp_url", label_visibility="collapsed")
                edit_temp_file = None

    with row2_col2:
        with st.container(border=True):
            st.markdown("📸 **4. Photos / Docs (Folder / Link)**")
            src_type_4 = st.segmented_control("Source Type 4", ["File Upload", "Remote Link"], default="File Upload", key="src_type_4", label_visibility="collapsed")
            if src_type_4 == "File Upload":
                media_files = st.file_uploader("Upload Images", accept_multiple_files=True, key="edit_media", type=['png', 'jpg', 'jpeg'], label_visibility="collapsed")
                media_url = None
            else:
                media_url = st.text_input("Cloud Drive Directory Link URL", placeholder="Paste Google Drive or OneDrive Folder URL Link", key="edit_media_url", label_visibility="collapsed")
                media_files = None

    resolved_edit_raw = resolve_file_source(edit_raw_file, edit_raw_url)
    resolved_edit_temp = resolve_file_source(edit_temp_file, edit_temp_url)
    
    existing_wbs = []
    if existing_wbs_raw:
        existing_wbs = existing_wbs_raw
    elif existing_wbs_url and existing_wbs_url.strip():
        resolved_wbs_zip = resolve_file_source(None, existing_wbs_url)
        if resolved_wbs_zip:
            try:
                with zipfile.ZipFile(resolved_wbs_zip) as z:
                    for name in z.namelist():
                        if name.endswith('.xlsx') and not name.split('/')[-1].startswith('.'):
                            existing_wbs.append(io.BytesIO(z.read(name)))
                            existing_wbs[-1].name = name.split('/')[-1]
            except Exception as e:
                st.error(f"Failed to unpack remote workbook package: {str(e)}")

    media_dict = {}
    if media_files:
        media_dict = {f.name: f for f in media_files}
    elif media_url and media_url.strip():
        resolved_media_data = resolve_file_source(None, media_url)
        if resolved_media_data:
            try:
                with zipfile.ZipFile(resolved_media_data) as z:
                    for name in z.namelist():
                        if any(ext in name.lower() for ext in ['.png', '.jpg', '.jpeg']) and not name.split('/')[-1].startswith('.'):
                            clean_name = name.split('/')[-1]
                            media_dict[clean_name] = StreamWrapper(z.read(name), clean_name)
            except zipfile.BadZipFile:
                media_dict["photo1.jpg"] = StreamWrapper(resolved_media_data.getvalue(), "photo1.jpg")

    if resolved_edit_raw and resolved_edit_temp and existing_wbs:
        df = pd.read_excel(resolved_edit_raw)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.columns = df.columns.str.strip().str.upper()
        headers = list(df.columns)
        
        if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns or "SITE ID" not in df.columns:
            st.error("ERROR: Raw data must contain 'TRADE AREA', 'SITE NAME', and 'SITE ID' columns.")
            st.stop()

        st.success(f"""
        🔍 **DISCOVERY PASSED:**
        ✅ Processed Trade Area Workbooks Data Feed
        ✅ Detected {len(df)} Rows Pending Inspection
        📸 System Media Cloud Drive Sync State Verified
        """)
        st.divider()

        template_wb = openpyxl.load_workbook(resolved_edit_temp)
        template_sheet = template_wb.active
        placeholders = get_placeholders(template_sheet)
        ph_coords = get_placeholder_coords(template_sheet) 

        st.markdown("### Selective Injection & Calibration Suite")
        
        active_mapping = {}
        for ph in placeholders:
            match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
            default_index = headers.index(match[0]) if match else 0
            
            with st.container(border=True):
                col1, col2, col3, col4 = st.columns([0.5, 1, 1, 1.5])
                with col1: update_check = st.checkbox(f"Update", key=f"chk_{ph}", value=False)
                with col2: st.markdown(f"**{{{{{ph}}}}}**")
                with col3: 
                    input_type = st.selectbox("Source", ["From Column", "Custom Value", "Image/Media Asset"], key=f"type_{ph}", label_visibility="collapsed", disabled=not update_check)
                with col4: 
                    if input_type == "From Column":
                        mapped_val = st.selectbox("Column target:", headers, index=default_index, key=f"map_edit_{ph}", label_visibility="collapsed", disabled=not update_check)
                    elif input_type == "Custom Value":
                        mapped_val = st.text_input("Global value text:", placeholder="e.g., June 1, 2026", key=f"custom_edit_{ph}", label_visibility="collapsed", disabled=not update_check)
                    else:
                        mapped_val = st.selectbox("Image reference column:", headers, index=default_index, key=f"image_edit_{ph}", label_visibility="collapsed", disabled=not update_check)

                m1, m2 = st.columns([1, 2])
                with m1:
                    sel_mask = st.selectbox("Data Formatting Mask Style", list(HUMAN_SPREADSHEET_MASKS.keys()), index=0, key=f"mask_edit_ui_{ph}", disabled=(not update_check or input_type=="Image/Media Asset"))
                
                with m2:
                    if input_type == "Image/Media Asset" and update_check:
                        with st.expander("🖼️ Interactive WYSIWYG Layout Alignment Matrix", expanded=True):
                            img_size = st.slider("Target Envelope Box Width (px)", 50, 800, 180, step=10, key=f"size_geo_{ph}")
                            st.markdown("<small><b>Interactive Placement Target Block Picker</b></small>", unsafe_allow_html=True)
                            
                            grid_pos = st.radio("Anchor Target Location", 
                                ["Top-Left (Standard Cell)", "Shift Left (-1 Col)", "Shift Right (+1 Col)", "Shift Up (-1 Row)", "Shift Down (+1 Row)"], 
                                index=0, horizontal=True, key=f"grid_geo_{ph}")
                            
                            col_off, row_off = 0, 0
                            if "Shift Left" in grid_pos: col_off = -1
                            elif "Shift Right" in grid_pos: col_off = 1
                            elif "Shift Up" in grid_pos: row_off = -1
                            elif "Shift Down" in grid_pos: row_off = 1
                            
                            st.caption(f"📍 **Dynamic Mapping Indicator:** Asset locks inside target cell coordinate path offset by **[{col_off} Col, {row_off} Row]**.")
                    else:
                        img_size, col_off, row_off = 180, 0, 0

                if update_check:
                    mask_id = HUMAN_SPREADSHEET_MASKS[sel_mask]
                    if input_type == "Image/Media Asset":
                        evaluated_preview = f"[Image Asset File Stream Loaded ──► Envelope Scale Bound to: {img_size}px]"
                    else:
                        mock_seed = generate_mock_value(mask_id) if input_type == "From Column" else mapped_val
                        evaluated_preview = format_with_mask(mock_seed, mask_id, ph)
                    
                    p_col1, p_col2 = st.columns([1, 2])
                    with p_col1:
                        st.markdown(f"    ↳ `<small>Live Output Visual:</small>`", unsafe_allow_html=True)
                    with p_col2:
                        st.markdown(f"`{evaluated_preview}`")

                data_origin_label = mapped_val if update_check else "None Assigned"
                st.markdown(f"<div style='text-align: right; opacity: 0.35; font-size: 10px; font-weight: bold;'>[Source: {data_origin_label}] ──► [Injected Token Var: {{{{ {ph} }}}}]</div>", unsafe_allow_html=True)

            if update_check:
                active_mapping[ph] = {
                    "type": input_type, 
                    "value": mapped_val, 
                    "mask": HUMAN_SPREADSHEET_MASKS[sel_mask],
                    "img_size": img_size,
                    "col_offset": col_off,
                    "row_offset": row_off
                }

        st.divider()
        action_placeholder = st.empty()

        if st.session_state.zip_data is None:
            if action_placeholder.button("Inject Data & Generate Updates", use_container_width=True):
                if not active_mapping:
                    st.warning("Please check at least one field to update.")
                else:
                    action_placeholder.empty()
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    zip_buffer = io.BytesIO()
                    existing_files_dict = {f.name: f for f in existing_wbs}
                    log_entries = []
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED)
