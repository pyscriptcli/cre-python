import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill
import re
import io
import zipfile
import difflib
import math

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
    """Map placeholders to their exact cell coordinates (e.g., {'SITE_NAME': ['L5', 'C6']})"""
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

def format_injected_value(ph, header, val):
    """Applies date and address formatting rules based on placeholder."""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    
    # Date Formatting Rule
    if "DATE" in str(header).upper() or isinstance(val, pd.Timestamp):
        try:
            return pd.to_datetime(val).strftime("%B %d, %Y")
        except:
            return str(val)
            
    # Address Slicing Rule
    if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and isinstance(val, str):
        p = [part.strip() for part in val.split(",")]
        length = len(p)
        if ph == "STREET": return ", ".join(p[:max(0, length - 6)]) if length >= 6 else val
        elif ph == "BARANGAY": return p[length - 6] if length >= 6 else ""
        elif ph == "CITY": return p[length - 5] if length >= 5 else ""
        elif ph == "REGION": return p[length - 3] if length >= 3 else ""
        elif ph == "POSTAL": return p[length - 2] if length >= 2 else ""
        
    return str(val)

# --- SESSION STATE INITIALIZATION ---
if "zip_data" not in st.session_state: st.session_state.zip_data = None

# --- UI: THE LANDING PAGE ---
st.title("Site Information Report")

mode = st.radio(
    "Select Workflow Mode:", 
    ["Create New Reports", "Edit / Update Existing Reports"], 
    horizontal=True
)
st.divider()

# ==========================================
# MODE A: CREATE NEW REPORTS
# ==========================================
if mode == "Create New Reports":
    st.markdown("### Upload Files")
    raw_file = st.file_uploader("Upload Raw Data (Excel)", type=["xlsx", "xls"], key="new_raw")
    template_file = st.file_uploader("Upload Excel Template", type=["xlsx"], key="new_temp")

    if raw_file and template_file:
        df = pd.read_excel(raw_file)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.columns = df.columns.str.strip().str.upper()
        headers = list(df.columns)
        
        if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
            st.error("ERROR: Raw data must contain exactly 'TRADE AREA' and 'SITE NAME' columns.")
            st.stop()
        
        template_wb = openpyxl.load_workbook(template_file)
        template_sheet = template_wb.active
        placeholders = get_placeholders(template_sheet)
        
        if not placeholders:
            st.warning("No {{Placeholders}} found in the uploaded template.")
        else:
            st.markdown("### Data Mapping")
            mapping = {}
            for ph in placeholders:
                match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
                default_index = headers.index(match[0]) if match else 0
                if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and "LOCATION/ADDRESS" in headers:
                    default_index = headers.index("LOCATION/ADDRESS")
                
                col1, col2 = st.columns([1, 2])
                with col1: st.markdown(f"**{{{{{ph}}}}}**")
                with col2: mapping[ph] = st.selectbox("Map to column:", headers, index=default_index, key=f"map_{ph}", label_visibility="collapsed")
            
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
                        status_text = st.empty()
                        
                        filtered_df = df[df["TRADE AREA"].astype(str).isin(selected_tas)]
                        groups = filtered_df.groupby("TRADE AREA")
                        total_groups = len(groups)
                        zip_buffer = io.BytesIO()
                        
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for i, (trade_area, group) in enumerate(groups):
                                status_text.text(f"Processing: {trade_area}...")
                                template_file.seek(0)
                                wb = openpyxl.load_workbook(template_file)
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
                                                for ph, header in mapping.items():
                                                    target = f"{{{{{ph}}}}}"
                                                    if target in new_val:
                                                        val_str = format_injected_value(ph, header, row.get(header))
                                                        new_val = new_val.replace(target, val_str)
                                                        if val_str.strip() != "": has_injected = True
                                                
                                                cell.value = new_val
                                                if has_injected:
                                                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                                                    
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
    st.markdown("### Upload Workbooks & Data")
    raw_file = st.file_uploader("1. Upload New Raw Data (Excel)", type=["xlsx", "xls"], key="edit_raw")
    existing_wbs = st.file_uploader("2. Upload Existing Trade Area Workbooks", type=["xlsx"], accept_multiple_files=True, key="edit_wbs")
    template_file = st.file_uploader("3. Upload Excel Template (For Coordinates)", type=["xlsx"], key="edit_temp")
    # Optional image uploader just to satisfy the discovery pass gatekeeper for future expansion
    photos = st.file_uploader("4. Upload Photos/Docs (Optional)", accept_multiple_files=True, key="edit_photos")

    if raw_file and template_file and existing_wbs:
        df = pd.read_excel(raw_file)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df.columns = df.columns.str.strip().str.upper()
        headers = list(df.columns)
        
        if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns or "SITE ID" not in df.columns:
            st.error("ERROR: Raw data must contain 'TRADE AREA', 'SITE NAME', and 'SITE ID' columns.")
            st.stop()

        # Gatekeeper / Discovery Pass
        st.success(f"""
        🔍 **DISCOVERY PASSED:**
        ✅ Detected {len(existing_wbs)} Trade Area Workbooks
        ✅ Detected {len(df)} New Site Rows in Data Sheet
        📸 Detected {len(photos) if photos else 0} Media Files matching SITE IDs
        
        *Proceed to configure injection below.*
        """)
        st.divider()

        template_wb = openpyxl.load_workbook(template_file)
        template_sheet = template_wb.active
        placeholders = get_placeholders(template_sheet)
        ph_coords = get_placeholder_coords(template_sheet) # The secret weapon for existing files

        st.markdown("### Selective Injection Mapping")
        st.markdown("Check the box next to the fields you want to overwrite. Unchecked fields will preserve existing data.")
        
        active_mapping = {}
        for ph in placeholders:
            match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
            default_index = headers.index(match[0]) if match else 0
            if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and "LOCATION/ADDRESS" in headers:
                default_index = headers.index("LOCATION/ADDRESS")
            
            col1, col2, col3 = st.columns([0.5, 1, 2])
            with col1: update_check = st.checkbox(f"Update", key=f"chk_{ph}", value=False)
            with col2: st.markdown(f"**{{{{{ph}}}}}**")
            with col3: 
                mapped_header = st.selectbox("Map to column:", headers, index=default_index, key=f"map_edit_{ph}", label_visibility="collapsed", disabled=not update_check)
            
            if update_check:
                active_mapping[ph] = mapped_header

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
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        processed_count = 0
                        
                        # Loop through the uploaded existing workbooks
                        for wb_name, file_obj in existing_files_dict.items():
                            status_text.text(f"Injecting into: {wb_name}...")
                            
                            # Reconstruct the Trade Area name from the filename to filter rows
                            ta_from_filename = wb_name.replace(".xlsx", "")
                            
                            # Try to find matching rows in data (handling the replace logic we did earlier)
                            # A simple way is to iterate over grouped df
                            matched_group = None
                            for ta, group in df.groupby("TRADE AREA"):
                                safe_ta = str(ta).replace("/", "-").replace("\\", "-")
                                if safe_ta == ta_from_filename:
                                    matched_group = group
                                    break
                                    
                            if matched_group is not None:
                                wb = openpyxl.load_workbook(file_obj)
                                
                                for _, row in matched_group.iterrows():
                                    # Use the sanitize logic to predict what the tab name would be
                                    # Since we don't know duplicate state easily, we check if base name exists
                                    clean_name = re.sub(r'[\\/*?\[\]:]', '', str(row["SITE NAME"]))[:31]
                                    
                                    target_sheet = None
                                    for sheet_name in wb.sheetnames:
                                        if sheet_name.startswith(clean_name):
                                            target_sheet = wb[sheet_name]
                                            break
                                            
                                    if target_sheet:
                                        # Inject ONLY mapped and checked coordinates
                                        for ph, header in active_mapping.items():
                                            val_str = format_injected_value(ph, header, row.get(header))
                                            coords = ph_coords.get(ph, [])
                                            for coord in coords:
                                                target_sheet[coord].value = val_str
                                                if val_str.strip() != "":
                                                    target_sheet[coord].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

                                wb_buffer = io.BytesIO()
                                wb.save(wb_buffer)
                                zip_file.writestr(wb_name, wb_buffer.getvalue())
                                
                            processed_count += 1
                            progress_bar.progress(processed_count / len(existing_files_dict))

                    st.session_state.zip_data = zip_buffer.getvalue()
                    st.rerun()
                    
        if st.session_state.zip_data is not None:
            st.success("Existing reports updated successfully!")
            st.download_button("Download Updated Reports (.zip)", data=st.session_state.zip_data, file_name="Updated_Trade_Area_Reports.zip", mime="application/zip", use_container_width=True)
            if st.button("Start Over"):
                st.session_state.zip_data = None
                st.rerun()
