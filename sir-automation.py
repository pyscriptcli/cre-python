import streamlit as st
import pandas as pd
import openpyxl
import re
import io
import zipfile
import difflib

# Must be the first Streamlit command
st.set_page_config(page_title="Report Generator", layout="centered")

def get_placeholders(sheet):
    """Extract all {{Placeholder}} variables from the Excel sheet."""
    placeholders = set()
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                placeholders.update(matches)
    return sorted(list(placeholders))

def sanitize_tab_name(name, existing_names):
    """Slice to 31 chars and prevent duplicate tab names within the same workbook."""
    base_name = str(name)[:31]
    if base_name not in existing_names:
        existing_names.add(base_name)
        return base_name
    
    counter = 1
    while True:
        suffix = f" ({counter})"
        max_len = 31 - len(suffix)
        new_name = f"{str(name)[:max_len]}{suffix}"
        if new_name not in existing_names:
            existing_names.add(new_name)
            return new_name
        counter += 1

# Reset session state for a fresh run if needed
if "zip_data" not in st.session_state:
    st.session_state.zip_data = None

st.title("Excel Report Generator")

# STEP 1: The Upload Section
raw_file = st.file_uploader("Upload Raw Data (Excel)", type=["xlsx", "xls"])
template_file = st.file_uploader("Upload Excel Template", type=["xlsx"])

if raw_file and template_file:
    # Read headers from Raw Data
    df = pd.read_excel(raw_file)
    headers = list(df.columns)
    
    # Read placeholders from Template
    template_wb = openpyxl.load_workbook(template_file)
    template_sheet = template_wb.active
    placeholders = get_placeholders(template_sheet)
    
    if not placeholders:
        st.warning("No {{Placeholders}} found in the uploaded template.")
    else:
        # STEP 2: The Smart Mapping Section
        st.markdown("### Smart Mapping")
        st.markdown("Verify the auto-matched data columns for your template placeholders.")
        
        mapping = {}
        for ph in placeholders:
            # Fuzzy match to guess the closest header
            match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
            default_index = headers.index(match[0]) if match else 0
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.markdown(f"**{{{{{ph}}}}}**")
            with col2:
                mapping[ph] = st.selectbox("Map to column:", headers, index=default_index, key=f"map_{ph}", label_visibility="collapsed")
        
        st.divider()

        # Action Area Placeholder
        action_placeholder = st.empty()
        
        # Determine what to show in the action area based on state
        if st.session_state.zip_data is None:
            if action_placeholder.button("Generate Reports", use_container_width=True):
                # STEP 3: The Engine
                action_placeholder.empty() # Hide the button
                
                # Check for mandatory columns
                if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
                    st.error("ERROR: Raw data must contain exactly 'TRADE AREA' and 'SITE NAME' columns.")
                    st.stop()

                progress_bar = st.progress(0)
                status_text = st.empty()
                
                groups = df.groupby("TRADE AREA")
                total_groups = len(groups)
                
                zip_buffer = io.BytesIO()
                
                # Start generating zip in memory
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for i, (trade_area, group) in enumerate(groups):
                        status_text.text(f"Processing Trade Area: {trade_area}...")
                        
                        # Load a fresh copy of the template workbook for this Trade Area
                        template_file.seek(0)
                        wb = openpyxl.load_workbook(template_file)
                        base_sheet = wb.active
                        base_sheet.title = "TEMPLATE_TO_DELETE"
                        existing_tabs = set()
                        
                        for _, row in group.iterrows():
                            # The Anchor (Tracking) & Tab Naming
                            site_id = row.get("SITE ID", "NO-ID")
                            site_name = row.get("SITE NAME", site_id)
                            safe_tab_name = sanitize_tab_name(site_name, existing_tabs)
                            
                            new_sheet = wb.copy_worksheet(base_sheet)
                            new_sheet.title = safe_tab_name
                            
                            # Inject Mapped Data
                            for row_cells in new_sheet.iter_rows():
                                for cell in row_cells:
                                    if isinstance(cell.value, str) and "{{" in cell.value:
                                        new_val = cell.value
                                        for ph, header in mapping.items():
                                            target = f"{{{{{ph}}}}}"
                                            if target in new_val:
                                                # Replace with mapped value, handling NaNs
                                                val = row.get(header)
                                                val_str = "" if pd.isna(val) else str(val)
                                                new_val = new_val.replace(target, val_str)
                                        cell.value = new_val
                                        
                        # Clean up the original template tab and save workbook to bytes
                        wb.remove(base_sheet)
                        wb_buffer = io.BytesIO()
                        wb.save(wb_buffer)
                        
                        # Write the workbook into the zip file
                        safe_filename = str(trade_area).replace("/", "-").replace("\\", "-")
                        zip_file.writestr(f"{safe_filename}.xlsx", wb_buffer.getvalue())
                        
                        # Update Progress
                        progress_bar.progress((i + 1) / total_groups)
                
                status_text.empty()
                progress_bar.empty()
                
                # Save to session state and trigger a UI refresh to show download button
                st.session_state.zip_data = zip_buffer.getvalue()
                st.rerun()

        # STEP 4: The Export Section
        if st.session_state.zip_data is not None:
            st.success("Reports generated successfully!")
            st.download_button(
                label="Download All Reports (.zip)",
                data=st.session_state.zip_data,
                file_name="Trade_Area_Reports.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            # Optional: Allow user to reset and generate a new batch
            if st.button("Start Over"):
                st.session_state.zip_data = None
                st.rerun()
