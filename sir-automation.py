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

# Reset session state for a fresh run if needed
if "zip_data" not in st.session_state:
    st.session_state.zip_data = None

st.title("Site Information Report Automation")

# STEP 1: The Upload Section
raw_file = st.file_uploader("Upload Raw Data (Excel)", type=["xlsx", "xls"])
template_file = st.file_uploader("Upload Excel Template", type=["xlsx"])

if raw_file and template_file:
    # Read headers from Raw Data & Auto-clean 'Unnamed' blank columns
    df = pd.read_excel(raw_file)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    headers = list(df.columns)
    
    # Mandatory Column Check (Moved up to prevent errors early)
    if "TRADE AREA" not in df.columns or "SITE NAME" not in df.columns:
        st.error("ERROR: Raw data must contain exactly 'TRADE AREA' and 'SITE NAME' columns.")
        st.stop()
    
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
            match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
            default_index = headers.index(match[0]) if match else 0
            
            col1, col2 = st.columns([1, 2])
            with col1:
                st.markdown(f"**{{{{{ph}}}}}**")
            with col2:
                if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and "LOCATION/ADDRESS" in headers:
                    default_index = headers.index("LOCATION/ADDRESS")
                
                mapping[ph] = st.selectbox("Map to column:", headers, index=default_index, key=f"map_{ph}", label_visibility="collapsed")
        
        st.divider()

        # STEP 2.5: Trade Area Selection
        st.markdown("### Select Trade Areas")
        st.markdown("Choose which Trade Areas to include in the batch generation.")
        
        # Get unique Trade Areas and cast to string for consistency
        unique_tas = sorted([str(ta) for ta in df["TRADE AREA"].dropna().unique()])
        
        # Select All / Clear All Buttons
        col_sel, col_clr, _ = st.columns([1, 1, 3])
        if col_sel.button("Select All", use_container_width=True):
            for ta in unique_tas:
                st.session_state[f"chk_{ta}"] = True
            st.rerun()
            
        if col_clr.button("Clear All", use_container_width=True):
            for ta in unique_tas:
                st.session_state[f"chk_{ta}"] = False
            st.rerun()

        # Render Checkboxes in a scrollable container
        selected_tas = []
        with st.container(height=250, border=True):
            for ta in unique_tas:
                # Initialize state to True by default if not set
                if f"chk_{ta}" not in st.session_state:
                    st.session_state[f"chk_{ta}"] = True
                
                # Render checkbox tied to session state
                if st.checkbox(ta, key=f"chk_{ta}"):
                    selected_tas.append(ta)
        
        st.divider()

        action_placeholder = st.empty()
        
        if st.session_state.zip_data is None:
            if action_placeholder.button("Generate Reports", use_container_width=True):
                
                if not selected_tas:
                    st.warning("Please select at least one Trade Area to generate reports.")
                else:
                    # STEP 3: The Engine
