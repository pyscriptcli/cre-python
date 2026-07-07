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
    page_icon="📊",
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

# --- CUSTOM GOOGLE WORKSPACE/MATERIAL DESIGN CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@300;400;500;700&display=swap');
    
    * { 
        font-family: 'Google Sans', 'Roboto', 'Segoe UI', sans-serif !important; 
    }
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    header {visibility: hidden !important;}
    button[title="View source"] {display: none !important;}
    .stAppDeployButton {display: none !important;}
    div[data-testid="stStatusWidget"] {display: none !important;}
    
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }
    
    /* Google Tonal Styled Export Button */
    .stButton > button, .stDownloadButton > button {
        background-color: #0b57d0 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 100px !important;
        padding: 0.5rem 1.5rem !important;
        font-size: 0.875rem !important;
        font-weight: 500 !important;
        min-height: 40px !important;
        height: 40px !important;
        width: 100% !important;
        box-shadow: 0 1px 2px 0 rgba(60,64,67,0.3), 0 1px 3px 1px rgba(60,64,67,0.15) !important;
        transition: background-color 0.2s, box-shadow 0.2s;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #0b4cb4 !important;
        box-shadow: 0 1px 3px 0 rgba(60,64,67,0.3), 0 4px 8px 3px rgba(60,64,67,0.15) !important;
    }
    
    /* Document Selectors Styled Frames */
    .stSelectbox label { 
        font-size: 0.75rem !important; 
        font-weight: 500 !important;
        color: #444746 !important;
        margin-bottom: 4px !important;
    }
    .stSelectbox > div > div {
        background-color: #fff !important;
        border: 1px solid #747775 !important;
        border-radius: 4px !important;
        min-height: 40px !important;
        height: 40px !important;
    }
    .stSelectbox > div > div > div { 
        padding-top: 4px !important; 
        font-size: 0.9rem !important; 
    }
    
    div[data-testid="stHorizontalBlock"] { 
        gap: 1rem !important; 
        align-items: flex-end !important; 
        background: #f0f4f9;
        padding: 1rem;
        border-radius: 16px;
        margin-bottom: 1rem;
    }
    
    .info-text { 
        font-size: 0.85rem; 
        color: #444746; 
        text-align: right; 
        margin: 0; 
        padding-bottom: 10px; 
        font-weight: 500; 
    }
    
    /* Document Display Blueprint Container */
    .excel-container {
        background-color: #ffffff !important;
        border-radius: 12px;
        padding: 1.5rem;
        border: 1px solid #c4c7c5;
        box-shadow: 0 1px 3px 0 rgba(0,0,0,0.1);
        overflow: auto;
        margin-top: 1rem;
        width: 100%;
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

def sanitize_tab_name(name, existing_names):
    clean_name = re.sub(r'[\\/*?\[\]:]', '', str(name))[:31]
    if not clean_name: clean_name = "Sheet"
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

def generate_trade_area_report(df, trade_area, template_bytes, placeholders):
    """Core engine for batch-building the multi-tab report, with hard regex token stripping."""
    ta_data = df[df["TRADE AREA"] == trade_area]
    wb = load_workbook(io.BytesIO(template_bytes))
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
                    # Replace found tokens
                    for ph in placeholders:
                        target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                        if re.search(target_regex, new_val):
                            raw_data_val = r_row.get(ph.upper(), "")
                            if pd.isna(raw_data_val) or raw_data_val is None: raw_data_val = ""
                            if isinstance(raw_data_val, float) and raw_data_val.is_integer(): val_str = str(int(raw_data_val))
                            elif hasattr(raw_data_val, 'strftime'): val_str = r_row.get(ph.upper(), "").strftime('%B %d, %Y')
                            else: val_str = str(raw_data_val)
                            new_val = re.sub(target_regex, val_str, new_val)
                            
                    # Hard-wipe any remaining {{PLACEHOLDER}} syntax that had no corresponding data
                    new_val = re.sub(r"\{\{.*?\}\}", "", new_val)
                    cell.value = new_val.strip() if new_val else ""
                    
    wb.remove(base_sheet)
    wb_buffer = io.BytesIO()
    wb.save(wb_buffer)
    return wb_buffer

# --- HTML TEMPLATE BLUEPRINT FROM EXPORT DATA ---
RAW_TEMPLATE_HTML = """
<style type="text/css">
    .ritz .waffle a { color: #15c; }
    .ritz .waffle .s0{border-bottom:1px SOLID #bfbfbf;border-right:1px SOLID #bfbfbf;background-color:#800000;text-align:center;font-weight:700;color:#ffffff;font-family:'Google Sans',Arial;font-size:14pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:6px 3px;}
    .ritz .waffle .s1{border-bottom:1px SOLID #bfbfbf;border-right:1px SOLID #bfbfbf;background-color:#f0f4f9;text-align:left;font-weight:700;color:#1f1f1f;font-family:'Google Sans',Arial;font-size:12pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s2{border-bottom:1px SOLID #e1e3e1;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s3{border-bottom:1px SOLID #e1e3e1;border-right:1px SOLID #e1e3e1;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s4{border-right:1px SOLID #e1e3e1;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s5{background-color:#ffffff;text-align:left;color:#444746;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:normal;word-break:break-word;direction:ltr;padding:6px 8px;}
    .ritz .waffle .s6{border:1px SOLID #c4c7c5;background-color:#f8f9fa;text-align:left;color:#1f1f1f;font-family:Arial;font-size:11pt;font-weight:500;vertical-align:middle;white-space:normal;word-break:break-word;direction:ltr;padding:6px 8px;}
    .ritz .waffle .s7{background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:normal;word-break:break-word;direction:ltr;padding:6px 8px;}
    .ritz .waffle .s8{border:1px SOLID #c4c7c5;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:normal;word-break:break-word;direction:ltr;padding:6px 8px;}
    .ritz .waffle .s9{border-bottom:1px SOLID transparent;border-right:1px SOLID transparent;background-color:#ffffff;text-align:left;color:#b3261e;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s10{border:1px SOLID #c4c7c5;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:normal;word-break:break-word;direction:ltr;padding:6px 8px;}
    .ritz .waffle .s11{background-color:#e1e3e1;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s12{border-bottom:1px SOLID #000000;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s13{border-bottom:1px SOLID #000000;border-right:1px SOLID #e1e3e1;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s14{background-color:#e1e3e1;text-align:left;font-weight:bold;color:#b3261e;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s15{background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s16{border-right:1px SOLID #000000;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s17{background-color:#f8f9fa;text-align:left;color:#b3261e;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s18{border-bottom:1px SOLID transparent;border-right:1px SOLID transparent;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s19{border-bottom:1px SOLID transparent;border-right:1px SOLID transparent;background-color:#f8f9fa;text-align:left;color:#b3261e;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s20{border-bottom:1px SOLID transparent;border-right:1px SOLID #000000;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s21{border-bottom:1px SOLID #000000;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s22{border-bottom:1px SOLID #000000;border-right:1px SOLID #000000;background-color:#f8f9fa;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s23{background-color:#ffffff;text-align:left;font-weight:bold;color:#1f1f1f;font-family:'Google Sans',Arial;font-size:12pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s24{border-bottom:1px SOLID transparent;border-right:1px SOLID transparent;background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle .s25{background-color:#ffffff;text-align:left;color:#000000;font-family:Arial;font-size:11pt;vertical-align:middle;white-space:nowrap;direction:ltr;padding:4px 6px;}
    .ritz .waffle td { border: 1px solid #e1e3e1; }
</style>
<div class="ritz grid-container" dir="ltr">
<table class="waffle" cellspacing="0" cellpadding="0" style="table-layout: fixed; width: 100%;">
    <colgroup>
        <col style="width:14%;"><col style="width:14%;"><col style="width:6%;"><col style="width:6%;"><col style="width:6%;"><col style="width:6%;"><col style="width:9%;"><col style="width:1%;"><col style="width:1%;"><col style="width:9%;"><col style="width:17%;"><col style="width:6%;"><col style="width:6%;"><col style="width:6%;"><col style="width:6%;"><col style="width:1%;">
    </colgroup>
    <tbody>
        <tr style="height: 32px;"><td class="s0" dir="ltr" colspan="16">SITE INFORMATION REPORT</td></tr>
        <tr style="height: 24px;"><td class="s1" dir="ltr" colspan="8">General Information</td><td class="s1"></td><td class="s1" dir="ltr" colspan="7">Location</td></tr>
        <tr style="height: 19px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 19px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Trade Area Name</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_TRADE_AREA_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Site Name</td><td class="s6" dir="ltr" colspan="5">_SITE_NAME_</td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Site Name:</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_SITE_NAME_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Unit #, Bldg/St # and St Name</td><td class="s6" dir="ltr" colspan="5">_UNIT_BLDG_ST_NAME_</td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Site Number:</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_SITE_NO_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Barangay/District Name</td><td class="s6" dir="ltr" colspan="5">_BARANGAY_DISTRICT_NAME_</td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Date Started</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_TIMESTAMP_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">City/Municipality</td><td class="s6" dir="ltr" colspan="5">_CITY_MUNICIPALITY_</td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Date Report Submitted</td><td class="s7" dir="ltr" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Region</td><td class="s6" dir="ltr" colspan="5">_REGION_</td></tr>
        <tr style="height: 24px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Postal Code</td><td class="s6" dir="ltr" colspan="5">_POSTAL_CODE_</td></tr>
        <tr style="height: 19px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 24px;"><td class="s1" dir="ltr" colspan="8">Terms</td><td class="s4"></td><td class="s1" dir="ltr" colspan="7">Rates</td></tr>
        <tr style="height: 19px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Site Availability Date</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Monthly Rental Rate (Php)</td><td class="s10" dir="ltr" colspan="5">_MONTHLY_RENTAL_RATE_</td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">COL Start Date</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Percentage Rent</td><td class="s10" colspan="5"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">COL End Date</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Minimum Guaranteed Rent</td><td class="s10" colspan="5"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Lease Terms</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Annual Escalation Rate (%)</td><td class="s10" dir="ltr" colspan="5">_ESCALATION_</td></tr>
        <tr style="height: 24px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Advance Rental (Php)</td><td class="s10" dir="ltr" colspan="5">_ADVANCE_RENTAL_</td></tr>
        <tr style="height: 24px;"><td class="s1" dir="ltr" colspan="8">Technical Info</td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Security Deposit Amount (Php)</td><td class="s10" dir="ltr" colspan="5">_SECURITY_DEPOSIT_</td></tr>
        <tr style="height: 24px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">CUSA Dues</td><td class="s10" dir="ltr" colspan="5">_CUSA_</td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Lot /Floor Area (in sqm)</td><td class="s6" dir="ltr" colspan="5">_LOT_FLOOR_AREA_SQM_</td><td class="s4"></td><td class="s4"></td><td class="s9" dir="ltr" colspan="2">Estimated Revenue Per Mo.</td><td class="s8" colspan="5"></td></tr>
        <tr style="height: 19px;"><td class="s5" dir="ltr">Frontage (in m)</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 19px;"><td class="s5" dir="ltr">Depth (in m)</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s1" dir="ltr" colspan="7">Provisions</td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Floor to Slab Height (in m) - if Bldg</td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">No. of Storeys (If Bldg Lessee)</td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Tenant is the Owner</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Type of Structure(if Bldg Lessee)</td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Lease Type</td><td class="s6" dir="ltr" colspan="4">_LEASE_TYPE_</td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Soil Profile</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Principal COL</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Supply Access:</td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Sub-Lease Provison</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Power</td><td class="s11"></td><td class="s5" dir="ltr">Aircon</td><td class="s11"></td><td class="s7" dir="ltr" colspan="2">LPG Fire Pro</td><td class="s11"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Pre-Term/Partial Term</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Water</td><td class="s11"></td><td class="s5" dir="ltr">Exhaust</td><td class="s11"></td><td class="s7" dir="ltr" colspan="2">Drainage TP</td><td class="s11"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Tripartite Agreement</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 19px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 24px;"><td class="s1" dir="ltr" colspan="8">Lessor and Tenant Details</td><td class="s4"></td><td class="s1" dir="ltr" colspan="7">If with Sub-Lessor/ Sub-Lessee</td></tr>
        <tr style="height: 19px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Name of Lessor</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_LESSOR_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Name of Sub-Lessor</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Contact No.</td><td class="s5"></td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Contact No.</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">E-mail Address</td><td class="s5"></td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">E-mail Address</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Type of Ownership</td><td class="s5"></td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Type of Ownership</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Company Name</td><td class="s5"></td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Company Name</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Developer Account Name</td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Developer Account Name</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Business Address</td><td class="s5"></td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Business Address</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Name of Authorized Representative</td><td class="s6" dir="ltr" colspan="5">_CONTACT_PERSON_SOURCE_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Name of Authorized Representative</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Residence Address of Authorized Representative</td><td class="s6" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Residence Address of Authorized Representative</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Contact No.</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_CONTACT_NUMBER_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Contact No.</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">E-mail Address</td><td class="s5"></td><td class="s6" dir="ltr" colspan="5">_EMAIL_ADDRESS_</td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">E-mail Address</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 19px;"><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Name of Lessee</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Name of Sub-Lessee</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Position</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Position</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Contact No.</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Contact No.</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">E-mail Address</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">E-mail Address</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s7" dir="ltr" colspan="2">Name of Authorized Representative</td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Name of Authorized Representative</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Business Address</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4"></td><td class="s4"></td><td class="s7" dir="ltr" colspan="2">Business Address</td><td class="s8" colspan="4"></td><td class="s4"></td></tr>
        <tr style="height: 19px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 20px;"><td class="s14" dir="ltr" colspan="16">Regulatory</td></tr>
        <tr style="height: 24px;">
            <td class="s17" dir="ltr">Setback Requirement</td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td>
            <td class="s18"></td><td class="s18"></td>
            <td class="s19" dir="ltr">Perm Traffic Re-Routing</td><td class="s15"></td><td class="s15"></td><td class="s15"></td>
            <td class="s17" dir="ltr">Future Development</td><td class="s15"></td><td class="s16"></td>
        </tr>
        <tr style="height: 24px;">
            <td class="s17" dir="ltr">Road Widening</td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td><td class="s15"></td>
            <td class="s18"></td><td class="s18"></td>
            <td class="s19" dir="ltr">Perm Road Closure</td><td class="s15"></td><td class="s15"></td><td class="s15"></td>
            <td class="s17" dir="ltr">Zoning Clearance</td><td class="s15"></td><td class="s16"></td>
        </tr>
        <tr style="height: 24px;">
            <td class="s17" dir="ltr">Pedestrian Overpass</td><td class="s21"></td><td class="s21"></td><td class="s21"></td><td class="s21"></td><td class="s21"></td><td class="s21"></td>
            <td class="s21"></td><td class="s21"></td>
            <td class="s21" dir="ltr">Infrastructure Programs</td><td class="s21"></td><td class="s21"></td><td class="s21"></td>
            <td class="s21" dir="ltr">Gas Station</td><td class="s21"></td><td class="s22"></td>
        </tr>
        <tr style="height: 19px;"><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td><td class="s4"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s2"></td><td class="s3"></td></tr>
        <tr style="height: 20px;"><td class="s23" dir="ltr" colspan="16">Site Acquirability:</td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Confidence Level</td><td class="s11" colspan="5"></td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Site Availability</td><td class="s24" colspan="5">_SITE_AVAILABILITY_CLASS_</td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
        <tr style="height: 24px;"><td class="s5" dir="ltr">Other Remarks:</td><td class="s25" colspan="5">_REMARKS_</td><td class="s5"></td><td class="s4"></td><td class="s4"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s5"></td><td class="s4"></td></tr>
    </tbody>
</table>
</div>
"""

# --- LOAD DATA ---
@st.cache_data(ttl=3600)
def load_data():
    source_data = download_file(SOURCE_URL)
    template_data = download_file(TEMPLATE_URL)
    if source_data is None or template_data is None: return None, None, None
    
    df = pd.read_excel(source_data)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    
    # Establish string formatter for Display Column ensuring Whole Number rendering of SITE NO
    df["SITE_DISPLAY"] = df.apply(
        lambda row: f"{int(row['SITE NO'])} - {row['SITE NAME']}" 
        if pd.notna(row.get('SITE NO')) and isinstance(row.get('SITE NO'), (int, float)) 
        else f"{row.get('SITE NO', '')} - {row.get('SITE NAME', '')}".strip(" -"), 
        axis=1
    )
    
    temp_wb = load_workbook(template_data)
    placeholders = get_placeholders(temp_wb.active)
    template_data.seek(0)
    return df, placeholders, template_data.getvalue()

with st.spinner("Loading Data Assets..."):
    df, placeholders, template_bytes_raw = load_data()

if df is None or template_bytes_raw is None:
    st.error("Failed to load data. Please check connection profiles.")
    st.stop()

# --- CONTROLS ROW ---
trade_areas = ["Select Trade Area..."] + sorted(df["TRADE AREA"].dropna().unique().tolist())
col1, col2, col3, col4, col5 = st.columns([1.5, 1.5, 0.7, 0.8, 0.5])

with col1:
    st.markdown("<p style='font-size:0.75rem; font-weight:500; color:#444746; margin:0;'>Trade Area</p>", unsafe_allow_html=True)
    selected_ta = st.selectbox("Trade Area", options=trade_areas, index=0, label_visibility="collapsed")
    
with col2:
    st.markdown("<p style='font-size:0.75rem; font-weight:500; color:#444746; margin:0;'>Site View</p>", unsafe_allow_html=True)
    if selected_ta and selected_ta != "Select Trade Area...":
        sites_in_ta = ["Select Site..."] + sorted(df[df["TRADE AREA"] == selected_ta]["SITE_DISPLAY"].dropna().unique().tolist())
    else:
        sites_in_ta = ["Select Site..."]
    selected_site_display = st.selectbox("Site Name", options=sites_in_ta, index=0, label_visibility="collapsed")

with col3:
    if st.button("Refresh Cache", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with col4:
    if selected_ta and selected_ta != "Select Trade Area...":
        with st.spinner("Generating..."):
            wb_buffer = generate_trade_area_report(df, selected_ta, template_bytes_raw, placeholders)
            st.download_button(
                label="Export", 
                data=wb_buffer.getvalue(), 
                file_name=f"{selected_ta}_Report.xlsx", 
                use_container_width=True
            )

with col5:
    st.markdown(f"<p class='info-text'>Total Registry<br><span style='font-size:1.1rem; color:#0b57d0;'>{len(df['SITE NAME'].dropna().unique())} Sites</span></p>", unsafe_allow_html=True)

# --- DIRECT HTML VIEW BLUEPRINT INJECTION LAYER ---
if selected_ta != "Select Trade Area..." and selected_site_display != "Select Site...":
    site_data = df[df["SITE_DISPLAY"] == selected_site_display]
    if not site_data.empty:
        site_row_data = site_data.iloc[0]
        try:
            # Data parsing formatter utilities matching explicit key bounds
            def process_val(key_string):
                val = site_row_data.get(key_string.upper(), "")
                if pd.isna(val) or val is None: return ""
                if isinstance(val, float) and val.is_integer(): return str(int(val))
                if hasattr(val, 'strftime'): return val.strftime('%B %d, %Y')
                return str(val).strip()

            # Dynamic string binding layer injecting row cells safely straight into the blueprint
            rendered_view = RAW_TEMPLATE_HTML
            rendered_view = rendered_view.replace("_TRADE_AREA_", process_val("TRADE AREA"))
            rendered_view = rendered_view.replace("_SITE_NAME_", process_val("SITE NAME"))
            rendered_view = rendered_view.replace("_SITE_NO_", process_val("SITE NO"))
            rendered_view = rendered_view.replace("_TIMESTAMP_", process_val("TIMESTAMP"))
            rendered_view = rendered_view.replace("_UNIT_BLDG_ST_NAME_", process_val("UNIT #, BLDG/ST # AND ST NAME"))
            rendered_view = rendered_view.replace("_BARANGAY_DISTRICT_NAME_", process_val("BARANGAY/DISTRICT NAME"))
            rendered_view = rendered_view.replace("_CITY_MUNICIPALITY_", process_val("CITY/MUNICIPALITY"))
            rendered_view = rendered_view.replace("_REGION_", process_val("REGION"))
            rendered_view = rendered_view.replace("_POSTAL_CODE_", process_val("POSTAL CODE"))
            rendered_view = rendered_view.replace("_MONTHLY_RENTAL_RATE_", process_val("MONTHLY RENTAL RATE"))
            rendered_view = rendered_view.replace("_ESCALATION_", process_val("ESCALATION"))
            rendered_view = rendered_view.replace("_ADVANCE_RENTAL_", process_val("ADVANCE RENTAL"))
            rendered_view = rendered_view.replace("_SECURITY_DEPOSIT_", process_val("SECURITY DEPOSIT"))
            rendered_view = rendered_view.replace("_CUSA_", process_val("CUSA"))
            rendered_view = rendered_view.replace("_LOT_FLOOR_AREA_SQM_", process_val("LOT/FLOOR AREA SQM"))
            rendered_view = rendered_view.replace("_LEASE_TYPE_", process_val("LEASE TYPE"))
            rendered_view = rendered_view.replace("_LESSOR_", process_val("LESSOR"))
            rendered_view = rendered_view.replace("_CONTACT_PERSON_SOURCE_", process_val("CONTACT PERSON/SOURCE"))
            rendered_view = rendered_view.replace("_CONTACT_NUMBER_", process_val("CONTACT NUMBER"))
            rendered_view = rendered_view.replace("_EMAIL_ADDRESS_", process_val("EMAIL ADDRESS"))
            rendered_view = rendered_view.replace("_SITE_AVAILABILITY_CLASS_", process_val("SITE AVAILABILITY CLASS"))
            rendered_view = rendered_view.replace("_REMARKS_", process_val("REMARKS"))
            
            # Final sweep to remove any remaining placeholder tags if fields were empty
            rendered_view = re.sub(r"_[A-Z0-9_]+_", "", rendered_view)
            
            st.markdown(f'<div class="excel-container">{rendered_view}</div>', unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"Error compiling visual blueprint frame layer: {str(e)}")
else:
    st.info("Please select a Trade Area and a Site to view the specific report.")
