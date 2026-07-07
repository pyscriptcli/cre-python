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
        width: auto !important;
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

# --- CONFIGURATIONSpreadsheets Routing ---
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
    if not clean_name:
        clean_name = "Sheet"
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

# --- HTML DOM COMPONENT ---
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
        <tr style="height: 24px;"><td class="s5" dir="ltr">Contact No.</td><td class="s5"></td><td class="s8" colspan="5"></td><td class="s4">
