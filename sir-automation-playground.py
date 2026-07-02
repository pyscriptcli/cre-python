import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
import re
import io
import zipfile
import difflib
import requests
from copy import copy
import json
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# --- SERVICE ACCOUNT CONFIGURATION ---
# Your service account credentials
SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "focused-studio-501200-f2",
    "private_key_id": "d5fa49f27824d4a2bbce318f88a1e379b2c2e122",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEuwIBADANBgkqhkiG9w0BAQEFAASCBKUwggShAgEAAoIBAQDHpyC4irELswst\nGF2SsnFdFj2R7CrdJRFEwhBdKVQA6Djlojd1d7uRLPO7T6DhohFZRPDhJ44UgjGg\n2sPffOfTTIJz8tCT3Qm/EKoA1i/BX5Z0TO0EIUd5R1RjOExTIABDwoTZoSifdSjT\nxDKGSOmBZPPxGKzoJ6PzcW/gvF0OxPBgGkDg+4CMupvutqEqO9hhe/LeVcxySLsF\ngUiDXBwBRW7wIX3k3XxB1ob5srD1FXoa7DsXbkgnkH/TgybkjCXXbnC/uvaR27xf\nN0FWdcfSG5cx2P8qdSD7FRg8J+QfInaWiE/BLFf12oFYU5BuP7Nk9jVxqCxxG+0E\n154JdaVhAgMBAAECgf8ZY9lD0exCBW4rOSk6Bq3NnY5zu3Axpdnt8vmk+PFRGQP8\n8AXT6OFWZqKrjeAMnoR/5CyRKNHNk3ql1V+J6sojxo8W5iiji6PTQ7za7mIt87ug\nGEhSeVXFlQFxiZW4D7gsGQi1quRBPLA8fhVbm0CKAnSjX8GQpIazgH5h1JKUPG20\nXNTAR7C0w2C8oGv1PUNkus9Fxjw1PGme6ujBfIfZEbXPeKpsjz2QUYjlg/B+8r1c\nwSMoNsxdTKMxCG4zzaSrOpsVkeodZbdwmIkHzDikb8r0NiBeQ6Bgp2BN4phVKRD5\nk/ZQXo5s+ENrTVXL5SQwYNmF0Ae3Jo2m2z/XPUECgYEA8xv/YC0LDjRHqUo5KI23\n0lM1Y1JOqGU+/q9LZpmCPKzsud9tCQ0CouR2m30vRnLhJv4+DJmc49HpvLYh3Lff\nOfdyn384/TmMWr1FUrNupkrXI+l29BC+UYXfg8P12MYcrUy2Q/GKvj/BqEWfC/aj\n9LYo9IDeQsXvJJ3z64HLTkECgYEA0j0+uEupvq3vdigEgW5hKO/o3UcOAm1LoEVD\ndOu0IvqCatGmDDLrfk+w01X7ZIW95gYrKW/mFs3hVCqJwplRGkp0nsYg+FWDkq7F\n2QCi9WAVu0Gqfavyg9asPMeGoTNVgj2gzjJc4fILw2N6PvyOOin0HmISLwnVZtQ2\nAd96zyECgYAIfMG9qdTo+gpGbsDwGYKBZUZH4We9mUtJuPT48AML+z2If7RezIV7\nCl7ZrtUnsHsL0XR5HCPOEFYIsJMeEY1JiMoHp3ll3cx4noL9ECacx6AbMNtmSe9b\nCUF0aDL9Dm2R30u9s4EUg0VPip6y3Dl9IZ7salNYIXDn5lvNrQpcAQKBgFZ/xFJn\nLwu43JEsnc3y8B67tn90QJtXBIqIdNyiLZdGomn4n+zc9m8dso8BDVGqhRsi5pdB\n5tTzGAZdChj6o5fBkoHQ2rfR1zR+nABQdrumMMq+lbrnB/yeUncfUJD6YfAYExVD\nO1vrDlPxldZcatgbcskdaIXZ8edA4IecvxaBAoGBAOLRIcIKa5cJ3sMmmSvuMGTL\n17GLxtsXZGN0OltqSgT6pu3/uODGVH1ljqm9SrDFSmcnRZwn8a2aZoAG7brPuoT0\n5OcYYnbwDMKG0hcQbXNtbshjYPQMXb9si6gacKfcCny2Q9KSIhEz6ioEsvrmYOYz\njmfWgXnpGmDWcsgk7nKF\n-----END PRIVATE KEY-----\n",
    "client_email": "report-generator@focused-studio-501200-f2.iam.gserviceaccount.com",
    "client_id": "118306996332253581955",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/report-generator%40focused-studio-501200-f2.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

# Your file IDs
SOURCE_SPREADSHEET_ID = "14nhO9u7zJRcOoux8I7l2IzwU7iQZNW9fRX6TCip47CE"
TEMPLATE_SPREADSHEET_ID = "1uS3xmnPi0o4c_EayQtURYDSMMPRDRGSb"

def get_access_token():
    """Get OAuth2 access token using service account"""
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                'https://www.googleapis.com/auth/drive.readonly']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            SERVICE_ACCOUNT_INFO, scope
        )
        creds.refresh(requests.Request())
        return creds.access_token
    except Exception as e:
        st.error(f"Failed to get access token: {str(e)}")
        return None

def download_google_sheet_as_excel(spreadsheet_id):
    """Download Google Sheet as Excel using Service Account"""
    try:
        token = get_access_token()
        if not token:
            return None
            
        url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
        headers = {
            'Authorization': f'Bearer {token}'
        }
        
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            return io.BytesIO(response.content)
        else:
            st.error(f"Failed to download spreadsheet. Status: {response.status_code}")
            st.error(f"Response: {response.text[:200]}")
            return None
    except Exception as e:
        st.error(f"Error downloading spreadsheet: {str(e)}")
        return None

# --- HUMAN READABLE MASK DEFINITIONS ---
HUMAN_SPREADSHEET_MASKS = {
    "Plain text": "TEXT",
    "Number": "NUMBER",
    "Percentage": "PERCENT",
    "Scientific": "SCIENTIFIC",
    "Accounting": "ACCOUNTING",
    "Financial": "FINANCIAL",
    "Currency (USD)": "CURRENCY_USD",
    "Currency (USD, Rounded)": "CURRENCY_USD_ROUND",
    "Currency (PHP)": "CURRENCY_PHP",
    "Currency (PHP, Rounded)": "CURRENCY_PHP_ROUND",
    "Date (Short)": "DATE_SHORT",
    "Time (Standard)": "TIME_STANDARD",
    "Date Time (Full)": "DATE_TIME_FULL",
    "Date (Month Day, Year)": "%B %d, %Y",
    "Date (ISO)": "%Y-%m-%d",
    "Date (Day Month Year)": "%d %b %Y",
    "Street Segment": "STREET_SEGMENT",
    "Barangay Segment": "BARANGAY_SEGMENT",
    "City Segment": "CITY_SEGMENT",
    "Region Segment": "REGION_SEGMENT",
    "Postal Segment": "POSTAL_SEGMENT"
}

INVERSE_MASK_LOOKUP = {v: k for k, v in HUMAN_SPREADSHEET_MASKS.items()}

# --- HELPER FUNCTIONS ---
def parse_token_signature(raw_token):
    if ":" in raw_token:
        parts = raw_token.split(":", 1)
        name = parts[0].strip()
        tag_type = parts[1].strip().upper()
        if tag_type in ["IMAGE", "PHOTO"]: return name, "IMAGE"
        if tag_type in ["TEXT", "STRING"]: return name, "TEXT"
        if tag_type in INVERSE_MASK_LOOKUP: return name, tag_type
    return raw_token.strip(), "TEXT"

def get_placeholders(sheet):
    placeholders = set()
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                for match in matches:
                    name, _ = parse_token_signature(match)
                    placeholders.add(name)
    return sorted(list(placeholders))

def get_template_initial_types(sheet):
    initial_types = {}
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                matches = re.findall(r"\{\{(.*?)\}\}", cell.value)
                for match in matches:
                    clean_name, tag_type = parse_token_signature(match)
                    initial_types[clean_name] = tag_type
    return initial_types

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

def format_with_mask(val, mask_pattern, placeholder_name):
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
st.set_page_config(page_title="Report Generator", layout="wide")
st.markdown("## Report Generator")
st.markdown("---")

# --- LOAD FILES ---
st.markdown("### Loading Files from Google Drive...")

@st.cache_resource
def load_files():
    with st.spinner("Connecting to Google Drive..."):
        source_data = download_google_sheet_as_excel(SOURCE_SPREADSHEET_ID)
        template_data = download_google_sheet_as_excel(TEMPLATE_SPREADSHEET_ID)
        return source_data, template_data

source_data, template_data = load_files()

if source_data is None or template_data is None:
    st.error("""
    Failed to load files from Google Drive. Please check:
    
    1. The service account has access to both files
    2. The file IDs are correct
    3. The files are shared with the service account email:
       report-generator@focused-studio-501200-f2.iam.gserviceaccount.com
    
    To share files:
    1. Open each file in Google Drive
    2. Click "Share"
    3. Add the service account email above
    4. Give "Viewer" permission
    """)
    st.stop()

st.success("Files loaded successfully!")

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
template_types_registry = get_template_initial_types(template_sheet)

if not placeholders:
    st.warning("No {{Placeholders}} found in the template.")
    st.stop()

# --- SESSION STATE ---
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None

# --- MAIN INTERFACE ---
st.markdown("### Data Mapping")

# Select All / Clear All
col_sel_all, col_clr_all = st.columns([1, 1, 3])
if col_sel_all.button("Select All", use_container_width=True):
    for ph in placeholders:
        st.session_state[f"chk_{ph}"] = True
    st.rerun()
if col_clr_all.button("Clear All", use_container_width=True):
    for ph in placeholders:
        st.session_state[f"chk_{ph}"] = False
    st.rerun()

mapping = {}
for ph in placeholders:
    match = difflib.get_close_matches(ph, headers, n=1, cutoff=0.3)
    default_index = headers.index(match[0]) if match else 0
    
    if ph in ["STREET", "BARANGAY", "CITY", "REGION", "POSTAL"] and "LOCATION/ADDRESS" in headers:
        default_index = headers.index("LOCATION/ADDRESS")
    
    with st.container(border=True):
        if f"chk_{ph}" not in st.session_state:
            st.session_state[f"chk_{ph}"] = True
        update_check = st.checkbox(f"Update {ph}", key=f"chk_{ph}")
        
        col1, col2 = st.columns(2)
        with col1:
            sel_col = st.selectbox("Header Reference", headers, index=default_index, key=f"map_{ph}", disabled=not update_check)
        with col2:
            inferred_type = template_types_registry.get(ph, "TEXT")
            default_mask_label = INVERSE_MASK_LOOKUP.get(inferred_type, "Plain text") if inferred_type != "IMAGE" else "Plain text"
            mask_index = list(HUMAN_SPREADSHEET_MASKS.keys()).index(default_mask_label)
            sel_mask = st.selectbox("Data Format", list(HUMAN_SPREADSHEET_MASKS.keys()), index=mask_index, key=f"mask_{ph}", disabled=not update_check)
        
        if update_check:
            mapping[ph] = {
                "column": sel_col,
                "mask": HUMAN_SPREADSHEET_MASKS[sel_mask],
                "inferred_type": inferred_type
            }

st.divider()
st.markdown("### Select Trade Areas")

unique_tas = sorted([str(ta) for ta in df["TRADE AREA"].dropna().unique()])

col_sel, col_clr = st.columns([1, 1, 3])
if col_sel.button("Select All Trade Areas", use_container_width=True):
    for ta in unique_tas:
        st.session_state[f"ta_{ta}"] = True
    st.rerun()
if col_clr.button("Clear All Trade Areas", use_container_width=True):
    for ta in unique_tas:
        st.session_state[f"ta_{ta}"] = False
    st.rerun()

selected_tas = []
with st.container(height=250, border=True):
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
        elif not mapping:
            st.warning("Please configure at least one field mapping.")
        else:
            with st.spinner("Generating reports..."):
                progress_bar = st.progress(0)
                
                filtered_df = df[df["TRADE AREA"].astype(str).isin(selected_tas)]
                groups = filtered_df.groupby("TRADE AREA")
                total_groups = len(groups)
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for i, (trade_area, group) in enumerate(groups):
                        # Reload template for each group
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
                                        for ph, map_conf in mapping.items():
                                            target_regex = r"\{\{\s*" + re.escape(ph) + r"(\s*:.*?)?\}\}"
                                            if re.search(target_regex, new_val):
                                                header = map_conf["column"]
                                                mask_patt = map_conf["mask"]
                                                raw_data_val = row.get(header)
                                                
                                                val_str = format_with_mask(raw_data_val, mask_patt, ph)
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
        st.info("Save to Google Drive:")
        st.caption("1. Download the file above")
        st.caption("2. Go to your Google Drive folder:")
        st.caption("https://drive.google.com/drive/folders/1MAo_8VYditz-BV3vGx3aX31-SLzxSAD8")
        st.caption("3. Click 'New' -> 'File Upload' and select the downloaded zip")
    
    if st.button("Start Over"):
        st.session_state.zip_data = None
        st.rerun()
