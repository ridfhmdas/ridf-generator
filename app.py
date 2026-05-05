import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from tinydb import TinyDB, Query
import datetime
import xlsxwriter

# # ==========================================
# # 1. INITIAL SETUP & THEME
# # ==========================================
st.set_page_config(page_title="PAGASA RIDF Calculator", page_icon="⛈️", layout="wide")

# Modern PAGASA Style Interface using CSS
st.markdown("""
    <style>
    /* Main Background */
    .stApp { background-color: #F8FAFC; }

    /* Titles & Text */
    h1 { color: #00508F; font-family: 'Arial Black', sans-serif; font-weight: 800; }
    h2, h3 { color: #2C3E50; font-family: 'Arial', sans-serif; }
    
    /* Buttons */
    .stButton>button { 
        background-color: #00508F; color: white; border-radius: 10px; 
        border: none; padding: 10px 24px; font-weight: bold; width: 100%;
    }
    .stButton>button:hover { background-color: #3498DB; color: white; }

    /* Sidebar Stat Cards */
    [data-testid="stSidebar"] { background-color: #003366; color: white; }
    [data-testid="stSidebar"] h1 { color: white; }
    .stat-card {
        background-color: rgba(255, 255, 255, 0.1); padding: 15px;
        border-radius: 10px; margin-bottom: 15px; border-left: 5px solid #3498DB;
    }
    .stat-title { font-size: 0.8rem; color: #CBD5E1; text-transform: uppercase; letter-spacing: 1px;}
    .stat-value { font-size: 1.8rem; color: white; font-weight: bold; }
    .stat-unit { font-size: 0.9rem; color: #94A3B8; }

    /* File Uploader */
    [data-testid="stFileUploader"] {
        border: 2px dashed #00508F; border-radius: 15px; padding: 20px; background-color: white;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize database (simulated with TinyDB for simplicity)
db = TinyDB('historical_records.json')
Rec = Query()

# # Configurable constants based on your pattern file
# DURATION_COLS = ['10min', '20min', '30min', '1hr', '2hr', '3hr', '6hr', '12hr', '24hr']

# ==========================================
# 2. CORE LOGIC FUNCTIONS
# ==========================================

def get_historical_stats():
    """Calculates summary statistics from the TinyDB."""
    all_records = db.all()
    if not all_records:
        return {}
    df_all = pd.DataFrame(all_records)
    
    stats = {
        'total_records': len(df_all),
        'unique_places': df_all['station_name'].nunique(),
        'place_most_data': df_all['station_name'].value_counts().idxmax(),
        'place_count_most': df_all['station_name'].value_counts().max(),
        'highest_rain_place': df_all.loc[df_all['24hr'].idxmax()]['station_name'],
        'highest_rain_val': df_all['24hr'].max(),
        'highest_rain_date': df_all.loc[df_all['24hr'].idxmax()]['date']
    }
    return stats


# --- SIDEBAR (HISTORICAL DASHBOARD) ---
with st.sidebar:
    st.markdown("<h1> Historical Records Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("---")
     
    # Get current stats from DB
    stats = get_historical_stats()
    
    if not stats:
        st.warning("Database is empty. Upload data to see statistics.")
    else:
        # Custom HTML Stat Cards
        st.markdown(f"""
            <div class="stat-card">
                <div class="stat-title">Total Records</div>
                <div class="stat-value">{stats['total_records']:,}</div>
                <div class="stat-unit">Daily Entries</div>
            </div>
            <div class="stat-card">
                <div class="stat-title">Unique Stations</div>
                <div class="stat-value">{stats['unique_places']}</div>
            </div>
            <div class="stat-card">
                <div class="stat-title">Most Recorded Station</div>
                <div class="stat-value">{stats['place_most_data']}</div>
                <div class="stat-unit">{stats['place_count_most']} Entries</div>
            </div>
            <div class="stat-card">
                <div class="stat-title">Highest 24hr Rain</div>
                <div class="stat-value">{stats['highest_rain_val']:.1f}</div>
                <div class="stat-unit">mm in {stats['highest_rain_place']} ({stats['highest_rain_date']})</div>
            </div>
        """, unsafe_allow_html=True)
    st.markdown("---")


import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from copy import copy
import io
import datetime
import re
import zipfile
import calendar


# --- 1. CORE LOGIC FUNCTIONS ---

def scan_files(uploaded_files):
    """Scans and returns the specific rows and filenames for the report."""
    error_log = {} # Dictionary to store filename: [rows]
    for up_file in uploaded_files:
        wb = openpyxl.load_workbook(up_file, data_only=True)
        ws = wb.active
        rows_found = []
        # Scan Column B (Index 2) from Row 7 onwards
        for r in range(7, ws.max_row + 1):
            prev = ws.cell(row=r-1, column=2).value
            curr = ws.cell(row=r, column=2).value
            if isinstance(prev, (int, float)) and isinstance(curr, (int, float)):
                if prev > curr:
                    rows_found.append(r)
        
        if rows_found:
            error_log[up_file.name] = rows_found
    return error_log

def process_batch_for_aggregation(uploaded_files, error_log):
    """Fixes errors, restores formulas, and identifies files for the audit ZIP."""
    final_batch = []
    fixed_files_only = []

    for up_file in uploaded_files:
        wb = openpyxl.load_workbook(up_file)
        ws = wb.active
        is_modified = False

        if up_file.name in error_log:
            is_modified = True
            for r in range(7, ws.max_row + 1):
                prev_b = ws.cell(row=r-1, column=2).value
                curr_b = ws.cell(row=r, column=2).value
                if isinstance(prev_b, (int, float)) and isinstance(curr_b, (int, float)):
                    if prev_b > curr_b:
                        ws.cell(row=r, column=2).value = prev_b # Apply Drag-Down [cite: 12]

        out = io.BytesIO()
        wb.save(out)
        
        file_data = {"name": up_file.name, "data": out.getvalue(), "wb": wb}
        final_batch.append(file_data)
        
        # Add to ZIP list only if it was an "Error" file 
        if is_modified:
            fixed_files_only.append(file_data)
            
    return final_batch, fixed_files_only

# --- 2. INTERFACE & RESET LOGIC ---

if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = 0

def reset_app():
    st.session_state['uploader_key'] += 1
    st.rerun()

st.title("Rain Intensity Duration Frequency (RIDF) Generator")
st.markdown("---")
st.subheader("Stage 1: Batch Data Validation & Correction")

if st.button("🧹 Clear & Reset Batch"):
    reset_app()

files = st.file_uploader(
    "Upload Batch HMDAS-12 Files", 
    type=['xlsx'], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}"
)

if files:
    # Consistency Check: Station, Year, Month
    name_pattern = r"\s+(.+)\s+(\d{4})(\d{2})(\d{2})"
    batch_meta = []
    valid_batch = True
    
    for f in files:
        match = re.search(name_pattern, f.name)
        if not match:
            st.error(f"❌ **Naming Error:** `{f.name}` does not follow naming convention.")
            valid_batch = False; break
        station, year, month, day = match.groups()
        batch_meta.append({'station': station.strip(), 'year': year, 'month': month, 'day': day, 'name': f.name})

    if valid_batch:
        # Check if all files belong to the same place and time 
        first = batch_meta[0]
        if any(m['station'] != first['station'] or m['month'] != first['month'] or m['year'] != first['year'] for m in batch_meta):
            st.error("🛑 **Batch Mismatch!** Ensure all files have the same Station, Month, and Year.")
            valid_batch = False

    if valid_batch:
        st.success(f"📂 **Validated Batch:** {first['station']} | {calendar.month_name[int(first['month'])]} {first['year']}")
        
        # Scan for Errors
        error_info = scan_files(files)
        
        if error_info:
            st.warning(f"⚠️ **Attention:** {len(error_info)} file(s) require correction.")
            for filename, rows in error_info.items():
                st.write(f"❌ **{filename}**: Negative dips at rows `{rows}`")
            
            # The Action Button
            if st.button(f"🚀 Fix {len(error_info)} Files & Prepare Stage 2"):
                clean_batch, fixed_only = process_batch_for_aggregation(files, error_info)
                st.session_state['ready_batch'] = clean_batch
                st.session_state['meta'] = batch_meta
                
                # --- ZIP DOWNLOAD CODE --- 
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for item in fixed_only:
                        zip_file.writestr(f"FIXED_{item['name']}", item['data'])
                
                st.success("Errors fixed. Download the audit ZIP below to verify.")
                st.download_button(
                    label="📥 Download Corrected Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"FIXED_AUDIT_{first['station']}_{first['month']}.zip",
                    mime="application/zip"
                )
        else:
            if st.button("🚀 All Clean - Proceed to Stage 2"):
                st.session_state['ready_batch'], _ = process_batch_for_aggregation(files, {})
                st.session_state['meta'] = batch_meta
                st.success("Batch ready for aggregation.")



# ... (KEEP ALL YOUR STAGE 1 CODE UNTOUCHED HERE) ...

# ==========================================
# 3. STAGE 2: THE AUTOMATED ARCHITECT
# ==========================================

# This section only appears once Stage 1 has successfully fixed or validated the batch
if st.session_state.get('ready_batch') is not None:
    st.markdown("---")
    st.header("Stage 2: Monthly Aggregation")
    st.info("The cleaned data from Stage 1 is loaded. We will now populate the HMDAS-14 Master Template.")
    
    if st.button("🏗️ Build Master Report (HMDAS-14)"):
        with st.spinner("Executing Master Template Injection..."):
            # 1. Load the Master XLSM (Preserving Macros)
            # Ensure 'HMDAS-14 Catarman 202110.xlsm' is in your server/folder
            try:
                wb_master = openpyxl.load_workbook("HMDAS-14 Catarman 202110.xlsm", keep_vba=True)
            except FileNotFoundError:
                st.error("🚨 Template file not found! Please ensure 'HMDAS-14 Catarman 202110.xlsm' is in the app directory.")
                st.stop()

            # 2. Extract Metadata from the Batch 
            meta = st.session_state['meta']
            station_val = meta[0]['station']
            year_val = meta[0]['year']
            # Convert month number to 3-letter uppercase (e.g., 10 -> 'OCT') 
            month_abbr = calendar.month_name[int(meta[0]['month'])][:3].upper()

            # --- A. REPLACE RED TEXT PLACEHOLDERS --- 
            
            # Update '3 Corr Monthly RR Max sheet' (Targets: 'Station' and 'YYYY') 
            if "3 Corr Monthly RR Max sheet" in wb_master.sheetnames:
                ws3 = wb_master["3 Corr Monthly RR Max"]
                for row in ws3.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str):
                            # Replace placeholders with batch-specific metadata 
                            cell.value = cell.value.replace("Station", station_val).replace("YYYY", year_val)

            # Update '1 Pattern 10min Chart Reading' (Target: 'MMM') 
            if "1 Pattern 10min Chart Reading" in wb_master.sheetnames:
                ws1 = wb_master["1 Pattern 10min Chart Reading"]
                for row in ws1.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and "MMM" in cell.value:
                            cell.value = cell.value.replace("MMM", month_abbr)

            # --- B. DATA POPULATION (Monthly chart Reading) --- 
            
            ws_monthly = wb_master["Monthly Rainfall Chart Reading"]
            # Loop through the cleaned batch (Error-free files from Stage 1) 
            for i, item in enumerate(st.session_state['ready_batch']):
                # Identify column based on the day (Day 01 = Column B) 
                day_num = int(meta[i]['day'])
                target_col = day_num + 1 
                
                # Load the fixed data from Stage 1 
                wb_fixed = openpyxl.load_workbook(io.BytesIO(item['data']), data_only=True)
                ws_fixed = wb_fixed.active
                
                # Copy Reading (Column B, Rows 6-150) into the Monthly Master 
                for r in range(6, 151):
                    val = ws_fixed.cell(row=r, column=2).value
                    ws_monthly.cell(row=r, column=target_col).value = val

            # --- C. FINALIZE & DOWNLOAD ---
            out_final = io.BytesIO()
            wb_master.save(out_final)
            
            st.success(f"Aggregation Complete for {station_val} ({month_abbr} {year_val})")
            
            st.download_button(
                label="📥 Download Aggregated HMDAS-14 Master",
                data=out_final.getvalue(),
                file_name=f"HMDAS-14_{station_val}_{year_val}_{month_abbr}.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )
            
            st.info("💡 **Ready for Macros:** Open the file and run `module 2` to complete the extraction. ")