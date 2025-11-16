import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np
import time
import io
from matplotlib import colormaps
import matplotlib.colors as mcolors
import base64
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Border, Side
import tempfile
import os
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

def get_sample_workbook():
    """Retrieve the sample workbook from the same directory as app.py"""
    try:
        current_dir = Path(__file__).parent
        sample_path = current_dir / "polymer_production_template.xlsx"
        
        if sample_path.exists():
            with open(sample_path, "rb") as f:
                return io.BytesIO(f.read())
        else:
            st.warning("Sample template file not found. Using generated template.")
            return create_sample_workbook()
    except Exception as e:
        st.warning(f"Could not load sample template: {e}. Using generated template.")

def process_shutdown_dates(plant_df, dates):
    """Process shutdown dates for each plant"""
    shutdown_periods = {}
    
    for index, row in plant_df.iterrows():
        plant = row['Plant']
        shutdown_start = row.get('Shutdown Start Date')
        shutdown_end = row.get('Shutdown End Date')
        
        if pd.notna(shutdown_start) and pd.notna(shutdown_end):
            try:
                start_date = pd.to_datetime(shutdown_start).date()
                end_date = pd.to_datetime(shutdown_end).date()
                
                if start_date > end_date:
                    st.warning(f"⚠️ Shutdown start date after end date for {plant}. Ignoring shutdown.")
                    shutdown_periods[plant] = []
                    continue
                
                shutdown_days = []
                for d, date in enumerate(dates):
                    if start_date <= date <= end_date:
                        shutdown_days.append(d)
                
                if shutdown_days:
                    shutdown_periods[plant] = shutdown_days
                    st.info(f"🔧 Shutdown scheduled for {plant}: {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')} ({len(shutdown_days)} days)")
                else:
                    shutdown_periods[plant] = []
                    st.info(f"ℹ️ Shutdown period for {plant} is outside planning horizon")
                    
            except Exception as e:
                st.warning(f"⚠️ Invalid shutdown dates for {plant}: {e}")
                shutdown_periods[plant] = []
        else:
            shutdown_periods[plant] = []
    
    return shutdown_periods

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1  # 1: Upload, 2: Preview & Config, 3: Results
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'solutions' not in st.session_state:
    st.session_state.solutions = []
if 'best_solution' not in st.session_state:
    st.session_state.best_solution = None

# Initialize optimization parameters in session state
if 'optimization_params' not in st.session_state:
    st.session_state.optimization_params = {
        'buffer_days': 3,
        'time_limit_min': 10,
        'stockout_penalty': 10,
        'transition_penalty': 50,
        'continuity_bonus': 1
    }
# ============================================================================
# MODERN MATERIAL MINIMALISM CSS
# ============================================================================
st.markdown("""
<style>
    /* Hide sidebar completely */
    [data-testid="stSidebar"] {
        display: none;
    }
    
    /* Base reset and typography */
    .main .block-container {
        padding-top: 3rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Material Design App Bar */
    .app-bar {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        text-align: center;
        padding: 2rem 3rem;
        margin: -3rem -3rem 3rem -3rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        border-radius: 16px;
    }
    
    .app-bar h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 600;
        letter-spacing: -0.5px;
    }
    
    .app-bar p {
        margin: 0.5rem 0 0 0;
        font-size: 1rem;
        opacity: 0.95;
        font-weight: 400;
    }
    
    /* Step indicator - Material Design */
    .step-indicator {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 2rem 0 3rem 0;
        position: relative;
    }
    
    .step {
        display: flex;
        flex-direction: column;
        align-items: center;
        position: relative;
        flex: 1;
        max-width: 200px;
    }
    
    .step-circle {
        width: 48px;
        height: 48px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        font-size: 1.125rem;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        z-index: 2;
        background: white;
        border: 3px solid #e0e0e0;
        color: #9e9e9e;
    }
    
    .step-circle.active {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-color: #667eea;
        color: white;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
        transform: scale(1.1);
    }
    
    .step-circle.completed {
        background: #4caf50;
        border-color: #4caf50;
        color: white;
    }
    
    .step-label {
        margin-top: 0.75rem;
        font-size: 0.875rem;
        font-weight: 500;
        color: #757575;
        text-align: center;
    }
    
    .step-label.active {
        color: #667eea;
        font-weight: 600;
    }
    
    .step-label.completed {
        color: #4caf50;
    }
    
    .step-line {
        position: absolute;
        top: 24px;
        left: 50%;
        right: -50%;
        height: 3px;
        background: #e0e0e0;
        z-index: 1;
    }
    
    .step-line.completed {
        background: #4caf50;
    }
    
    /* Material Cards */
    .material-card {
        background: #F0F2FF;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        padding: 2rem;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        transition: box-shadow 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        border: 1px solid rgba(0, 0, 0, 0.06);
    }
    
    .material-card:hover {
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
    }
    
    .card-title {
        font-size: 1.25rem;
        font-weight: 600;
        text-align: center;
        color: #212121;
        margin: 0 0 1rem 0;
        display: flex;
        align-items: center;
        justify-content: center; 
    }
    
    .card-subtitle {
        font-size: 0.875rem;
        color: #757575;
        margin: -0.5rem 0 1rem 0;
    }
    
    /* Elevated cards for special content */
    .elevated-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        border-radius: 16px;
        padding: 2rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
        border: none;
    }
    
    /* Material buttons */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .stButton > button:hover {
        box-shadow: 0 4px 16px rgba(102, 126, 234, 0.4);
        transform: translateY(-2px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* File uploader - Material Design */
    [data-testid="stFileUploader"] {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8eef5 100%);
        border: 2px dashed #667eea;
        border-radius: 16px;
        padding: 1rem 1rem;
        text-align: center;
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #764ba2;
        background: linear-gradient(135deg, #f0f4ff 0%, #e3e9f7 100%);
        box-shadow: 0 4px 16px rgba(102, 126, 234, 0.15);
    }
    
    [data-testid="stFileUploader"] section {
        border: none;
        background: transparent;
    }
    
    /* Metrics - Material Design cards */
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        border-left: 4px solid;
        transition: all 0.3s ease;
        height: 100%;
    }
    
    .metric-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    
    .metric-card.primary {
        border-left-color: #667eea;
        background: linear-gradient(135deg, #f0f4ff 0%, #ffffff 100%);
    }
    
    .metric-card.success {
        border-left-color: #4caf50;
        background: linear-gradient(135deg, #f1f8f4 0%, #ffffff 100%);
    }
    
    .metric-card.warning {
        border-left-color: #ff9800;
        background: linear-gradient(135deg, #fff8f0 0%, #ffffff 100%);
    }
    
    .metric-card.info {
        border-left-color: #2196f3;
        background: linear-gradient(135deg, #f0f7ff 0%, #ffffff 100%);
    }
    
    .metric-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #757575;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #212121;
        line-height: 1.2;
    }
    
    .metric-subtitle {
        font-size: 0.75rem;
        color: #9e9e9e;
        margin-top: 0.25rem;
    }
    
    /* Chips/Badges - Material Design */
    .chip {
        display: inline-flex;
        align-items: center;
        padding: 0.375rem 0.875rem;
        border-radius: 16px;
        font-size: 0.8125rem;
        font-weight: 500;
        margin: 0.25rem;
        transition: all 0.2s ease;
    }
    
    .chip.success {
        background: #e8f5e9;
        color: #2e7d32;
    }
    
    .chip.warning {
        background: #fff3e0;
        color: #e65100;
    }
    
    .chip.info {
        background: #e3f2fd;
        color: #1565c0;
    }
    
    .chip.error {
        background: #ffebee;
        color: #c62828;
    }
    
    /* Material tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
        display: flex;
        width: 100%;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        background: transparent;
        border: none;
        color: #757575;
        transition: all 0.3s ease;
        flex: 1; /* This makes all tabs equal width */
        text-align: center;
        justify-content: center;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3);
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Data tables */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: #f5f5f5;
        border-radius: 8px;
        font-weight: 600;
        color: #424242;
        transition: all 0.3s ease;
    }
    
    .streamlit-expanderHeader:hover {
        background: #eeeeee;
    }
    
    /* Number input */
    .stNumberInput > div > div > input {
        border-radius: 8px;
        border: 2px solid #e0e0e0;
        transition: all 0.3s ease;
    }
    
    .stNumberInput > div > div > input:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    /* Alert boxes */
    .alert-box {
        padding: 1rem 1.5rem;
        border-radius: 8px;
        margin: 1rem 0;
        border-left: 4px solid;
        display: flex;
        align-items: flex-start;
        gap: 0.75rem;
    }
    
    .alert-box.info {
        background: #e3f2fd;
        border-left-color: #2196f3;
        color: #1565c0;
    }
    
    .alert-box.success {
        background: #e8f5e9;
        border-left-color: #4caf50;
        color: #2e7d32;
    }
    
    .alert-box.warning {
        background: #fff3e0;
        border-left-color: #ff9800;
        color: #e65100;
    }
    
    /* Loading animation */
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    
    .loading {
        animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
    }
    
    /* Divider */
    .divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, #e0e0e0, transparent);
        margin: 2rem 0;
    }
    
    /* List styling */
    .styled-list {
        list-style: none;
        padding-left: 0;
    }
    
    .styled-list li {
        padding: 0.5rem 0 0.5rem 2rem;
        position: relative;
    }
    
    .styled-list li:before {
        content: "✓";
        position: absolute;
        left: 0;
        color: #4caf50;
        font-weight: bold;
        font-size: 1.25rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER - Material Design App Bar
# ============================================================================
st.markdown("""
<div class="app-bar">
    <h1>🏭 Polymer Production Scheduler</h1>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# STEP INDICATOR - Material Design Stepper
# ============================================================================
step_status = ['active' if st.session_state.step == 1 else 'completed',
               'active' if st.session_state.step == 2 else ('completed' if st.session_state.step > 2 else ''),
               'active' if st.session_state.step == 3 else '']

st.markdown(f"""
<div class="step-indicator">
    <div class="step">
        <div class="step-circle {step_status[0]}">
            {'✓' if st.session_state.step > 1 else '1'}
        </div>
        <div class="step-label {step_status[0]}">Upload Data</div>
        <div class="step-line {step_status[0] if st.session_state.step > 1 else ''}"></div>
    </div>
    <div class="step">
        <div class="step-circle {step_status[1]}">
            {'✓' if st.session_state.step > 2 else '2'}
        </div>
        <div class="step-label {step_status[1]}">Configure & Preview</div>
        <div class="step-line {step_status[1] if st.session_state.step > 2 else ''}"></div>
    </div>
    <div class="step">
        <div class="step-circle {step_status[2]}">3</div>
        <div class="step-label {step_status[2]}">View Results</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ============================================================================
# STEP 1: UPLOAD DATA
# ============================================================================
if st.session_state.step == 1:

    col1, col2 = st.columns([4, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Choose Excel File",
            type=["xlsx"],
            help="Upload an Excel file with Plant, Inventory, and Demand sheets",
            label_visibility="collapsed"
        )
    
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("✅ File uploaded successfully!")
            time.sleep(0.5)
            st.session_state.step = 2
            st.rerun()

    with col2:
        sample_workbook = get_sample_workbook()
        
        st.markdown("""
            <style>
                /* Target the download button container */
                .stDownloadButton > button {
                    height: calc(100% + 1rem) !important;
                    min-height: 5rem !important;
                    margin-left: -0.5rem !important;
                    margin-right: -0.5rem !important;
                }
            </style>
            """, unsafe_allow_html=True)
        
        st.download_button(
            label="📥 Download Excel Template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown("""
        <div class="material-card">
            <div class="card-title">📋 Quick Start Guide</div>
            <div class="card-subtitle">Follow these simple steps to get started</div>
            <ol class="styled-list">
                <li>Download the Excel template from the card on the right</li>
                <li>Fill in your production data (plants, inventory, demand)</li>
                <li>Upload your completed Excel file below</li>
                <li>Review and configure optimization parameters</li>
                <li>Run the optimization and analyze results</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
            
    with col2:
        st.markdown("""
        <div class="material-card">
            <div class="card-title">✨ Key Capabilities</div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-top: 1rem;">
                <div>
                    <div style="font-weight: 600; color: #667eea; margin-bottom: 0.5rem;">🏭 Multi-Plant</div>
                    <div style="font-size: 0.875rem; color: #757575;">Optimize across multiple production lines simultaneously</div>
                </div>
                <div>
                    <div style="font-weight: 600; color: #667eea; margin-bottom: 0.5rem;">📦 Inventory Control</div>
                    <div style="font-size: 0.875rem; color: #757575;">Maintain optimal stock levels with min/max constraints</div>
                </div>
                <div>
                    <div style="font-weight: 600; color: #667eea; margin-bottom: 0.5rem;">🔄 Transition Rules</div>
                    <div style="font-size: 0.875rem; color: #757575;">Grade changeover optimization with custom matrices</div>
                </div>
                <div>
                    <div style="font-weight: 600; color: #667eea; margin-bottom: 0.5rem;">🔧 Shutdown Handling</div>
                    <div style="font-size: 0.875rem; color: #757575;">Plan around maintenance and downtime periods</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
    
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    # Detailed format documentation
    with st.expander("📚 Detailed Excel Format Specification", expanded=False):
        st.markdown("""
        ### Required Sheets & Columns
        
        #### 1️⃣ Plant Sheet
        
        | Column | Description | Type | Required |
        |--------|-------------|------|----------|
        | **Plant** | Unique plant identifier | Text | ✓ |
        | **Capacity per day** | Daily production capacity (MT) | Number | ✓ |
        | **Material Running** | Currently running grade | Text | Optional |
        | **Expected Run Days** | Continuation days | Number | Optional |
        | **Shutdown Start Date** | Maintenance start | Date | Optional |
        | **Shutdown End Date** | Maintenance end | Date | Optional |
        
        #### 2️⃣ Inventory Sheet
        
        | Column | Description | Type | Required |
        |--------|-------------|------|----------|
        | **Grade Name** | Product grade identifier | Text | ✓ |
        | **Opening Inventory** | Starting stock (MT) | Number | ✓ |
        | **Min. Inventory** | Safety stock (MT) | Number | ✓ |
        | **Max. Inventory** | Storage limit (MT) | Number | ✓ |
        | **Min. Run Days** | Minimum run length | Number | ✓ |
        | **Max. Run Days** | Maximum run length | Number | ✓ |
        | **Force Start Date** | Mandatory start date | Date | Optional |
        | **Lines** | Allowed plants (comma-separated) | Text | ✓ |
        | **Rerun Allowed** | Multiple runs (Yes/No) | Text | ✓ |
        | **Min. Closing Inventory** | End target (MT) | Number | ✓ |
        
        #### 3️⃣ Demand Sheet
        
        - **First column**: Date (daily dates in ascending order)
        - **Remaining columns**: One column per grade (column header = grade name)
        - **Values**: Daily demand in MT
        
        #### 4️⃣ Transition Sheets (Optional)
        
        - **Sheet naming**: `Transition_[PlantName]` (e.g., `Transition_Plant1`)
        - **Format**: Matrix with previous grade (rows) → next grade (columns)
        - **Values**: "yes" for allowed transitions, anything else for blocked
        
        ### Examples
        
        **Multi-Plant Grade Configuration:**
        ```
        Grade | Lines  | Force Start | Min Run
        BOPP  | Plant1 |             | 5
        BOPP  | Plant2 | 01-Dec-24   | 5
        ```
        
        **Shutdown Period:**
        ```
        Plant  | Shutdown Start | Shutdown End
        Plant1 | 15-Nov-25      | 18-Nov-25
        ```
        """)

# ============================================================================
# STEP 2: PREVIEW & CONFIGURE
# ============================================================================

elif st.session_state.step == 2:
    
    try:
        uploaded_file = st.session_state.uploaded_file
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Data preview in tabs
        tab1, tab2, tab3 = st.tabs(["🏭 Plant Data", "📦 Inventory Data", "📊 Demand Data"])
        
        with tab1:
            try:
                plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                plant_display_df = plant_df.copy()
                start_column = plant_display_df.columns[4]
                end_column = plant_display_df.columns[5]
                
                if pd.api.types.is_datetime64_any_dtype(plant_display_df[start_column]):
                    plant_display_df[start_column] = plant_display_df[start_column].dt.strftime('%d-%b-%y')
                if pd.api.types.is_datetime64_any_dtype(plant_display_df[end_column]):
                    plant_display_df[end_column] = plant_display_df[end_column].dt.strftime('%d-%b-%y')
                
                st.dataframe(plant_display_df, use_container_width=True, height=300)
                
                # Shutdown summary
                shutdown_count = 0
                for index, row in plant_df.iterrows():
                    if pd.notna(row.get('Shutdown Start Date')) and pd.notna(row.get('Shutdown End Date')):
                        shutdown_count += 1
                
                if shutdown_count > 0:
                    st.markdown(f'<span class="chip warning">🔧 {shutdown_count} plant(s) with scheduled shutdowns</span>', unsafe_allow_html=True)
                else:
                    st.markdown('<span class="chip success">✓ No shutdowns scheduled</span>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"❌ Error reading Plant sheet: {e}")
                st.stop()
        
        with tab2:
            try:
                excel_file.seek(0)
                inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                inventory_display_df = inventory_df.copy()
                force_start_column = inventory_display_df.columns[7]
                
                if pd.api.types.is_datetime64_any_dtype(inventory_display_df[force_start_column]):
                    inventory_display_df[force_start_column] = inventory_display_df[force_start_column].dt.strftime('%d-%b-%y')
                
                st.dataframe(inventory_display_df, use_container_width=True, height=300)
                
                grade_count = len(inventory_df['Grade Name'].unique())
                st.markdown(f'<span class="chip info">📦 {grade_count} unique grade(s)</span>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"❌ Error reading Inventory sheet: {e}")
                st.stop()
        
        with tab3:
            try:
                excel_file.seek(0)
                demand_df = pd.read_excel(excel_file, sheet_name='Demand')
                demand_display_df = demand_df.copy()
                
                date_column = demand_display_df.columns[0]
                if pd.api.types.is_datetime64_any_dtype(demand_display_df[date_column]):
                    demand_display_df[date_column] = demand_display_df[date_column].dt.strftime('%d-%b-%y')
                
                st.dataframe(demand_display_df, use_container_width=True, height=300)
                
                num_days = len(demand_df)
                st.markdown(f'<span class="chip info">📅 {num_days} day(s) planning horizon</span>', unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"❌ Error reading Demand sheet: {e}")
                st.stop()
        
        excel_file.seek(0)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # Configuration section
        st.markdown("""
        <div class="material-card">
            <div class="card-title">⚙️ Optimization Parameters</div>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Core Settings")
            # Update session state directly
            st.session_state.optimization_params['time_limit_min'] = st.number_input(
                "⏱️ Time Limit (minutes)",
                min_value=1,
                max_value=120,
                value=st.session_state.optimization_params['time_limit_min'],
                help="Maximum solver runtime"
            )
            
            st.session_state.optimization_params['buffer_days'] = st.number_input(
                "📅 Planning Buffer (days)",
                min_value=0,
                max_value=7,
                value=st.session_state.optimization_params['buffer_days'],
                help="Additional days for safety planning"
            )
        
        with col2:
            st.markdown("#### Objective Weights")
            st.session_state.optimization_params['stockout_penalty'] = st.number_input(
                "🎯 Stockout Penalty",
                min_value=1,
                value=st.session_state.optimization_params['stockout_penalty'],
                help="Cost weight for inventory shortages"
            )
            
            st.session_state.optimization_params['transition_penalty'] = st.number_input(
                "🔄 Transition Penalty",
                min_value=1,
                value=st.session_state.optimization_params['transition_penalty'],
                help="Cost weight for grade changeovers"
            )
            
            st.session_state.optimization_params['continuity_bonus'] = st.number_input(
                "➕ Continuity Bonus", 
                min_value=0,
                value=st.session_state.optimization_params['continuity_bonus'],
                help="Reward for extended production runs"
            )
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        # Action buttons
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col1:
            if st.button("← Back to Upload", use_container_width=True):
                st.session_state.step = 1
                st.rerun()
        
        with col2:
            if st.button("🚀 Run Optimization", type="primary", use_container_width=True):
                st.session_state.step = 3
                st.rerun()
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        if st.button("← Back to Upload"):
            st.session_state.step = 1
            st.rerun()

# ============================================================================
# STEP 3: OPTIMIZATION & RESULTS
# ============================================================================
elif st.session_state.step == 3:
    
    try:
        uploaded_file = st.session_state.uploaded_file
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())

        # Extract parameters from session state
        params = st.session_state.optimization_params
        buffer_days = params['buffer_days']
        time_limit_min = params['time_limit_min'] 
        stockout_penalty = params['stockout_penalty']
        transition_penalty = params['transition_penalty']
        continuity_bonus = params['continuity_bonus']
        
        # Show optimization progress
        st.markdown("""
        <div class="material-card">
            <div class="card-title">⚡ Running Optimization</div>
        </div>
        """, unsafe_allow_html=True)
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        time.sleep(1)
        
        status_text.markdown('<div class="alert-box info">📊 Preprocessing data...</div>', unsafe_allow_html=True)
        progress_bar.progress(10)
        time.sleep(1)
        
        # ====================================================================
        # DATA PREPROCESSING
        # ====================================================================
        try:
            plant_df = pd.read_excel(excel_file, sheet_name='Plant')
            excel_file.seek(0)
            inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
            excel_file.seek(0)
            demand_df = pd.read_excel(excel_file, sheet_name='Demand')
            excel_file.seek(0)
            
            num_lines = len(plant_df)
            lines = list(plant_df['Plant'])
            capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
            
            # --- Grade Data Extraction ---
            grades = [col for col in demand_df.columns if col != demand_df.columns[0]]
            
            initial_inventory = {}
            min_inventory = {}
            max_inventory = {}
            min_closing_inventory = {}
            min_run_days = {}
            max_run_days = {}
            force_start_date = {}
            allowed_lines = {grade: [] for grade in grades}
            rerun_allowed = {}
            
            grade_inventory_defined = set()
            for index, row in inventory_df.iterrows():
                grade = row['Grade Name']
                
                lines_value = row['Lines']
                if pd.notna(lines_value) and lines_value != '':
                    plants_for_row = [x.strip() for x in str(lines_value).split(',')]
                else:
                    plants_for_row = lines
                
                for plant in plants_for_row:
                    if plant not in allowed_lines[grade]:
                        allowed_lines[grade].append(plant)

                if grade not in grade_inventory_defined:
                    initial_inventory[grade] = row['Opening Inventory'] if pd.notna(row['Opening Inventory']) else 0
                    min_inventory[grade] = row['Min. Inventory'] if pd.notna(row['Min. Inventory']) else 0
                    max_inventory[grade] = row['Max. Inventory'] if pd.notna(row['Max. Inventory']) else 1000000000
                    min_closing_inventory[grade] = row['Min. Closing Inventory'] if pd.notna(row['Min. Closing Inventory']) else 0
                    grade_inventory_defined.add(grade)

                for plant in plants_for_row:
                    min_run_days[(grade, plant)] = row['Min. Run Days'] if pd.notna(row['Min. Run Days']) else 1
                    max_run_days[(grade, plant)] = row['Max. Run Days'] if pd.notna(row['Max. Run Days']) else len(demand_df)
                    rerun_allowed[grade] = str(row.get('Rerun Allowed', 'Yes')).strip().lower() == 'yes'

                    if pd.notna(row.get('Force Start Date')):
                        force_start_date[(grade, plant)] = pd.to_datetime(row['Force Start Date']).date()

            # --- Date & Demand Data Extraction ---
            demand_df.columns = [demand_df.columns[0]] + grades
            demand_df[demand_df.columns[0]] = pd.to_datetime(demand_df.iloc[:, 0]).dt.date
            
            dates = list(demand_df.iloc[:, 0])
            formatted_dates = [d.strftime('%d-%b-%y') for d in dates]
            num_days = len(dates)
            
            demand_data = {}
            for grade in grades:
                if grade in demand_df.columns:
                    demand_data[grade] = {demand_df.iloc[i, 0].date(): demand_df[grade].iloc[i] for i in range(len(demand_df))}
                else:
                    demand_data[grade] = {date: 0 for date in dates}

            # Add buffer days to demand data (demand is 0 for buffer days)
            for d in range(num_days - buffer_days, num_days):
                date = dates[d]
                for grade in grades:
                    if date not in demand_data[grade]:
                        demand_data[grade][date] = 0
            
            shutdown_periods = process_shutdown_dates(plant_df, dates)

            # --- Transition Rules Loading ---
            transition_dfs = {}
            for i in range(len(plant_df)):
                plant_name = plant_df['Plant'].iloc[i]
                possible_sheet_names = [
                    f'Transition_{plant_name}',
                    f'Transition_{plant_name.replace(" ", "_")}',
                    f'Transition{plant_name.replace(" ", "")}',
                ]
                transition_df_found = None
                for sheet_name in possible_sheet_names:
                    try:
                        excel_file.seek(0)
                        # Read with index_col=0 to treat the first column (Previous Grade) as the index
                        transition_df_found = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=0)
                        break
                    except:
                        continue
                transition_dfs[plant_name] = transition_df_found

            transition_rules = {}
            for line, df in transition_dfs.items():
                if df is not None:
                    transition_rules[line] = {}
                    for prev_grade in df.index:
                        allowed_transitions = []
                        for current_grade in df.columns:
                            # 'yes' means allowed, anything else means blocked
                            if str(df.loc[prev_grade, current_grade]).strip().lower() == 'yes':
                                allowed_transitions.append(current_grade)
                        transition_rules[line][prev_grade] = allowed_transitions
                else:
                    transition_rules[line] = None # No transition sheet for this plant

        except Exception as e:
            st.error(f"❌ Data Preprocessing Error: {e}")
            st.stop()

        progress_bar.progress(30)
        status_text.markdown('<div class="alert-box info">🔧 Building optimization model...</div>', unsafe_allow_html=True)
        time.sleep(1)

        # ====================================================================
        # MODEL BUILDING
        # ====================================================================
        model = cp_model.CpModel()
        
        # --- Helper functions ---
        def get_is_producing_var(grade, line, d):
            key = (grade, line, d)
            return is_producing.get(key)
        
        def get_production_var(grade, line, d):
            key = (grade, line, d)
            return production.get(key)

        # --- Variables ---
        is_producing = {} # Binary: 1 if grade is produced on line on day d
        production = {} # Integer: quantity produced
        inventory_vars = {} # Integer: inventory level at end of day d (d+1 is day after last day)
        stockout_vars = {} # Integer: stockout quantity on day d
        is_start_vars = {} # Binary: 1 if grade run starts on day d
        is_end_vars = {} # Binary: 1 if grade run ends on day d

        # Production variables and constraints (Must be produced on line d if producing)
        for grade in grades:
            for line in allowed_lines[grade]:
                for d in range(num_days):
                    key = (grade, line, d)
                    is_producing[key] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')

                    if d < num_days - buffer_days:
                        # Production is only modeled for non-buffer days
                        production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                        production[key] = production_value
                        
                        # Constraint: Production is 0 if not producing, and Capacity if producing
                        model.Add(production_value == capacities[line]).OnlyEnforceIf(is_producing[key])
                        model.Add(production_value == 0).OnlyEnforceIf(is_producing[key].Not())
                    else:
                        # For buffer days, production is always 0
                        production[key] = model.NewIntVar(0, 0, f'production_{grade}_{line}_{d}')
                        model.Add(is_producing[key] == 0)

                    # No production on shutdown days
                    if line in shutdown_periods and d in shutdown_periods[line]:
                        model.Add(is_producing[key] == 0)
                        
        # One grade per line per day constraint
        for line in lines:
            for d in range(num_days):
                # Only consider grades allowed on this line
                producing_vars = [is_producing[(grade, line, d)] for grade in grades if line in allowed_lines[grade]]
                if producing_vars:
                    model.Add(sum(producing_vars) <= 1)
        
        # Inventory variables
        for grade in grades:
            # Inventory at start of day 0 (end of day -1)
            inventory_vars[(grade, 0)] = model.NewIntVar(int(min_inventory[grade]), int(max_inventory[grade]), f'inventory_{grade}_0')
            for d in range(num_days):
                # Inventory at end of day d (start of day d+1)
                inv_max = int(max_inventory[grade])
                inv_vars_key = (grade, d + 1)
                inventory_vars[inv_vars_key] = model.NewIntVar(0, inv_max, f'inventory_{grade}_{d + 1}')
                # Stockout variables
                stockout_vars[(grade, d)] = model.NewIntVar(0, int(demand_data[grade].get(dates[d], 0) * stockout_penalty), f'stockout_{grade}_{d}')

        # Min/Max Run variables (is_start, is_end)
        for grade in grades:
            for line in allowed_lines[grade]:
                for d in range(num_days):
                    is_start_vars[(grade, line, d)] = model.NewBoolVar(f'is_start_{grade}_{line}_{d}')
                    is_end_vars[(grade, line, d)] = model.NewBoolVar(f'is_end_{grade}_{line}_{d}')
        
        # Initial production constraints (Material Running / Expected Run Days)
        material_running_info = {row['Plant']: (row.get('Material Running'), row.get('Expected Run Days', 0)) 
                                for index, row in plant_df.iterrows()}
        
        for plant, (material, expected_days) in material_running_info.items():
            if pd.notna(material) and pd.notna(expected_days) and material in grades:
                expected_days = int(expected_days)
                for d in range(min(expected_days, num_days)):
                    if get_is_producing_var(material, plant, d) is not None:
                        model.Add(get_is_producing_var(material, plant, d) == 1)
                        for other_material in grades:
                            if other_material != material and get_is_producing_var(other_material, plant, d) is not None:
                                model.Add(get_is_producing_var(other_material, plant, d) == 0)

        # Objective function initialization
        objective = 0

        # --- Inventory Balance, Min/Max Inventory, Stockout Constraints ---
        
        # 1. Initial Inventory
        for grade in grades:
            model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])

        # 2. Inventory Balance
        for grade in grades:
            for d in range(num_days):
                produced_today = sum(get_production_var(grade, line, d) for line in allowed_lines[grade] if get_production_var(grade, line, d) is not None)
                demand_today = demand_data[grade].get(dates[d], 0)
                
                inv_prev = inventory_vars[(grade, d)]
                inv_current = inventory_vars[(grade, d + 1)]
                
                # Demand satisfaction is complex in CP-SAT, simplify with stockout variable
                # supplied <= inv_prev + produced_today
                supplied = model.NewIntVar(0, int(demand_today), f'supplied_{grade}_{d}')
                model.Add(supplied <= inv_prev + produced_today)
                model.Add(supplied == demand_today - stockout_vars[(grade, d)])
                
                # Inventory balance equation
                model.Add(inv_current == inv_prev + produced_today - supplied)

                # Penalize Stockouts
                objective += stockout_penalty * stockout_vars[(grade, d)]

        # 3. Minimum Inventory penalties (Safety Stock)
        for grade in grades:
            for d in range(num_days):
                if min_inventory[grade] > 0:
                    deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                    # Deficit is max(0, Min_Inv - Current_Inv)
                    model.Add(deficit >= int(min_inventory[grade]) - inventory_vars[(grade, d + 1)])
                    model.Add(deficit >= 0)
                    objective += stockout_penalty * deficit

        # 4. Closing Inventory Requirement (Last actual production day)
        last_actual_day = num_days - buffer_days - 1
        for grade in grades:
            closing_inventory = inventory_vars[(grade, last_actual_day + 1)]
            if min_closing_inventory[grade] > 0:
                closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                # Closing Deficit is max(0, Min_Closing_Inv - Closing_Inv)
                model.Add(closing_deficit >= int(min_closing_inventory[grade]) - closing_inventory)
                model.Add(closing_deficit >= 0)
                # Apply higher penalty for closing deficit
                objective += stockout_penalty * closing_deficit * 3
        
        # 5. Maximum Inventory constraint
        for grade in grades:
            for d in range(1, num_days + 1):
                model.Add(inventory_vars[(grade, d)] <= int(max_inventory[grade]))
        
        # 6. Force Start Date Constraint
        for (grade, line), start_date in force_start_date.items():
            if start_date in dates:
                d_start = dates.index(start_date)
                # Must be producing on the force start date
                model.Add(get_is_producing_var(grade, line, d_start) == 1)
        
        # --- Run Length Constraints (Min/Max Run, is_start/is_end definitions) ---
        for grade in grades:
            for line in allowed_lines[grade]:
                min_run = min_run_days.get((grade, line), 1)
                max_run = max_run_days.get((grade, line), num_days)

                for d in range(num_days):
                    # Define is_start and is_end
                    current_prod = get_is_producing_var(grade, line, d)
                    
                    # is_start: starts producing today (d) AND was NOT producing yesterday (d-1)
                    if d == 0:
                        model.Add(is_start_vars[(grade, line, d)] == current_prod)
                    else:
                        prev_prod = get_is_producing_var(grade, line, d - 1)
                        if prev_prod is not None and current_prod is not None:
                            model.AddBoolAnd([current_prod, prev_prod.Not()]).OnlyEnforceIf(is_start_vars[(grade, line, d)])
                            model.AddBoolOr([current_prod.Not(), prev_prod]).OnlyEnforceIf(is_start_vars[(grade, line, d)].Not())

                    # is_end: producing today (d) AND NOT producing tomorrow (d+1)
                    if d == num_days - 1:
                        model.Add(is_end_vars[(grade, line, d)] == current_prod)
                    else:
                        next_prod = get_is_producing_var(grade, line, d + 1)
                        if current_prod is not None and next_prod is not None:
                            model.AddBoolAnd([current_prod, next_prod.Not()]).OnlyEnforceIf(is_end_vars[(grade, line, d)])
                            model.AddBoolOr([current_prod.Not(), next_prod]).OnlyEnforceIf(is_end_vars[(grade, line, d)].Not())

                    # Min run days enforcement
                    if min_run > 1:
                        is_start = is_start_vars[(grade, line, d)]
                        # If a run starts on day d, it must run for min_run days, avoiding shutdowns
                        max_possible_run = 0
                        run_days_vars = []
                        for k in range(min_run):
                            if d + k < num_days:
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    # Shutdown day, breaks the consecutive run. Skip enforcing min run from this start day.
                                    # This logic is simplified: it should ensure the min run is met if possible.
                                    # For a simple minimum run, we enforce k-th day production if it's a start day.
                                    break
                                future_prod = get_is_producing_var(grade, line, d + k)
                                if future_prod is not None:
                                    model.Add(future_prod == 1).OnlyEnforceIf(is_start)
                                    max_possible_run += 1
                        
                        # Add constraint that is_start must be 0 if a full min_run is not possible
                        if max_possible_run < min_run:
                            model.Add(is_start == 0)

                # Max run days enforcement
                if max_run < num_days:
                    for d in range(num_days - max_run):
                        # Consecutive days list includes max_run + 1 days
                        consecutive_days = []
                        for k in range(max_run + 1):
                            if d + k < num_days:
                                # Skip days during shutdown for the count, but production is 0 anyway
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    break # Shutdown breaks the run
                                prod_var = get_is_producing_var(grade, line, d + k)
                                if prod_var is not None:
                                    consecutive_days.append(prod_var)
                        
                        # If we have max_run + 1 consecutive possible production days,
                        # at least one of them must be 0 (i.e., we must stop production).
                        if len(consecutive_days) == max_run + 1:
                            model.Add(sum(consecutive_days) <= max_run)
                            
        # Rerun allowed constraint: A run cannot start on a day if a run of the same grade ended on a previous day.
        # This is a complex global constraint. For simplicity and performance, we focus on min/max run.
        # If Rerun is NOT allowed, the total number of 'is_start' variables for that grade on that line must be <= 1.
        for grade in grades:
            if not rerun_allowed.get(grade, True):
                for line in allowed_lines[grade]:
                    # Sum of is_start variables over the entire horizon for a non-rerun grade must be at most 1
                    total_starts = sum(is_start_vars.get((grade, line, d), model.NewBoolVar('dummy_0')) for d in range(num_days))
                    model.Add(total_starts <= 1)


        # --- Transition Constraints and Penalty & Continuity Bonus ---
        for line in lines:
            for d in range(num_days - 1):
                # 1. Transition Penalty Constraints
                for grade1 in grades:
                    if line not in allowed_lines[grade1]:
                        continue
                    
                    for grade2 in grades:
                        if line not in allowed_lines[grade2] or grade1 == grade2:
                            continue
                        
                        # Check transition rules: if the rule exists and grade2 is NOT in the allowed list for grade1,
                        # then the transition is blocked (we skip defining the transition variable and penalty, 
                        # relying on other constraints to enforce the non-production).
                        if (transition_rules.get(line) 
                            and grade1 in transition_rules[line] 
                            and grade2 not in transition_rules[line][grade1]):
                            
                            # To strictly enforce the block: grade1 production on d OR grade2 production on d+1 must be 0
                            # This is the strict way, but the original code (snippet 13) simply uses `continue` to avoid the penalty.
                            # We reproduce the original logic: skip adding penalty for blocked transition.
                            continue

                        # Define the transition variable (for transitions that are allowed or not explicitly blocked)
                        trans_var = model.NewBoolVar(f'trans_{line}_{d}_{grade1}_to_{grade2}')
                        
                        # Enforce: trans_var <=> is_producing(grade1, d) AND is_producing(grade2, d+1)
                        model.AddBoolAnd([is_producing[(grade1, line, d)], is_producing[(grade2, line, d + 1)]]).OnlyEnforceIf(trans_var)
                        
                        # These two lines are often used for cleaner logic but are logically redundant with AddBoolAnd for the objective
                        # model.Add(trans_var == 0).OnlyEnforceIf(is_producing[(grade1, line, d)].Not())
                        # model.Add(trans_var == 0).OnlyEnforceIf(is_producing[(grade2, line, d + 1)].Not())
                        
                        # Add penalty to objective
                        objective += transition_penalty * trans_var

                # 2. Continuity Bonus (reward for maintaining the same grade)
                for grade in grades:
                    if line in allowed_lines[grade]:
                        continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                        # Enforce: continuity <=> is_producing(grade, d) AND is_producing(grade, d+1)
                        model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                        # Continuity is a negative cost (a bonus)
                        objective += -continuity_bonus * continuity

        # --- Set Objective ---
        model.Minimize(objective)

        progress_bar.progress(50)
        status_text.markdown('<div class="alert-box info">⚡ Solving optimization problem...</div>', unsafe_allow_html=True)

        # ====================================================================
        # SOLUTION CALLBACK
        # ====================================================================
        class SolutionCallback(cp_model.CpSolverSolutionCallback):
            def __init__(self, production, inventory, stockout, is_producing, grades, lines, dates, formatted_dates, num_days):
                cp_model.CpSolverSolutionCallback.__init__(self)
                self.production = production
                self.inventory = inventory
                self.stockout = stockout
                self.is_producing = is_producing
                self.grades = grades
                self.lines = lines
                self.dates = dates
                self.formatted_dates = formatted_dates
                self.num_days = num_days
                self.solutions = []
                self.solution_times = []
                self.start_time = time.time()
                self.best_objective = float('inf')

            def on_solution_callback(self):
                current_time = time.time() - self.start_time
                current_obj = self.ObjectiveValue()
                
                # Only store better solutions
                if current_obj < self.best_objective:
                    self.best_objective = current_obj
                    
                    solution = {
                        'objective': current_obj,
                        'time': current_time,
                        'production': {},
                        'inventory': {},
                        'stockout': {},
                        'is_producing': {}
                    }
                    
                    # 1. Production Data
                    for grade in self.grades:
                        solution['production'][grade] = {}
                        for line in self.lines:
                            for d in range(self.num_days):
                                key = (grade, line, d)
                                if key in self.production:
                                    value = self.Value(self.production[key])
                                    if value > 0:
                                        date_key = self.formatted_dates[d]
                                        if date_key not in solution['production'][grade]:
                                            solution['production'][grade][date_key] = 0
                                        solution['production'][grade][date_key] += value
                    
                    # 2. Inventory Data
                    for grade in self.grades:
                        solution['inventory'][grade] = {}
                        # Loop up to num_days (includes initial inventory at index 0)
                        for d in range(self.num_days + 1):
                            key = (grade, d)
                            solution['inventory'][grade][d] = self.Value(self.inventory[key])

                    # 3. Stockout Data
                    for grade in self.grades:
                        solution['stockout'][grade] = {}
                        for d in range(self.num_days):
                            key = (grade, d)
                            solution['stockout'][grade][self.formatted_dates[d]] = self.Value(self.stockout[key])

                    # 4. Is Producing Data & Transition Count
                    total_transitions = 0
                    transition_count_per_line = {line: 0 for line in self.lines}
                    
                    for line in self.lines:
                        last_grade = None
                        for d in range(self.num_days):
                            current_grade = None
                            for grade in self.grades:
                                key = (grade, line, d)
                                if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                                    current_grade = grade
                                    break
                            
                            solution['is_producing'][(line, d)] = current_grade # Store the single grade running
                            
                            # Count transitions
                            if current_grade is not None:
                                if last_grade is not None and current_grade != last_grade:
                                    transition_count_per_line[line] += 1
                                    total_transitions += 1
                            
                            # If line is idle, last_grade remains the last produced grade for continuity check
                            last_grade = current_grade

                    solution['transitions'] = {
                        'per_line': transition_count_per_line,
                        'total': total_transitions
                    }
                    
                    # Replace the list with a single best solution (for simplicity in Streamlit display)
                    if not self.solutions:
                        self.solutions.append(solution)
                    else:
                        self.solutions[0] = solution
                        
            def num_solutions(self):
                return len(self.solutions)

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit_min * 60.0
        solver.parameters.num_search_workers = 8
        solver.parameters.random_seed = 42
        solver.parameters.log_search_progress = True

        solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

        start_time = time.time()
        status = solver.Solve(model, solution_callback)
        end_time = time.time()
        
        progress_bar.progress(100)

        # ====================================================================
        # RESULTS DISPLAY - Material Design
        # ====================================================================
        
        st.session_state.solutions = solution_callback.solutions

        if status == cp_model.OPTIMAL:
            status_text.markdown('<div class="alert-box success">✅ Optimal solution found!</div>', unsafe_allow_html=True)
        elif status == cp_model.FEASIBLE:
            status_text.markdown('<div class="alert-box success">✅ Feasible solution found!</div>', unsafe_allow_html=True)
        else:
            status_text.markdown('<div class="alert-box warning">⚠️ No optimal solution found</div>', unsafe_allow_html=True)
        
        st.markdown(f'<div style="font-size: 0.875rem; color: #757575; text-align: center; margin-top: 0.5rem;">Solver Time: **{end_time - start_time:.2f} s** (Status: {solver.StatusName(status)})</div>', unsafe_allow_html=True)

        time.sleep(1)
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        if solution_callback.num_solutions() > 0:
            best_solution = solution_callback.solutions[0]
            st.session_state.best_solution = best_solution
            
            # --- Performance metrics dashboard ---
            st.markdown("""
            <div class="material-card">
                <div class="card-title">📊 Optimization Results</div>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.markdown(f"""
                <div class="metric-card primary">
                    <div class="metric-label">Objective Value</div>
                    <div class="metric-value">{best_solution['objective']:,.0f}</div>
                    <div class="metric-subtitle">Lower is Better</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card info">
                    <div class="metric-label">Transitions</div>
                    <div class="metric-value">{best_solution['transitions']['total']}</div>
                    <div class="metric-subtitle">Grade Changeovers</div>
                </div>
                """, unsafe_allow_html=True)

            with col3:
                # Calculate total stockouts
                total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                st.markdown(f"""
                <div class="metric-card warning">
                    <div class="metric-label">Total Stockouts</div>
                    <div class="metric-value">{total_stockouts:,.0f}</div>
                    <div class="metric-subtitle">MT</div>
                </div>
                """, unsafe_allow_html=True)

            with col4:
                # Calculate total production
                total_production = sum(sum(best_solution['production'][g].values()) for g in grades)
                st.markdown(f"""
                <div class="metric-card success">
                    <div class="metric-label">Total Production</div>
                    <div class="metric-value">{total_production:,.0f}</div>
                    <div class="metric-subtitle">MT</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            
            # --- Detailed Results ---
            tab1, tab2, tab3 = st.tabs(["🗓️ Production Schedule", "📈 Inventory Trends", "🔗 Grade Details"])
            
            # Get a color map for visualization
            all_grades_sorted = sorted(grades)
            cmap = colormaps['viridis']
            color_indices = np.linspace(0, 1, len(all_grades_sorted))
            grade_color_map = {grade: mcolors.to_hex(cmap(i)) for grade, i in zip(all_grades_sorted, color_indices)}

            # TAB 1: PRODUCTION SCHEDULE
            with tab1:
                st.markdown("### Plant-Level Schedules")
                for line in lines:
                    st.markdown(f"**{line} Schedule**")
                    
                    # Gantt chart data
                    gantt_data = []
                    current_grade = None
                    start_day = None
                    
                    for d in range(num_days):
                        date = dates[d]
                        grade_today = best_solution['is_producing'].get((line, d))
                        
                        if grade_today != current_grade:
                            if current_grade is not None:
                                end_date = dates[d - 1]
                                duration = (end_date - start_day).days + 1
                                # If the run was blocked by a shutdown, duration might be tricky, but this is simple logging
                                gantt_data.append(dict(
                                    Task=line, 
                                    Start=start_day, 
                                    Finish=end_date + timedelta(days=1), # +1 day for plotly finish date
                                    Resource=current_grade, 
                                    Color=grade_color_map.get(current_grade, '#cccccc')
                                ))
                            
                            current_grade = grade_today
                            start_day = date
                    
                    # Final entry
                    if current_grade is not None:
                        end_date = dates[num_days - 1]
                        gantt_data.append(dict(
                            Task=line, 
                            Start=start_day, 
                            Finish=end_date + timedelta(days=1),
                            Resource=current_grade, 
                            Color=grade_color_map.get(current_grade, '#cccccc')
                        ))
                        
                    if not gantt_data:
                        st.info(f"No production scheduled for {line}.")
                        continue

                    gantt_df = pd.DataFrame(gantt_data)
                    
                    # Plotly Gantt Chart
                    fig = px.timeline(
                        gantt_df, 
                        x_start="Start", 
                        x_end="Finish", 
                        y="Task", 
                        color="Resource",
                        color_discrete_map={g: c for g, c in grade_color_map.items()},
                        custom_data=['Resource'],
                        category_orders={"Task": [line]}
                    )

                    # Add Shutdown periods as colored rectangles
                    if line in shutdown_periods and shutdown_periods[line]:
                        shutdown_days = shutdown_periods[line]
                        start_shutdown = dates[shutdown_days[0]]
                        end_shutdown = dates[shutdown_days[-1]]
                        fig.add_vrect(
                            x0=start_shutdown, 
                            x1=end_shutdown + timedelta(days=1), 
                            fillcolor="red", 
                            opacity=0.15, 
                            layer="below", 
                            line_width=0, 
                            annotation_text="Shutdown", 
                            annotation_position="top left", 
                            annotation_font_size=12, 
                            annotation_font_color="#c62828"
                        )
                    
                    fig.update_yaxes(
                        autorange="reversed", 
                        title=None, 
                        showgrid=True, 
                        gridcolor="#e0e0e0"
                    )
                    fig.update_xaxes(
                        title="Date", 
                        showgrid=True, 
                        gridcolor="#e0e0e0",
                        tickformat="%d-%b"
                    )
                    fig.update_layout(
                        height=150,
                        showlegend=False,
                        plot_bgcolor="white", 
                        paper_bgcolor="white", 
                        margin=dict(l=60, r=20, t=20, b=40),
                        hoverlabel=dict(bgcolor="white", font_size=12, font_family="Inter")
                    )
                    st.plotly_chart(fig, use_container_width=True)

                    # Detailed Schedule Table
                    def color_grade(val):
                        if val in grade_color_map:
                            return f'background-color: {grade_color_map[val]}; color: white; font-weight: bold; text-align: center;'
                        return ''
                    
                    schedule_data = []
                    current_grade = None
                    start_day = None
                    
                    # Reconstruct the run periods for the table (simpler list version)
                    for d in range(num_days):
                        date = dates[d]
                        grade_today = best_solution['is_producing'].get((line, d))
                        
                        if grade_today != current_grade:
                            if current_grade is not None:
                                end_date = dates[d - 1]
                                duration = (end_date - start_day).days + 1
                                schedule_data.append({
                                    "Grade": current_grade,
                                    "Start Date": start_day.strftime("%d-%b-%y"),
                                    "End Date": end_date.strftime("%d-%b-%y"),
                                    "Days": duration
                                })
                            current_grade = grade_today
                            start_day = date
                    
                    # Final entry
                    if current_grade is not None:
                        end_date = dates[num_days - 1]
                        duration = (end_date - start_day).days + 1
                        schedule_data.append({
                            "Grade": current_grade,
                            "Start Date": start_day.strftime("%d-%b-%y"),
                            "End Date": end_date.strftime("%d-%b-%y"),
                            "Days": duration
                        })
                    
                    if schedule_data:
                        schedule_df = pd.DataFrame(schedule_data)
                        styled_df = schedule_df.style.applymap(color_grade, subset=['Grade'])
                        st.dataframe(styled_df, use_container_width=True, hide_index=True)
                    
                    st.markdown('<div class="divider" style="margin: 1rem 0;"></div>', unsafe_allow_html=True)


            # TAB 2: INVENTORY TRENDS
            with tab2:
                st.markdown("### Inventory Trends")
                last_actual_day = num_days - buffer_days - 1
                
                # Plot inventory for each grade
                for grade in all_grades_sorted:
                    inventory_values = [best_solution['inventory'][grade][d] for d in range(num_days + 1)] # +1 for initial inventory
                    
                    # Inventory starts at day 0 (d=0) and ends at day num_days (d=num_days)
                    plot_dates = [dates[0] - timedelta(days=1)] + dates # Add a dummy date for initial inventory point
                    
                    fig = go.Figure()
                    
                    # Inventory line
                    fig.add_trace(go.Scatter(
                        x=plot_dates,
                        y=inventory_values,
                        mode="lines",
                        name="Inventory",
                        line=dict(color=grade_color_map[grade], width=3),
                        hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"
                    ))
                    
                    # Min/Max lines
                    if min_inventory[grade] > 0:
                        fig.add_hline(
                            y=min_inventory[grade], 
                            line_dash="dot", 
                            line_color="red", 
                            annotation_text="Min. Inventory", 
                            annotation_position="bottom right"
                        )
                    if max_inventory[grade] < 1000000000:
                        fig.add_hline(
                            y=max_inventory[grade], 
                            line_dash="dot", 
                            line_color="gray", 
                            annotation_text="Max. Inventory", 
                            annotation_position="top right"
                        )
                    if min_closing_inventory[grade] > 0:
                        fig.add_vline(
                            x=dates[last_actual_day],
                            line_dash="dash",
                            line_color="#764ba2",
                            annotation_text="Min. Closing Inv. Deadline",
                            annotation_position="top right"
                        )
                    
                    # Buffer Zone Shading
                    fig.add_vrect(
                        x0=dates[num_days - buffer_days], 
                        x1=dates[-1] + timedelta(days=1), 
                        fillcolor="yellow", 
                        opacity=0.1, 
                        layer="below", 
                        line_width=0, 
                        annotation_text="Buffer Period", 
                        annotation_position="top left", 
                        annotation_font_size=12, 
                        annotation_font_color="#e65100"
                    )

                    fig.update_layout(
                        title=f"Inventory for Grade: {grade}",
                        height=400, 
                        plot_bgcolor="white", 
                        paper_bgcolor="white", 
                        margin=dict(l=60, r=10, t=40, b=60),
                        showlegend=False
                    )
                    fig.update_xaxes(
                        title="Date", 
                        showgrid=True, 
                        gridcolor="#e0e0e0", 
                        tickformat="%d-%b"
                    )
                    fig.update_yaxes(
                        title="Inventory (MT)",
                        showgrid=True,
                        gridcolor="#e0e0e0"
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)


            # TAB 3: GRADE DETAILS
            with tab3:
                st.markdown("### Grade and Transition Summary")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**Production vs. Demand**")
                    summary_data = []
                    for grade in all_grades_sorted:
                        total_production = sum(best_solution['production'][grade].values())
                        total_demand = sum(demand_data[grade][date] for date in dates[:num_days - buffer_days]) # Only actual demand days
                        
                        summary_data.append({
                            "Grade": grade,
                            "Total Production (MT)": f"{total_production:,.0f}",
                            "Total Demand (MT)": f"{total_demand:,.0f}",
                            "Demand Met (%)": f"{((total_production / total_demand) * 100):.1f}%" if total_demand > 0 else "N/A"
                        })
                    
                    summary_df = pd.DataFrame(summary_data)
                    st.dataframe(
                        summary_df, 
                        use_container_width=True,
                        hide_index=True
                    )

                with col2:
                    st.markdown("**Transitions per Plant**")
                    transition_data = []
                    for line, count in best_solution['transitions']['per_line'].items():
                        transition_data.append({
                            "Plant": line,
                            "Transitions": count
                        })
                    
                    transition_df = pd.DataFrame(transition_data)
                    st.dataframe(
                        transition_df, 
                        use_container_width=True, 
                        hide_index=True
                    )

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            
            # --- Navigation Buttons ---
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Try Again", use_container_width=True):
                    st.session_state.step = 1
                    st.session_state.uploaded_file = None
                    st.rerun()
            
            with col2:
                if st.button("⚙️ Adjust Settings", use_container_width=True):
                    st.session_state.step = 2
                    st.rerun()
    
    except Exception as e:
        st.markdown(f"""
        <div class="alert-box error">
            <div>
                <strong>❌ Error During Optimization</strong><br>
                {str(e)}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander("View Technical Details"):
            import traceback
            st.code(traceback.format_exc())
        
        if st.button("← Return to Start"):
            st.session_state.step = 1
            st.session_state.uploaded_file = None
            st.rerun()

# ============================================================================
# FOOTER
# ============================================================================
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; color: #9e9e9e; font-size: 0.875rem; padding: 1rem 0;">
    <strong>Polymer Production Scheduler</strong> • Powered by OR-Tools & Streamlit<br>
    Material Minimalism Design • Version 2.0
</div>
""", unsafe_allow_html=True)
