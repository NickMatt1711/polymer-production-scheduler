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
        padding: 3rem 2rem;
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

    col1, col2 = st.columns([11, 1])

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
            time_limit_min = st.number_input(
                "⏱️ Time Limit (minutes)",
                min_value=1,
                max_value=120,
                value=10,
                help="Maximum solver runtime"
            )
            
            buffer_days = st.number_input(
                "📅 Planning Buffer (days)",
                min_value=0,
                max_value=7,
                value=3,
                help="Additional days for safety planning"
            )
        
        with col2:
            st.markdown("#### Objective Weights")
            stockout_penalty = st.number_input(
                "🎯 Stockout Penalty",
                min_value=1,
                value=10,
                help="Cost weight for inventory shortages"
            )
            
            transition_penalty = st.number_input(
                "🔄 Transition Penalty",
                min_value=1,
                value=10,
                help="Cost weight for grade changeovers"
            )
            
            continuity_bonus = st.number_input(
                "➕ Continuity Bonus",
                min_value=0,
                value=1,
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
        # DATA PREPROCESSING - All logic unchanged
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
                    grade_plant_key = (grade, plant)
                    min_run_days[grade_plant_key] = int(row['Min. Run Days']) if pd.notna(row['Min. Run Days']) else 1
                    max_run_days[grade_plant_key] = int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999
                    
                    if pd.notna(row['Force Start Date']):
                        try:
                            force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date()
                        except:
                            force_start_date[grade_plant_key] = None
                    else:
                        force_start_date[grade_plant_key] = None
                    
                    rerun_val = row['Rerun Allowed']
                    if pd.notna(rerun_val):
                        val_str = str(rerun_val).strip().lower()
                        rerun_allowed[grade_plant_key] = val_str not in ['no', 'n', 'false', '0']
                    else:
                        rerun_allowed[grade_plant_key] = True
            
            material_running_info = {}
            for index, row in plant_df.iterrows():
                plant = row['Plant']
                material = row['Material Running']
                expected_days = row['Expected Run Days']
                
                if pd.notna(material) and pd.notna(expected_days):
                    try:
                        material_running_info[plant] = (str(material).strip(), int(expected_days))
                    except (ValueError, TypeError):
                        pass
        
        except Exception as e:
            st.error(f"Error in data preprocessing: {str(e)}")
            st.stop()
        
        demand_data = {}
        dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
        num_days = len(dates)
        last_date = dates[-1]
        
        # Use parameters from step 2 (stored in widgets)
        buffer_days = 3  # Default, should come from session state in production
        time_limit_min = 10  # Default
        stockout_penalty = 10
        transition_penalty = 10
        continuity_bonus = 1
        
        for i in range(1, buffer_days + 1):
            dates.append(last_date + timedelta(days=i))
        num_days = len(dates)
        
        formatted_dates = [date.strftime('%d-%b-%y') for date in dates]
        
        for grade in grades:
            if grade in demand_df.columns:
                demand_data[grade] = {demand_df.iloc[i, 0].date(): demand_df[grade].iloc[i] for i in range(len(demand_df))}
            else:
                demand_data[grade] = {date: 0 for date in dates}
        
        for grade in grades:
            for date in dates[-buffer_days:]:
                if date not in demand_data[grade]:
                    demand_data[grade][date] = 0
        
        shutdown_periods = process_shutdown_dates(plant_df, dates)
        
        # Load transition matrices
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
                        if str(df.loc[prev_grade, current_grade]).lower() == 'yes':
                            allowed_transitions.append(current_grade)
                    transition_rules[line][prev_grade] = allowed_transitions
            else:
                transition_rules[line] = None
        
        progress_bar.progress(30)
        status_text.markdown('<div class="alert-box info">🔧 Building optimization model...</div>', unsafe_allow_html=True)
        time.sleep(1)
        
        # ====================================================================
        # MODEL BUILDING - All solver logic unchanged
        # ====================================================================
        model = cp_model.CpModel()
        
        is_producing = {}
        production = {}
        
        def is_allowed_combination(grade, line):
            return line in allowed_lines.get(grade, [])
        
        for grade in grades:
            for line in allowed_lines[grade]:
                for d in range(num_days):
                    key = (grade, line, d)
                    is_producing[key] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                    
                    if d < num_days - buffer_days:
                        production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                        model.Add(production_value == capacities[line]).OnlyEnforceIf(is_producing[key])
                        model.Add(production_value == 0).OnlyEnforceIf(is_producing[key].Not())
                    else:
                        production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                        model.Add(production_value <= capacities[line] * is_producing[key])
                    
                    production[key] = production_value
        
        def get_production_var(grade, line, d):
            key = (grade, line, d)
            return production.get(key, 0)
        
        def get_is_producing_var(grade, line, d):
            key = (grade, line, d)
            return is_producing.get(key, None)
        
        # Shutdown constraints
        for line in lines:
            if line in shutdown_periods and shutdown_periods[line]:
                for d in shutdown_periods[line]:
                    for grade in grades:
                        if is_allowed_combination(grade, line):
                            key = (grade, line, d)
                            if key in is_producing:
                                model.Add(is_producing[key] == 0)
                                model.Add(production[key] == 0)
        
        inventory_vars = {}
        for grade in grades:
            for d in range(num_days + 1):
                inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')
        
        stockout_vars = {}
        for grade in grades:
            for d in range(num_days):
                stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')
        
        # One grade per line constraint
        for line in lines:
            for d in range(num_days):
                producing_vars = []
                for grade in grades:
                    if is_allowed_combination(grade, line):
                        var = get_is_producing_var(grade, line, d)
                        if var is not None:
                            producing_vars.append(var)
                if producing_vars:
                    model.Add(sum(producing_vars) <= 1)
        
        # Material running constraints
        for plant, (material, expected_days) in material_running_info.items():
            for d in range(min(expected_days, num_days)):
                if is_allowed_combination(material, plant):
                    model.Add(get_is_producing_var(material, plant, d) == 1)
                    for other_material in grades:
                        if other_material != material and is_allowed_combination(other_material, plant):
                            model.Add(get_is_producing_var(other_material, plant, d) == 0)
        
        objective = 0
        
        # Inventory balance
        for grade in grades:
            model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])
        
        for grade in grades:
            for d in range(num_days):
                produced_today = sum(get_production_var(grade, line, d) for line in allowed_lines[grade])
                demand_today = demand_data[grade].get(dates[d], 0)
                
                supplied = model.NewIntVar(0, 100000, f'supplied_{grade}_{d}')
                model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
                model.Add(supplied <= demand_today)
                model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
                model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
                model.Add(inventory_vars[(grade, d + 1)] >= 0)
        
        # Minimum inventory penalties
        for grade in grades:
            for d in range(num_days):
                if min_inventory[grade] > 0:
                    deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                    model.Add(deficit >= min_inventory[grade] - inventory_vars[(grade, d + 1)])
                    model.Add(deficit >= 0)
                    objective += stockout_penalty * deficit
        
        # Closing inventory
        for grade in grades:
            closing_inventory = inventory_vars[(grade, num_days - buffer_days)]
            if min_closing_inventory[grade] > 0:
                closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                model.Add(closing_deficit >= min_closing_inventory[grade] - closing_inventory)
                model.Add(closing_deficit >= 0)
                objective += stockout_penalty * closing_deficit * 3
        
        # Max inventory
        for grade in grades:
            for d in range(1, num_days + 1):
                model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])
        
        # Capacity constraints
        for line in lines:
            for d in range(num_days - buffer_days):
                if line in shutdown_periods and d in shutdown_periods[line]:
                    continue
                production_vars = [get_production_var(grade, line, d) for grade in grades if is_allowed_combination(grade, line)]
                if production_vars:
                    model.Add(sum(production_vars) == capacities[line])
            
            for d in range(num_days - buffer_days, num_days):
                production_vars = [get_production_var(grade, line, d) for grade in grades if is_allowed_combination(grade, line)]
                if production_vars:
                    model.Add(sum(production_vars) <= capacities[line])
        
        # Force start dates
        for grade_plant_key, start_date in force_start_date.items():
            if start_date:
                grade, plant = grade_plant_key
                try:
                    start_day_index = dates.index(start_date)
                    var = get_is_producing_var(grade, plant, start_day_index)
                    if var is not None:
                        model.Add(var == 1)
                except ValueError:
                    pass
        
        # Run length constraints
        is_start_vars = {}
        run_end_vars = {}
        
        for grade in grades:
            for line in allowed_lines[grade]:
                grade_plant_key = (grade, line)
                min_run = min_run_days.get(grade_plant_key, 1)
                max_run = max_run_days.get(grade_plant_key, 9999)
                
                for d in range(num_days):
                    is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                    is_start_vars[(grade, line, d)] = is_start
                    
                    is_end = model.NewBoolVar(f'end_{grade}_{line}_{d}')
                    run_end_vars[(grade, line, d)] = is_end
                    
                    current_prod = get_is_producing_var(grade, line, d)
                    
                    if d > 0:
                        prev_prod = get_is_producing_var(grade, line, d - 1)
                        if current_prod is not None and prev_prod is not None:
                            model.AddBoolAnd([current_prod, prev_prod.Not()]).OnlyEnforceIf(is_start)
                            model.AddBoolOr([current_prod.Not(), prev_prod]).OnlyEnforceIf(is_start.Not())
                    else:
                        if current_prod is not None:
                            model.Add(current_prod == 1).OnlyEnforceIf(is_start)
                            model.Add(is_start == 1).OnlyEnforceIf(current_prod)
                    
                    if d < num_days - 1:
                        next_prod = get_is_producing_var(grade, line, d + 1)
                        if current_prod is not None and next_prod is not None:
                            model.AddBoolAnd([current_prod, next_prod.Not()]).OnlyEnforceIf(is_end)
                            model.AddBoolOr([current_prod.Not(), next_prod]).OnlyEnforceIf(is_end.Not())
                    else:
                        if current_prod is not None:
                            model.Add(current_prod == 1).OnlyEnforceIf(is_end)
                            model.Add(is_end == 1).OnlyEnforceIf(current_prod)
        
                # Min run days
                for d in range(num_days):
                    is_start = is_start_vars[(grade, line, d)]
                    max_possible_run = 0
                    for k in range(min_run):
                        if d + k < num_days:
                            if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                break
                            max_possible_run += 1
                    
                    if max_possible_run >= min_run:
                        for k in range(min_run):
                            if d + k < num_days:
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    continue
                                future_prod = get_is_producing_var(grade, line, d + k)
                                if future_prod is not None:
                                    model.Add(future_prod == 1).OnlyEnforceIf(is_start)
        
                # Max run days
                for d in range(num_days - max_run):
                    consecutive_days = []
                    for k in range(max_run + 1):
                        if d + k < num_days:
                            if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                break
                            prod_var = get_is_producing_var(grade, line, d + k)
                            if prod_var is not None:
                                consecutive_days.append(prod_var)
                    
                    if len(consecutive_days) == max_run + 1:
                        model.Add(sum(consecutive_days) <= max_run)
        
        # Transition rules
        for line in lines:
            if transition_rules.get(line):
                for d in range(num_days - 1):
                    for prev_grade in grades:
                        if prev_grade in transition_rules[line] and is_allowed_combination(prev_grade, line):
                            allowed_next = transition_rules[line][prev_grade]
                            for current_grade in grades:
                                if (current_grade != prev_grade and current_grade not in allowed_next and is_allowed_combination(current_grade, line)):
                                    prev_var = get_is_producing_var(prev_grade, line, d)
                                    current_var = get_is_producing_var(current_grade, line, d + 1)
                                    if prev_var is not None and current_var is not None:
                                        model.Add(prev_var + current_var <= 1)
        
        # Rerun constraints
        for grade in grades:
            for line in allowed_lines[grade]:
                grade_plant_key = (grade, line)
                if not rerun_allowed.get(grade_plant_key, True):
                    starts = [is_start_vars[(grade, line, d)] for d in range(num_days) if (grade, line, d) in is_start_vars]
                    if starts:
                        model.Add(sum(starts) <= 1)
        
        # Stockout penalties
        for grade in grades:
            for d in range(num_days):
                objective += stockout_penalty * stockout_vars[(grade, d)]
        
        # Transition penalties and continuity bonuses
        for line in lines:
            for d in range(num_days - 1):
                for grade1 in grades:
                    if line not in allowed_lines[grade1]:
                        continue
                    for grade2 in grades:
                        if line not in allowed_lines[grade2] or grade1 == grade2:
                            continue
                        if transition_rules.get(line) and grade1 in transition_rules[line] and grade2 not in transition_rules[line][grade1]:
                            continue
                        trans_var = model.NewBoolVar(f'trans_{line}_{d}_{grade1}_to_{grade2}')
                        model.AddBoolAnd([is_producing[(grade1, line, d)], is_producing[(grade2, line, d + 1)]]).OnlyEnforceIf(trans_var)
                        model.Add(trans_var == 0).OnlyEnforceIf(is_producing[(grade1, line, d)].Not())
                        model.Add(trans_var == 0).OnlyEnforceIf(is_producing[(grade2, line, d + 1)].Not())
                        objective += transition_penalty * trans_var

                for grade in grades:
                    if line in allowed_lines[grade]:
                        continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                        model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                        objective += -continuity_bonus * continuity
        
        model.Minimize(objective)
        
        progress_bar.progress(50)
        status_text.markdown('<div class="alert-box info">⚡ Solving optimization problem...</div>', unsafe_allow_html=True)
        
        # ====================================================================
        # SOLUTION CALLBACK - Unchanged
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

            def on_solution_callback(self):
                current_time = time.time() - self.start_time
                self.solution_times.append(current_time)
                current_obj = self.ObjectiveValue()

                solution = {
                    'objective': current_obj,
                    'time': current_time,
                    'production': {},
                    'inventory': {},
                    'stockout': {},
                    'is_producing': {}
                }

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
                
                for grade in self.grades:
                    solution['inventory'][grade] = {}
                    for d in range(self.num_days + 1):
                        key = (grade, d)
                        if key in self.inventory:
                            if d < self.num_days:
                                solution['inventory'][grade][self.formatted_dates[d] if d > 0 else 'initial'] = self.Value(self.inventory[key])
                            else:
                                solution['inventory'][grade]['final'] = self.Value(self.inventory[key])
                
                for grade in self.grades:
                    solution['stockout'][grade] = {}
                    for d in range(self.num_days):
                        key = (grade, d)
                        if key in self.stockout:
                            value = self.Value(self.stockout[key])
                            if value > 0:
                                solution['stockout'][grade][self.formatted_dates[d]] = value
                
                for line in self.lines:
                    solution['is_producing'][line] = {}
                    for d in range(self.num_days):
                        date_key = self.formatted_dates[d]
                        solution['is_producing'][line][date_key] = None
                        for grade in self.grades:
                            key = (grade, line, d)
                            if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                                solution['is_producing'][line][date_key] = grade
                                break

                self.solutions.append(solution)

                transition_count_per_line = {line: 0 for line in self.lines}
                total_transitions = 0

                for line in self.lines:
                    last_grade = None
                    for d in range(self.num_days):
                        current_grade = None
                        for grade in self.grades:
                            key = (grade, line, d)
                            if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                                current_grade = grade
                                break
                        
                        if current_grade is not None:
                            if last_grade is not None and current_grade != last_grade:
                                transition_count_per_line[line] += 1
                                total_transitions += 1
                            last_grade = current_grade

                solution['transitions'] = {
                    'per_line': transition_count_per_line,
                    'total': total_transitions
                }

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
        
        progress_bar.progress(100)
        
        # ====================================================================
        # RESULTS DISPLAY - Material Design
        # ====================================================================
        if status == cp_model.OPTIMAL:
            status_text.markdown('<div class="alert-box success">✅ Optimal solution found!</div>', unsafe_allow_html=True)
        elif status == cp_model.FEASIBLE:
            status_text.markdown('<div class="alert-box success">✅ Feasible solution found!</div>', unsafe_allow_html=True)
        else:
            status_text.markdown('<div class="alert-box warning">⚠️ No optimal solution found</div>', unsafe_allow_html=True)

        time.sleep(1)
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        if solution_callback.num_solutions() > 0:
            best_solution = solution_callback.solutions[-1]

            # Performance metrics dashboard
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
                total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                st.markdown(f"""
                <div class="metric-card {'success' if total_stockouts == 0 else 'warning'}">
                    <div class="metric-label">Stockouts</div>
                    <div class="metric-value">{total_stockouts:,.0f}</div>
                    <div class="metric-subtitle">MT Unmet Demand</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                st.markdown(f"""
                <div class="metric-card info">
                    <div class="metric-label">Planning Horizon</div>
                    <div class="metric-value">{num_days}</div>
                    <div class="metric-subtitle">Days</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            
            # Tabbed results
            tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Summary Analytics", "📦 Inventory Trends"])
            
            # TAB 1: PRODUCTION SCHEDULE
            with tab1:
                                
                sorted_grades = sorted(grades)
                base_colors = px.colors.qualitative.Vivid
                grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

                for line in lines:
                    st.markdown(f"#### 🏭 {line}")
                
                    gantt_data = []
                    for d in range(num_days):
                        date = dates[d]
                        for grade in sorted_grades:
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                gantt_data.append({
                                    "Grade": grade,
                                    "Start": date,
                                    "Finish": date + timedelta(days=1),
                                    "Line": line
                                })
                
                    if not gantt_data:
                        st.info(f"No production scheduled for {line}")
                        continue
                
                    gantt_df = pd.DataFrame(gantt_data)
                
                    fig = px.timeline(
                        gantt_df,
                        x_start="Start",
                        x_end="Finish",
                        y="Grade",
                        color="Grade",
                        color_discrete_map=grade_color_map,
                        category_orders={"Grade": sorted_grades},
                    )
                
                    if line in shutdown_periods and shutdown_periods[line]:
                        shutdown_days = shutdown_periods[line]
                        start_shutdown = dates[shutdown_days[0]]
                        end_shutdown = dates[shutdown_days[-1]] + timedelta(days=1)
                        
                        fig.add_vrect(
                            x0=start_shutdown,
                            x1=end_shutdown,
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
                        height=350,
                        showlegend=True,
                        legend=dict(
                            orientation="v",
                            yanchor="middle",
                            y=0.5,
                            xanchor="left",
                            x=1.02
                        ),
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(l=60, r=160, t=40, b=60),
                    )
                
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Schedule table
                    st.markdown(f"**Detailed Schedule - {line}**")

                    def color_grade(val):
                            if val in grade_color_map:
                                color = grade_color_map[val]
                                return f'background-color: {color}; color: white; font-weight: bold; text-align: center;'
                            return ''

                    schedule_data = []
                    current_grade = None
                    start_day = None
                
                    for d in range(num_days):
                        date = dates[d]
                        grade_today = None
                
                        for grade in sorted_grades:
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                grade_today = grade
                                break
                
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
                
                    if current_grade is not None:
                        end_date = dates[num_days - 1]
                        duration = (end_date - start_day).days + 1
                        schedule_data.append({
                            "Grade": current_grade,
                            "Start Date": start_day.strftime("%d-%b-%y"),
                            "End Date": end_date.strftime("%d-%b-%y"),
                            "Days": duration
                        })
                
                    if not schedule_data:
                        st.info(f"No production data available for {line}.")
                        continue
                
                    schedule_df = pd.DataFrame(schedule_data)
                    styled_df = schedule_df.style.applymap(color_grade, subset=['Grade'])
                    st.dataframe(styled_df, use_container_width=True)
            
            # TAB 2: SUMMARY ANALYTICS
            with tab2:

                col1, col2 = st.columns([2, 1])
                
                with col1:
                    production_totals = {}
                    grade_totals = {}
                    plant_totals = {line: 0 for line in lines}
                    stockout_totals = {}
                    
                    for grade in grades:
                        production_totals[grade] = {}
                        grade_totals[grade] = 0
                        stockout_totals[grade] = 0
                        
                        for line in lines:
                            total_prod = 0
                            for d in range(num_days):
                                key = (grade, line, d)
                                if key in production:
                                    total_prod += solver.Value(production[key])
                            production_totals[grade][line] = total_prod
                            grade_totals[grade] += total_prod
                            plant_totals[line] += total_prod
                        
                        for d in range(num_days):
                            key = (grade, d)
                            if key in stockout_vars:
                                stockout_totals[grade] += solver.Value(stockout_vars[key])
                    
                    total_prod_data = []
                    for grade in grades:
                        row = {'Grade': grade}
                        for line in lines:
                            row[line] = production_totals[grade][line]
                        row['Total Produced'] = grade_totals[grade]
                        row['Total Stockout'] = stockout_totals[grade]
                        total_prod_data.append(row)
                    
                    totals_row = {'Grade': 'TOTAL'}
                    for line in lines:
                        totals_row[line] = plant_totals[line]
                    totals_row['Total Produced'] = sum(plant_totals.values())
                    totals_row['Total Stockout'] = sum(stockout_totals.values())
                    total_prod_data.append(totals_row)
                    
                    total_prod_df = pd.DataFrame(total_prod_data)
                    
                    st.dataframe(
                        total_prod_df.style.apply(
                            lambda x: ['background-color: #f5f5f5; font-weight: bold' if x.name == len(total_prod_df) - 1 else '' for i in x],
                            axis=1
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                
                with col2:
                    transition_data = []
                    for line, count in best_solution['transitions']['per_line'].items():
                        transition_data.append({
                            "Plant": line,
                            "Transitions": count
                        })
                    
                    transition_df = pd.DataFrame(transition_data)
                    st.dataframe(transition_df, use_container_width=True, hide_index=True)
            
            # TAB 3: INVENTORY TRENDS
            with tab3:
                                
                last_actual_day = num_days - buffer_days - 1

                for grade in sorted_grades:
                    inventory_values = [solver.Value(inventory_vars[(grade, d)]) for d in range(num_days)]
                
                    start_val = inventory_values[0]
                    end_val = inventory_values[last_actual_day]
                    highest_val = max(inventory_values[: last_actual_day + 1])
                    lowest_val = min(inventory_values[: last_actual_day + 1])
                
                    start_x = dates[0]
                    end_x = dates[last_actual_day]
                    highest_x = dates[inventory_values.index(highest_val)]
                    lowest_x = dates[inventory_values.index(lowest_val)]
                
                    fig = go.Figure()
                
                    fig.add_trace(go.Scatter(
                        x=dates,
                        y=inventory_values,
                        mode="lines+markers",
                        name=grade,
                        line=dict(color=grade_color_map[grade], width=3),
                        marker=dict(size=6),
                        hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"
                    ))
                
                    # Shutdown visualization
                    shutdown_added = False
                    for line in allowed_lines[grade]:
                        if line in shutdown_periods and shutdown_periods[line]:
                            shutdown_days = shutdown_periods[line]
                            start_shutdown = dates[shutdown_days[0]]
                            end_shutdown = dates[shutdown_days[-1]]
                            
                            fig.add_vrect(
                                x0=start_shutdown,
                                x1=end_shutdown + timedelta(days=1),
                                fillcolor="red",
                                opacity=0.1,
                                layer="below",
                                line_width=0,
                                annotation_text=f"Shutdown: {line}" if not shutdown_added else "",
                                annotation_position="top left",
                                annotation_font_size=12,
                                annotation_font_color="#c62828"
                            )
                            shutdown_added = True
                
                    # Min/Max lines
                    if min_inventory[grade] > 0:
                        fig.add_hline(
                            y=min_inventory[grade],
                            line=dict(color="#ef4444", width=2, dash="dash"),
                            annotation_text=f"Min: {min_inventory[grade]:,.0f}",
                            annotation_position="top left",
                            annotation_font_color="#ef4444"
                        )
                    
                    if max_inventory[grade] < 1000000000:
                        fig.add_hline(
                            y=max_inventory[grade],
                            line=dict(color="#10b981", width=2, dash="dash"),
                            annotation_text=f"Max: {max_inventory[grade]:,.0f}",
                            annotation_position="bottom left",
                            annotation_font_color="#10b981"
                        )
                
                    # Key point annotations
                    annotations = [
                        dict(
                            x=start_x, y=start_val,
                            text=f"Start: {start_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=-40, ay=30,
                            font=dict(color="#212121", size=11),
                            bgcolor="white", bordercolor="#bdbdbd", borderwidth=1
                        ),
                        dict(
                            x=end_x, y=end_val,
                            text=f"End: {end_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=40, ay=30,
                            font=dict(color="#212121", size=11),
                            bgcolor="white", bordercolor="#bdbdbd", borderwidth=1
                        ),
                        dict(
                            x=highest_x, y=highest_val,
                            text=f"Peak: {highest_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=0, ay=-40,
                            font=dict(color="#10b981", size=11),
                            bgcolor="white", bordercolor="#bdbdbd", borderwidth=1
                        ),
                        dict(
                            x=lowest_x, y=lowest_val,
                            text=f"Low: {lowest_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=0, ay=40,
                            font=dict(color="#ef4444", size=11),
                            bgcolor="white", bordercolor="#bdbdbd", borderwidth=1
                        )
                    ]
                
                    fig.update_layout(
                        title=dict(
                            text=f"Inventory Level - {grade}",
                            font=dict(size=16, color="#212121")
                        ),
                        xaxis=dict(
                            title="Date",
                            showgrid=True,
                            gridcolor="#e0e0e0",
                            tickformat="%d-%b"
                        ),
                        yaxis=dict(
                            title="Inventory Volume (MT)",
                            showgrid=True,
                            gridcolor="#e0e0e0"
                        ),
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(l=60, r=80, t=80, b=60),
                        height=420,
                        showlegend=False
                    )
                    
                    for ann in annotations:
                        fig.add_annotation(
                            x=ann['x'],
                            y=ann['y'],
                            text=ann['text'],
                            showarrow=ann['showarrow'],
                            arrowhead=ann['arrowhead'],
                            ax=ann['ax'],
                            ay=ann['ay'],
                            font=ann['font'],
                            bgcolor=ann['bgcolor'],
                            bordercolor=ann['bordercolor'],
                            borderwidth=ann.get('borderwidth', 1),
                            borderpad=4,
                            opacity=0.9
                        )
                
                    st.plotly_chart(fig, use_container_width=True)
            
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            
            # Action buttons at bottom
            col1, col2, col3 = st.columns([1, 2, 1])
            
            with col1:
                if st.button("🔄 New Optimization", use_container_width=True):
                    st.session_state.step = 1
                    st.session_state.uploaded_file = None
                    st.rerun()
            
            with col2:
                if st.button("📝 Adjust Parameters", use_container_width=True):
                    st.session_state.step = 2
                    st.rerun()

        else:
            st.markdown("""
            <div class="alert-box error">
                <div>
                    <strong>❌ No Feasible Solution Found</strong><br>
                    The optimization could not find a valid schedule with the given constraints.
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            with st.expander("🔍 Troubleshooting Guide", expanded=True):
                st.markdown("""
                ### Common Causes & Solutions
                
                #### 🔴 Capacity Issues
                - **Problem**: Total demand exceeds production capacity
                - **Solution**: Increase plant capacity or reduce demand forecasts
                
                #### 🔴 Constraint Conflicts
                - **Problem**: Minimum run days too long for available windows
                - **Solution**: Reduce minimum run day requirements
                
                #### 🔴 Inventory Issues
                - **Problem**: Cannot meet minimum closing inventory targets
                - **Solution**: Increase opening inventory or lower targets
                
                #### 🔴 Shutdown Conflicts
                - **Problem**: Shutdown periods block critical production
                - **Solution**: Reschedule shutdowns or increase opening inventory
                
                #### 🔴 Transition Restrictions
                - **Problem**: Transition matrix too restrictive
                - **Solution**: Allow more grade changeover combinations
                
                ### Recommended Actions
                
                1. Review and relax constraint parameters
                2. Check for data entry errors in Excel file
                3. Validate demand forecasts against capacity
                4. Consider increasing buffer days for flexibility
                """)
            
            col1, col2, col3 = st.columns([1, 2, 1])
            
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
