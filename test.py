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
# PAGE CONFIGURATION - Unchanged
# ============================================================================
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state for process tracking - Unchanged
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0
if 'solutions' not in st.session_state:
    st.session_state.solutions = []
if 'best_solution' not in st.session_state:
    st.session_state.best_solution = None

# ============================================================================
# MODERN CSS STYLING - Completely Redesigned
# ============================================================================
st.markdown("""
<style>
    /* Reset and base styles */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }
    
    /* Modern header - clean, minimal */
    .app-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
    }
    
    .app-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    
    .app-header p {
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
        opacity: 0.9;
        font-weight: 400;
    }
    
    /* Clean section headers */
    .section-title {
        font-size: 1.5rem;
        font-weight: 600;
        color: #1a202c;
        margin: 2rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #e2e8f0;
    }
    
    /* Modern card design - subtle, clean */
    .info-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        transition: box-shadow 0.2s ease;
    }
    
    .info-card:hover {
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
    }
    
    /* Metric cards - dashboard style */
    .metric-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.25);
        height: 100%;
    }
    
    .metric-label {
        font-size: 0.875rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.9;
        margin-bottom: 0.5rem;
    }
    
    .metric-value {
        font-size: 2.25rem;
        font-weight: 700;
        line-height: 1.2;
    }
    
    .metric-subtitle {
        font-size: 0.75rem;
        opacity: 0.8;
        margin-top: 0.25rem;
    }
    
    /* Status badges - clean indicators */
    .status-badge {
        display: inline-block;
        padding: 0.375rem 0.75rem;
        border-radius: 6px;
        font-size: 0.875rem;
        font-weight: 500;
        margin: 0.25rem;
    }
    
    .status-success {
        background: #d1fae5;
        color: #065f46;
        border: 1px solid #a7f3d0;
    }
    
    .status-warning {
        background: #fef3c7;
        color: #92400e;
        border: 1px solid #fde68a;
    }
    
    .status-info {
        background: #dbeafe;
        color: #1e40af;
        border: 1px solid #bfdbfe;
    }
    
    /* Data preview tables */
    .dataframe {
        font-size: 0.875rem;
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Streamlit component styling */
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        padding: 0.625rem 1.25rem;
        font-weight: 600;
        font-size: 1rem;
        border: none;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        box-shadow: 0 2px 4px rgba(102, 126, 234, 0.2);
        transition: all 0.2s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(102, 126, 234, 0.3);
    }
    
    /* File uploader - cleaner design */
    [data-testid="stFileUploader"] {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 1.5rem;
    }
    
    [data-testid="stFileUploader"] section {
        border: none;
        background: transparent;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: #f8fafc;
    }
    
    [data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }
    
    /* Tabs - modern, clean */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: #f8fafc;
        padding: 0.5rem;
        border-radius: 12px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        background: white;
        border: 1px solid #e2e8f0;
        color: #64748b;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-color: transparent;
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Expander styling */
    .streamlit-expanderHeader {
        font-weight: 600;
        font-size: 1rem;
        background: #f8fafc;
        border-radius: 8px;
    }
    
    /* Remove extra padding */
    .element-container {
        margin-bottom: 0;
    }
    
    /* Info boxes */
    .guide-box {
        background: linear-gradient(135deg, #e0e7ff 0%, #e0f2fe 100%);
        border-left: 4px solid #3b82f6;
        padding: 1.25rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .guide-box h4 {
        margin: 0 0 0.75rem 0;
        color: #1e40af;
        font-size: 1.125rem;
    }
    
    .guide-box ul {
        margin: 0.5rem 0;
        padding-left: 1.25rem;
    }
    
    .guide-box li {
        margin: 0.25rem 0;
        color: #1e40af;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# HEADER - Redesigned with clean, modern look
# ============================================================================
st.markdown("""
<div class="app-header">
    <h1>🏭 Polymer Production Scheduler</h1>
    <p>Multi-Plant Optimization with Shutdown Management & Inventory Control</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# SIDEBAR - Restructured for better workflow
# ============================================================================
with st.sidebar:
    st.markdown("### 📁 Data Input")
    
    uploaded_file = st.file_uploader(
        "Upload Production Data",
        type=["xlsx"],
        help="Excel file with Plant, Inventory, and Demand sheets"
    )
    
    if uploaded_file:
        st.success("✅ File loaded successfully")
        
        st.markdown("---")
        st.markdown("### ⚙️ Optimization Settings")
        
        # Basic parameters - always visible
        st.markdown("#### Core Parameters")
        time_limit_min = st.number_input(
            "Time Limit (minutes)",
            min_value=1,
            max_value=120,
            value=10,
            help="Maximum solver runtime"
        )
        
        buffer_days = st.number_input(
            "Planning Buffer (days)",
            min_value=0,
            max_value=7,
            value=3,
            help="Extra days for safety stock planning"
        )
        
        # Advanced parameters in expander
        with st.expander("🎯 Advanced Weights"):
            st.markdown("Fine-tune objective function priorities:")
            
            stockout_penalty = st.number_input(
                "Stockout Penalty",
                min_value=1,
                value=10,
                help="Cost weight for inventory shortages"
            )
            
            transition_penalty = st.number_input(
                "Transition Penalty",
                min_value=1,
                value=10,
                help="Cost weight for grade changeovers"
            )
            
            continuity_bonus = st.number_input(
                "Continuity Bonus",
                min_value=0,
                value=1,
                help="Reward for extended production runs"
            )
        
        st.markdown("---")
        st.markdown("### 📊 Optimization Status")
        if st.session_state.current_step == 0:
            st.info("Ready to optimize")
        elif st.session_state.current_step == 2:
            st.warning("Optimization running...")
        else:
            st.success("Optimization complete")

# ============================================================================
# MAIN CONTENT AREA
# ============================================================================
if uploaded_file:
    try:
        # Read uploaded file
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        st.markdown("---")
        
        # ========================================================================
        # DATA PREVIEW SECTION - Redesigned with cleaner cards
        # ========================================================================
        st.markdown('<div class="section-title">📋 Data Overview & Validation</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            try:
                plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                plant_display_df = plant_df.copy()
                start_column = plant_display_df.columns[4]
                end_column = plant_display_df.columns[5]
                
                if pd.api.types.is_datetime64_any_dtype(plant_display_df[start_column]):
                    plant_display_df[start_column] = plant_display_df[start_column].dt.strftime('%d-%b-%y')
                if pd.api.types.is_datetime64_any_dtype(plant_display_df[end_column]):
                    plant_display_df[end_column] = plant_display_df[end_column].dt.strftime('%d-%b-%y')

                st.markdown("**🏭 Plant Configuration**")
                st.dataframe(plant_display_df, use_container_width=True, height=250)
            except Exception as e:
                st.error(f"Error reading Plant sheet: {e}")
                st.stop()
        
        with col2:
            try:
                excel_file.seek(0)
                inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                inventory_display_df = inventory_df.copy()
                force_start_column = inventory_display_df.columns[7]
                
                if pd.api.types.is_datetime64_any_dtype(inventory_display_df[force_start_column]):
                    inventory_display_df[force_start_column] = inventory_display_df[force_start_column].dt.strftime('%d-%b-%y')
                
                st.markdown("**📦 Inventory Rules**")
                st.dataframe(inventory_display_df, use_container_width=True, height=250)
            except Exception as e:
                st.error(f"Error reading Inventory sheet: {e}")
                st.stop()
        
        with col3:
            try:
                excel_file.seek(0)
                demand_df = pd.read_excel(excel_file, sheet_name='Demand')
                
                demand_display_df = demand_df.copy()
                
                date_column = demand_display_df.columns[0]
                if pd.api.types.is_datetime64_any_dtype(demand_display_df[date_column]):
                    demand_display_df[date_column] = demand_display_df[date_column].dt.strftime('%d-%b-%y')
                
                st.markdown("**📈 Demand Forecast**")
                st.dataframe(demand_display_df, use_container_width=True, height=250)
            except Exception as e:
                st.error(f"Error reading Demand sheet: {e}")
                st.stop()
        
        excel_file.seek(0)
        
        # ========================================================================
        # SHUTDOWN INFORMATION - Cleaner display
        # ========================================================================
        st.markdown("---")
        st.markdown('<div class="section-title">🔧 Planned Shutdowns</div>', unsafe_allow_html=True)
        
        shutdown_found = False
        shutdown_badges = []
        
        for index, row in plant_df.iterrows():
            plant = row['Plant']
            shutdown_start = row.get('Shutdown Start Date')
            shutdown_end = row.get('Shutdown End Date')
            
            if pd.notna(shutdown_start) and pd.notna(shutdown_end):
                try:
                    start_date = pd.to_datetime(shutdown_start).date()
                    end_date = pd.to_datetime(shutdown_end).date()
                    duration = (end_date - start_date).days + 1
                    
                    if start_date > end_date:
                        shutdown_badges.append(f'<span class="status-badge status-warning">⚠️ {plant}: Invalid shutdown dates</span>')
                    else:
                        shutdown_badges.append(f'<span class="status-badge status-info">🔧 {plant}: {start_date.strftime("%d-%b-%y")} to {end_date.strftime("%d-%b-%y")} ({duration} days)</span>')
                        shutdown_found = True
                except Exception as e:
                    shutdown_badges.append(f'<span class="status-badge status-warning">⚠️ {plant}: Invalid dates</span>')
        
        if shutdown_found:
            st.markdown(" ".join(shutdown_badges), unsafe_allow_html=True)
        else:
            st.markdown('<span class="status-badge status-success">✓ No shutdowns scheduled</span>', unsafe_allow_html=True)
        
        # ========================================================================
        # TRANSITION MATRICES - Load (logic unchanged)
        # ========================================================================
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
                    st.info(f"✅ Loaded transition matrix for {plant_name} from sheet '{sheet_name}'")
                    break
                except:
                    continue
            
            if transition_df_found is not None:
                transition_dfs[plant_name] = transition_df_found
            else:
                st.info(f"ℹ️ No transition matrix found for {plant_name}. Assuming no transition constraints.")
                transition_dfs[plant_name] = None
        
        st.markdown("---")
        
        # ========================================================================
        # OPTIMIZATION BUTTON - Prominent, clear call-to-action
        # ========================================================================
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            run_optimization = st.button("🚀 Run Optimization", type="primary", use_container_width=True)
        
        if run_optimization:
            st.session_state.current_step = 2
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            if 'solutions' not in st.session_state:
                st.session_state.solutions = []
            if 'best_solution' not in st.session_state:
                st.session_state.best_solution = None

            time.sleep(1)
            
            status_text.markdown('<span class="status-badge status-info">📊 Preprocessing data...</span>', unsafe_allow_html=True)
            progress_bar.progress(10)

            time.sleep(2)

            # ====================================================================
            # DATA PREPROCESSING - Logic completely unchanged
            # ====================================================================
            try:
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
                        st.warning(f"⚠️ Lines for grade '{grade}' (row {index}) are not specified, allowing all lines")
                    
                    for plant in plants_for_row:
                        if plant not in allowed_lines[grade]:
                            allowed_lines[grade].append(plant)
                    
                    if grade not in grade_inventory_defined:
                        if pd.notna(row['Opening Inventory']):
                            initial_inventory[grade] = row['Opening Inventory']
                        else:
                            initial_inventory[grade] = 0
                        
                        if pd.notna(row['Min. Inventory']):
                            min_inventory[grade] = row['Min. Inventory']
                        else:
                            min_inventory[grade] = 0
                        
                        if pd.notna(row['Max. Inventory']):
                            max_inventory[grade] = row['Max. Inventory']
                        else:
                            max_inventory[grade] = 1000000000
                        
                        if pd.notna(row['Min. Closing Inventory']):
                            min_closing_inventory[grade] = row['Min. Closing Inventory']
                        else:
                            min_closing_inventory[grade] = 0
                        
                        grade_inventory_defined.add(grade)
                    
                    for plant in plants_for_row:
                        grade_plant_key = (grade, plant)
                        
                        if pd.notna(row['Min. Run Days']):
                            min_run_days[grade_plant_key] = int(row['Min. Run Days'])
                        else:
                            min_run_days[grade_plant_key] = 1
                        
                        if pd.notna(row['Max. Run Days']):
                            max_run_days[grade_plant_key] = int(row['Max. Run Days'])
                        else:
                            max_run_days[grade_plant_key] = 9999
                        
                        if pd.notna(row['Force Start Date']):
                            try:
                                force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date()
                            except:
                                force_start_date[grade_plant_key] = None
                                st.warning(f"⚠️ Invalid Force Start Date for grade '{grade}' on plant '{plant}'")
                        else:
                            force_start_date[grade_plant_key] = None
                        
                        rerun_val = row['Rerun Allowed']
                        if pd.notna(rerun_val):
                            val_str = str(rerun_val).strip().lower()
                            if val_str in ['no', 'n', 'false', '0']:
                                rerun_allowed[grade_plant_key] = False
                            else:
                                rerun_allowed[grade_plant_key] = True
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
                            st.warning(f"⚠️ Invalid Material Running or Expected Run Days for plant '{plant}'")
            
            except Exception as e:
                st.error(f"Error in data preprocessing: {str(e)}")
                import traceback
                st.error(f"Traceback: {traceback.format_exc()}")
                st.stop()
            
            # Process demand data
            demand_data = {}
            dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            num_days = len(dates)
            last_date = dates[-1]
            for i in range(1, buffer_days + 1):
                dates.append(last_date + timedelta(days=i))
            num_days = len(dates)
            
            formatted_dates = [date.strftime('%d-%b-%y') for date in dates]
            
            for grade in grades:
                if grade in demand_df.columns:
                    demand_data[grade] = {demand_df.iloc[i, 0].date(): demand_df[grade].iloc[i] for i in range(len(demand_df))}
                else:
                    st.warning(f"Demand data not found for grade '{grade}'. Assuming zero demand.")
                    demand_data[grade] = {date: 0 for date in dates}
            
            for grade in grades:
                for date in dates[-buffer_days:]:
                    if date not in demand_data[grade]:
                        demand_data[grade][date] = 0
            
            shutdown_periods = process_shutdown_dates(plant_df, dates)
            
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
            status_text.markdown('<span class="status-badge status-info">🔧 Building model...</span>', unsafe_allow_html=True)

            time.sleep(2)
            
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
                if key not in production:
                    return 0
                return production[key]
            
            def get_is_producing_var(grade, line, d):
                key = (grade, line, d)
                if key not in is_producing:
                    return None
                return is_producing[key]
            
            # SHUTDOWN CONSTRAINTS
            for line in lines:
                if line in shutdown_periods and shutdown_periods[line]:
                    for d in shutdown_periods[line]:
                        for grade in grades:
                            if is_allowed_combination(grade, line):
                                key = (grade, line, d)
                                if key in is_producing:
                                    model.Add(is_producing[key] == 0)
                                    model.Add(production[key] == 0)

            shutdown_demand = {}
            for grade in grades:
                shutdown_demand[grade] = 0
                for line in allowed_lines[grade]:
                    if line in shutdown_periods:
                        for d in shutdown_periods[line]:
                            shutdown_demand[grade] += demand_data[grade].get(dates[d], 0)
            
            for grade, total_shutdown_demand in shutdown_demand.items():
                if total_shutdown_demand > initial_inventory[grade]:
                    st.warning(f"⚠️ Grade '{grade}': Shutdown periods require {total_shutdown_demand} MT from inventory (current: {initial_inventory[grade]} MT). Consider increasing opening inventory or adjusting shutdown schedule.")
            
            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')
            
            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')
            
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
            
            for plant, (material, expected_days) in material_running_info.items():
                for d in range(min(expected_days, num_days)):
                    if is_allowed_combination(material, plant):
                        model.Add(get_is_producing_var(material, plant, d) == 1)
                        for other_material in grades:
                            if other_material != material and is_allowed_combination(other_material, plant):
                                model.Add(get_is_producing_var(other_material, plant, d) == 0)
            
            objective = 0
            
            for grade in grades:
                model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])
            
            for grade in grades:
                for d in range(num_days):
                    produced_today = sum(
                        get_production_var(grade, line, d) 
                        for line in allowed_lines[grade]
                    )
                    demand_today = demand_data[grade].get(dates[d], 0)
                    
                    supplied = model.NewIntVar(0, 100000, f'supplied_{grade}_{d}')
                    model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
                    model.Add(supplied <= demand_today)
                    
                    model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
                    
                    model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
                    
                    model.Add(inventory_vars[(grade, d + 1)] >= 0)
            
            for grade in grades:
                for d in range(num_days):
                    if min_inventory[grade] > 0:
                        min_inv_value = int(min_inventory[grade])
                        inventory_tomorrow = inventory_vars[(grade, d + 1)]
                        
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        model.Add(deficit >= min_inv_value - inventory_tomorrow)
                        model.Add(deficit >= 0)
                        objective += stockout_penalty * deficit
            
            for grade in grades:
                closing_inventory = inventory_vars[(grade, num_days - buffer_days)]
                min_closing = min_closing_inventory[grade]
                
                if min_closing > 0:
                    closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                    model.Add(closing_deficit >= min_closing - closing_inventory)
                    model.Add(closing_deficit >= 0)
                    objective += stockout_penalty * closing_deficit * 3
            
            for grade in grades:
                for d in range(1, num_days + 1):
                    model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])
            
            for line in lines:
                for d in range(num_days - buffer_days):
                    if line in shutdown_periods and d in shutdown_periods[line]:
                        continue
                    production_vars = [
                        get_production_var(grade, line, d) 
                        for grade in grades 
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:
                        model.Add(sum(production_vars) == capacities[line])
                
                for d in range(num_days - buffer_days, num_days):
                    production_vars = [
                        get_production_var(grade, line, d) 
                        for grade in grades 
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:
                        model.Add(sum(production_vars) <= capacities[line])
            
            for grade_plant_key, start_date in force_start_date.items():
                if start_date:
                    grade, plant = grade_plant_key
                    try:
                        start_day_index = dates.index(start_date)
                        var = get_is_producing_var(grade, plant, start_day_index)
                        if var is not None:
                            model.Add(var == 1)
                            st.info(f"✅ Enforced force start date for grade '{grade}' on plant '{plant}' at day {start_date.strftime('%d-%b-%y')}")
                        else:
                            st.warning(f"⚠️ Cannot enforce force start date for grade '{grade}' on plant '{plant}' - combination not allowed")
                    except ValueError:
                        st.warning(f"⚠️ Force start date '{start_date.strftime('%d-%b-%y')}' for grade '{grade}' on plant '{plant}' not found in demand dates")
            
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
            
            for line in lines:
                if transition_rules.get(line):
                    for d in range(num_days - 1):
                        for prev_grade in grades:
                            if prev_grade in transition_rules[line] and is_allowed_combination(prev_grade, line):
                                allowed_next = transition_rules[line][prev_grade]
                                for current_grade in grades:
                                    if (current_grade != prev_grade and 
                                        current_grade not in allowed_next and 
                                        is_allowed_combination(current_grade, line)):
                                        
                                        prev_var = get_is_producing_var(prev_grade, line, d)
                                        current_var = get_is_producing_var(current_grade, line, d + 1)
                                        
                                        if prev_var is not None and current_var is not None:
                                            model.Add(prev_var + current_var <= 1)

            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    if not rerun_allowed.get(grade_plant_key, True):
                        starts = [is_start_vars[(grade, line, d)] for d in range(num_days) 
                                 if (grade, line, d) in is_start_vars]
                        if starts:
                            model.Add(sum(starts) <= 1)

            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

            for line in lines:
                for d in range(num_days - 1):
                    transition_vars = []
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
                            transition_vars.append(trans_var)
                            objective += transition_penalty * trans_var

                    for grade in grades:
                        if line in allowed_lines[grade]:
                            continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                            model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                            objective += -continuity_bonus * continuity

            model.Minimize(objective)

            progress_bar.progress(50)
            status_text.markdown('<span class="status-badge status-info">⚡ Running solver...</span>', unsafe_allow_html=True)

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
            # RESULTS DISPLAY - Completely Redesigned UI
            # ====================================================================
            if status == cp_model.OPTIMAL:
                status_text.markdown('<span class="status-badge status-success">✅ Optimal solution found!</span>', unsafe_allow_html=True)
            elif status == cp_model.FEASIBLE:
                status_text.markdown('<span class="status-badge status-success">✅ Feasible solution found!</span>', unsafe_allow_html=True)
            else:
                status_text.markdown('<span class="status-badge status-warning">⚠️ No optimal solution found</span>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown('<div class="section-title">📊 Optimization Results</div>', unsafe_allow_html=True)

            if solution_callback.num_solutions() > 0:
                best_solution = solution_callback.solutions[-1]

                # ============================================================
                # KEY METRICS DASHBOARD - Redesigned with clean cards
                # ============================================================
                st.markdown("#### Performance Metrics")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">Objective Value</div>
                            <div class="metric-value">{best_solution['objective']:,.0f}</div>
                            <div class="metric-subtitle">Lower is Better</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">Total Transitions</div>
                            <div class="metric-value">{best_solution['transitions']['total']}</div>
                            <div class="metric-subtitle">Grade Changeovers</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                    st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">Total Stockouts</div>
                            <div class="metric-value">{total_stockouts:,.0f}</div>
                            <div class="metric-subtitle">MT Unmet Demand</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                        <div class="metric-container">
                            <div class="metric-label">Planning Horizon</div>
                            <div class="metric-value">{num_days}</div>
                            <div class="metric-subtitle">Days</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # ============================================================
                # TABBED RESULTS - Clean, organized navigation
                # ============================================================
                tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Production Summary", "📦 Inventory Analysis"])
                
                # TAB 1: PRODUCTION SCHEDULE
                with tab1:
                    st.markdown("### Visual Production Schedule")
                    
                    sorted_grades = sorted(grades)
                    base_colors = px.colors.qualitative.Vivid
                    grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

                    for line in lines:
                        st.markdown(f"#### {line}")
                    
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
                                annotation_font_color="red"
                            )
                    
                        fig.update_yaxes(
                            autorange="reversed",
                            title=None,
                            showgrid=True,
                            gridcolor="#e2e8f0"
                        )
                    
                        fig.update_xaxes(
                            title="Date",
                            showgrid=True,
                            gridcolor="#e2e8f0",
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
                        
                        # Schedule table below chart
                        st.markdown(f"**Detailed Schedule - {line}**")
                        
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
                    
                        if schedule_data:
                            schedule_df = pd.DataFrame(schedule_data)
                            st.dataframe(schedule_df, use_container_width=True, hide_index=True)
                        
                        st.markdown("---")
                
                # TAB 2: PRODUCTION SUMMARY
                with tab2:
                    st.markdown("### Production Volume by Grade and Plant")
                    
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
                            lambda x: ['background-color: #f8fafc; font-weight: bold' if x.name == len(total_prod_df) - 1 else '' for i in x],
                            axis=1
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    # Transition details
                    st.markdown("### Transition Details by Plant")
                    transition_data = []
                    for line, count in best_solution['transitions']['per_line'].items():
                        transition_data.append({
                            "Plant": line,
                            "Transitions": count
                        })
                    
                    transition_df = pd.DataFrame(transition_data)
                    st.dataframe(transition_df, use_container_width=True, hide_index=True)
                
                # TAB 3: INVENTORY ANALYSIS
                with tab3:
                    st.markdown("### Inventory Trajectory by Grade")
                    
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
                                    annotation_font_color="red"
                                )
                                shutdown_added = True
                    
                        # Min/Max inventory lines
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
                                font=dict(color="#1a202c", size=11),
                                bgcolor="white", bordercolor="#cbd5e1", borderwidth=1
                            ),
                            dict(
                                x=end_x, y=end_val,
                                text=f"End: {end_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=40, ay=30,
                                font=dict(color="#1a202c", size=11),
                                bgcolor="white", bordercolor="#cbd5e1", borderwidth=1
                            ),
                            dict(
                                x=highest_x, y=highest_val,
                                text=f"Peak: {highest_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=0, ay=-40,
                                font=dict(color="#10b981", size=11),
                                bgcolor="white", bordercolor="#cbd5e1", borderwidth=1
                            ),
                            dict(
                                x=lowest_x, y=lowest_val,
                                text=f"Low: {lowest_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=0, ay=40,
                                font=dict(color="#ef4444", size=11),
                                bgcolor="white", bordercolor="#cbd5e1", borderwidth=1
                            )
                        ]
                    
                        fig.update_layout(
                            title=dict(
                                text=f"Inventory Level - {grade}",
                                font=dict(size=16, color="#1a202c")
                            ),
                            xaxis=dict(
                                title="Date",
                                showgrid=True,
                                gridcolor="#e2e8f0",
                                tickformat="%d-%b"
                            ),
                            yaxis=dict(
                                title="Inventory Volume (MT)",
                                showgrid=True,
                                gridcolor="#e2e8f0"
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

            else:
                st.error("❌ No feasible solution found")
                
                with st.expander("🔍 Troubleshooting Guide"):
                    st.markdown("""
                    ### Common Causes of Infeasibility:
                    
                    **Capacity Issues:**
                    - Total demand exceeds available production capacity
                    - Shutdown periods reduce capacity below demand requirements
                    
                    **Constraint Conflicts:**
                    - Minimum run days too long for available time windows
                    - Force start dates conflict with other constraints
                    - Minimum closing inventory cannot be achieved
                    
                    **Transition Restrictions:**
                    - Transition matrix too restrictive
                    - No valid production sequence exists
                    
                    ### Suggested Actions:
                    
                    1. **Reduce Constraints:**
                       - Lower minimum run days requirements
                       - Reduce minimum inventory targets
                       - Remove or adjust force start dates
                    
                    2. **Increase Capacity:**
                       - Add production capacity
                       - Reduce or reschedule shutdown periods
                    
                    3. **Adjust Demand:**
                       - Review demand forecast accuracy
                       - Consider phased delivery schedules
                    
                    4. **Relax Transitions:**
                       - Allow more grade changeover combinations
                       - Review transition matrix restrictions
                    """)

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        with st.expander("View Error Details"):
            import traceback
            st.code(traceback.format_exc())
        st.info("💡 Ensure your Excel file contains sheets: 'Plant', 'Inventory', and 'Demand' with proper formatting")

else:
    # ========================================================================
    # WELCOME SCREEN - Redesigned with clean, informative layout
    # ========================================================================
    st.markdown('<div class="section-title">🚀 Getting Started</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        <div class="info-card">
            <h3 style="margin-top: 0; color: #1a202c;">Welcome to Polymer Production Scheduler</h3>
            <p style="color: #64748b; line-height: 1.6;">
                This application uses advanced optimization algorithms to create efficient production schedules 
                for multi-plant polymer manufacturing operations. It considers capacity constraints, inventory 
                management, demand fulfillment, and grade transition rules.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("""
        <div class="guide-box">
            <h4>📋 Quick Start Steps</h4>
            <ol style="line-height: 1.8;">
                <li><strong>Download</strong> the sample Excel template (see right panel)</li>
                <li><strong>Fill in</strong> your production data (plants, inventory, demand)</li>
                <li><strong>Upload</strong> your Excel file using the sidebar</li>
                <li><strong>Configure</strong> optimization parameters</li>
                <li><strong>Run</strong> the optimization and review results</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="info-card">
            <h4 style="margin-top: 0; color: #1a202c;">📥 Sample Template</h4>
            <p style="color: #64748b; font-size: 0.875rem; margin-bottom: 1rem;">
                Download our pre-configured template with sample data and proper formatting.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        sample_workbook = get_sample_workbook()
        
        st.download_button(
            label="📥 Download Template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    st.markdown("---")
    
    # Feature highlights
    st.markdown('<div class="section-title">✨ Key Features</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="info-card">
            <h4 style="color: #667eea;">🏭 Multi-Plant Support</h4>
            <ul style="color: #64748b; font-size: 0.875rem; line-height: 1.6;">
                <li>Manage multiple production lines</li>
                <li>Grade-specific plant assignments</li>
                <li>Capacity optimization across facilities</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="info-card">
            <h4 style="color: #667eea;">📦 Inventory Management</h4>
            <ul style="color: #64748b; font-size: 0.875rem; line-height: 1.6;">
                <li>Min/max inventory constraints</li>
                <li>Closing inventory targets</li>
                <li>Stockout minimization</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="info-card">
            <h4 style="color: #667eea;">🔧 Advanced Constraints</h4>
            <ul style="color: #64748b; font-size: 0.875rem; line-height: 1.6;">
                <li>Shutdown period handling</li>
                <li>Grade transition rules</li>
                <li>Run length constraints</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # File format details
    with st.expander("📄 Required Excel File Format", expanded=False):
        st.markdown("""
        ### Sheet Structure
        
        Your Excel workbook must contain these sheets:
        
        #### 1. Plant Sheet
        | Column | Description | Required |
        |--------|-------------|----------|
        | Plant | Plant identifier | Yes |
        | Capacity per day | Daily production capacity (MT) | Yes |
        | Material Running | Currently running grade | Optional |
        | Expected Run Days | Expected continuation days | Optional |
        | Shutdown Start Date | Maintenance start date | Optional |
        | Shutdown End Date | Maintenance end date | Optional |
        
        #### 2. Inventory Sheet
        | Column | Description | Required |
        |--------|-------------|----------|
        | Grade Name | Product grade identifier | Yes |
        | Opening Inventory | Starting stock (MT) | Yes |
        | Min. Inventory | Minimum safety stock (MT) | Yes |
        | Max. Inventory | Storage capacity (MT) | Yes |
        | Min. Run Days | Minimum production run length | Yes |
        | Max. Run Days | Maximum production run length | Yes |
        | Force Start Date | Mandatory production start | Optional |
        | Lines | Allowed production lines | Yes |
        | Rerun Allowed | Allow multiple runs (Yes/No) | Yes |
        | Min. Closing Inventory | End period target (MT) | Yes |
        
        #### 3. Demand Sheet
        - First column: Date (daily dates)
        - Remaining columns: Demand for each grade (column name = grade name)
        
        #### 4. Transition Sheets (Optional)
        - Sheet name: `Transition_[PlantName]`
        - Matrix format: Previous grade (rows) → Next grade (columns)
        - Values: "yes" for allowed transitions
        
        ### Multi-Plant Grade Configuration
        
        For grades that can run on multiple plants, create multiple rows:
        ```
        Grade Name | Lines  | Force Start Date | Min. Run Days
        BOPP       | Plant1 |                  | 5
        BOPP       | Plant2 | 01-Dec-24        | 5
        ```
        
        ### Shutdown Period Example
        ```
        Plant  | Shutdown Start Date | Shutdown End Date
        Plant1 | 15-Nov-25           | 18-Nov-25
        ```
        During shutdowns, the plant will have zero production capacity.
        """)

# ============================================================================
# FOOTER
# ============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #94a3b8; font-size: 0.875rem; padding: 1rem 0;">
    Polymer Production Scheduler • Built with Streamlit & OR-Tools • Version 2.0
</div>
""", unsafe_allow_html=True)
