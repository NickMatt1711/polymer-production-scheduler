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

st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'plant_df' not in st.session_state:
    st.session_state.plant_df = None
if 'inventory_df' not in st.session_state:
    st.session_state.inventory_df = None
if 'demand_df' not in st.session_state:
    st.session_state.demand_df = None
if 'transition_dfs' not in st.session_state:
    st.session_state.transition_dfs = None
if 'optimization_run' not in st.session_state:
    st.session_state.optimization_run = False
if 'results' not in st.session_state:
    st.session_state.results = None

# Modern CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8eef5 100%);
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .top-nav {
        background: rgba(255, 255, 255, 0.98);
        backdrop-filter: blur(20px);
        border-bottom: 1px solid rgba(0, 0, 0, 0.06);
        padding: 1.25rem 2rem;
        margin: -1rem -1rem 2rem -1rem;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.04);
    }
    
    .nav-content {
        display: flex;
        justify-content: space-between;
        align-items: center;
        max-width: 1400px;
        margin: 0 auto;
    }
    
    .logo-section {
        display: flex;
        align-items: center;
        gap: 0.75rem;
    }
    
    .logo-text {
        font-size: 1.25rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .modern-card {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.04);
        border: 1px solid rgba(0, 0, 0, 0.04);
        margin-bottom: 1.5rem;
    }
    
    .card-title {
        font-size: 1.125rem;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 1rem;
    }
    
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .metric-card-modern {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.04);
        border: 1px solid rgba(0, 0, 0, 0.04);
        transition: all 0.3s ease;
    }
    
    .metric-card-modern:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
    }
    
    .metric-label-modern {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94a3b8;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .metric-value-modern {
        font-size: 2rem;
        font-weight: 700;
        color: #1e293b;
    }
    
    .stButton > button {
        border-radius: 12px;
        border: none;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        box-shadow: 0 4px 16px rgba(102, 126, 234, 0.3);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    
    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.5rem;
        background: transparent;
        border-radius: 8px;
        font-weight: 600;
        color: #64748b;
        border: none;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    .upload-section {
        background: white;
        border-radius: 20px;
        padding: 3rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06);
        margin: 2rem auto;
        max-width: 800px;
    }
    
    .upload-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.5rem;
    }
    
    .upload-subtitle {
        font-size: 1rem;
        color: #64748b;
        margin-bottom: 2rem;
    }
    
    .param-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 1.5rem;
        margin: 1.5rem 0;
    }
    
    .param-card {
        background: #f8fafc;
        border-radius: 12px;
        padding: 1.25rem;
        border: 1px solid #e2e8f0;
    }
    
    .param-header {
        font-size: 0.875rem;
        font-weight: 600;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 1rem;
    }
    
    .dataframe {
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
    }
</style>
""", unsafe_allow_html=True)

# Top Navigation
st.markdown("""
<div class="top-nav">
    <div class="nav-content">
        <div class="logo-section">
            <span style="font-size: 1.5rem;">🏭</span>
            <div class="logo-text">Polymer Production Scheduler</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Create tabs
tab1, tab2, tab3 = st.tabs(["📤 Upload Data", "⚙️ Configure", "🎯 Optimize & Results"])

with tab1:
    st.markdown("""
    <div class="upload-section">
        <h1 class="upload-title">Welcome to Production Optimization</h1>
        <p class="upload-subtitle">Upload your Excel file to begin optimizing your polymer production schedule</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        uploaded_file = st.file_uploader(
            "Choose Excel File",
            type=["xlsx"],
            help="Upload an Excel file with Plant, Inventory, and Demand sheets",
            key="file_uploader"
        )
        
        if uploaded_file:
            try:
                # Read and store data in session state
                uploaded_file.seek(0)
                excel_file = io.BytesIO(uploaded_file.read())
                
                st.session_state.plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                excel_file.seek(0)
                st.session_state.inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                excel_file.seek(0)
                st.session_state.demand_df = pd.read_excel(excel_file, sheet_name='Demand')
                
                # Load transition matrices
                st.session_state.transition_dfs = {}
                for i in range(len(st.session_state.plant_df)):
                    plant_name = st.session_state.plant_df['Plant'].iloc[i]
                    possible_sheet_names = [
                        f'Transition_{plant_name}',
                        f'Transition_{plant_name.replace(" ", "_")}',
                        f'Transition{plant_name.replace(" ", "")}',
                    ]
                    
                    for sheet_name in possible_sheet_names:
                        try:
                            excel_file.seek(0)
                            st.session_state.transition_dfs[plant_name] = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=0)
                            break
                        except:
                            continue
                    
                    if plant_name not in st.session_state.transition_dfs:
                        st.session_state.transition_dfs[plant_name] = None
                
                st.session_state.uploaded_file = uploaded_file
                st.success("✅ File uploaded successfully! Please go to the 'Configure' tab.")
                
            except Exception as e:
                st.error(f"Error reading file: {str(e)}")
    
    st.markdown("---")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown("""
        <div class="modern-card">
            <div class="card-title">📋 Quick Start Guide</div>
            <ol style="color: #64748b; line-height: 1.8;">
                <li>Download the sample template or prepare your own Excel file</li>
                <li>Ensure it contains Plant, Inventory, and Demand sheets</li>
                <li>Upload your file using the button above</li>
                <li>Configure optimization parameters in the next tab</li>
                <li>Run optimization and view results</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        sample_workbook = get_sample_workbook()
        st.download_button(
            label="📥 Download Template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

with tab2:
    if st.session_state.uploaded_file is None:
        st.info("⬅️ Please upload a file in the 'Upload Data' tab first")
    else:
        st.markdown('<div class="modern-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">⚙️ Optimization Parameters</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="param-grid">', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="param-card">', unsafe_allow_html=True)
            st.markdown('<div class="param-header">Basic Settings</div>', unsafe_allow_html=True)
            time_limit_min = st.number_input(
                "Time limit (minutes)",
                min_value=1,
                max_value=120,
                value=10,
                help="Maximum time to run the optimization"
            )
            buffer_days = st.number_input(
                "Buffer days",
                min_value=0,
                max_value=7,
                value=3,
                help="Additional days for planning buffer"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="param-card">', unsafe_allow_html=True)
            st.markdown('<div class="param-header">Objective Weights</div>', unsafe_allow_html=True)
            stockout_penalty = st.number_input(
                "Stockout penalty",
                min_value=1,
                value=10,
                help="Penalty weight for stockouts"
            )
            transition_penalty = st.number_input(
                "Transition penalty", 
                min_value=1,
                value=10,
                help="Penalty weight for transitions"
            )
            continuity_bonus = st.number_input(
                "Continuity bonus",
                min_value=0,
                value=1,
                help="Bonus for continuing same grade"
            )
            st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Store parameters in session state
        st.session_state.time_limit_min = time_limit_min
        st.session_state.buffer_days = buffer_days
        st.session_state.stockout_penalty = stockout_penalty
        st.session_state.transition_penalty = transition_penalty
        st.session_state.continuity_bonus = continuity_bonus
        
        st.success("✅ Parameters configured! Go to 'Optimize & Results' tab to run optimization.")

with tab3:
    if st.session_state.uploaded_file is None:
        st.info("⬅️ Please upload a file in the 'Upload Data' tab first")
    else:
        # Data Preview
        st.markdown('<div class="modern-card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📊 Data Preview</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**🏭 Plant Data**")
            plant_display = st.session_state.plant_df.copy()
            if len(plant_display.columns) > 4:
                if pd.api.types.is_datetime64_any_dtype(plant_display.iloc[:, 4]):
                    plant_display.iloc[:, 4] = plant_display.iloc[:, 4].dt.strftime('%d-%b-%y')
            if len(plant_display.columns) > 5:
                if pd.api.types.is_datetime64_any_dtype(plant_display.iloc[:, 5]):
                    plant_display.iloc[:, 5] = plant_display.iloc[:, 5].dt.strftime('%d-%b-%y')
            st.dataframe(plant_display, use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("**📦 Inventory Data**")
            st.dataframe(st.session_state.inventory_df, use_container_width=True, hide_index=True)
        
        with col3:
            st.markdown("**📊 Demand Data**")
            demand_display = st.session_state.demand_df.copy()
            if pd.api.types.is_datetime64_any_dtype(demand_display.iloc[:, 0]):
                demand_display.iloc[:, 0] = demand_display.iloc[:, 0].dt.strftime('%d-%b-%y')
            st.dataframe(demand_display, use_container_width=True, hide_index=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Run Optimization Button
        if st.button("🎯 Run Production Optimization", type="primary", use_container_width=True):
            
            progress_container = st.empty()
            status_container = st.empty()
            
            with progress_container.container():
                progress_bar = st.progress(0)
            
            with status_container.container():
                st.info("🔄 Starting optimization...")
            
            try:
                # Get parameters from session state
                time_limit_min = st.session_state.get('time_limit_min', 10)
                buffer_days = st.session_state.get('buffer_days', 3)
                stockout_penalty = st.session_state.get('stockout_penalty', 10)
                transition_penalty = st.session_state.get('transition_penalty', 10)
                continuity_bonus = st.session_state.get('continuity_bonus', 1)
                
                plant_df = st.session_state.plant_df
                inventory_df = st.session_state.inventory_df
                demand_df = st.session_state.demand_df
                transition_dfs = st.session_state.transition_dfs
                
                status_container.info("🔄 Preprocessing data...")
                progress_bar.progress(10)
                
                # Process inventory data
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
                        demand_data[grade] = {date: 0 for date in dates}
                
                for grade in grades:
                    for date in dates[-buffer_days:]:
                        if date not in demand_data[grade]:
                            demand_data[grade][date] = 0
                
                # Process shutdown dates
                shutdown_periods = {}
                for index, row in plant_df.iterrows():
                    plant = row['Plant']
                    shutdown_start = row.get('Shutdown Start Date')
                    shutdown_end = row.get('Shutdown End Date')
                    
                    if pd.notna(shutdown_start) and pd.notna(shutdown_end):
                        try:
                            start_date = pd.to_datetime(shutdown_start).date()
                            end_date = pd.to_datetime(shutdown_end).date()
                            
                            if start_date <= end_date:
                                shutdown_days = []
                                for d, date in enumerate(dates):
                                    if start_date <= date <= end_date:
                                        shutdown_days.append(d)
                                shutdown_periods[plant] = shutdown_days
                            else:
                                shutdown_periods[plant] = []
                        except:
                            shutdown_periods[plant] = []
                    else:
                        shutdown_periods[plant] = []
                
                # Process transition rules
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
                status_container.info("🔧 Building optimization model...")
                
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
                    return is_producing.get(key)
                
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
                
                # One grade per line per day
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
                            var = get_is_producing_var(material, plant, d)
                            if var:
                                model.Add(var == 1)
                
                objective = 0
                
                # Inventory balance
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
                
                # Minimum inventory constraints
                for grade in grades:
                    for d in range(num_days):
                        if min_inventory[grade] > 0:
                            deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                            model.Add(deficit >= int(min_inventory[grade]) - inventory_vars[(grade, d + 1)])
                            model.Add(deficit >= 0)
                            objective += stockout_penalty * deficit
                
                # Minimum closing inventory
                for grade in grades:
                    closing_inventory = inventory_vars[(grade, num_days - buffer_days)]
                    min_closing = min_closing_inventory[grade]
                    
                    if min_closing > 0:
                        closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                        model.Add(closing_deficit >= min_closing - closing_inventory)
                        model.Add(closing_deficit >= 0)
                        objective += stockout_penalty * closing_deficit * 3
                
                # Maximum inventory
                for grade in grades:
                    for d in range(1, num_days + 1):
                        model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])
                
                # Capacity constraints
                for line in lines:
                    for d in range(num_days - buffer_days):
                        if line not in shutdown_periods or d not in shutdown_periods.get(line, []):
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
                
                # Min/Max run days
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
                        
                        # Minimum run days
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
                        
                        # Maximum run days
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
                                        if (current_grade != prev_grade and 
                                            current_grade not in allowed_next and 
                                            is_allowed_combination(current_grade, line)):
                                            
                                            prev_var = get_is_producing_var(prev_grade, line, d)
                                            current_var = get_is_producing_var(current_grade, line, d + 1)
                                            
                                            if prev_var is not None and current_var is not None:
                                                model.Add(prev_var + current_var <= 1)
                
                # Rerun allowed
                for grade in grades:
                    for line in allowed_lines[grade]:
                        grade_plant_key = (grade, line)
                        if not rerun_allowed.get(grade_plant_key, True):
                            starts = [is_start_vars[(grade, line, d)] for d in range(num_days) 
                                     if (grade, line, d) in is_start_vars]
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
                                
                                key1 = (grade1, line, d)
                                key2 = (grade2, line, d + 1)
                                if key1 in is_producing and key2 in is_producing:
                                    trans_var = model.NewBoolVar(f'trans_{line}_{d}_{grade1}_to_{grade2}')
                                    model.AddBoolAnd([is_producing[key1], is_producing[key2]]).OnlyEnforceIf(trans_var)
                                    model.Add(trans_var == 0).OnlyEnforceIf(is_producing[key1].Not())
                                    model.Add(trans_var == 0).OnlyEnforceIf(is_producing[key2].Not())
                                    objective += transition_penalty * trans_var
                        
                        for grade in grades:
                            if line in allowed_lines[grade]:
                                key1 = (grade, line, d)
                                key2 = (grade, line, d + 1)
                                if key1 in is_producing and key2 in is_producing:
                                    continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                                    model.AddBoolAnd([is_producing[key1], is_producing[key2]]).OnlyEnforceIf(continuity)
                                    objective += -continuity_bonus * continuity
                
                model.Minimize(objective)
                
                progress_bar.progress(50)
                status_container.info("⚡ Running optimization solver...")
                
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = time_limit_min * 60.0
                solver.parameters.num_search_workers = 8
                solver.parameters.random_seed = 42
                
                status = solver.Solve(model)
                
                progress_bar.progress(100)
                
                if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                    status_container.success("✅ Optimization completed successfully!")
                    
                    # Calculate results
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
                    
                    # Count transitions
                    total_transitions = 0
                    for line in lines:
                        last_grade = None
                        for d in range(num_days):
                            current_grade = None
                            for grade in grades:
                                key = (grade, line, d)
                                if key in is_producing and solver.Value(is_producing[key]) == 1:
                                    current_grade = grade
                                    break
                            
                            if current_grade is not None:
                                if last_grade is not None and current_grade != last_grade:
                                    total_transitions += 1
                                last_grade = current_grade
                    
                    # Store results
                    st.session_state.results = {
                        'solver': solver,
                        'production': production,
                        'inventory_vars': inventory_vars,
                        'stockout_vars': stockout_vars,
                        'is_producing': is_producing,
                        'grades': grades,
                        'lines': lines,
                        'dates': dates,
                        'formatted_dates': formatted_dates,
                        'num_days': num_days,
                        'buffer_days': buffer_days,
                        'production_totals': production_totals,
                        'grade_totals': grade_totals,
                        'plant_totals': plant_totals,
                        'stockout_totals': stockout_totals,
                        'total_transitions': total_transitions,
                        'objective_value': solver.ObjectiveValue(),
                        'allowed_lines': allowed_lines,
                        'min_inventory': min_inventory,
                        'max_inventory': max_inventory,
                        'shutdown_periods': shutdown_periods
                    }
                    
                    st.session_state.optimization_run = True
                    st.rerun()
                    
                else:
                    status_container.error("❌ Optimization failed to find a feasible solution")
                    st.error("No feasible solution found. Please check your constraints.")
                    
            except Exception as e:
                status_container.error(f"❌ Error during optimization: {str(e)}")
                st.error(f"Error: {str(e)}")
                import traceback
                st.error(traceback.format_exc())
        
        # Display results if optimization has been run
        if st.session_state.optimization_run and st.session_state.results:
            st.markdown("---")
            
            results = st.session_state.results
            solver = results['solver']
            grades = results['grades']
            lines = results['lines']
            dates = results['dates']
            formatted_dates = results['formatted_dates']
            num_days = results['num_days']
            buffer_days = results['buffer_days']
            
            # Key Metrics
            st.markdown("""
            <div class="metric-grid">
                <div class="metric-card-modern">
                    <div class="metric-label-modern">Objective Value</div>
                    <div class="metric-value-modern">{:,.0f}</div>
                </div>
                <div class="metric-card-modern">
                    <div class="metric-label-modern">Total Transitions</div>
                    <div class="metric-value-modern">{}</div>
                </div>
                <div class="metric-card-modern">
                    <div class="metric-label-modern">Total Stockouts</div>
                    <div class="metric-value-modern">{:,.0f} MT</div>
                </div>
                <div class="metric-card-modern">
                    <div class="metric-label-modern">Planning Horizon</div>
                    <div class="metric-value-modern">{} days</div>
                </div>
            </div>
            """.format(
                results['objective_value'],
                results['total_transitions'],
                sum(results['stockout_totals'].values()),
                num_days
            ), unsafe_allow_html=True)
            
            # Results tabs
            result_tab1, result_tab2, result_tab3 = st.tabs(["📅 Production Schedule", "📊 Summary", "📦 Inventory"])
            
            with result_tab1:
                sorted_grades = sorted(grades)
                base_colors = px.colors.qualitative.Vivid
                grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}
                
                st.subheader("Production Visualization")
                
                for line in lines:
                    st.markdown(f"### Production Schedule - {line}")
                    
                    gantt_data = []
                    for d in range(num_days):
                        date = dates[d]
                        for grade in sorted_grades:
                            key = (grade, line, d)
                            if key in results['is_producing'] and solver.Value(results['is_producing'][key]) == 1:
                                gantt_data.append({
                                    "Grade": grade,
                                    "Start": date,
                                    "Finish": date + timedelta(days=1),
                                    "Line": line
                                })
                    
                    if not gantt_data:
                        st.info(f"No production data available for {line}.")
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
                        title=f"Production Schedule - {line}"
                    )
                    
                    # Add shutdown visualization
                    if line in results['shutdown_periods'] and results['shutdown_periods'][line]:
                        shutdown_days = results['shutdown_periods'][line]
                        start_shutdown = dates[shutdown_days[0]]
                        end_shutdown = dates[shutdown_days[-1]] + timedelta(days=1)
                        
                        fig.add_vrect(
                            x0=start_shutdown,
                            x1=end_shutdown,
                            fillcolor="red",
                            opacity=0.2,
                            layer="below",
                            line_width=0,
                            annotation_text="Shutdown",
                            annotation_position="top left"
                        )
                    
                    fig.update_yaxes(autorange="reversed", title=None)
                    fig.update_xaxes(title="Date", tickformat="%d-%b")
                    fig.update_layout(height=350, showlegend=True)
                    
                    st.plotly_chart(fig, use_container_width=True)
                
                st.subheader("Production Schedule by Line")
                
                def color_grade(val):
                    if val in grade_color_map:
                        color = grade_color_map[val]
                        return f'background-color: {color}; color: white; font-weight: bold; text-align: center;'
                    return ''
                
                for line in lines:
                    st.markdown(f"### 🏭 {line}")
                    
                    schedule_data = []
                    current_grade = None
                    start_day = None
                    
                    for d in range(num_days):
                        date = dates[d]
                        grade_today = None
                        
                        for grade in sorted_grades:
                            key = (grade, line, d)
                            if key in results['is_producing'] and solver.Value(results['is_producing'][key]) == 1:
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
            
            with result_tab2:
                st.subheader("Production Summary")
                
                total_prod_data = []
                for grade in grades:
                    row = {'Grade': grade}
                    for line in lines:
                        row[line] = results['production_totals'][grade][line]
                    row['Total Produced'] = results['grade_totals'][grade]
                    row['Total Stockout'] = results['stockout_totals'][grade]
                    total_prod_data.append(row)
                
                totals_row = {'Grade': 'Total'}
                for line in lines:
                    totals_row[line] = results['plant_totals'][line]
                totals_row['Total Produced'] = sum(results['plant_totals'].values())
                totals_row['Total Stockout'] = sum(results['stockout_totals'].values())
                total_prod_data.append(totals_row)
                
                total_prod_df = pd.DataFrame(total_prod_data)
                st.dataframe(total_prod_df, use_container_width=True)
            
            with result_tab3:
                st.subheader("Inventory Levels")
                
                last_actual_day = num_days - buffer_days - 1
                
                for grade in sorted_grades:
                    inventory_values = [solver.Value(results['inventory_vars'][(grade, d)]) for d in range(num_days)]
                    
                    start_val = inventory_values[0]
                    end_val = inventory_values[last_actual_day]
                    highest_val = max(inventory_values[: last_actual_day + 1])
                    lowest_val = min(inventory_values[: last_actual_day + 1])
                    
                    fig = go.Figure()
                    
                    fig.add_trace(go.Scatter(
                        x=dates,
                        y=inventory_values,
                        mode="lines+markers",
                        name=grade,
                        line=dict(color=grade_color_map[grade], width=3),
                        marker=dict(size=6)
                    ))
                    
                    # Add shutdown periods
                    for line in results['allowed_lines'][grade]:
                        if line in results['shutdown_periods'] and results['shutdown_periods'][line]:
                            shutdown_days = results['shutdown_periods'][line]
                            start_shutdown = dates[shutdown_days[0]]
                            end_shutdown = dates[shutdown_days[-1]]
                            
                            fig.add_vrect(
                                x0=start_shutdown,
                                x1=end_shutdown + timedelta(days=1),
                                fillcolor="red",
                                opacity=0.1,
                                layer="below",
                                line_width=0
                            )
                    
                    fig.add_hline(y=results['min_inventory'][grade], line=dict(color="red", width=2, dash="dash"))
                    fig.add_hline(y=results['max_inventory'][grade], line=dict(color="green", width=2, dash="dash"))
                    
                    fig.update_layout(
                        title=f"Inventory Level - {grade}",
                        xaxis_title="Date",
                        yaxis_title="Inventory Volume (MT)",
                        height=420
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
