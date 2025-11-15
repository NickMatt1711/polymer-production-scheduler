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
            st.warning("Sample template file not found.")
            return None
    except Exception as e:
        st.warning(f"Could not load sample template: {e}")
        return None

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
                    st.warning(f"Invalid shutdown period for {plant}")
                    shutdown_periods[plant] = []
                    continue
                
                shutdown_days = []
                for d, date in enumerate(dates):
                    if start_date <= date <= end_date:
                        shutdown_days.append(d)
                
                shutdown_periods[plant] = shutdown_days
                    
            except Exception as e:
                st.warning(f"Invalid shutdown dates for {plant}: {e}")
                shutdown_periods[plant] = []
        else:
            shutdown_periods[plant] = []
    
    return shutdown_periods

st.set_page_config(
    page_title="Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'optimization_complete' not in st.session_state:
    st.session_state.optimization_complete = False

# Clean, minimal CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', system-ui, -apple-system, sans-serif;
    }
    
    .main {
        background-color: #FFFFFF;
    }
    
    /* Header */
    .clean-header {
        background: #FFFFFF;
        padding: 2rem 0 1rem 0;
        border-bottom: 1px solid #E5E7EB;
        margin-bottom: 2rem;
    }
    
    .clean-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1F2937;
        margin: 0;
    }
    
    .clean-subtitle {
        font-size: 0.95rem;
        color: #6B7280;
        margin-top: 0.25rem;
    }
    
    /* Cards */
    .clean-card {
        background: #F8F9FA;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 1.5rem;
        margin-bottom: 1rem;
    }
    
    .clean-card-white {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Metrics */
    .metric-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin: 1.5rem 0;
    }
    
    .metric-box {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 1.25rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .metric-value {
        font-size: 1.75rem;
        font-weight: 700;
        color: #2563EB;
        margin: 0.5rem 0;
    }
    
    .metric-label {
        font-size: 0.875rem;
        color: #6B7280;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Buttons */
    .stButton>button {
        background: #2563EB;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.625rem 1.5rem;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.2s;
    }
    
    .stButton>button:hover {
        background: #1D4ED8;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2);
    }
    
    /* Section headers */
    .section-title {
        font-size: 1.25rem;
        font-weight: 600;
        color: #1F2937;
        margin: 2rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #E5E7EB;
    }
    
    /* Status indicators */
    .status-success {
        background: #D1FAE5;
        border: 1px solid #059669;
        color: #047857;
        padding: 0.75rem 1rem;
        border-radius: 8px;
        font-weight: 500;
        margin: 1rem 0;
    }
    
    .status-info {
        background: #DBEAFE;
        border: 1px solid #2563EB;
        color: #1E40AF;
        padding: 0.75rem 1rem;
        border-radius: 8px;
        font-weight: 500;
        margin: 1rem 0;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: #F8F9FA;
        padding: 0.5rem;
        border-radius: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 6px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        color: #6B7280;
    }
    
    .stTabs [aria-selected="true"] {
        background: #2563EB;
        color: white;
        border-color: #2563EB;
    }
    
    /* Progress bar */
    .stProgress > div > div > div > div {
        background: #2563EB;
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        background: #F8F9FA;
        border: 2px dashed #E5E7EB;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Dataframes */
    .dataframe {
        font-size: 0.875rem;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: #F8F9FA;
        border-radius: 8px;
        font-weight: 600;
    }
    
    /* Clean spacing */
    .element-container {
        margin-bottom: 0.5rem;
    }
    
    /* Number inputs */
    .stNumberInput > div > div > input {
        border-radius: 8px;
        border: 1px solid #E5E7EB;
    }
    
    /* Hide sidebar */
    [data-testid="stSidebar"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="clean-header">
    <div class="clean-title">🏭 Production Scheduler</div>
    <div class="clean-subtitle">Multi-plant optimization with intelligent scheduling</div>
</div>
""", unsafe_allow_html=True)

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

# Check if file is uploaded
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

# Landing page with upload
if st.session_state.uploaded_file is None:
    st.markdown("""
    <div class="clean-card-white">
        <h3 style="margin-top: 0; color: #1F2937;">Getting Started</h3>
        <p style="color: #6B7280;">Upload your production data to begin optimization</p>
        
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-top: 1.5rem;">
            <div>
                <h4 style="color: #1F2937; font-size: 1rem; margin-bottom: 0.5rem;">📊 What You'll Need</h4>
                <ul style="color: #6B7280; font-size: 0.9rem; line-height: 1.6;">
                    <li>Plant capacity data</li>
                    <li>Inventory levels & constraints</li>
                    <li>Demand forecasts</li>
                    <li>Transition rules (optional)</li>
                </ul>
            </div>
            <div>
                <h4 style="color: #1F2937; font-size: 1rem; margin-bottom: 0.5rem;">⚡ What You'll Get</h4>
                <ul style="color: #6B7280; font-size: 0.9rem; line-height: 1.6;">
                    <li>Optimized production schedule</li>
                    <li>Minimized transitions</li>
                    <li>Inventory management</li>
                    <li>Stockout prevention</li>
                </ul>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    st.success(f"File uploaded: {st.session_state.uploaded_file.name}")
    
    st.markdown("---")
    
    sample_workbook = get_sample_workbook()
    
    if sample_workbook:
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown("""
            <div class="clean-card">
                <h4 style="margin-top: 0; color: #1F2937;">📥 Download Sample Template</h4>
                <p style="color: #6B7280; font-size: 0.9rem;">
                    Start with our pre-formatted template that includes all required sheets and sample data.
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.download_button(
                label="Download Template",
                data=sample_workbook,
                file_name="production_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    st.markdown("---")
    st.markdown('<div class="section-title">📤 Upload Your Data</div>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Select Excel file with Plant, Inventory, and Demand sheets", 
        type=["xlsx"],
        help="Upload Excel file containing production data"
    )
    
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file
        st.rerun()

else:
    # File is uploaded, show main interface
    uploaded_file = st.session_state.uploaded_file
    
    try:
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Data preview section
        st.markdown('<div class="section-title">📊 Data Overview</div>', unsafe_allow_html=True)
        
        tab1, tab2, tab3 = st.tabs(["Plant", "Inventory", "Demand"])
        
        with tab1:
            plant_df = pd.read_excel(excel_file, sheet_name='Plant')
            plant_display_df = plant_df.copy()
            
            # Format dates for display
            for col_idx in [4, 5]:
                if col_idx < len(plant_display_df.columns):
                    col = plant_display_df.columns[col_idx]
                    if pd.api.types.is_datetime64_any_dtype(plant_display_df[col]):
                        plant_display_df[col] = plant_display_df[col].dt.strftime('%d-%b-%y')
            
            st.dataframe(plant_display_df, use_container_width=True, height=200)
        
        with tab2:
            excel_file.seek(0)
            inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
            inventory_display_df = inventory_df.copy()
            
            if len(inventory_display_df.columns) > 7:
                col = inventory_display_df.columns[7]
                if pd.api.types.is_datetime64_any_dtype(inventory_display_df[col]):
                    inventory_display_df[col] = inventory_display_df[col].dt.strftime('%d-%b-%y')
            
            st.dataframe(inventory_display_df, use_container_width=True, height=200)
        
        with tab3:
            excel_file.seek(0)
            demand_df = pd.read_excel(excel_file, sheet_name='Demand')
            demand_display_df = demand_df.copy()
            
            date_column = demand_display_df.columns[0]
            if pd.api.types.is_datetime64_any_dtype(demand_display_df[date_column]):
                demand_display_df[date_column] = demand_display_df[date_column].dt.strftime('%d-%b-%y')
            
            st.dataframe(demand_display_df, use_container_width=True, height=200)
        
        # Shutdown info
        shutdown_info = []
        for index, row in plant_df.iterrows():
            plant = row['Plant']
            shutdown_start = row.get('Shutdown Start Date')
            shutdown_end = row.get('Shutdown End Date')
            
            if pd.notna(shutdown_start) and pd.notna(shutdown_end):
                try:
                    start_date = pd.to_datetime(shutdown_start).date()
                    end_date = pd.to_datetime(shutdown_end).date()
                    duration = (end_date - start_date).days + 1
                    
                    if start_date <= end_date:
                        shutdown_info.append(f"**{plant}**: {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')} ({duration} days)")
                except:
                    pass
        
        if shutdown_info:
            st.markdown('<div class="status-info">🔧 Scheduled shutdowns:<br>' + '<br>'.join(shutdown_info) + '</div>', unsafe_allow_html=True)
        
        # Load transition matrices
        excel_file.seek(0)
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
        
        # Parameters section
        st.markdown("---")
        st.markdown('<div class="section-title">⚙️ Optimization Parameters</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            time_limit_min = st.number_input(
                "Time limit (minutes)",
                min_value=1,
                max_value=120,
                value=10,
                help="Maximum time to run the optimization"
            )
        
        with col2:
            buffer_days = st.number_input(
                "Buffer days",
                min_value=0,
                max_value=7,
                value=3,
                help="Additional days for planning buffer"
            )
        
        with col3:
            st.write("")  # Spacer
        
        with st.expander("Advanced Parameters"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                stockout_penalty = st.number_input(
                    "Stockout penalty",
                    min_value=1,
                    value=10,
                    help="Penalty weight for stockouts in objective function"
                )
            
            with col2:
                transition_penalty = st.number_input(
                    "Transition penalty", 
                    min_value=1,
                    value=10,
                    help="Penalty weight for production line transitions"
                )
            
            with col3:
                continuity_bonus = st.number_input(
                    "Continuity bonus",
                    min_value=0,
                    value=1,
                    help="Bonus for continuing the same grade"
                )
        
        # Optimization button
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            run_optimization = st.button("🎯 Run Optimization", type="primary", use_container_width=True)
        
        if run_optimization:
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.markdown('<div class="status-info">📊 Processing data...</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            time.sleep(1)

            try:
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
            
            except Exception as e:
                st.error(f"Error in data preprocessing: {str(e)}")
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
            status_text.markdown('<div class="status-info">🔧 Building model...</div>', unsafe_allow_html=True)
            time.sleep(1)
            
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
                return production[key] if key in production else 0
            
            def get_is_producing_var(grade, line, d):
                key = (grade, line, d)
                return is_producing[key] if key in is_producing else None
            
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
                        min_inv_value = int(min_inventory[grade])
                        inventory_tomorrow = inventory_vars[(grade, d + 1)]
                        
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        model.Add(deficit >= min_inv_value - inventory_tomorrow)
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
            
            for grade in grades:
                for d in range(1, num_days + 1):
                    model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])
            
            # Capacity constraints
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

            # Rerun allowed constraints
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    if not rerun_allowed.get(grade_plant_key, True):
                        starts = [is_start_vars[(grade, line, d)] for d in range(num_days) 
                                 if (grade, line, d) in is_start_vars]
                        if starts:
                            model.Add(sum(starts) <= 1)

            # Objective function
            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

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
            status_text.markdown('<div class="status-info">⚡ Running solver...</div>', unsafe_allow_html=True)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit_min * 60.0
            solver.parameters.num_search_workers = 8
            solver.parameters.random_seed = 42
            
            solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

            status = solver.Solve(model, solution_callback)
            
            progress_bar.progress(100)
            
            if status == cp_model.OPTIMAL:
                status_text.markdown('<div class="status-success">✓ Optimal solution found</div>', unsafe_allow_html=True)
            elif status == cp_model.FEASIBLE:
                status_text.markdown('<div class="status-success">✓ Feasible solution found</div>', unsafe_allow_html=True)
            else:
                status_text.markdown('<div class="status-info">⚠ No solution found</div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown('<div class="section-title">📊 Results</div>', unsafe_allow_html=True)

            if solution_callback.num_solutions() > 0:
                best_solution = solution_callback.solutions[-1]
                st.session_state.optimization_complete = True

                # Metrics
                col1, col2, col3, col4 = st.columns(4)
                
                total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                
                with col1:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-label">Objective</div>
                            <div class="metric-value">{best_solution['objective']:,.0f}</div>
                            <div class="metric-label">Lower is better</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-label">Transitions</div>
                            <div class="metric-value">{best_solution['transitions']['total']}</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-label">Stockouts</div>
                            <div class="metric-value">{total_stockouts:,.0f}</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-label">Horizon</div>
                            <div class="metric-value">{num_days}-{buffer_days}days</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Results tabs
                tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Summary", "📦 Inventory"])
                
                with tab1:
                    sorted_grades = sorted(grades)
                    base_colors = px.colors.qualitative.Vivid
                    grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}
                    
                    cmap = colormaps.get_cmap('tab20')
                    grade_colors = {}
                    for idx, grade in enumerate(grades):
                        grade_colors[grade] = cmap(idx % 20)

                    st.subheader("Production Visualization")
                    
                    for line in lines:
                        st.markdown(f"### Production Schedule - {line}")
                    
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
                    
                        # Add shutdown period visualization
                        if line in shutdown_periods and shutdown_periods[line]:
                            shutdown_days = shutdown_periods[line]
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
                                annotation_position="top left",
                                annotation_font_size=14,
                                annotation_font_color="red"
                            )
                    
                        fig.update_yaxes(
                            autorange="reversed",
                            title=None,
                            showgrid=True,
                            gridcolor="lightgray",
                            gridwidth=1
                        )
                    
                        fig.update_xaxes(
                            title="Date",
                            showgrid=True,
                            gridcolor="lightgray",
                            gridwidth=1,
                            tickvals=dates,
                            tickformat="%d-%b",
                            dtick="D1"
                        )
                    
                        fig.update_layout(
                            height=350,
                            bargap=0.2,
                            showlegend=True,
                            legend_title_text="Grade",
                            legend=dict(
                                traceorder="normal",
                                orientation="v",
                                yanchor="middle",
                                y=0.5,
                                xanchor="left",
                                x=1.02,
                                bgcolor="rgba(255,255,255,0)",
                                bordercolor="lightgray",
                                borderwidth=0
                            ),
                            xaxis=dict(showline=True, showticklabels=True),
                            yaxis=dict(showline=True),
                            margin=dict(l=60, r=160, t=60, b=60),
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            font=dict(size=12),
                        )
                    
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

                with tab2:
                    st.subheader("Production Summary")
                    
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
                    
                    totals_row = {'Grade': 'Total'}
                    for line in lines:
                        totals_row[line] = plant_totals[line]
                    totals_row['Total Produced'] = sum(plant_totals.values())
                    totals_row['Total Stockout'] = sum(stockout_totals.values())
                    total_prod_data.append(totals_row)
                    
                    total_prod_df = pd.DataFrame(total_prod_data)
                    st.dataframe(total_prod_df, use_container_width=True)

                with tab3:
                    st.subheader("Inventory Levels")
                    
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
                    
                        # Add shutdown periods for plants that produce this grade
                        shutdown_added = False
                        for line in allowed_lines[grade]:
                            if line in shutdown_periods and shutdown_periods[line]:
                                shutdown_days = shutdown_periods[line]
                                start_shutdown = dates[shutdown_days[0]]
                                end_shutdown = dates[shutdown_days[-1]]
                                
                                # Add vertical shaded regions for shutdown periods
                                fig.add_vrect(
                                    x0=start_shutdown,
                                    x1=end_shutdown + timedelta(days=1),
                                    fillcolor="red",
                                    opacity=0.1,
                                    layer="below",
                                    line_width=0,
                                    annotation_text=f"Shutdown: {line}" if not shutdown_added else "",
                                    annotation_position="top left",
                                    annotation_font_size=14,
                                    annotation_font_color="red"
                                )
                                shutdown_added = True
                    
                        fig.add_hline(
                            y=min_inventory[grade],
                            line=dict(color="red", width=2, dash="dash"),
                            annotation_text=f"Min: {min_inventory[grade]:,.0f}",
                            annotation_position="top left",
                            annotation_font_color="red"
                        )
                        fig.add_hline(
                            y=max_inventory[grade],
                            line=dict(color="green", width=2, dash="dash"),
                            annotation_text=f"Max: {max_inventory[grade]:,.0f}",
                            annotation_position="bottom left",
                            annotation_font_color="green"
                        )
                    
                        annotations = [
                            dict(
                                x=start_x, y=start_val,
                                text=f"Start: {start_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=-40, ay=30,
                                font=dict(color="black", size=11),
                                bgcolor="white", bordercolor="gray"
                            ),
                            dict(
                                x=end_x, y=end_val,
                                text=f"End: {end_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=40, ay=30,
                                font=dict(color="black", size=11),
                                bgcolor="white", bordercolor="gray"
                            ),
                            dict(
                                x=highest_x, y=highest_val,
                                text=f"High: {highest_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=0, ay=-40,
                                font=dict(color="darkgreen", size=11),
                                bgcolor="white", bordercolor="gray"
                            ),
                            dict(
                                x=lowest_x, y=lowest_val,
                                text=f"Low: {lowest_val:.0f}",
                                showarrow=True, arrowhead=2,
                                ax=0, ay=40,
                                font=dict(color="firebrick", size=11),
                                bgcolor="white", bordercolor="gray"
                            )
                        ]
                    
                        fig.update_layout(
                            title=f"Inventory Level - {grade}",
                            xaxis=dict(
                                title="Date",
                                showgrid=True,
                                gridcolor="lightgray",
                                tickvals=dates,
                                tickformat="%d-%b",
                                dtick="D1"
                            ),
                            yaxis=dict(
                                title="Inventory Volume (MT)",
                                showgrid=True,
                                gridcolor="lightgray"
                            ),
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(l=60, r=80, t=80, b=60),
                            font=dict(size=12),
                            height=420,
                            showlegend=False
                        )
                        
                        # Add annotations one by one with explicit parameters
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
                                borderwidth=1,
                                borderpad=4,
                                opacity=0.9
                            )
                    
                        st.plotly_chart(fig, use_container_width=True)

            else:
                st.error("No solutions found during optimization. Please check your constraints and data.")

    except Exception as e:
        st.error(f"Error: {str(e)}")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #9CA3AF; font-size: 0.875rem;'>Production Scheduler • Built with Streamlit</div>",
    unsafe_allow_html=True
)
