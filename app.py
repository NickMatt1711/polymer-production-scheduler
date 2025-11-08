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

# Set page configuration
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #1f77b4;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

# Main title
st.markdown('<div class="main-header">🏭 Polymer Production Scheduler</div>', unsafe_allow_html=True)

# Solution callback class
class SolutionCallback(cp_model.CpSolverSolutionCallback):
    def __init__(self, production, inventory, stockout, is_producing, grades, lines, dates, num_days):
        cp_model.CpSolverSolutionCallback.__init__(self)
        self.production = production
        self.inventory = inventory
        self.stockout = stockout
        self.is_producing = is_producing
        self.grades = grades
        self.lines = lines
        self.dates = dates
        self.num_days = num_days
        self.solutions = []
        self.solution_times = []
        self.start_time = time.time()

    def on_solution_callback(self):
        current_time = time.time() - self.start_time
        self.solution_times.append(current_time)
        current_obj = self.ObjectiveValue()

        # Store the current solution
        solution = {
            'objective': current_obj,
            'time': current_time,
            'production': {},
            'inventory': {},
            'stockout': {},
            'is_producing': {}
        }

        # Extract production values
        for grade in self.grades:
            solution['production'][grade] = {}
            for line in self.lines:
                for d in range(self.num_days):
                    key = (grade, line, d)
                    if key in self.production:
                        value = self.Value(self.production[key])
                        if value > 0:
                            date_key = self.dates[d]
                            if date_key not in solution['production'][grade]:
                                solution['production'][grade][date_key] = 0
                            solution['production'][grade][date_key] += value

        # Extract inventory values
        for grade in self.grades:
            solution['inventory'][grade] = {}
            for d in range(self.num_days + 1):
                key = (grade, d)
                if key in self.inventory:
                    if d < self.num_days:
                        solution['inventory'][grade][self.dates[d] if d > 0 else 'initial'] = self.Value(self.inventory[key])
                    else:
                        solution['inventory'][grade]['final'] = self.Value(self.inventory[key])

        # Extract stockout values
        for grade in self.grades:
            solution['stockout'][grade] = {}
            for d in range(self.num_days):
                key = (grade, d)
                if key in self.stockout:
                    value = self.Value(self.stockout[key])
                    if value > 0:
                        solution['stockout'][grade][self.dates[d]] = value

        # Extract production schedule
        for line in self.lines:
            solution['is_producing'][line] = {}
            for d in range(self.num_days):
                date_key = self.dates[d]
                solution['is_producing'][line][date_key] = None
                for grade in self.grades:
                    key = (grade, line, d)
                    if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                        solution['is_producing'][line][date_key] = grade
                        break

        self.solutions.append(solution)

        # Calculate transitions
        transition_count_per_line = {line: 0 for line in self.lines}
        total_transitions = 0

        for line in self.lines:
            last_grade = None
            for date in self.dates:
                current_grade = solution['is_producing'][line].get(date)
                if current_grade is not None and last_grade is not None:
                    if current_grade != last_grade:
                        transition_count_per_line[line] += 1
                        total_transitions += 1
                if current_grade is not None:
                    last_grade = current_grade

        solution['transitions'] = {
            'per_line': transition_count_per_line,
            'total': total_transitions
        }

    def num_solutions(self):
        return len(self.solutions)

# Sidebar for file upload and parameters
with st.sidebar:
    st.header("📁 Data Input")
    
    uploaded_file = st.file_uploader(
        "Upload Excel File", 
        type=["xlsx"],
        help="Upload an Excel file with Plant, Inventory, and Demand sheets"
    )
    
    if uploaded_file:
        st.success("✅ File uploaded successfully!")
        
        # Debug information
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 File Debug Info")
        st.sidebar.write(f"File name: {uploaded_file.name}")
        st.sidebar.write(f"File size: {uploaded_file.size} bytes")
        
        # Check available sheets
        try:
            uploaded_file.seek(0)
            excel_file = io.BytesIO(uploaded_file.read())
            available_sheets = pd.ExcelFile(excel_file).sheet_names
            st.sidebar.write("Available sheets:", available_sheets)
            uploaded_file.seek(0)  # Reset for future reads
        except Exception as e:
            st.sidebar.error(f"Error reading sheet names: {e}")
        
        st.header("⚙️ Optimization Parameters")
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
        
        stockout_penalty = st.number_input(
            "Stockout penalty",
            min_value=1,
            value=10,
            help="Penalty weight for stockouts in objective function"
        )
        
        transition_penalty = st.number_input(
            "Transition penalty", 
            min_value=1,
            value=10,
            help="Penalty weight for production line transitions"
        )
        
        continuity_bonus = st.number_input(
            "Continuity bonus",
            min_value=0,
            value=1,
            help="Bonus for continuing the same grade (negative penalty)"
        )

# Main content area
if uploaded_file:
    try:
        # Load data - reset the file pointer first
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Show data preview
        st.markdown('<div class="section-header">📊 Data Preview</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            try:
                plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                st.subheader("Plant Data")
                st.dataframe(plant_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Plant sheet: {e}")
                st.stop()
        
        with col2:
            try:
                excel_file.seek(0)
                inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                st.subheader("Inventory Data")
                st.dataframe(inventory_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Inventory sheet: {e}")
                st.stop()
        
        with col3:
            try:
                excel_file.seek(0)
                demand_df = pd.read_excel(excel_file, sheet_name='Demand')
                st.subheader("Demand Data")
                st.dataframe(demand_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Demand sheet: {e}")
                st.stop()
        
        # Reset file pointer for transition matrices
        excel_file.seek(0)
        
        # Load transition matrices
        transition_dfs = {}
        for i in range(len(plant_df)):
            plant_name = plant_df['Plant'].iloc[i]
            sheet_name = f'Transition_{plant_name}'
            try:
                excel_file.seek(0)
                transition_dfs[plant_name] = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=0)
                st.info(f"✅ Loaded transition matrix for {plant_name}")
            except Exception as e:
                st.info(f"ℹ️ No transition matrix found for {plant_name}. Assuming no transition constraints.")
                transition_dfs[plant_name] = None
        
        # Run optimization button
        st.markdown('<div class="section-header">🚀 Optimization</div>', unsafe_allow_html=True)
        
        if st.button("Run Production Optimization", type="primary", use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_placeholder = st.empty()
            
            # Initialize session state for solutions
            if 'solutions' not in st.session_state:
                st.session_state.solutions = []
            if 'best_solution' not in st.session_state:
                st.session_state.best_solution = None
            
            # Data preprocessing
            status_text.markdown('<div class="info-box">🔄 Preprocessing data...</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            
            # Data preprocessing
            num_lines = len(plant_df)
            lines = list(plant_df['Plant'])
            capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
            grades = list(inventory_df['Grade Name'])
            initial_inventory = {row['Grade Name']: row['Opening Inventory'] for index, row in inventory_df.iterrows()}
            min_inventory = {row['Grade Name']: row['Min. Inventory'] for index, row in inventory_df.iterrows()}
            max_inventory = {row['Grade Name']: row['Max. Inventory'] for index, row in inventory_df.iterrows()}
            min_run_days = {row['Grade Name']: int(row['Min. Run Days']) if pd.notna(row['Min. Run Days']) else 1 for index, row in inventory_df.iterrows()}
            force_start_date = {row['Grade Name']: pd.to_datetime(row['Force Start Date']).date() if pd.notna(row['Force Start Date']) else None for index, row in inventory_df.iterrows()}
            allowed_lines = {
                row['Grade Name']: [x.strip() for x in str(row['Plant']).split(',')] if pd.notna(row['Plant']) else lines
                for index, row in inventory_df.iterrows()
            }
            
            rerun_allowed = {}
            for index, row in inventory_df.iterrows():
                rerun_val = row['Rerun Allowed']
                if isinstance(rerun_val, str) and rerun_val.strip().lower() == 'yes':
                    rerun_allowed[row['Grade Name']] = True
                else:
                    rerun_allowed[row['Grade Name']] = False
            
            max_run_days = {row['Grade Name']: int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999 for index, row in inventory_df.iterrows()}
            min_closing_inventory = {row['Grade Name']: row['Min. Closing Inventory'] if pd.notna(row['Min. Closing Inventory']) else 0 for index, row in inventory_df.iterrows()}
            
            # Process demand data
            demand_data = {}
            dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            num_days = len(dates)
            last_date = dates[-1]
            for i in range(1, buffer_days + 1):
                dates.append(last_date + timedelta(days=i))
            num_days = len(dates)
            
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
            status_text.markdown('<div class="info-box">🔧 Building optimization model...</div>', unsafe_allow_html=True)
            
            # --- Create the CP-SAT Model ---
            model = cp_model.CpModel()

            # Decision Variables - FIXED VERSION
            is_producing = {}
            production = {}
            for grade in grades:
                allowed = allowed_lines[grade]
                for line in allowed:
                    for d in range(num_days):
                        is_producing[(grade, line, d)] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                        
                        # Create production variable based on whether it's a buffer day or not
                        if d < num_days - buffer_days:
                            # For non-buffer days, production equals capacity when producing
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value == capacities[line]).OnlyEnforceIf(is_producing[(grade, line, d)])
                            model.Add(production_value == 0).OnlyEnforceIf(is_producing[(grade, line, d)].Not())
                        else:
                            # For buffer days, production can be up to capacity
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value <= capacities[line] * is_producing[(grade, line, d)])
                        
                        production[(grade, line, d)] = production_value

            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')

            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')

            # [Rest of your optimization code remains the same...]
            # Continue with the rest of your constraints and solving logic

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.info("Please make sure your Excel file has the required sheets: 'Plant', 'Inventory', and 'Demand'")

else:
    # Welcome message when no file is uploaded
    st.markdown("""
    <div class="info-box">
    <h3>Welcome to the Polymer Production Scheduler! 🏭</h3>
    <p>This application helps optimize your polymer production schedule by:</p>
    <ul>
        <li>📊 Analyzing plant capacities and inventory constraints</li>
        <li>⚡ Optimizing production sequences to minimize transitions</li>
        <li>📈 Balancing inventory levels and meeting demand</li>
        <li>💾 Generating detailed production schedules and reports</li>
    </ul>
    <p><strong>To get started:</strong></p>
    <ol>
        <li>Upload an Excel file with the required sheets (Plant, Inventory, Demand)</li>
        <li>Configure optimization parameters in the sidebar</li>
        <li>Run the optimization and view results</li>
        <li>Download the production schedule report</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
