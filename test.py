# test.py
# Fintech-style UI redesign (Stripe-like) — UI only. Solver & visualizations unchanged.
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

# --------------------------
# Helper functions (unchanged behavior)
# --------------------------
def get_sample_workbook():
    """Retrieve the sample workbook from the same directory as app.py"""
    try:
        # Get the directory where app.py is located
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
        
        # Check if both start and end dates are provided
        if pd.notna(shutdown_start) and pd.notna(shutdown_end):
            try:
                start_date = pd.to_datetime(shutdown_start).date()
                end_date = pd.to_datetime(shutdown_end).date()
                
                # Validate date range
                if start_date > end_date:
                    st.warning(f"⚠️ Shutdown start date after end date for {plant}. Ignoring shutdown.")
                    shutdown_periods[plant] = []
                    continue
                
                # Find day indices for shutdown period
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

# ---------------------------
# Page config & session state
# ---------------------------
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# session state keys used by the app (initialize if missing)
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0
if 'solutions' not in st.session_state:
    st.session_state.solutions = []
if 'best_solution' not in st.session_state:
    st.session_state.best_solution = None

# ---------------------------
# Stripe-style CSS (Fintech dashboard)
# - Minimal, structured, ample whitespace
# - Neutral surfaces, subtle borders, crisp typography
# ---------------------------
st.markdown("""
<style>
/* Design tokens */
:root {
  --bg: #f7fafc;
  --surface: #ffffff;
  --muted: #68707d;
  --text: #0b1220;
  --primary: #2563eb; /* blue-600 */
  --accent: #6d28d9;  /* purple-700 */
  --danger: #ef4444;
  --success: #10b981;
  --border: rgba(11,18,32,0.06);
}

/* App background & base font */
.reportview-container, .main .block-container {
  background: var(--bg);
  color: var(--text);
  font-family: Inter, ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial;
}

/* Header: left aligned brand, small action area on right */
.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 18px 22px;
  background: linear-gradient(90deg, rgba(255,255,255,0.9), rgba(255,255,255,0.9));
  border-bottom: 1px solid var(--border);
  margin-bottom: 18px;
}
.brand {
  display: flex;
  align-items: center;
  gap: 12px;
}
.logo {
  background: linear-gradient(135deg, var(--primary), var(--accent));
  color: white;
  height: 42px;
  width: 42px;
  border-radius: 8px;
  display:flex;
  align-items:center;
  justify-content:center;
  font-weight:800;
  box-shadow: 0 6px 18px rgba(37,99,235,0.08);
  font-size:18px;
}
.title {
  font-size: 18px;
  font-weight: 700;
  letter-spacing: -0.2px;
}
.subtitle {
  font-size: 12px;
  color: var(--muted);
  margin-top: 2px;
}

/* Action area */
.header-actions {
  display:flex;
  align-items:center;
  gap: 10px;
}
.kpi {
  text-align: right;
  font-weight:700;
  font-size: 14px;
}
.kpi-sub { font-size:12px; color:var(--muted); font-weight:600; }

/* Sidebar - full height column look */
[data-testid="stSidebar"] {
  background: var(--surface);
  padding: 18px 16px 24px 16px;
  border-right: 1px solid var(--border);
  min-width: 320px;
  box-shadow: none;
}

/* Sidebar groups */
.sidebar-group {
  border: 1px solid var(--border);
  background: linear-gradient(180deg, #ffffff, #ffffff);
  padding: 14px;
  border-radius: 10px;
  margin-bottom: 12px;
}

/* Buttons */
.stButton>button {
  background: var(--primary);
  color: white;
  border-radius: 8px;
  padding: 0.6rem 12px;
  border: none;
  font-weight:700;
  box-shadow: 0 6px 18px rgba(37,99,235,0.07);
}
.stButton>button:hover { transform: translateY(-2px); }

/* Cards and sections */
.panel {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 16px;
  box-shadow: 0 6px 20px rgba(11,18,32,0.02);
  margin-bottom: 16px;
}

/* Section titles */
.h2 {
  font-size: 15px;
  font-weight: 700;
  margin-bottom: 10px;
  color: var(--text);
}

/* Metric tiles (inline) */
.metrics {
  display:flex;
  gap:12px;
  align-items:center;
  margin-bottom: 12px;
}
.metric-tile {
  background: linear-gradient(180deg,#fff,#fff);
  border: 1px solid var(--border);
  padding: 10px 12px;
  border-radius: 8px;
  min-width: 140px;
}
.metric-label { color: var(--muted); font-size: 12px; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; }
.metric-value { font-size:18px; font-weight:800; margin-top:6px; }

/* DataFrames look */
.dataframe { border-radius: 8px; overflow: hidden; border: 1px solid var(--border); }
tbody tr:nth-child(odd) td { background: #ffffff; }
tbody tr:nth-child(even) td { background: #fbfcfe; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
  gap: 8px;
  background-color: transparent;
  padding: 6px;
}
.stTabs [data-baseweb="tab"] {
  background-color: #fff;
  border-radius: 999px;
  padding: 8px 14px;
  font-weight:700;
  border: 1px solid var(--border);
}
.stTabs [aria-selected="true"] {
  background: linear-gradient(90deg, var(--primary), var(--accent));
  color: #fff !important;
  box-shadow: 0 6px 18px rgba(37,99,235,0.08);
}

/* Footer */
.app-footer { margin-top: 18px; text-align:center; color:var(--muted); font-size:13px; }

/* Responsive */
@media (max-width: 900px) {
  [data-testid="stSidebar"] { min-width: 100% !important; border-right: none; border-bottom: 1px solid var(--border); }
  .metrics { flex-direction: column; align-items:flex-start; }
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Header (Stripe-style)
# ---------------------------
st.markdown("""
<div class="app-header">
  <div class="brand">
    <div class="logo">PP</div>
    <div>
      <div class="title">Polymer Production Scheduler</div>
      <div class="subtitle">Shutdown-aware multi-plant optimization</div>
    </div>
  </div>
  <div class="header-actions">
    <div style="text-align:right;">
      <div class="kpi">UI: Fintech Redesign</div>
      <div class="kpi-sub">Solver unchanged</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------
# Solution callback (unchanged logic)
# ---------------------------
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

        # FIXED: Consistent transition counting using day indices
        transition_count_per_line = {line: 0 for line in self.lines}
        total_transitions = 0

        for line in self.lines:
            last_grade = None
            
            for d in range(self.num_days):
                current_grade = None
                # Find which grade is producing (use consistent indexing)
                for grade in self.grades:
                    key = (grade, line, d)
                    if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                        current_grade = grade
                        break
                
                # Only count transitions on consecutive production days
                if current_grade is not None:
                    if last_grade is not None and current_grade != last_grade:
                        transition_count_per_line[line] += 1
                        total_transitions += 1
                    last_grade = current_grade
                # Note: Don't reset last_grade if no production - this is correct for shutdown handling

        solution['transitions'] = {
            'per_line': transition_count_per_line,
            'total': total_transitions
        }

    def num_solutions(self):
        return len(self.solutions)

# ---------------------------
# Sidebar: file upload & parameters (re-styled containers only)
# ---------------------------
with st.sidebar:
    st.markdown("<div class='sidebar-group'>", unsafe_allow_html=True)
    st.markdown("<div style='font-weight:700; margin-bottom:8px;'>📥 Data</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Upload Excel File", 
        type=["xlsx"],
        help="Upload an Excel file with Plant, Inventory, and Demand sheets"
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='sidebar-group'>", unsafe_allow_html=True)
    st.markdown("<div style='font-weight:700; margin-bottom:8px;'>⚙️ Optimization Parameters</div>", unsafe_allow_html=True)

    if uploaded_file:
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
        
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        st.markdown("<div style='font-weight:600; margin-bottom:6px;'>Objective weights</div>", unsafe_allow_html=True)

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

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div style='padding:4px 0; text-align:center;'>", unsafe_allow_html=True)
        st.button("🎯 Run Production Optimization", type="primary", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        # Minimal guidance when file not uploaded
        st.markdown("<div style='color:var(--muted); font-size:13px;'>Upload your Excel file to enable parameters and run optimization.</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------
# Main body — same workflow and logic as original file
# (I kept solver, model, and all plotting code the same; only UI wrappers and classes changed)
# ---------------------------
if uploaded_file:
    try:
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Data preview area
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<div class='h2'>📈 Data preview & validation</div>", unsafe_allow_html=True)
        
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

                st.subheader("🏭 Plant Data")
                st.dataframe(plant_display_df, use_container_width=True)
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
                    inventory_display_df[force_start_column] = inventory_display_df[start_column].dt.strftime('%d-%b-%y')
                
                st.subheader("📦 Inventory Data")
                st.dataframe(inventory_display_df, use_container_width=True)
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
                
                st.subheader("📊 Demand Data")
                st.dataframe(demand_display_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Demand sheet: {e}")
                st.stop()
        
        st.markdown("</div>", unsafe_allow_html=True)

        # Shutdown info (kept logic)
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<div class='h2'>🔧 Shutdowns</div>", unsafe_allow_html=True)
        shutdown_found = False
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
                        st.warning(f"⚠️ Invalid shutdown period for {plant}: Start date is after end date")
                    else:
                        st.info(f"**{plant}**: Scheduled for shutdown from {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')} ({duration} days)")
                        shutdown_found = True
                except Exception as e:
                    st.warning(f"⚠️ Invalid shutdown dates for {plant}: {e}")
        
        if not shutdown_found:
            st.info("ℹ️ No plant shutdowns scheduled")
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Transition matrices (unchanged behavior)
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
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<div class='h2'>Run Optimization</div>", unsafe_allow_html=True)

        if st.button("Run Optimization", type="primary", use_container_width=False):
            # Keep process & logic identical
            st.session_state.current_step = 2  # Optimization running
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_placeholder = st.empty()
            
            if 'solutions' not in st.session_state:
                st.session_state.solutions = []
            if 'best_solution' not in st.session_state:
                st.session_state.best_solution = None

            time.sleep(0.7)
            status_text.markdown('<div style="padding:8px; border-radius:8px; background:#f1f5f9; color:#0b1220">📄 Preprocessing data...</div>', unsafe_allow_html=True)
            progress_bar.progress(10)

            time.sleep(1.2)

            try:
                # Process inventory data with grade-plant combinations
                num_lines = len(plant_df)
                lines = list(plant_df['Plant'])
                capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
                
                # Get unique grades from demand sheet
                grades = [col for col in demand_df.columns if col != demand_df.columns[0]]
                
                # Store inventory parameters per (grade, plant) combination
                initial_inventory = {}  # Global per grade
                min_inventory = {}  # Global per grade
                max_inventory = {}  # Global per grade
                min_closing_inventory = {}  # Global per grade
                min_run_days = {}  # Per (grade, plant)
                max_run_days = {}  # Per (grade, plant)
                force_start_date = {}  # Per (grade, plant)
                allowed_lines = {grade: [] for grade in grades}  # List of lines per grade
                rerun_allowed = {}  # Per (grade, plant)
                
                # Track which grades have global inventory settings defined
                grade_inventory_defined = set()
                
                for index, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    
                    # Process Lines column
                    lines_value = row['Lines']
                    if pd.notna(lines_value) and lines_value != '':
                        plants_for_row = [x.strip() for x in str(lines_value).split(',')]
                    else:
                        plants_for_row = lines
                        st.warning(f"⚠️ Lines for grade '{grade}' (row {index}) are not specified, allowing all lines")
                    
                    # Add plants to allowed_lines for this grade
                    for plant in plants_for_row:
                        if plant not in allowed_lines[grade]:
                            allowed_lines[grade].append(plant)
                    
                    # Global inventory parameters (only set once per grade)
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
                    
                    # Plant-specific parameters
                    for plant in plants_for_row:
                        grade_plant_key = (grade, plant)
                        
                        # Min Run Days
                        if pd.notna(row['Min. Run Days']):
                            min_run_days[grade_plant_key] = int(row['Min. Run Days'])
                        else:
                            min_run_days[grade_plant_key] = 1
                        
                        # Max Run Days
                        if pd.notna(row['Max. Run Days']):
                            max_run_days[grade_plant_key] = int(row['Max. Run Days'])
                        else:
                            max_run_days[grade_plant_key] = 9999
                        
                        # Force Start Date
                        if pd.notna(row['Force Start Date']):
                            try:
                                force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date()
                            except:
                                force_start_date[grade_plant_key] = None
                                st.warning(f"⚠️ Invalid Force Start Date for grade '{grade}' on plant '{plant}'")
                        else:
                            force_start_date[grade_plant_key] = None
                        
                        # FIXED: Improved Rerun Allowed parsing
                        rerun_val = row['Rerun Allowed']
                        if pd.notna(rerun_val):
                            val_str = str(rerun_val).strip().lower()
                            if val_str in ['no', 'n', 'false', '0']:
                                rerun_allowed[grade_plant_key] = False
                            else:
                                rerun_allowed[grade_plant_key] = True
                        else:
                            rerun_allowed[grade_plant_key] = True
                
                # Material running info
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
            
            # Process demand data (unchanged)
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
            
            # Process shutdown dates
            shutdown_periods = process_shutdown_dates(plant_df, dates)
            
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
            status_text.markdown('<div style="padding:8px; border-radius:8px; background:#f1f5f9; color:#0b1220">🔧 Building optimization model...</div>', unsafe_allow_html=True)

            time.sleep(1.6)
            
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
            
            # SHUTDOWN CONSTRAINTS: No production during shutdown periods
            for line in lines:
                if line in shutdown_periods and shutdown_periods[line]:
                    for d in shutdown_periods[line]:
                        for grade in grades:
                            if is_allowed_combination(grade, line):
                                key = (grade, line, d)
                                if key in is_producing:
                                    # Force no production during shutdown
                                    model.Add(is_producing[key] == 0)
                                    model.Add(production[key] == 0)

            # FIXED: Document shutdown impact for validation
            shutdown_demand = {}
            for grade in grades:
                shutdown_demand[grade] = 0
                for line in allowed_lines[grade]:
                    if line in shutdown_periods:
                        for d in shutdown_periods[line]:
                            shutdown_demand[grade] += demand_data[grade].get(dates[d], 0)
            
            # Add warning if shutdown causes potential issues
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
            
            # FIXED CORRECTED INVENTORY BALANCE with proper stockout handling
            for grade in grades:
                model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])
            
            for grade in grades:
                for d in range(num_days):
                    produced_today = sum(
                        get_production_var(grade, line, d) 
                        for line in allowed_lines[grade]
                    )
                    demand_today = demand_data[grade].get(dates[d], 0)
                    
                    # Step 1: Calculate available inventory (expression, no variable needed)
                    # available = inventory_vars[(grade, d)] + produced_today
                    
                    # Step 2: Determine what can be supplied
                    supplied = model.NewIntVar(0, 100000, f'supplied_{grade}_{d}')
                    model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
                    model.Add(supplied <= demand_today)
                    
                    # Step 3: Calculate stockout based on unmet demand
                    model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
                    
                    # Step 4: Calculate closing inventory
                    model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
                    
                    # Ensure non-negativity
                    model.Add(inventory_vars[(grade, d + 1)] >= 0)
            
            # Minimum inventory constraints (as soft constraints with penalties)
            for grade in grades:
                for d in range(num_days):
                    if min_inventory[grade] > 0:
                        min_inv_value = int(min_inventory[grade])
                        inventory_tomorrow = inventory_vars[(grade, d + 1)]
                        
                        # Create a deficit variable that is positive when below minimum
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        model.Add(deficit >= min_inv_value - inventory_tomorrow)
                        model.Add(deficit >= 0)
                        objective += stockout_penalty * deficit
            
            # Minimum Closing Inventory constraint
            for grade in grades:
                closing_inventory = inventory_vars[(grade, num_days - buffer_days)]
                min_closing = min_closing_inventory[grade]
                
                if min_closing > 0:
                    closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                    model.Add(closing_deficit >= min_closing - closing_inventory)
                    model.Add(closing_deficit >= 0)
                    objective += stockout_penalty * closing_deficit * 3  # Higher penalty for closing inventory
            
            for grade in grades:
                for d in range(1, num_days + 1):
                    model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])
            
            for line in lines:
                for d in range(num_days - buffer_days):
                    # Skip shutdown days for full capacity requirement
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
            
            # Force Start Date per (grade, plant) combination
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
            
            # Minimum & Maximum Run Days per (grade, plant) - account for shutdown interruptions
            is_start_vars = {}
            run_end_vars = {}
            
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    min_run = min_run_days.get(grade_plant_key, 1)
                    max_run = max_run_days.get(grade_plant_key, 9999)
                    
                    # Create start and end variables
                    for d in range(num_days):
                        is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                        is_start_vars[(grade, line, d)] = is_start
                        
                        is_end = model.NewBoolVar(f'end_{grade}_{line}_{d}')
                        run_end_vars[(grade, line, d)] = is_end
                        
                        current_prod = get_is_producing_var(grade, line, d)
                        
                        # Start definition: producing today but not yesterday (or today is day 0)
                        if d > 0:
                            prev_prod = get_is_producing_var(grade, line, d - 1)
                            if current_prod is not None and prev_prod is not None:
                                model.AddBoolAnd([current_prod, prev_prod.Not()]).OnlyEnforceIf(is_start)
                                model.AddBoolOr([current_prod.Not(), prev_prod]).OnlyEnforceIf(is_start.Not())
                        else:
                            if current_prod is not None:
                                model.Add(current_prod == 1).OnlyEnforceIf(is_start)
                                model.Add(is_start == 1).OnlyEnforceIf(current_prod)
                        
                        # End definition: producing today but not tomorrow (or today is last day)
                        if d < num_days - 1:
                            next_prod = get_is_producing_var(grade, line, d + 1)
                            if current_prod is not None and next_prod is not None:
                                model.AddBoolAnd([current_prod, next_prod.Not()]).OnlyEnforceIf(is_end)
                                model.AddBoolOr([current_prod.Not(), next_prod]).OnlyEnforceIf(is_end.Not())
                        else:
                            if current_prod is not None:
                                model.Add(current_prod == 1).OnlyEnforceIf(is_end)
                                model.Add(is_end == 1).OnlyEnforceIf(current_prod)
            
                    # MINIMUM RUN DAYS: If we start a run, it must continue for at least min_run days
                    # (unless interrupted by shutdown)
                    for d in range(num_days):
                        is_start = is_start_vars[(grade, line, d)]
                        
                        # Check how many consecutive non-shutdown days we have from day d
                        max_possible_run = 0
                        for k in range(min_run):
                            if d + k < num_days:
                                # Check if this day is a shutdown day
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    break
                                max_possible_run += 1
                        
                        # Only enforce if we have enough consecutive days available
                        if max_possible_run >= min_run:
                            # Force production for the next min_run days (if no shutdown)
                            for k in range(min_run):
                                if d + k < num_days:
                                    # Skip if this is a shutdown day
                                    if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                        continue
                                    future_prod = get_is_producing_var(grade, line, d + k)
                                    if future_prod is not None:
                                        model.Add(future_prod == 1).OnlyEnforceIf(is_start)
            
                    # FIXED: MAXIMUM RUN DAYS - Single correct implementation
                    # Use a sliding window approach with consecutive production counting
                    for d in range(num_days - max_run):
                        # Count consecutive days of production starting from day d
                        consecutive_days = []
                        for k in range(max_run + 1):
                            if d + k < num_days:
                                # Skip shutdown days - they break the run naturally
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    break
                                prod_var = get_is_producing_var(grade, line, d + k)
                                if prod_var is not None:
                                    consecutive_days.append(prod_var)
                        
                        # If we have max_run+1 consecutive days available, prevent all being active
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

            # Rerun Allowed Constraints
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    if not rerun_allowed.get(grade_plant_key, True):
                        # If rerun not allowed, don't start multiple runs
                        starts = [is_start_vars[(grade, line, d)] for d in range(num_days) 
                                 if (grade, line, d) in is_start_vars]
                        if starts:
                            model.Add(sum(starts) <= 1)

            # Stockout penalties in objective
            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

            # Transition penalties and continuity bonuses
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
            status_text.markdown('<div style="padding:8px; border-radius:8px; background:#f1f5f9; color:#0b1220">⚡ Running optimization solver...</div>', unsafe_allow_html=True)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit_min * 60.0
            solver.parameters.num_search_workers = 8
            solver.parameters.random_seed = 42  # FIXED: Add for repeatability
            solver.parameters.log_search_progress = True  # Optional: for debugging
            
            solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

            start_time = time.time()
            status = solver.Solve(model, solution_callback)
            
            progress_bar.progress(100)
            
            # Check solver status
            if status == cp_model.OPTIMAL:
                status_text.markdown('<div style="padding:8px; border-radius:8px; background:#ecfdf5; color:#065f46">✅ Optimization completed optimally!</div>', unsafe_allow_html=True)
            elif status == cp_model.FEASIBLE:
                status_text.markdown('<div style="padding:8px; border-radius:8px; background:#ecfdf5; color:#065f46">✅ Optimization completed with feasible solution!</div>', unsafe_allow_html=True)
            else:
                status_text.markdown('<div style="padding:8px; border-radius:8px; background:#fff7ed; color:#92400e">⚠️ Optimization ended without proven optimal solution.</div>', unsafe_allow_html=True)

            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
            st.markdown("<div class='panel'>", unsafe_allow_html=True)
            st.markdown("<div class='h2'>📈 Results</div>", unsafe_allow_html=True)

            if solution_callback.num_solutions() > 0:
                best_solution = solution_callback.solutions[-1]

                # KPI row (Stripe-style minimal tiles)
                st.markdown("<div class='metrics'>", unsafe_allow_html=True)
                col1, col2, col3, col4 = st.columns([1,1,1,1])
                with col1:
                    st.markdown("<div class='metric-tile'>", unsafe_allow_html=True)
                    st.markdown("<div class='metric-label'>Objective Value</div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='metric-value'>{best_solution['objective']:,.0f}</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                with col2:
                    st.markdown("<div class='metric-tile'>", unsafe_allow_html=True)
                    st.markdown("<div class='metric-label'>Total Transitions</div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='metric-value'>{best_solution['transitions']['total']}</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                with col3:
                    total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                    st.markdown("<div class='metric-tile'>", unsafe_allow_html=True)
                    st.markdown("<div class='metric-label'>Total Stockouts</div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='metric-value'>{total_stockouts:,.0f} MT</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                with col4:
                    st.markdown("<div class='metric-tile'>", unsafe_allow_html=True)
                    st.markdown("<div class='metric-label'>Planning Horizon</div>", unsafe_allow_html=True)
                    st.markdown(f"<div class='metric-value'>{num_days} days</div>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

                # Tabs for production / summary / inventory (visualization code unchanged)
                tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Summary", "📦 Inventory"])
                
                with tab1:
                    sorted_grades = sorted(grades)
                    base_colors = px.colors.qualitative.Vivid
                    grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

                    cmap = colormaps.get_cmap('tab20')
                    grade_colors = {}
                    for idx, grade in enumerate(grades):
                        grade_colors[grade] = cmap(idx % 20)
                    
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
                        
                        # Calculate total stockout for this grade
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
                st.info("""
                **Common issues that cause infeasibility:**
                - Demand is too high compared to production capacity
                - Minimum run days are too long for the available production days
                - Minimum closing inventory requirements are too high
                - Shutdown periods conflict with mandatory production requirements
                - Transition rules are too restrictive
                
                **Suggestions:**
                - Reduce minimum run days requirements
                - Lower minimum closing inventory targets  
                - Increase production capacity
                - Reduce demand forecasts
                - Adjust shutdown periods
                """)
            
            st.markdown("</div>", unsafe_allow_html=True)  # results panel
        st.markdown("</div>", unsafe_allow_html=True)  # run panel

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.info("Please make sure your Excel file has the required sheets: 'Plant', 'Inventory', and 'Demand'")

else:
    # Landing page content (Stripe-style)
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='h2'>Welcome</div>", unsafe_allow_html=True)
    st.markdown("""
    <div style="display:flex; gap:18px; align-items:flex-start; flex-wrap:wrap;">
      <div style="flex:1; min-width:260px;">
        <div style="font-weight:700; margin-bottom:6px;">Quick start</div>
        <div style="color:#68707d;">Download the sample template, prepare Plant/Inventory/Demand sheets, upload and run the optimizer.</div>
      </div>
      <div style="flex:1; min-width:260px;">
        <div style="font-weight:700; margin-bottom:6px;">What this tool does</div>
        <ul style="margin:0; padding-left:16px; color:#68707d;">
          <li>Balance production across plants</li>
          <li>Respect run-lengths & transition rules</li>
          <li>Handle scheduled shutdowns</li>
        </ul>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    sample_workbook = get_sample_workbook()

    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.markdown("<div class='h2'>📥 Sample template</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([2,1])
    with col1:
        st.markdown("""
        **Download the sample template to get started quickly.**
        It contains example Plant, Inventory and Demand sheets pre-configured.
        """)
    with col2:
        st.download_button(
            label="Download template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    st.markdown("</div>", unsafe_allow_html=True)

# Footer
st.markdown("<div class='app-footer'>Polymer Production Scheduler • Built with Streamlit • UI redesign: Fintech Dashboard</div>", unsafe_allow_html=True)
