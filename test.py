# app.py
import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta, datetime
import matplotlib.pyplot as plt
import numpy as np
import time
import io
import base64
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go

# -------------------------
# Solution Callback (preserved)
# -------------------------
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

        # Consistent transition counting using day indices
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

# -------------------------
# Helper: process shutdown dates (preserved)
# -------------------------
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

# -------------------------
# Page config & CSS (dark Material + glass morphism)
# -------------------------
st.set_page_config(page_title="Polymer Production Scheduler (Dark)", page_icon="🏭", layout="wide", initial_sidebar_state="collapsed")

# Dark UI CSS
st.markdown(
    """
    <style>
    :root{
        --bg:#0b1020;
        --card:#071024cc;
        --glass:#0f172433;
        --muted:#9aa6bf;
        --accent1:#7c4dff;
        --accent2:#00bcd4;
        --accent3:#4f46e5;
        --glass-border: rgba(255,255,255,0.06);
    }
    html,body,#root, .appview-container {
        background: linear-gradient(180deg, var(--bg), #071229) !important;
        color: #e6eef8;
    }
    .stApp, .css-1b0z5mo {
        background: transparent;
    }
    /* Header */
    .header {
        background: linear-gradient(90deg, rgba(79,70,229,0.18), rgba(7,120,150,0.12));
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 16px;
        box-shadow: 0 8px 30px rgba(2,6,23,0.7), inset 0 1px 0 rgba(255,255,255,0.02);
        backdrop-filter: blur(6px) saturate(120%);
        border: 1px solid var(--glass-border);
    }
    .header h1 {
        margin: 0;
        font-size: 1.6rem;
        letter-spacing: 0.2px;
    }
    .subtle {
        color: var(--muted);
        font-size: 0.9rem;
    }
    /* Sidebar card style */
    .side-card {
        background: linear-gradient(90deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 12px;
        border: 1px solid var(--glass-border);
        box-shadow: 0 6px 20px rgba(3,6,23,0.6);
        backdrop-filter: blur(6px);
    }
    .uploader {
        padding: 10px;
        border-radius: 10px;
        border: 1px dashed rgba(255,255,255,0.06) !important;
        background: linear-gradient(180deg, rgba(255,255,255,0.012), rgba(255,255,255,0.006));
    }
    .valid-badge {
        display:inline-block;
        background: linear-gradient(90deg, #21d28a, #00b894);
        color: #071024;
        padding: 4px 8px;
        border-radius: 999px;
        font-weight:600;
        font-size:0.8rem;
    }
    .warn-badge {
        display:inline-block;
        background: linear-gradient(90deg, #ffb366, #ff7a66);
        color: #071024;
        padding: 4px 8px;
        border-radius: 999px;
        font-weight:600;
        font-size:0.8rem;
    }
    /* Minimal metric cards */
    .metric {
        background: linear-gradient(90deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));
        padding: 12px;
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,0.03);
        text-align:center;
    }
    .metric h3 { margin:0; font-size:1.25rem; }
    .metric p { margin:0; color: var(--muted); }
    /* Buttons */
    .primary-btn {
        background: linear-gradient(90deg, var(--accent1), var(--accent3));
        border:none;
        color:white;
        padding:10px 18px;
        border-radius:10px;
        font-weight:700;
        cursor:pointer;
    }
    .primary-btn:hover { transform: translateY(-2px); box-shadow: 0 10px 30px rgba(79,70,229,0.18); }
    /* Tabs dark */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(255,255,255,0.02);
        border-radius:10px;
        padding:6px;
        display:flex;
        gap:8px;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #cfe6ff;
        border-radius:8px;
        padding:8px 16px;
        font-weight:600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, var(--accent1), var(--accent2));
        color:#071024;
    }
    /* Dataframe dark adjustments */
    .dataframe tr:hover { background: rgba(255,255,255,0.02) !important; }
    /* compact spacing */
    .compact { padding:6px; margin:6px; }
    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------
# Sidebar: compact card-based parameter panel (upload + params)
# -------------------------
with st.sidebar:
    st.markdown("<div class='side-card'>", unsafe_allow_html=True)
    st.markdown("<div style='display:flex;align-items:center;gap:12px'><div style='font-size:1.3rem;'>🏭</div><div><strong>Polymer Scheduler</strong><div class='subtle'>Dark Material • Desktop only</div></div></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Upload card
    st.markdown("<div class='side-card'><strong>📥 Upload</strong><div class='subtle'>Excel file with sheets: Plant, Inventory, Demand</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["xlsx"], key="uploader", help="Upload the Excel workbook", label_visibility="collapsed")
    if uploaded_file:
        st.markdown("<div style='margin-top:8px;'><span class='valid-badge'>File received</span></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div style='margin-top:8px;'><span class='warn-badge'>No file</span></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Parameters card (compact, expandable)
    st.markdown("<div class='side-card'><strong>⚙️ Parameters</strong>", unsafe_allow_html=True)
    with st.expander("Basic", expanded=True):
        time_limit_min = st.number_input("Time limit (min)", min_value=1, max_value=120, value=10, help="Solver max runtime")
        buffer_days = st.number_input("Buffer days", min_value=0, max_value=7, value=3, help="Extra horizon days (buffer)")
    with st.expander("Objective", expanded=False):
        stockout_penalty = st.number_input("Stockout penalty", min_value=1, value=10)
        transition_penalty = st.number_input("Transition penalty", min_value=1, value=10)
        continuity_bonus = st.number_input("Continuity bonus", min_value=0, value=1)
    st.markdown("</div>", unsafe_allow_html=True)

    # Quick actions
    st.markdown("<div class='side-card'>", unsafe_allow_html=True)
    st.button("⚡ One-click Optimize", key="run_topbar", help="Shortcut to run optimization (also available in main area)")
    st.markdown("<div class='subtle' style='margin-top:8px'>Keyboard: N/A (desktop-focused)</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Sample download (load from repository - same folder as app.py)
    st.markdown("<div class='side-card'>", unsafe_allow_html=True)
    try:
        current_dir = Path(__file__).parent
        template_path = current_dir / "polymer_production_template.xlsx"
        if template_path.exists():
            with open(template_path, "rb") as f:
                template_bytes = f.read()
            st.download_button(
                "📥 Download Template",
                data=template_bytes,
                file_name="polymer_production_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("Template file 'polymer_production_template.xlsx' not found in repository.")
    except Exception as e:
        st.error(f"Unable to load template file: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------
# Main header / topology
# -------------------------
st.markdown("<div class='header'><h1>Polymer Production Scheduler</h1><div class='subtle'>Dark Material • Glass morphism • Desktop only — streamlined workflow</div></div>", unsafe_allow_html=True)

# Use session state to track steps and data
if 'uploaded' not in st.session_state:
    st.session_state.uploaded = False
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'processing' not in st.session_state:
    st.session_state.processing = False
if 'last_status' not in st.session_state:
    st.session_state.last_status = ""

# -------------------------
# Data preview + validation (single page, tabbed cards)
# -------------------------
data_ready = False
plant_df = inventory_df = demand_df = None
transition_dfs = {}

if uploaded_file:
    try:
        excel_file_bytes = uploaded_file.read()
        excel_file = io.BytesIO(excel_file_bytes)

        # Quick validation: check required sheets
        xls = pd.ExcelFile(excel_file)
        sheets = xls.sheet_names
        have_plant = 'Plant' in sheets
        have_inventory = 'Inventory' in sheets
        have_demand = 'Demand' in sheets

        # immediate feedback badges
        validation_cols = st.columns(3)
        with validation_cols[0]:
            st.markdown("<div class='metric'><h3>{}</h3><p>Plant sheet</p></div>".format("✅" if have_plant else "⚠️"), unsafe_allow_html=True)
        with validation_cols[1]:
            st.markdown("<div class='metric'><h3>{}</h3><p>Inventory sheet</p></div>".format("✅" if have_inventory else "⚠️"), unsafe_allow_html=True)
        with validation_cols[2]:
            st.markdown("<div class='metric'><h3>{}</h3><p>Demand sheet</p></div>".format("✅" if have_demand else "⚠️"), unsafe_allow_html=True)

        # If mandatory sheets exist, load them for preview; else show helpful message
        if not (have_plant and have_inventory and have_demand):
            st.error("Missing required sheets. Please ensure your Excel has 'Plant', 'Inventory', and 'Demand' sheets.")
            st.info("Refer to the template for exact headers. (See download in the sidebar.)")
        else:
            # Read sheets
            excel_file.seek(0)
            plant_df = pd.read_excel(excel_file, sheet_name='Plant')
            excel_file.seek(0)
            inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
            excel_file.seek(0)
            demand_df = pd.read_excel(excel_file, sheet_name='Demand', parse_dates=[0])

            # detect transition sheets
            excel_file.seek(0)
            xls2 = pd.ExcelFile(io.BytesIO(excel_file_bytes))
            for sn in xls2.sheet_names:
                if sn.lower().startswith("transition"):
                    try:
                        df = pd.read_excel(io.BytesIO(excel_file_bytes), sheet_name=sn, index_col=0)
                        transition_dfs[sn] = df
                    except Exception:
                        transition_dfs[sn] = None

            st.session_state.uploaded = True
            st.session_state.data_loaded = True
            data_ready = True

            # Tabbed preview
            tabs = st.tabs(["📋 Data Preview", "⚠️ Validation", "🔧 Shutdowns & Transitions"])
            with tabs[0]:
                st.subheader("Data Preview (compact)")
                c1, c2 = st.columns([1, 1])
                with c1:
                    st.markdown("**Plant**")
                    st.dataframe(plant_df, use_container_width=True)
                with c2:
                    st.markdown("**Inventory**")
                    st.dataframe(inventory_df, use_container_width=True)
                st.markdown("**Demand (first 20 rows)**")
                st.dataframe(demand_df.head(20), use_container_width=True)

            with tabs[1]:
                st.subheader("Validation")
                # Quick checks for obvious problems
                issues = []
                if 'Plant' in plant_df.columns:
                    if plant_df['Plant'].duplicated().any():
                        issues.append("Duplicate plant names – please ensure unique Plant identifiers.")
                if 'Grade Name' in inventory_df.columns:
                    if inventory_df['Grade Name'].isnull().any():
                        issues.append("Empty Grade Name found in Inventory.")
                # Show color-coded indicators
                if issues:
                    for it in issues:
                        st.warning(it)
                else:
                    st.success("Basic validation passed. No obvious issues found.")

            with tabs[2]:
                st.subheader("Shutdowns & Transition Matrices")
                # Shutdown cards
                for idx, row in plant_df.iterrows():
                    plant = row['Plant']
                    ss = row.get('Shutdown Start Date')
                    se = row.get('Shutdown End Date')
                    if pd.notna(ss) and pd.notna(se):
                        try:
                            ssd = pd.to_datetime(ss).date()
                            sed = pd.to_datetime(se).date()
                            st.info(f"**{plant}** shutdown: {ssd.strftime('%d-%b-%y')} → {sed.strftime('%d-%b-%y')}")
                        except Exception:
                            st.warning(f"Invalid shutdown dates for {plant}")
                    else:
                        st.markdown(f"**{plant}** — no scheduled shutdowns", unsafe_allow_html=True)

                if transition_dfs:
                    for name, df in transition_dfs.items():
                        st.markdown(f"**{name}**")
                        if df is not None:
                            st.dataframe(df, use_container_width=True)
                        else:
                            st.markdown("_Could not parse transition matrix._")
    except Exception as e:
        st.error(f"Error reading uploaded file: {e}")
        st.stop()
else:
    st.info("Upload an Excel workbook via the left panel to begin. Use the sample template if you need a quick start.")
    data_ready = False

# -------------------------
# Main optimization control and runner
# -------------------------
st.markdown("---")
st.markdown("## ⚙️ Optimization")
col_run, col_status = st.columns([2, 3])

with col_run:
    run_btn = st.button("🎯 Run Production Optimization", key="run_main", help="Run the optimizer with current parameters")

with col_status:
    status_box = st.empty()
    status_box.markdown("<div class='subtle'>Status: idle</div>", unsafe_allow_html=True)

if run_btn:
    if not data_ready:
        st.error("No valid data to run. Upload a workbook with required sheets first.")
    else:
        # Start processing
        st.session_state.processing = True
        status_box.markdown("<div class='subtle'>Starting preprocessing...</div>", unsafe_allow_html=True)
        progress = st.progress(0)
        time.sleep(0.5)
        progress.progress(5)

        # --- Begin: Data preprocessing and model build (logic preserved) ---
        try:
            # Recreate byte stream for repeated reads
            excel_file = io.BytesIO(excel_file_bytes)

            # Quick mapping and parameter extraction
            num_lines = len(plant_df)
            lines = list(plant_df['Plant'])
            capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}

            # Get unique grades from demand sheet (columns except first date column)
            grades = [col for col in demand_df.columns if col != demand_df.columns[0]]

            # Inventory & related dictionaries
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

                # Lines parsing
                lines_value = row.get('Lines') if 'Lines' in inventory_df.columns else None
                if pd.notna(lines_value) and str(lines_value).strip() != "":
                    plants_for_row = [x.strip() for x in str(lines_value).split(',')]
                else:
                    plants_for_row = lines
                    st.warning(f"⚠️ Lines for grade '{grade}' (row {index}) are not specified, allowing all lines")

                for plant in plants_for_row:
                    if plant not in allowed_lines[grade]:
                        allowed_lines[grade].append(plant)

                if grade not in grade_inventory_defined:
                    initial_inventory[grade] = row['Opening Inventory'] if pd.notna(row.get('Opening Inventory')) else 0
                    min_inventory[grade] = row['Min. Inventory'] if pd.notna(row.get('Min. Inventory')) else 0
                    max_inventory[grade] = row['Max. Inventory'] if pd.notna(row.get('Max. Inventory')) else 1000000000
                    min_closing_inventory[grade] = row['Min. Closing Inventory'] if pd.notna(row.get('Min. Closing Inventory')) else 0
                    grade_inventory_defined.add(grade)

                for plant in plants_for_row:
                    grade_plant_key = (grade, plant)
                    min_run_days[grade_plant_key] = int(row['Min. Run Days']) if pd.notna(row.get('Min. Run Days')) else 1
                    max_run_days[grade_plant_key] = int(row['Max. Run Days']) if pd.notna(row.get('Max. Run Days')) else 9999
                    if pd.notna(row.get('Force Start Date')):
                        try:
                            force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date()
                        except Exception:
                            force_start_date[grade_plant_key] = None
                            st.warning(f"⚠️ Invalid Force Start Date for grade '{grade}' on plant '{plant}'")
                    else:
                        force_start_date[grade_plant_key] = None
                    # Rerun parsing
                    rerun_val = row.get('Rerun Allowed')
                    if pd.notna(rerun_val):
                        val_str = str(rerun_val).strip().lower()
                        if val_str in ['no', 'n', 'false', '0']:
                            rerun_allowed[grade_plant_key] = False
                        else:
                            rerun_allowed[grade_plant_key] = True
                    else:
                        rerun_allowed[grade_plant_key] = True

            # Material running info from plant sheet
            material_running_info = {}
            for index, row in plant_df.iterrows():
                plant = row['Plant']
                material = row.get('Material Running')
                expected_days = row.get('Expected Run Days')
                if pd.notna(material) and pd.notna(expected_days):
                    try:
                        material_running_info[plant] = (str(material).strip(), int(expected_days))
                    except (ValueError, TypeError):
                        st.warning(f"⚠️ Invalid Material Running or Expected Run Days for plant '{plant}'")

            progress.progress(20)
            status_box.markdown("<div class='subtle'>Preparing demand & horizon...</div>", unsafe_allow_html=True)
            time.sleep(0.5)

            # Demand data processing (preserved)
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

            # Shutdowns
            shutdown_periods = process_shutdown_dates(plant_df, dates)

            # Transition rules parsing
            transition_rules = {}
            for name, df in transition_dfs.items():
                plant_name = name.replace("Transition_", "").replace("Transition-", "").replace("transition_", "")
                try:
                    transition_rules[plant_name] = {}
                    for prev_grade in df.index:
                        allowed_transitions = []
                        for current_grade in df.columns:
                            if str(df.loc[prev_grade, current_grade]).lower() == 'yes':
                                allowed_transitions.append(current_grade)
                        transition_rules[plant_name][prev_grade] = allowed_transitions
                except Exception:
                    transition_rules[plant_name] = None

            progress.progress(35)
            status_box.markdown("<div class='subtle'>Building CP-SAT model...</div>", unsafe_allow_html=True)
            time.sleep(0.5)

            # --- Model building (the original solver logic is preserved exactly) ---
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

            # Shutdown constraints (preserved)
            for line in lines:
                if line in shutdown_periods and shutdown_periods[line]:
                    for d in shutdown_periods[line]:
                        for grade in grades:
                            if is_allowed_combination(grade, line):
                                key = (grade, line, d)
                                if key in is_producing:
                                    model.Add(is_producing[key] == 0)
                                    model.Add(production[key] == 0)

            # shutdown_demand warning (preserved)
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

            # Inventory & stockout variables (preserved)
            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')

            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')

            # Only one grade per line per day
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

            # Material running enforcement
            for plant, (material, expected_days) in material_running_info.items():
                for d in range(min(expected_days, num_days)):
                    if is_allowed_combination(material, plant):
                        model.Add(get_is_producing_var(material, plant, d) == 1)
                        for other_material in grades:
                            if other_material != material and is_allowed_combination(other_material, plant):
                                model.Add(get_is_producing_var(other_material, plant, d) == 0)

            # Objective building & inventory balance (preserved)
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

            # Soft minimum inventory constraint with deficit variables (preserved)
            for grade in grades:
                for d in range(num_days):
                    if min_inventory[grade] > 0:
                        min_inv_value = int(min_inventory[grade])
                        inventory_tomorrow = inventory_vars[(grade, d + 1)]
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        model.Add(deficit >= min_inv_value - inventory_tomorrow)
                        model.Add(deficit >= 0)
                        objective += stockout_penalty * deficit

            # Minimum Closing Inventory
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

            # Production capacity enforcement (preserved)
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

            # Force Start Dates
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

            # Min/Max run days (preserved)
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

            # Transition rules enforcement (preserved)
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

            # Rerun allowed constraints (preserved)
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    if not rerun_allowed.get(grade_plant_key, True):
                        starts = [is_start_vars[(grade, line, d)] for d in range(num_days) if (grade, line, d) in is_start_vars]
                        if starts:
                            model.Add(sum(starts) <= 1)

            # Stockout penalties in objective
            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

            # Transition penalties & continuity (preserved)
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

            progress.progress(60)
            status_box.markdown("<div class='subtle'>Solver starting...</div>", unsafe_allow_html=True)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit_min * 60.0
            solver.parameters.num_search_workers = 8
            solver.parameters.random_seed = 42

            solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

            start_time = time.time()
            status = solver.Solve(model, solution_callback)
            elapsed = time.time() - start_time

            progress.progress(100)
            if status == cp_model.OPTIMAL:
                status_box.markdown("<div class='subtle'>✅ Optimization completed optimally</div>", unsafe_allow_html=True)
            elif status == cp_model.FEASIBLE:
                status_box.markdown("<div class='subtle'>✅ Optimization found feasible solution</div>", unsafe_allow_html=True)
            else:
                status_box.markdown("<div class='subtle'>⚠️ Solver ended without proven optimal solution</div>", unsafe_allow_html=True)

            # -------------------------
            # Results display (tabs) — preserve Plotly logic, adapt dark layout
            # -------------------------
            st.markdown("---")
            st.markdown("## 📈 Results")
            if solution_callback.num_solutions() > 0:
                best_solution = solution_callback.solutions[-1]

                # Key metrics
                col_a, col_b, col_c, col_d = st.columns(4)
                with col_a:
                    st.markdown(f"<div class='metric'><h3>{best_solution['objective']:,.0f}</h3><p>Objective Value</p></div>", unsafe_allow_html=True)
                with col_b:
                    st.markdown(f"<div class='metric'><h3>{best_solution['transitions']['total']}</h3><p>Total Transitions</p></div>", unsafe_allow_html=True)
                with col_c:
                    total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                    st.markdown(f"<div class='metric'><h3>{total_stockouts:,.0f} MT</h3><p>Total Stockouts</p></div>", unsafe_allow_html=True)
                with col_d:
                    st.markdown(f"<div class='metric'><h3>{num_days}</h3><p>Planning Horizon (days)</p></div>", unsafe_allow_html=True)

                # Tabs: schedule / summary / inventory
                rtab1, rtab2, rtab3 = st.tabs(["📅 Production Schedule", "📊 Summary", "📦 Inventory"])

                with rtab1:
                    sorted_grades = sorted(grades)
                    base_colors = px.colors.qualitative.Vivid
                    grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

                    # Keep original Gantt & shutdown visualization code but set dark backgrounds
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
                        # Shutdown visualization preserved
                        if line in shutdown_periods and shutdown_periods[line]:
                            shutdown_days = shutdown_periods[line]
                            start_shutdown = dates[shutdown_days[0]]
                            end_shutdown = dates[shutdown_days[-1]] + timedelta(days=1)
                            fig.add_vrect(
                                x0=start_shutdown,
                                x1=end_shutdown,
                                fillcolor="red",
                                opacity=0.18,
                                layer="below",
                                line_width=0,
                                annotation_text="Shutdown",
                                annotation_position="top left",
                                annotation_font_size=12,
                                annotation_font_color="white"
                            )
                        fig.update_yaxes(autorange="reversed", title=None, showgrid=True, gridcolor="rgba(255,255,255,0.03)")
                        fig.update_xaxes(title="Date", showgrid=True, gridcolor="rgba(255,255,255,0.03)", tickvals=dates, tickformat="%d-%b", dtick="D1")
                        fig.update_layout(
                            height=340,
                            bargap=0.2,
                            showlegend=True,
                            legend_title_text="Grade",
                            legend=dict(traceorder="normal", orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02, bgcolor="rgba(0,0,0,0)"),
                            xaxis=dict(showline=True, showticklabels=True),
                            yaxis=dict(showline=True),
                            margin=dict(l=60, r=160, t=60, b=60),
                            plot_bgcolor="rgba(0,0,0,0)",
                            paper_bgcolor="rgba(0,0,0,0)",
                            font=dict(size=12, color="#e6eef8"),
                        )
                        st.plotly_chart(fig, use_container_width=True)

                    # Production schedule by line (table)
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
                        st.dataframe(schedule_df, use_container_width=True)

                with rtab2:
                    # Production summary (preserved)
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

                with rtab3:
                    # Inventory charts (preserved, dark themed)
                    st.subheader("Inventory Levels")
                    last_actual_day = num_days - buffer_days - 1
                    for grade in sorted(grades):
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
                                    opacity=0.12,
                                    layer="below",
                                    line_width=0,
                                    annotation_text=f"Shutdown: {line}" if not shutdown_added else "",
                                    annotation_position="top left",
                                    annotation_font_size=12,
                                    annotation_font_color="white"
                                )
                                shutdown_added = True

                        fig.add_hline(
                            y=min_inventory[grade],
                            line=dict(color="red", width=2, dash="dash"),
                            annotation_text=f"Min: {min_inventory[grade]:,.0f}",
                            annotation_position="top left",
                            annotation_font_color="white"
                        )
                        fig.add_hline(
                            y=max_inventory[grade],
                            line=dict(color="green", width=2, dash="dash"),
                            annotation_text=f"Max: {max_inventory[grade]:,.0f}",
                            annotation_position="bottom left",
                            annotation_font_color="white"
                        )

                        annotations = [
                            dict(x=start_x, y=start_val, text=f"Start: {start_val:.0f}", showarrow=True, arrowhead=2, ax=-40, ay=30, font=dict(color="white", size=11), bgcolor="#071024", bordercolor="#2b3748"),
                            dict(x=end_x, y=end_val, text=f"End: {end_val:.0f}", showarrow=True, arrowhead=2, ax=40, ay=30, font=dict(color="white", size=11), bgcolor="#071024", bordercolor="#2b3748"),
                            dict(x=highest_x, y=highest_val, text=f"High: {highest_val:.0f}", showarrow=True, arrowhead=2, ax=0, ay=-40, font=dict(color="white", size=11), bgcolor="#071024", bordercolor="#2b3748"),
                            dict(x=lowest_x, y=lowest_val, text=f"Low: {lowest_val:.0f}", showarrow=True, arrowhead=2, ax=0, ay=40, font=dict(color="white", size=11), bgcolor="#071024", bordercolor="#2b3748"),
                        ]

                        fig.update_layout(
                            title=f"Inventory Level - {grade}",
                            xaxis=dict(title="Date", showgrid=True, gridcolor="rgba(255,255,255,0.03)", tickvals=dates, tickformat="%d-%b", dtick="D1"),
                            yaxis=dict(title="Inventory Volume (MT)", showgrid=True, gridcolor="rgba(255,255,255,0.03)"),
                            plot_bgcolor="rgba(0,0,0,0)",
                            paper_bgcolor="rgba(0,0,0,0)",
                            margin=dict(l=60, r=80, t=80, b=60),
                            font=dict(size=12, color="#e6eef8"),
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
                                borderwidth=1,
                                borderpad=4,
                                opacity=0.95
                            )
                        st.plotly_chart(fig, use_container_width=True)

            else:
                st.error("No solutions found during optimization. Please check constraints and data.")
                st.info("""
                    Common causes:
                    - Demand > capacity
                    - Min run days too strict
                    - Closing inventory needs too high
                    - Shutdowns conflict with forced production
                    - Transition constraints too restrictive
                    """)
        except Exception as e:
            st.error(f"Error during optimization: {e}")
            import traceback
            st.text(traceback.format_exc())
        finally:
            st.session_state.processing = False

# -------------------------
# Footer
# -------------------------
st.markdown("---")
st.markdown("<div class='subtle' style='text-align:center'>Polymer Production Scheduler • Dark Material UI • Keep solver rules & plot logic unchanged</div>", unsafe_allow_html=True)
