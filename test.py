# app.py
import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta, datetime
import numpy as np
import time
import io
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

# -------------------------
# Streamlit config + Light Material + Glassmorphism CSS
# -------------------------
st.set_page_config(page_title="Polymer Production Scheduler", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
    <style>
    :root{
        --bg:#f6f9fc;
        --card:#ffffff;
        --glass: rgba(255,255,255,0.6);
        --muted:#6b7280;
        --accent1:#3b82f6; /* blue */
        --accent2:#06b6d4; /* teal */
        --accent3:#7c3aed; /* purple */
        --glass-border: rgba(0,0,0,0.06);
    }
    html,body,#root, .appview-container {
        background: linear-gradient(180deg, #f6f9fc, #eef6ff) !important;
        color: #0f172a;
        font-family: "Inter", system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial;
    }
    .stApp {
        background: transparent;
    }
    /* Header */
    .header {
        background: linear-gradient(90deg, rgba(59,130,246,0.06), rgba(124,58,237,0.04));
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 12px;
        box-shadow: 0 6px 20px rgba(15,23,42,0.06), inset 0 1px 0 rgba(255,255,255,0.6);
        backdrop-filter: blur(6px) saturate(120%);
        border: 1px solid var(--glass-border);
    }
    .header h1 { margin:0; font-size:1.6rem; }
    .subtle { color: var(--muted); font-size:0.95rem; }
    /* Sidebar card style */
    .side-card {
        background: var(--card);
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 12px;
        border: 1px solid var(--glass-border);
        box-shadow: 0 6px 18px rgba(15,23,42,0.04);
    }
    .uploader {
        padding: 10px;
        border-radius: 10px;
        border: 1px dashed rgba(15,23,42,0.06) !important;
        background: linear-gradient(180deg, rgba(255,255,255,0.8), rgba(255,255,255,0.94));
    }
    .valid-badge {
        display:inline-block;
        background: linear-gradient(90deg, #10b981, #34d399);
        color: white;
        padding: 4px 8px;
        border-radius: 999px;
        font-weight:600;
        font-size:0.8rem;
    }
    .warn-badge {
        display:inline-block;
        background: linear-gradient(90deg, #f97316, #fb923c);
        color: white;
        padding: 4px 8px;
        border-radius: 999px;
        font-weight:600;
        font-size:0.8rem;
    }
    /* Top tabs: make children stretch evenly */
    .stTabs [data-baseweb="tab-list"] {
        display:flex !important;
        gap:8px;
    }
    .stTabs [data-baseweb="tab"] {
        flex:1 1 auto !important;
        text-align:center;
        padding:10px 16px !important;
        border-radius:10px;
        background: transparent;
        font-weight:600;
        color: #0f172a;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, var(--accent1), var(--accent3));
        color: white !important;
        box-shadow: 0 8px 20px rgba(59,130,246,0.12);
    }
    /* Metrics */
    .metric {
        background: linear-gradient(180deg, rgba(255,255,255,0.9), rgba(255,255,255,0.98));
        padding: 12px;
        border-radius: 10px;
        border: 1px solid rgba(15,23,42,0.04);
        text-align:center;
    }
    .metric h3 { margin:0; font-size:1.2rem; }
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
    .primary-btn:hover { transform: translateY(-2px); box-shadow: 0 10px 30px rgba(59,130,246,0.12); }
    /* Dataframe hover */
    .dataframe tr:hover { background: rgba(59,130,246,0.04) !important; }
    /* compact spacing */
    .compact { padding:6px; margin:6px; }
    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------
# Sidebar: parameters & upload simplified
# -------------------------
with st.sidebar:
    st.markdown("<div class='side-card'>", unsafe_allow_html=True)
    st.markdown("<div style='display:flex;align-items:center;gap:12px'><div style='font-size:1.2rem;'>🏭</div><div><strong>Polymer Scheduler</strong><div class='subtle'>Light Material • Desktop</div></div></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='side-card'><strong>📥 Upload</strong><div class='subtle'>Excel workbook with sheets: Plant, Inventory, Demand</div>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["xlsx"], key="uploader", help="Upload the Excel workbook", label_visibility="collapsed")
    if uploaded_file:
        st.markdown("<div style='margin-top:8px;'><span class='valid-badge'>File received</span></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div style='margin-top:8px;'><span class='warn-badge'>No file</span></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='side-card'><strong>⚙️ Parameters</strong>", unsafe_allow_html=True)
    with st.expander("Basic", expanded=True):
        time_limit_min = st.number_input("Time limit (min)", min_value=1, max_value=120, value=10, help="Solver max runtime")
        buffer_days = st.number_input("Buffer days", min_value=0, max_value=7, value=3, help="Extra horizon days (buffer)")
    with st.expander("Objective", expanded=False):
        stockout_penalty = st.number_input("Stockout penalty", min_value=1, value=10)
        transition_penalty = st.number_input("Transition penalty", min_value=1, value=10)
        continuity_bonus = st.number_input("Continuity bonus", min_value=0, value=1)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='side-card'>", unsafe_allow_html=True)
    st.button("⚡ One-click Optimize", key="run_topbar", help="Shortcut to run optimization (same as Optimize tab)")
    st.markdown("</div>", unsafe_allow_html=True)

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
            st.error("Template 'polymer_production_template.xlsx' not found next to app.py.")
    except Exception as e:
        st.error(f"Unable to load template file: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------
# Header + top tabs (Upload / Optimize / Results)
# -------------------------
st.markdown("<div class='header'><h1>Polymer Production Scheduler</h1><div class='subtle'>Light Material • Glassmorphism • Desktop-first — Upload → Optimize → Results</div></div>", unsafe_allow_html=True)
tab_upload, tab_optimize, tab_results = st.tabs(["Upload", "Optimize", "Results"])

# session state flags
if 'uploaded' not in st.session_state:
    st.session_state.uploaded = False
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'last_solution' not in st.session_state:
    st.session_state.last_solution = None
if 'excel_bytes' not in st.session_state:
    st.session_state.excel_bytes = None

# -------------------------
# UPLOAD TAB — merged preview + validation inline
# -------------------------
with tab_upload:
    st.subheader("1) Upload & Preview")
    st.markdown("Upload your Excel workbook (sheets: **Plant**, **Inventory**, **Demand**). Basic validation will be shown below.")
    col_left, col_right = st.columns([2, 1])

    with col_left:
        if uploaded_file:
            try:
                excel_file_bytes = uploaded_file.read()
                st.session_state.excel_bytes = excel_file_bytes
                excel_io = io.BytesIO(excel_file_bytes)

                xls = pd.ExcelFile(excel_io)
                sheets = xls.sheet_names
                have_plant = 'Plant' in sheets
                have_inventory = 'Inventory' in sheets
                have_demand = 'Demand' in sheets

                # Show basic status badges
                st.write("")
                c1, c2, c3 = st.columns(3)
                c1.metric("Plant sheet", "Found" if have_plant else "Missing", "" if have_plant else "Upload template")
                c2.metric("Inventory sheet", "Found" if have_inventory else "Missing", "" if have_inventory else "Upload template")
                c3.metric("Demand sheet", "Found" if have_demand else "Missing", "" if have_demand else "Upload template")

                if not (have_plant and have_inventory and have_demand):
                    st.error("Missing required sheets. Use the template from the sidebar or fix headers.")
                    st.stop()
                else:
                    # load dataframes
                    excel_io.seek(0)
                    plant_df = pd.read_excel(excel_io, sheet_name='Plant')
                    excel_io.seek(0)
                    inventory_df = pd.read_excel(excel_io, sheet_name='Inventory')
                    excel_io.seek(0)
                    demand_df = pd.read_excel(excel_io, sheet_name='Demand', parse_dates=[0])

                    # detect transition sheets
                    excel_io.seek(0)
                    xls2 = pd.ExcelFile(io.BytesIO(excel_file_bytes))
                    transition_dfs = {}
                    for sn in xls2.sheet_names:
                        if sn.lower().startswith("transition"):
                            try:
                                df = pd.read_excel(io.BytesIO(excel_file_bytes), sheet_name=sn, index_col=0)
                                transition_dfs[sn] = df
                            except Exception:
                                transition_dfs[sn] = None

                    st.session_state.uploaded = True
                    st.session_state.data_loaded = True

                    st.markdown("#### Plant")
                    st.dataframe(plant_df, use_container_width=True)
                    st.markdown("#### Inventory")
                    st.dataframe(inventory_df, use_container_width=True)
                    st.markdown("#### Demand (first 20 rows)")
                    st.dataframe(demand_df.head(20), use_container_width=True)

                    # Inline validation checks (merged — no separate validation tab)
                    st.markdown("#### Quick validation")
                    issues = []
                    if 'Plant' in plant_df.columns:
                        if plant_df['Plant'].duplicated().any():
                            issues.append("Duplicate Plant names detected — please ensure Plant identifiers are unique.")
                    if 'Grade Name' in inventory_df.columns:
                        if inventory_df['Grade Name'].isnull().any():
                            issues.append("Empty Grade Name found in Inventory sheet.")
                    if demand_df.shape[1] < 2:
                        issues.append("Demand sheet appears to have no grades (only date column).")

                    if issues:
                        for it in issues:
                            st.warning(it)
                    else:
                        st.success("Basic validation passed. No obvious issues found.")

                    # Show shutdown summary compactly
                    st.markdown("#### Shutdowns (overview)")
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

            except Exception as e:
                st.error(f"Error reading file: {e}")
                st.stop()
        else:
            st.info("Please upload a workbook using the left panel (or download the template).")

    with col_right:
        st.markdown("### Helpful")
        st.markdown("- Use the sample template in the sidebar if unsure of headers.")
        st.markdown("- Files must have sheet names: **Plant**, **Inventory**, **Demand**.")
        st.markdown("- Transitions sheets can be named `Transition_<plant>` (optional).")
        st.markdown("### Quick actions")
        if st.button("Clear uploaded file"):
            st.session_state.uploaded = False
            st.session_state.data_loaded = False
            st.session_state.excel_bytes = None
            st.experimental_rerun()

# -------------------------
# OPTIMIZE TAB — controls + run area
# -------------------------
with tab_optimize:
    st.subheader("2) Optimize")
    st.markdown("Configure solver parameters and run the production optimizer. Progress and status appear below.")

    cols = st.columns([2, 1])
    with cols[0]:
        st.markdown("### Solver settings")
        # show the same input widget values (they are declared in sidebar too)
        st.info("These parameters can also be set in the sidebar (for quick access).")
        time_limit_min = st.number_input("Time limit (min) — main", min_value=1, max_value=120, value=int(time_limit_min), help="Solver max runtime")
        buffer_days = st.number_input("Buffer days — main", min_value=0, max_value=7, value=int(buffer_days), help="Extra horizon days (buffer)")

        st.markdown("### Objective tuning")
        stockout_penalty = st.number_input("Stockout penalty — main", min_value=1, value=int(stockout_penalty))
        transition_penalty = st.number_input("Transition penalty — main", min_value=1, value=int(transition_penalty))
        continuity_bonus = st.number_input("Continuity bonus — main", min_value=0, value=int(continuity_bonus))

    with cols[1]:
        st.markdown("### Quick status")
        if st.session_state.get('data_loaded', False):
            st.success("Data loaded and ready.")
        else:
            st.warning("No data loaded. Upload a workbook first.")

        st.markdown("### Run")
        run_btn = st.button("🎯 Run Optimization (Start)", key="run_main")
        if st.button("Reset last solution"):
            st.session_state.last_solution = None

    # Map topbar quick button to run as well
    if st.session_state.get("run_topbar", False):
        run_btn = True

    if run_btn:
        if not st.session_state.get('data_loaded', False):
            st.error("No data to run. Upload a valid workbook in the Upload tab.")
        else:
            # Begin processing (this preserves all original logic)
            st.session_state.processing = True
            progress_bar = st.progress(0)
            status_text = st.empty()
            status_text.info("Preprocessing data...")

            try:
                excel_bytes_local = st.session_state.excel_bytes
                excel_io = io.BytesIO(excel_bytes_local)

                xls = pd.ExcelFile(excel_io)
                plant_df = pd.read_excel(excel_io, sheet_name='Plant')
                excel_io.seek(0)
                inventory_df = pd.read_excel(excel_io, sheet_name='Inventory')
                excel_io.seek(0)
                demand_df = pd.read_excel(excel_io, sheet_name='Demand', parse_dates=[0])

                # transition sheets (optional)
                excel_io.seek(0)
                xls2 = pd.ExcelFile(io.BytesIO(excel_bytes_local))
                transition_dfs = {}
                for sn in xls2.sheet_names:
                    if sn.lower().startswith("transition"):
                        try:
                            df = pd.read_excel(io.BytesIO(excel_bytes_local), sheet_name=sn, index_col=0)
                            transition_dfs[sn] = df
                        except Exception:
                            transition_dfs[sn] = None

                progress_bar.progress(10)
                status_text.info("Mapping plants, grades, and inventory...")

                # Mapping & parameters (preserved)
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
                        rerun_val = row.get('Rerun Allowed')
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
                    material = row.get('Material Running')
                    expected_days = row.get('Expected Run Days')
                    if pd.notna(material) and pd.notna(expected_days):
                        try:
                            material_running_info[plant] = (str(material).strip(), int(expected_days))
                        except (ValueError, TypeError):
                            st.warning(f"⚠️ Invalid Material Running or Expected Run Days for plant '{plant}'")

                progress_bar.progress(25)
                status_text.info("Preparing demand & horizon...")

                # Demand processing (preserved)
                demand_data = {}
                dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
                num_days = len(dates)
                last_date = dates[-1]
                for i in range(1, int(buffer_days) + 1):
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
                    for date in dates[-int(buffer_days):]:
                        if date not in demand_data[grade]:
                            demand_data[grade][date] = 0

                # Shutdowns and transition rules
                shutdown_periods = process_shutdown_dates(plant_df, dates)
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

                progress_bar.progress(40)
                status_text.info("Building CP-SAT model...")

                # --- Model building (preserved logic) ---
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
                            
                            if d < num_days - int(buffer_days):
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

                # shutdown_demand checks
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

                # Inventory & stockout variables
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

                # Objective & inventory balance
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

                # Soft minimum inventory constraint with deficit variables
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
                    closing_inventory = inventory_vars[(grade, num_days - int(buffer_days))]
                    min_closing = min_closing_inventory[grade]
                    if min_closing > 0:
                        closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                        model.Add(closing_deficit >= min_closing - closing_inventory)
                        model.Add(closing_deficit >= 0)
                        objective += stockout_penalty * closing_deficit * 3

                for grade in grades:
                    for d in range(1, num_days + 1):
                        model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])

                # Production capacity enforcement
                for line in lines:
                    for d in range(num_days - int(buffer_days)):
                        if line in shutdown_periods and d in shutdown_periods[line]:
                            continue
                        production_vars = [
                            get_production_var(grade, line, d) 
                            for grade in grades 
                            if is_allowed_combination(grade, line)
                        ]
                        if production_vars:
                            model.Add(sum(production_vars) == capacities[line])
                    for d in range(num_days - int(buffer_days), num_days):
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

                # Transition rules enforcement
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
                            starts = [is_start_vars[(grade, line, d)] for d in range(num_days) if (grade, line, d) in is_start_vars]
                            if starts:
                                model.Add(sum(starts) <= 1)

                # Stockout penalties
                for grade in grades:
                    for d in range(num_days):
                        objective += stockout_penalty * stockout_vars[(grade, d)]

                # Transition penalties & continuity
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

                progress_bar.progress(60)
                status_text.info("Solver starting...")

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = int(time_limit_min) * 60.0
                solver.parameters.num_search_workers = 8
                solver.parameters.random_seed = 42

                solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

                start_time = time.time()
                status = solver.Solve(model, solution_callback)
                elapsed = time.time() - start_time

                progress_bar.progress(100)
                if status == cp_model.OPTIMAL:
                    status_text.success("✅ Optimization completed optimally")
                elif status == cp_model.FEASIBLE:
                    status_text.success("✅ Optimization found feasible solution")
                else:
                    status_text.warning("⚠️ Solver ended without proven optimal solution")

                # Store last solution for Results tab
                if solution_callback.num_solutions() > 0:
                    st.session_state.last_solution = {
                        'callback': solution_callback,
                        'solver': solver,
                        'grades': grades,
                        'lines': lines,
                        'dates': dates,
                        'formatted_dates': formatted_dates,
                        'num_days': num_days,
                        'buffer_days': int(buffer_days),
                        'production': production,
                        'inventory_vars': inventory_vars,
                        'stockout_vars': stockout_vars,
                        'is_producing': is_producing,
                        'min_inventory': min_inventory,
                        'max_inventory': max_inventory,
                        'min_closing_inventory': min_closing_inventory,
                        'allowed_lines': allowed_lines,
                        'shutdown_periods': shutdown_periods,
                        'transition_dfs': transition_dfs
                    }
                    st.success("Solution saved. View Results tab.")
                else:
                    st.error("No solutions captured by callback.")

            except Exception as e:
                st.error(f"Error during optimization: {e}")
                import traceback
                st.text(traceback.format_exc())
            finally:
                st.session_state.processing = False

# -------------------------
# RESULTS TAB — dashboards & Plotly (preserved)
# -------------------------
with tab_results:
    st.subheader("3) Results")
    if not st.session_state.get('last_solution', None):
        st.info("No solution available. Run optimization from the Optimize tab first.")
    else:
        data = st.session_state.last_solution
        solution_callback = data['callback']
        solver = data['solver']
        grades = data['grades']
        lines = data['lines']
        dates = data['dates']
        formatted_dates = data['formatted_dates']
        num_days = data['num_days']
        buffer_days = data['buffer_days']
        production = data['production']
        inventory_vars = data['inventory_vars']
        stockout_vars = data['stockout_vars']
        is_producing = data['is_producing']
        min_inventory = data['min_inventory']
        max_inventory = data['max_inventory']
        min_closing_inventory = data['min_closing_inventory']
        allowed_lines = data['allowed_lines']
        shutdown_periods = data['shutdown_periods']

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

        # Results tabs across full width
        rtab1, rtab2, rtab3 = st.tabs(["Production Schedule", "Summary", "Inventory"])

        with rtab1:
            st.markdown("### Production schedule (Gantt by line)")
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
                    st.info(f"No production data for {line}.")
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
                if line in shutdown_periods and shutdown_periods[line]:
                    shutdown_days = shutdown_periods[line]
                    start_shutdown = dates[shutdown_days[0]]
                    end_shutdown = dates[shutdown_days[-1]] + timedelta(days=1)
                    fig.add_vrect(
                        x0=start_shutdown,
                        x1=end_shutdown,
                        fillcolor="rgba(239,68,68,0.12)",
                        opacity=0.5,
                        layer="below",
                        line_width=0,
                        annotation_text="Shutdown",
                        annotation_position="top left",
                        annotation_font_size=12,
                        annotation_font_color="#7f1d1d"
                    )
                fig.update_yaxes(autorange="reversed", title=None, showgrid=True)
                fig.update_xaxes(title="Date", tickformat="%d-%b")
                fig.update_layout(height=360, margin=dict(l=60, r=20, t=60, b=60))
                st.plotly_chart(fig, use_container_width=True)

        with rtab2:
            st.markdown("### Production summary")
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
            st.markdown("### Inventory charts")
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
                            fillcolor="rgba(239,68,68,0.12)",
                            opacity=0.5,
                            layer="below",
                            line_width=0,
                            annotation_text=f"Shutdown: {line}" if not shutdown_added else "",
                            annotation_position="top left",
                            annotation_font_size=12,
                            annotation_font_color="#7f1d1d"
                        )
                        shutdown_added = True

                fig.add_hline(
                    y=min_inventory[grade],
                    line=dict(color="rgba(239,68,68,0.8)", width=2, dash="dash"),
                    annotation_text=f"Min: {min_inventory[grade]:,.0f}",
                    annotation_position="top left"
                )
                fig.add_hline(
                    y=max_inventory[grade],
                    line=dict(color="rgba(16,185,129,0.8)", width=2, dash="dash"),
                    annotation_text=f"Max: {max_inventory[grade]:,.0f}",
                    annotation_position="bottom left"
                )

                annotations = [
                    dict(x=start_x, y=start_val, text=f"Start: {start_val:.0f}", showarrow=True, arrowhead=2, ax=-40, ay=30),
                    dict(x=end_x, y=end_val, text=f"End: {end_val:.0f}", showarrow=True, arrowhead=2, ax=40, ay=30),
                    dict(x=highest_x, y=highest_val, text=f"High: {highest_val:.0f}", showarrow=True, arrowhead=2, ax=0, ay=-40),
                    dict(x=lowest_x, y=lowest_val, text=f"Low: {lowest_val:.0f}", showarrow=True, arrowhead=2, ax=0, ay=40),
                ]

                fig.update_layout(
                    title=f"Inventory Level - {grade}",
                    xaxis=dict(title="Date", showgrid=True, tickformat="%d-%b"),
                    yaxis=dict(title="Inventory Volume (MT)", showgrid=True),
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    margin=dict(l=60, r=80, t=60, b=60),
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
                        opacity=0.95
                    )
                st.plotly_chart(fig, use_container_width=True)

# -------------------------
# Footer
# -------------------------
st.markdown("---")
st.markdown("<div style='text-align:center;color:#475569;font-size:0.95rem'>Polymer Production Scheduler • Light Material UI • Solver & plot logic preserved</div>", unsafe_allow_html=True)
