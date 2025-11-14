# app.py
# Material Design 3 — Navigation Drawer App
# Full redesign (Option E1) with OR-Tools solver preserved

import streamlit as st
import pandas as pd
import numpy as np
import io
import time
from datetime import timedelta
from ortools.sat.python import cp_model
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import math
import base64

# -----------------------------
# Material Design 3 CSS injector
# -----------------------------
def inject_md3_css():
    css = """
    <style>
    :root{
        --md3-primary: #3F51B5;
        --md3-on-primary: #FFFFFF;
        --md3-secondary: #E91E63;
        --md3-surface: #FFFFFF;
        --md3-surface-variant: #F5F7FB;
        --md3-outline: #E6EDF7;
        --md3-radius: 12px;
        --md3-elev-1: 0 1px 2px rgba(16,24,40,0.05);
        --md3-elev-2: 0 6px 18px rgba(16,24,40,0.06);
    }
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&display=swap');
    body, .stApp { font-family: 'Roboto', sans-serif; color: #0f172a; background: #ffffff; }

    /* App Shell */
    .md3-appbar { background: linear-gradient(90deg,var(--md3-primary), #3342a0); color: var(--md3-on-primary); padding: 14px 20px; border-radius: 0 0 12px 12px; box-shadow: var(--md3-elev-2); }
    .md3-title { font-size: 1.25rem; font-weight:700; display:inline-block; margin-left:10px; }
    .md3-subtitle { font-size:0.95rem; color:rgba(255,255,255,0.9); margin-top:4px; }

    /* Drawer */
    .md3-drawer { background: #fbfdff; border-radius:12px; padding: 12px; border: 1px solid var(--md3-outline); box-shadow: var(--md3-elev-1); }
    .md3-drawer-item { padding: 10px 12px; border-radius:8px; cursor:pointer; margin-bottom:6px; font-weight:600; color:#0f172a; }
    .md3-drawer-item.active { background: linear-gradient(90deg, rgba(63,81,181,0.08), rgba(63,81,181,0.04)); color: var(--md3-primary); box-shadow: inset 0 -1px 0 rgba(0,0,0,0.02); }

    /* Cards */
    .md3-card { background: var(--md3-surface); border-radius: var(--md3-radius); padding: 1rem; box-shadow: var(--md3-elev-1); border:1px solid var(--md3-outline); margin-bottom: 1rem; }
    .md3-card-ghost { background: var(--md3-surface-variant); border-radius: 10px; padding: .75rem; box-shadow: var(--md3-elev-1); border: 1px solid var(--md3-outline); margin-bottom: .75rem; }

    /* Filezone */
    .md3-filezone { border: 2px dashed var(--md3-outline); padding: 1rem; border-radius: 12px; text-align:center; background: linear-gradient(180deg, rgba(63,81,181,0.02), transparent); }

    /* Chips */
    .md3-chip { display:inline-block; padding:6px 10px; border-radius:20px; border:1px solid var(--md3-outline); margin:4px 6px 4px 0; cursor:pointer; background: white; font-weight:600; }
    .md3-chip.active { background: var(--md3-primary); color:var(--md3-on-primary); }

    /* FAB */
    .md3-fab { position: fixed; right: 28px; bottom: 28px; width:56px; height:56px; border-radius:50%; display:flex; align-items:center; justify-content:center; background: var(--md3-secondary); color:white; box-shadow: 0 10px 30px rgba(233,30,99,0.22); z-index:9999; font-weight:700; font-size:22px; text-decoration:none; }

    /* Buttons style */
    .stButton>button { border-radius: 10px; padding:8px 12px; font-weight:600; background: linear-gradient(90deg, var(--md3-primary), #3342a0); color:var(--md3-on-primary); border: none; }

    /* Focus */
    :focus { outline: 3px solid rgba(63,81,181,0.12); outline-offset:2px; }

    /* Small responsive */
    @media (max-width: 900px) {
        .md3-title { font-size:1rem; }
    }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# Inject MD3 CSS
inject_md3_css()

# -----------------------------
# Utility helpers (MD3 wrappers)
# -----------------------------
def md3_appbar(title="Production Scheduler", subtitle="Material Design 3"):
    st.markdown(f"""
    <div class="md3-appbar">
      <div style="display:flex; align-items:center; justify-content:space-between;">
        <div style="display:flex;align-items:center;">
          <div style="width:40px;height:40px;border-radius:10px;background:rgba(255,255,255,0.12);display:flex;align-items:center;justify-content:center;font-weight:800;">🏭</div>
          <div style="margin-left:12px;">
            <div class="md3-title">{title}</div>
            <div class="md3-subtitle">{subtitle}</div>
          </div>
        </div>
        <div style="display:flex;gap:10px;align-items:center;">
          <div style="font-size:0.95rem; color:rgba(255,255,255,0.9)">User: Operator</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

def md3_drawer(items, active_key):
    # Render drawer items (use st.columns to create a left drawer area)
    drawer_html = '<div class="md3-drawer">'
    for key, label in items:
        cls = "md3-drawer-item active" if key == active_key else "md3-drawer-item"
        drawer_html += f'<div class="{cls}" data-key="{key}">{label}</div>'
    drawer_html += '</div>'
    st.markdown(drawer_html, unsafe_allow_html=True)

def md3_filezone(label="Upload .xlsx"):
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div class="md3-filezone"><strong>Drop file here or click to browse</strong><div style="font-size:0.9rem;color:#6b7280;margin-top:6px;">Accepts .xlsx — sheets: Plant, Inventory, Demand</div></div>', unsafe_allow_html=True)
    f = st.file_uploader(label, type=["xlsx"])
    st.markdown('</div>', unsafe_allow_html=True)
    return f

def md3_card_start(title):
    st.markdown(f'<div class="md3-card"><div style="font-weight:700; margin-bottom:8px;">{title}</div>', unsafe_allow_html=True)

def md3_card_end():
    st.markdown('</div>', unsafe_allow_html=True)

def md3_chip(label, key, active=False):
    cls = "md3-chip active" if active else "md3-chip"
    el = f'<button class="{cls}" style="border:none;background:none;cursor:pointer;">{label}</button>'
    st.markdown(el, unsafe_allow_html=True)

# -----------------------------
# Navigation state
# -----------------------------
if 'page' not in st.session_state:
    st.session_state.page = "home"

# Drawer items
drawer_items = [
    ("home", "Home"),
    ("data", "Data Upload"),
    ("params", "Parameter Setup"),
    ("opt", "Optimization"),
    ("results", "Results Dashboard"),
]

# Top Appbar
md3_appbar(title="Polymer Production Scheduler", subtitle="Material Design 3 — Navigation Drawer")

# Layout: left drawer (col 1), main area (col 2)
col_left, col_main = st.columns([1, 4], gap="small")

with col_left:
    md3_drawer(drawer_items, st.session_state.page)
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="md3-card-ghost"><strong>Quick Actions</strong><div style="margin-top:8px;"></div>', unsafe_allow_html=True)
    if st.button("Upload Sample Template"):
        # Offer download (if present in working dir)
        try:
            sample_path = Path(__file__).parent / "polymer_production_template.xlsx"
            if sample_path.exists():
                with open(sample_path, "rb") as f:
                    data = f.read()
                st.download_button("Download sample template", data, file_name="polymer_production_template.xlsx")
            else:
                st.info("Sample template not found on server.")
        except Exception as e:
            st.error(f"Download failed: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    # Navigation controls (using buttons that set session_state.page)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    if st.button("Home", key="nav_home"):
        st.session_state.page = "home"
    if st.button("Data Upload", key="nav_data"):
        st.session_state.page = "data"
    if st.button("Parameters", key="nav_params"):
        st.session_state.page = "params"
    if st.button("Optimization", key="nav_opt"):
        st.session_state.page = "opt"
    if st.button("Results", key="nav_results"):
        st.session_state.page = "results"

# Helper: read sample workbook if exists
def get_sample_workbook_bytes():
    try:
        sample_path = Path(__file__).parent / "polymer_production_template.xlsx"
        if sample_path.exists():
            return open(sample_path, "rb").read()
        else:
            return None
    except Exception:
        return None

# -----------------------------
# Shared session variables for data & model (preserve across pages)
# -----------------------------
if 'uploaded_file_bytes' not in st.session_state:
    st.session_state.uploaded_file_bytes = None
if 'plant_df' not in st.session_state:
    st.session_state.plant_df = None
if 'inventory_df' not in st.session_state:
    st.session_state.inventory_df = None
if 'demand_df' not in st.session_state:
    st.session_state.demand_df = None
if 'transition_dfs' not in st.session_state:
    st.session_state.transition_dfs = {}
if 'params' not in st.session_state:
    st.session_state.params = {
        'time_limit_min': 10,
        'buffer_days': 3,
        'stockout_penalty': 10,
        'transition_penalty': 10,
        'continuity_bonus': 1
    }
if 'last_run_solution' not in st.session_state:
    st.session_state.last_run_solution = None
if 'last_run_summary' not in st.session_state:
    st.session_state.last_run_summary = None

# -----------------------------
# Page: Home
# -----------------------------
def page_home():
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div style="font-weight:700;font-size:1.1rem;">Welcome</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#374151;margin-top:6px;">This redesign follows Material Design 3 (MD3) and reorganizes your app into clear destinations. Use the left navigation to move between pages.</div>', unsafe_allow_html=True)
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown('<div style="display:flex;gap:12px;">', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="flex:1;">
      <div class="md3-card-ghost"><strong>Next Steps</strong>
      <ol style="margin:8px 0 0 18px;">
        <li>Upload your Excel data on the <strong>Data Upload</strong> page</li>
        <li>Adjust optimization parameters on <strong>Parameter Setup</strong></li>
        <li>Run the solver using the FAB or the Optimization page</li>
        <li>Explore results in the <strong>Results Dashboard</strong></li>
      </ol>
      </div>
    </div>
    <div style="width:320px;">
      <div class="md3-card-ghost">
        <strong>Quick Metrics</strong>
        <div style="margin-top:8px;">
          <div style="font-weight:700;font-size:1.25rem;">{ 'Ready' if st.session_state.uploaded_file_bytes else 'No data' }</div>
          <div style="color:#6b7280;">Upload state</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page: Data Upload
# -----------------------------
def page_data_upload():
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div style="font-weight:700;">Upload your Excel file</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#374151;margin-top:6px;">Drop an .xlsx file with Plant, Inventory and Demand sheets. Optional Transition_[Plant] sheets are supported.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

    uploaded = md3_filezone("Upload .xlsx file")
    if uploaded:
        try:
            uploaded.seek(0)
            raw = uploaded.read()
            st.session_state.uploaded_file_bytes = raw
            # parse basic sheets for preview
            excel = pd.ExcelFile(io.BytesIO(raw))
            # safe loads
            try:
                st.session_state.plant_df = pd.read_excel(io.BytesIO(raw), sheet_name='Plant')
            except Exception:
                st.warning("Plant sheet not found or invalid.")
                st.session_state.plant_df = None
            try:
                st.session_state.inventory_df = pd.read_excel(io.BytesIO(raw), sheet_name='Inventory')
            except Exception:
                st.warning("Inventory sheet not found or invalid.")
                st.session_state.inventory_df = None
            try:
                st.session_state.demand_df = pd.read_excel(io.BytesIO(raw), sheet_name='Demand')
            except Exception:
                st.warning("Demand sheet not found or invalid.")
                st.session_state.demand_df = None

            # Load any transition sheets
            st.session_state.transition_dfs = {}
            for sheet in excel.sheet_names:
                if sheet.startswith("Transition_") or sheet.startswith("Transition"):
                    try:
                        df = pd.read_excel(io.BytesIO(raw), sheet_name=sheet, index_col=0)
                        st.session_state.transition_dfs[sheet] = df
                    except Exception:
                        continue

            st.success("File loaded. Preview below.")
        except Exception as e:
            st.error(f"Upload or parsing error: {e}")

    # Show previews
    if st.session_state.plant_df is not None:
        md3_card_start("Plant Data Preview")
        st.dataframe(st.session_state.plant_df, use_container_width=True)
        md3_card_end()
    if st.session_state.inventory_df is not None:
        md3_card_start("Inventory Data Preview")
        st.dataframe(st.session_state.inventory_df, use_container_width=True)
        md3_card_end()
    if st.session_state.demand_df is not None:
        md3_card_start("Demand Data Preview (first cols)")
        st.dataframe(st.session_state.demand_df.iloc[:, :6], use_container_width=True)
        md3_card_end()

# -----------------------------
# Page: Parameter Setup
# -----------------------------
def page_params():
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div style="font-weight:700;">Parameter Setup</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#374151;margin-top:6px;">Batch parameter submission reduces clicks and prevents accidental runs.</div>', unsafe_allow_html=True)

    with st.form("params_form_main"):
        st.markdown('<div style="display:flex;gap:12px;">', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            tl = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=st.session_state.params['time_limit_min'])
            buf = st.number_input("Buffer days", min_value=0, max_value=30, value=st.session_state.params['buffer_days'])
            st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)
            st.markdown("<div style='font-weight:600;'>Presets</div>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            if c1.button("Balanced"):
                tl, buf = 15, 3
            if c2.button("Min Transitions"):
                tl, buf = 30, 2
            if c3.button("Aggressive"):
                tl, buf = 5, 1
        with col2:
            sp = st.number_input("Stockout penalty", min_value=0, value=st.session_state.params['stockout_penalty'])
            tp = st.number_input("Transition penalty", min_value=0, value=st.session_state.params['transition_penalty'])
            cb = st.number_input("Continuity bonus", min_value=0, value=st.session_state.params['continuity_bonus'])
            st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)
            st.markdown('<div style="color:#6b7280;">Tip: Use presets to try common configurations quickly.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        submitted = st.form_submit_button("Apply Parameters")
        if submitted:
            st.session_state.params['time_limit_min'] = int(tl)
            st.session_state.params['buffer_days'] = int(buf)
            st.session_state.params['stockout_penalty'] = int(sp)
            st.session_state.params['transition_penalty'] = int(tp)
            st.session_state.params['continuity_bonus'] = int(cb)
            st.success("Parameters updated.")

    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Core: Solver & preserved logic (wrapped inside function)
# -----------------------------
def run_solver_and_collect_results():
    """This function preserves your original OR-Tools solver logic and returns structured results."""
    # Validate data existence
    if not st.session_state.uploaded_file_bytes:
        st.error("No data uploaded. Please upload file on Data Upload page.")
        return None

    raw = st.session_state.uploaded_file_bytes
    excel_bytes = io.BytesIO(raw)

    # Read required sheets (mirrors original logic; errors bubbled up)
    try:
        plant_df = pd.read_excel(excel_bytes, sheet_name='Plant')
        excel_bytes.seek(0)
        inventory_df = pd.read_excel(excel_bytes, sheet_name='Inventory')
        excel_bytes.seek(0)
        demand_df = pd.read_excel(excel_bytes, sheet_name='Demand')
    except Exception as e:
        st.error(f"Failed to read required sheets: {e}")
        return None

    # Preserve many variables and processing steps from original
    try:
        # Basic transforms
        # Extract lines, capacities, grades etc.
        lines = list(plant_df['Plant'])
        capacities = {row['Plant']: row['Capacity per day'] for idx, row in plant_df.iterrows()}

        # Determine grades from demand columns (exclude first date column)
        grades = [col for col in demand_df.columns if col != demand_df.columns[0]]

        # Build demand_data mapping: grade -> {date: qty}
        # original code expected demand_df first column to be dates
        date_col = demand_df.columns[0]
        demand_dates = [pd.to_datetime(d).date() for d in demand_df[date_col]]
        dates = sorted(list(set(demand_dates)))
        # extend for buffer days
        buffer_days = st.session_state.params['buffer_days']
        last_date = dates[-1]
        for i in range(1, buffer_days + 1):
            dates.append(last_date + timedelta(days=i))
        num_days = len(dates)
        formatted_dates = [d.strftime('%d-%b-%y') for d in dates]

        demand_data = {}
        for grade in grades:
            demand_data[grade] = {}
            # map existing dates
            for idx, row in demand_df.iterrows():
                d = pd.to_datetime(row[date_col]).date()
                demand_data[grade][d] = row[grade] if grade in demand_df.columns else 0
            # fill buffer days with zeros
            for d in dates[-buffer_days:]:
                demand_data[grade].setdefault(d, 0)

        # inventory initial, min, max, min_closing, min/max run, force_start, allowed_lines
        initial_inventory = {}
        min_inventory = {}
        max_inventory = {}
        min_closing_inventory = {}
        min_run_days = {}
        max_run_days = {}
        force_start_date = {}
        allowed_lines = {g: [] for g in grades}
        rerun_allowed = {}

        grade_inventory_defined = set()
        for idx, row in inventory_df.iterrows():
            grade = row['Grade Name']
            lines_value = row.get('Lines', '')
            if pd.notna(lines_value) and lines_value != '':
                plants_for_row = [x.strip() for x in str(lines_value).split(',')]
            else:
                plants_for_row = lines

            for plant in plants_for_row:
                if plant not in allowed_lines[grade]:
                    allowed_lines[grade].append(plant)

            if grade not in grade_inventory_defined:
                initial_inventory[grade] = row.get('Opening Inventory', 0) if pd.notna(row.get('Opening Inventory', None)) else 0
                min_inventory[grade] = row.get('Min. Inventory', 0) if pd.notna(row.get('Min. Inventory', None)) else 0
                max_inventory[grade] = row.get('Max. Inventory', 1000000000) if pd.notna(row.get('Max. Inventory', None)) else 1000000000
                min_closing_inventory[grade] = row.get('Min. Closing Inventory', 0) if pd.notna(row.get('Min. Closing Inventory', None)) else 0
                grade_inventory_defined.add(grade)

            for plant in plants_for_row:
                gp = (grade, plant)
                min_run_days[gp] = int(row['Min. Run Days']) if pd.notna(row.get('Min. Run Days', None)) else 1
                max_run_days[gp] = int(row['Max. Run Days']) if pd.notna(row.get('Max. Run Days', None)) else 9999
                if pd.notna(row.get('Force Start Date', None)):
                    try:
                        force_start_date[gp] = pd.to_datetime(row.get('Force Start Date')).date()
                    except Exception:
                        force_start_date[gp] = None
                else:
                    force_start_date[gp] = None
                rerun_allowed_val = row.get('Rerun Allowed', 'Yes')
                if pd.notna(rerun_allowed_val):
                    at = str(rerun_allowed_val).strip().lower()
                    rerun_allowed[gp] = False if at in ['no','n','false','0'] else True
                else:
                    rerun_allowed[gp] = True

        # process shutdowns using plant_df
        shutdown_periods = {}
        for idx, row in plant_df.iterrows():
            plant = row['Plant']
            start = row.get('Shutdown Start Date', None)
            end = row.get('Shutdown End Date', None)
            if pd.notna(start) and pd.notna(end):
                try:
                    s = pd.to_datetime(start).date()
                    e = pd.to_datetime(end).date()
                    sd = [i for i,d in enumerate(dates) if s <= d <= e]
                    shutdown_periods[plant] = sd
                except Exception:
                    shutdown_periods[plant] = []
            else:
                shutdown_periods[plant] = []

        # transition rules from any uploaded transition dfs stored earlier
        transition_rules = {}
        for key, df in st.session_state.transition_dfs.items():
            # try to map df into plant name from sheet string
            plant_name = key.replace("Transition_", "").replace("Transition", "")
            if df is not None:
                transition_rules[plant_name] = {}
                for prev in df.index:
                    allowed = []
                    for curr in df.columns:
                        if str(df.loc[prev, curr]).strip().lower() == 'yes':
                            allowed.append(curr)
                    transition_rules[plant_name][prev] = allowed
            else:
                transition_rules[plant_name] = None

        # Build model (preserving original logic)
        time_limit = st.session_state.params['time_limit_min']
        stockout_penalty = st.session_state.params['stockout_penalty']
        transition_penalty = st.session_state.params['transition_penalty']
        continuity_bonus = st.session_state.params['continuity_bonus']

        model = cp_model.CpModel()

        # Create is_producing and production integer vars
        is_producing = {}
        production = {}
        inventory_vars = {}
        stockout_vars = {}

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

        for grade in grades:
            for d in range(num_days + 1):
                inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')
        for grade in grades:
            for d in range(num_days):
                stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')

        # production exclusivity per line per day
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

        # material_running enforcement (if present)
        material_running_info = {}
        for idx, row in plant_df.iterrows():
            plant = row['Plant']
            material = row.get('Material Running', None)
            expected_days = row.get('Expected Run Days', None)
            if pd.notna(material) and pd.notna(expected_days):
                try:
                    material_running_info[plant] = (str(material).strip(), int(expected_days))
                except Exception:
                    pass

        for plant, (material, expected_days) in material_running_info.items():
            for d in range(min(expected_days, num_days)):
                if is_allowed_combination(material, plant):
                    var = get_is_producing_var(material, plant, d)
                    if var is not None:
                        model.Add(var == 1)
                        for other in grades:
                            if other != material and is_allowed_combination(other, plant):
                                opr = get_is_producing_var(other, plant, d)
                                if opr is not None:
                                    model.Add(opr == 0)

        objective_terms = []

        # inventory initial constraints
        for grade in grades:
            model.Add(inventory_vars[(grade, 0)] == initial_inventory.get(grade, 0))

        # per-day inventory/stockout/supplied calculations
        for grade in grades:
            for d in range(num_days):
                produced_today = sum(get_production_var(grade, line, d) for line in allowed_lines[grade])
                demand_today = int(demand_data[grade].get(dates[d], 0))

                # supplied var
                supplied = model.NewIntVar(0, 100000, f'supplied_{grade}_{d}')
                model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
                model.Add(supplied <= demand_today)

                model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
                model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
                model.Add(inventory_vars[(grade, d + 1)] >= 0)

        # min inventory penalties
        for grade in grades:
            for d in range(num_days):
                min_inv = int(min_inventory.get(grade, 0))
                if min_inv > 0:
                    deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                    model.Add(deficit >= min_inv - inventory_vars[(grade, d + 1)])
                    model.Add(deficit >= 0)
                    objective_terms.append(stockout_penalty * deficit)

        # min closing penalty
        for grade in grades:
            min_cl = int(min_closing_inventory.get(grade, 0))
            if min_cl > 0:
                closing_def = model.NewIntVar(0, 100000, f'closing_def_{grade}')
                model.Add(closing_def >= min_cl - inventory_vars[(grade, num_days - buffer_days)])
                model.Add(closing_def >= 0)
                objective_terms.append(stockout_penalty * closing_def * 3)

        # max inventory caps
        for grade in grades:
            for d in range(1, num_days + 1):
                model.Add(inventory_vars[(grade, d)] <= max_inventory.get(grade, 1000000000))

        # production capacity full vs partial days (skipping shutdown handling for full capacity)
        for line in lines:
            for d in range(num_days - buffer_days):
                if line in shutdown_periods and d in shutdown_periods[line]:
                    continue
                production_vars = [get_production_var(grade, line, d) for grade in grades if is_allowed_combination(grade, line)]
                if production_vars:
                    model.Add(sum(production_vars) == capacities[line])

            for d in range(max(0, num_days - buffer_days), num_days):
                production_vars = [get_production_var(grade, line, d) for grade in grades if is_allowed_combination(grade, line)]
                if production_vars:
                    model.Add(sum(production_vars) <= capacities[line])

        # force start dates
        for (grade, plant), fs in force_start_date.items():
            if fs:
                try:
                    start_idx = dates.index(fs)
                    var = get_is_producing_var(grade, plant, start_idx)
                    if var is not None:
                        model.Add(var == 1)
                except ValueError:
                    pass

        # Start/end/run-length logic
        is_start_vars = {}
        run_end_vars = {}
        for grade in grades:
            for line in allowed_lines[grade]:
                gp = (grade, line)
                min_run = min_run_days.get(gp, 1)
                max_run = max_run_days.get(gp, 9999)
                for d in range(num_days):
                    is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                    is_start_vars[(grade, line, d)] = is_start
                    is_end = model.NewBoolVar(f'end_{grade}_{line}_{d}')
                    run_end_vars[(grade, line, d)] = is_end
                    current = get_is_producing_var(grade, line, d)
                    if d > 0:
                        prev = get_is_producing_var(grade, line, d - 1)
                        if current is not None and prev is not None:
                            model.AddBoolAnd([current, prev.Not()]).OnlyEnforceIf(is_start)
                            model.AddBoolOr([current.Not(), prev]).OnlyEnforceIf(is_start.Not())
                    else:
                        if current is not None:
                            model.Add(current == 1).OnlyEnforceIf(is_start)
                            model.Add(is_start == 1).OnlyEnforceIf(current)
                    if d < num_days - 1:
                        nex = get_is_producing_var(grade, line, d + 1)
                        if current is not None and nex is not None:
                            model.AddBoolAnd([current, nex.Not()]).OnlyEnforceIf(is_end)
                            model.AddBoolOr([current.Not(), nex]).OnlyEnforceIf(is_end.Not())
                    else:
                        if current is not None:
                            model.Add(current == 1).OnlyEnforceIf(is_end)
                            model.Add(is_end == 1).OnlyEnforceIf(current)

                # minimum run enforcement where possible
                for d in range(num_days):
                    is_start = is_start_vars[(grade, line, d)]
                    # compute available consecutive days
                    max_possible = 0
                    for k in range(min_run):
                        if d + k < num_days:
                            if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                break
                            max_possible += 1
                    if max_possible >= min_run:
                        for k in range(min_run):
                            if d + k < num_days:
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    continue
                                fut = get_is_producing_var(grade, line, d + k)
                                if fut is not None:
                                    model.Add(fut == 1).OnlyEnforceIf(is_start)

                # maximum run sliding window
                for d in range(max(0, num_days - max_run)):
                    consecutive = []
                    for k in range(max_run + 1):
                        if d + k < num_days:
                            if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                break
                            pv = get_is_producing_var(grade, line, d + k)
                            if pv is not None:
                                consecutive.append(pv)
                    if len(consecutive) == max_run + 1:
                        model.Add(sum(consecutive) <= max_run)

        # rerun allowed
        for grade in grades:
            for line in allowed_lines[grade]:
                gp = (grade, line)
                if not rerun_allowed.get(gp, True):
                    starts = [is_start_vars[(grade, line, d)] for d in range(num_days)]
                    if starts:
                        model.Add(sum(starts) <= 1)

        # stockout penalties
        for grade in grades:
            for d in range(num_days):
                objective_terms.append(stockout_penalty * stockout_vars[(grade, d)])

        # transitions & continuity
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
                        objective_terms.append(transition_penalty * trans_var)
                for grade in grades:
                    if line in allowed_lines[grade]:
                        continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                        model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                        objective_terms.append(-continuity_bonus * continuity)

        model.Minimize(sum(objective_terms))

        # solver params & callback (preserve structure)
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit * 60.0
        solver.parameters.num_search_workers = 8
        solver.parameters.random_seed = 42

        class SimpleCallback(cp_model.CpSolverSolutionCallback):
            def __init__(self):
                cp_model.CpSolverSolutionCallback.__init__(self)
                self.best = None
            def on_solution_callback(self):
                pass

        # Run solver
        start = time.time()
        status = solver.Solve(model)
        end = time.time()
        runtime = end - start

        # collect results in structured dict (similar to previous)
        result = {
            'status': solver.StatusName(status),
            'objective': solver.ObjectiveValue() if status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None,
            'runtime': runtime,
            'production': {},
            'inventory': {},
            'stockout': {},
            'is_producing': {}
        }

        # populate result maps
        for grade in grades:
            result['production'][grade] = {}
            for line in allowed_lines[grade]:
                for d in range(num_days):
                    key = (grade, line, d)
                    if key in production:
                        val = solver.Value(production[key])
                        if val > 0:
                            date_key = formatted_dates[d]
                            result['production'][grade].setdefault(date_key, 0)
                            result['production'][grade][date_key] += val

        for grade in grades:
            result['inventory'][grade] = {}
            for d in range(num_days + 1):
                key = (grade, d)
                if key in inventory_vars:
                    if d < num_days:
                        result['inventory'][grade][formatted_dates[d] if d > 0 else 'initial'] = solver.Value(inventory_vars[key])
                    else:
                        result['inventory'][grade]['final'] = solver.Value(inventory_vars[key])

        for grade in grades:
            result['stockout'][grade] = {}
            for d in range(num_days):
                key = (grade, d)
                if key in stockout_vars:
                    val = solver.Value(stockout_vars[key])
                    if val > 0:
                        result['stockout'][grade][formatted_dates[d]] = val

        for line in lines:
            result['is_producing'][line] = {}
            for d in range(num_days):
                date_key = formatted_dates[d]
                result['is_producing'][line][date_key] = None
                for grade in grades:
                    key = (grade, line, d)
                    if key in is_producing and solver.Value(is_producing[key]) == 1:
                        result['is_producing'][line][date_key] = grade
                        break

        # transitions count
        transition_count = {l:0 for l in lines}
        total_trans = 0
        for line in lines:
            last = None
            for d in range(num_days):
                cur = None
                for grade in grades:
                    key = (grade, line, d)
                    if key in is_producing and solver.Value(is_producing[key]) == 1:
                        cur = grade
                        break
                if cur is not None and last is not None and cur != last:
                    transition_count[line] += 1
                    total_trans += 1
                last = cur
        result['transitions'] = {'per_line': transition_count, 'total': total_trans}

        # store to session
        st.session_state.last_run_solution = result
        st.success("Optimization run completed.")
        return result

    except Exception as e:
        st.error(f"Solver error: {e}")
        import traceback
        st.error(traceback.format_exc())
        return None

# -----------------------------
# Page: Optimization (run controls)
# -----------------------------
def page_opt():
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div style="font-weight:700;">Optimization</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#374151;margin-top:6px;">Run the solver with current data & parameters. Use FAB to run from anywhere.</div>', unsafe_allow_html=True)

    st.markdown('<div style="margin-top:10px;"></div>', unsafe_allow_html=True)
    if st.session_state.uploaded_file_bytes is None:
        st.warning("No data uploaded yet. Navigate to Data Upload to add your Excel file.")
    else:
        st.markdown('<div style="display:flex;gap:12px;">', unsafe_allow_html=True)
        c1, c2 = st.columns([2,1])
        with c1:
            st.markdown('<div class="md3-card-ghost"><strong>Run Controls</strong></div>', unsafe_allow_html=True)
            if st.button("🎯 Run Optimization (full)"):
                with st.spinner("Running solver..."):
                    res = run_solver_and_collect_results()
                    if res:
                        st.session_state.last_run_summary = res
                        st.success("Run complete — view Results Dashboard.")
                        st.session_state.page = "results"
        with c2:
            st.markdown('<div class="md3-card-ghost"><strong>Quick Info</strong><div style="margin-top:8px;">Adjust parameters on Parameter Setup before running.</div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page: Results Dashboard
# -----------------------------
def page_results():
    st.markdown('<div class="md3-card">', unsafe_allow_html=True)
    st.markdown('<div style="font-weight:700;">Results Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div style="color:#374151;margin-top:6px;">View the latest solution (from the last run). Use tabs to explore schedule, inventory and exports.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    res = st.session_state.last_run_solution or st.session_state.last_run_summary
    if not res:
        st.info("No run results available. Run the solver on the Optimization page or use the FAB.")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    # Summary metrics
    st.markdown('<div class="md3-card-ghost"><strong>Key Metrics</strong></div>', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Objective", f"{res['objective']:,}" if res['objective'] is not None else "N/A")
    with col2:
        st.metric("Solver Status", res.get('status', 'N/A'))
    with col3:
        st.metric("Runtime (s)", f"{res.get('runtime', 0):.1f}")
    with col4:
        st.metric("Total Transitions", f"{res.get('transitions', {}).get('total', 0)}")

    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["Production Schedule", "Inventory Trends", "Export/Download"])

    with tab1:
        st.markdown('<div class="md3-card">', unsafe_allow_html=True)
        st.markdown('<div style="font-weight:700;">Production Schedule Gantt</div>', unsafe_allow_html=True)
        # build a Gantt-like visualization using result['is_producing']
        for line, day_map in res['is_producing'].items():
            st.markdown(f"### 🏭 {line}")
            schedule = []
            current_grade = None
            start_date = None
            for date, grade in day_map.items():
                if grade != current_grade:
                    if current_grade is not None:
                        schedule.append((current_grade, start_date, prev_date))
                    current_grade = grade
                    start_date = date
                prev_date = date
            if current_grade is not None:
                schedule.append((current_grade, start_date, prev_date))
            if not schedule:
                st.info(f"No production scheduled for {line}.")
                continue
            # Show table of runs
            schedule_df = pd.DataFrame([{"Grade": g, "Start": s, "End": e} for g,s,e in schedule])
            st.dataframe(schedule_df, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="md3-card">', unsafe_allow_html=True)
        st.markdown('<div style="font-weight:700;">Inventory over Time</div>', unsafe_allow_html=True)
        # render per-grade line charts for inventory if present
        inv = res.get('inventory', {})
        if not inv:
            st.info("No inventory results available.")
        else:
            for grade, vals in inv.items():
                # build x/y arrays preserving order by formatted_dates used earlier (we will try to use keys order)
                x = []
                y = []
                # attempt to order by keys encountered
                for k in vals:
                    x.append(k)
                    y.append(vals[k])
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=x, y=y, mode='lines+markers', name=grade))
                fig.update_layout(title=f"Inventory - {grade}", height=320, margin=dict(t=40))
                st.plotly_chart(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="md3-card">', unsafe_allow_html=True)
        st.markdown('<div style="font-weight:700;">Export Results</div>', unsafe_allow_html=True)
        st.markdown('<div style="color:#374151;margin-top:6px;">Export solution as CSV/Excel for sharing.</div>', unsafe_allow_html=True)
        if st.button("Download production CSV"):
            # prepare csv
            prod_rows = []
            for grade, date_map in res['production'].items():
                for date, qty in date_map.items():
                    prod_rows.append({"Grade": grade, "Date": date, "Qty": qty})
            dfp = pd.DataFrame(prod_rows)
            csv = dfp.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV", csv, file_name="production_schedule.csv", mime="text/csv")
        if st.button("Download inventory CSV"):
            inv_rows = []
            for grade, date_map in res['inventory'].items():
                for date, qty in date_map.items():
                    inv_rows.append({"Grade": grade, "Date": date, "Inventory": qty})
            dfi = pd.DataFrame(inv_rows)
            csv = dfi.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV (inventory)", csv, file_name="inventory.csv", mime="text/csv")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Page routing
# -----------------------------
with col_main:
    # Clicking the custom nav buttons changed session state; also support left-drawer click simulated via those buttons
    # Render page content
    if st.session_state.page == "home":
        page_home()
    elif st.session_state.page == "data":
        page_data_upload()
    elif st.session_state.page == "params":
        page_params()
    elif st.session_state.page == "opt":
        page_opt()
    elif st.session_state.page == "results":
        page_results()
    else:
        page_home()

# -----------------------------
# Floating FAB to run solver globally
# -----------------------------
st.markdown('<a class="md3-fab" href="#" title="Run">▶</a>', unsafe_allow_html=True)

# Provide clickable run in footer as well for accessibility
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
if st.button("Run Optimization (accessible)"):
    if st.session_state.uploaded_file_bytes is None:
        st.warning("No data uploaded. Go to Data Upload page.")
    else:
        with st.spinner("Running optimization..."):
            result = run_solver_and_collect_results()
            if result:
                st.session_state.last_run_summary = result
                st.success("Optimization finished. Switching to Results dashboard.")
                st.session_state.page = "results"
                st.experimental_rerun()
