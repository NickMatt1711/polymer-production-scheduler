# app.py
# Wizard-style (B1) — Full-screen assistant flow
# Material-like styling; preserves original OR-Tools solver & Plotly code.

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

# -----------------------------
# CSS injector (Material-lite)
# -----------------------------
def inject_css():
    css = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    body, .stApp { font-family: 'Roboto', sans-serif; background: #f6f8fb; color:#0f172a; }
    .wizard-shell { max-width:1200px; margin: 18px auto; }
    .md-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 6px 18px rgba(16,24,40,0.06); border: 1px solid #e6edf7; }
    .wizard-steps { display:flex; gap:8px; align-items:center; margin-bottom:16px; }
    .step-pill { padding:8px 12px; border-radius:999px; background:#eef2ff; color:#3342a0; font-weight:600; }
    .step-pill.active { background:linear-gradient(90deg,#3F51B5,#3342a0); color:#fff; box-shadow:0 6px 18px rgba(63,81,181,0.14); }
    .muted { color:#6b7280; font-size:0.95rem; }
    .file-zone { border: 2px dashed #e6edf7; border-radius:12px; padding:28px; text-align:center; background: linear-gradient(180deg, rgba(63,81,181,0.02), transparent); }
    .chip { display:inline-block; padding:6px 12px; border-radius:20px; border:1px solid #e6edf7; margin-right:8px; cursor:pointer; background:#fff; }
    .chip.active { background:#3F51B5; color:#fff; border: none; }
    .kpi { display:inline-block; padding:12px 14px; border-radius:10px; background:#fff; border:1px solid #eef2ff; box-shadow: 0 2px 6px rgba(16,24,40,0.04); margin-right:10px; }
    .footer-actions { display:flex; gap:8px; justify-content:flex-end; margin-top:16px; }
    .md-button { padding:8px 14px; border-radius:10px; background:linear-gradient(90deg,#3F51B5,#3342a0); color:#fff; border:none; font-weight:600; }
    .md-button.secondary { background:#fff; color:#0f172a; border:1px solid #e6edf7; }
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

inject_css()

# -----------------------------
# Wizard steps state
# -----------------------------
STEPS = [
    {"key": "welcome", "label": "Welcome"},
    {"key": "upload", "label": "Upload"},
    {"key": "validate", "label": "Validate"},
    {"key": "params", "label": "Parameters"},
    {"key": "run", "label": "Run"},
    {"key": "results", "label": "Results"},
    {"key": "download", "label": "Download"},
]

if 'step_index' not in st.session_state:
    st.session_state.step_index = 0

# persistent storage for uploaded bytes and parsed tables
if 'uploaded_bytes' not in st.session_state:
    st.session_state.uploaded_bytes = None
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
if 'last_result' not in st.session_state:
    st.session_state.last_result = None

# -----------------------------
# Utility functions
# -----------------------------
def nav_forward():
    if st.session_state.step_index < len(STEPS)-1:
        st.session_state.step_index += 1

def nav_back():
    if st.session_state.step_index > 0:
        st.session_state.step_index -= 1

def go_to_step_key(key):
    for i, s in enumerate(STEPS):
        if s['key'] == key:
            st.session_state.step_index = i
            return

def step_bar():
    parts = []
    for i, s in enumerate(STEPS):
        cls = "step-pill active" if i == st.session_state.step_index else "step-pill"
        parts.append(f'<div class="{cls}">{i}. {s["label"]}</div>')
    html = '<div class="wizard-steps">' + "".join(parts) + '</div>'
    st.markdown(html, unsafe_allow_html=True)

# Safe Excel read helper
def try_read_sheets(bytes_buf):
    excel = pd.ExcelFile(io.BytesIO(bytes_buf))
    plant = inventory = demand = None
    transition_map = {}
    if 'Plant' in excel.sheet_names:
        try:
            plant = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Plant')
        except Exception:
            plant = None
    if 'Inventory' in excel.sheet_names:
        try:
            inventory = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Inventory')
        except Exception:
            inventory = None
    if 'Demand' in excel.sheet_names:
        try:
            demand = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Demand')
        except Exception:
            demand = None
    # read transition sheets if any
    for sheet in excel.sheet_names:
        if sheet.startswith("Transition"):
            try:
                df = pd.read_excel(io.BytesIO(bytes_buf), sheet_name=sheet, index_col=0)
                transition_map[sheet] = df
            except Exception:
                continue
    return plant, inventory, demand, transition_map

# -----------------------------
# SOLVER: preserved core logic (adapted into a function)
# -----------------------------
def run_solver_from_session():
    """
    This function preserves the original OR-Tools model logic.
    It reads session_state uploaded bytes and parameters, and returns a result dict.
    """
    bytes_buf = st.session_state.uploaded_bytes
    if not bytes_buf:
        st.error("No file uploaded.")
        return None

    try:
        excel = pd.ExcelFile(io.BytesIO(bytes_buf))
        plant_df = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Plant')
        inventory_df = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Inventory')
        demand_df = pd.read_excel(io.BytesIO(bytes_buf), sheet_name='Demand')
    except Exception as e:
        st.error(f"Error reading sheets: {e}")
        return None

    # Keep copies in session for inspection
    st.session_state.plant_df = plant_df
    st.session_state.inventory_df = inventory_df
    st.session_state.demand_df = demand_df

    # Basic derived data (mirrors original structure)
    lines = list(plant_df['Plant'])
    capacities = {row['Plant']: row['Capacity per day'] for idx, row in plant_df.iterrows()}

    # Determine grades from demand columns (exclude first date column)
    date_col = demand_df.columns[0]
    grades = [col for col in demand_df.columns if col != date_col]

    # Build date list and extend buffer days
    demand_dates = [pd.to_datetime(d).date() for d in demand_df[date_col]]
    dates = sorted(list(set(demand_dates)))
    buffer_days = st.session_state.params['buffer_days']
    last_date = dates[-1]
    for i in range(1, buffer_days + 1):
        dates.append(last_date + timedelta(days=i))
    num_days = len(dates)
    formatted_dates = [d.strftime('%d-%b-%y') for d in dates]

    # Build demand_data per grade
    demand_data = {}
    for grade in grades:
        demand_data[grade] = {}
        for idx, row in demand_df.iterrows():
            d = pd.to_datetime(row[date_col]).date()
            demand_data[grade][d] = int(row[grade]) if grade in demand_df.columns and not pd.isna(row[grade]) else 0
        # Add buffer days with zero demand
        for d in dates[-buffer_days:]:
            demand_data[grade].setdefault(d, 0)

    # Inventory and per-grade settings
    initial_inventory = {}
    min_inventory = {}
    max_inventory = {}
    min_closing_inventory = {}
    min_run_days = {}
    max_run_days = {}
    force_start_date = {}
    allowed_lines = {g: list(lines) for g in grades}  # default allow all
    rerun_allowed = {}

    # Try to parse inventory sheet rows
    for idx, row in inventory_df.iterrows():
        grade = row['Grade Name']
        lines_val = row.get('Lines', '')
        if pd.notna(lines_val) and str(lines_val).strip() != '':
            plants_for_row = [x.strip() for x in str(lines_val).split(',')]
        else:
            plants_for_row = lines
        # set allowed lines
        allowed_lines[grade] = plants_for_row
        initial_inventory[grade] = row.get('Opening Inventory', 0) if pd.notna(row.get('Opening Inventory', None)) else 0
        min_inventory[grade] = row.get('Min. Inventory', 0) if pd.notna(row.get('Min. Inventory', None)) else 0
        max_inventory[grade] = row.get('Max. Inventory', 1000000000) if pd.notna(row.get('Max. Inventory', None)) else 1000000000
        min_closing_inventory[grade] = row.get('Min. Closing Inventory', 0) if pd.notna(row.get('Min. Closing Inventory', None)) else 0

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
            rerun_val = row.get('Rerun Allowed', 'Yes')
            if pd.notna(rerun_val):
                v = str(rerun_val).strip().lower()
                rerun_allowed[gp] = False if v in ['no','n','false','0'] else True
            else:
                rerun_allowed[gp] = True

    # Shutdown processing from plant_df
    shutdown_periods = {}
    for idx, row in plant_df.iterrows():
        plant = row['Plant']
        start = row.get('Shutdown Start Date', None)
        end = row.get('Shutdown End Date', None)
        if pd.notna(start) and pd.notna(end):
            try:
                s = pd.to_datetime(start).date()
                e = pd.to_datetime(end).date()
                shutdown_days = [i for i, d in enumerate(dates) if s <= d <= e]
                shutdown_periods[plant] = shutdown_days
            except Exception:
                shutdown_periods[plant] = []
        else:
            shutdown_periods[plant] = []

    # Transition rules: read any transition dfs loaded earlier in session
    transition_rules = {}
    for sheet_name, df in st.session_state.transition_dfs.items():
        # map sheet to plant name heuristically
        plant_name = sheet_name.replace("Transition_", "").replace("Transition", "")
        transition_rules[plant_name] = {}
        if df is not None:
            for prev in df.index:
                allowed = []
                for curr in df.columns:
                    if str(df.loc[prev, curr]).strip().lower() == 'yes':
                        allowed.append(curr)
                transition_rules[plant_name][prev] = allowed

    # Parameters
    time_limit = st.session_state.params['time_limit_min']
    stockout_penalty = st.session_state.params['stockout_penalty']
    transition_penalty = st.session_state.params['transition_penalty']
    continuity_bonus = st.session_state.params['continuity_bonus']

    # Build OR-Tools model (preserve earlier logic)
    model = cp_model.CpModel()

    is_producing = {}
    production = {}
    inventory_vars = {}
    stockout_vars = {}

    def is_allowed_combination(grade, line):
        return line in allowed_lines.get(grade, [])

    # Create variables
    for grade in grades:
        for line in allowed_lines[grade]:
            for d in range(num_days):
                key = (grade, line, d)
                is_producing[key] = model.NewBoolVar(f'isprod_{grade}_{line}_{d}')
                if d < num_days - buffer_days:
                    prod_val = model.NewIntVar(0, capacities[line], f'prod_{grade}_{line}_{d}')
                    model.Add(prod_val == capacities[line]).OnlyEnforceIf(is_producing[key])
                    model.Add(prod_val == 0).OnlyEnforceIf(is_producing[key].Not())
                else:
                    prod_val = model.NewIntVar(0, capacities[line], f'prod_{grade}_{line}_{d}')
                    model.Add(prod_val <= capacities[line] * is_producing[key])
                production[key] = prod_val

    for grade in grades:
        for d in range(num_days + 1):
            inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inv_{grade}_{d}')
    for grade in grades:
        for d in range(num_days):
            stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'sout_{grade}_{d}')

    # One grade per line per day
    for line in lines:
        for d in range(num_days):
            candidates = []
            for grade in grades:
                if is_allowed_combination(grade, line):
                    v = is_producing.get((grade, line, d))
                    if v is not None:
                        candidates.append(v)
            if candidates:
                model.Add(sum(candidates) <= 1)

    # material running enforcement if present
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
    for plant, (material, expected) in material_running_info.items():
        for d in range(min(expected, num_days)):
            if is_allowed_combination(material, plant):
                var = is_producing.get((material, plant, d))
                if var is not None:
                    model.Add(var == 1)
                    for other in grades:
                        if other != material and is_allowed_combination(other, plant):
                            op = is_producing.get((other, plant, d))
                            if op is not None:
                                model.Add(op == 0)

    objective_terms = []

    # initial inventory
    for grade in grades:
        model.Add(inventory_vars[(grade, 0)] == initial_inventory.get(grade, 0))

    # daily balance
    for grade in grades:
        for d in range(num_days):
            produced_today = sum(production.get((grade, line, d), 0) for line in allowed_lines[grade])
            demand_today = int(demand_data[grade].get(dates[d], 0))
            supplied = model.NewIntVar(0, 100000, f'sup_{grade}_{d}')
            model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
            model.Add(supplied <= demand_today)
            model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
            model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
            model.Add(inventory_vars[(grade, d + 1)] >= 0)

    # min inventory deficits as penalty
    for grade in grades:
        for d in range(num_days):
            min_inv = int(min_inventory.get(grade, 0))
            if min_inv > 0:
                deficit = model.NewIntVar(0, 100000, f'def_{grade}_{d}')
                model.Add(deficit >= min_inv - inventory_vars[(grade, d + 1)])
                model.Add(deficit >= 0)
                objective_terms.append(stockout_penalty * deficit)

    # min closing inventory penalty
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

    # production capacity constraints
    for line in lines:
        for d in range(num_days - buffer_days):
            if line in shutdown_periods and d in shutdown_periods[line]:
                continue
            prod_vars = [production.get((grade, line, d), 0) for grade in grades if is_allowed_combination(grade, line)]
            if prod_vars:
                model.Add(sum(prod_vars) == capacities[line])

        for d in range(max(0, num_days - buffer_days), num_days):
            prod_vars = [production.get((grade, line, d), 0) for grade in grades if is_allowed_combination(grade, line)]
            if prod_vars:
                model.Add(sum(prod_vars) <= capacities[line])

    # force start dates
    for (grade, plant), fs in force_start_date.items():
        if fs:
            try:
                idx = dates.index(fs)
                var = is_producing.get((grade, plant, idx))
                if var is not None:
                    model.Add(var == 1)
            except ValueError:
                pass

    # start/end/run-length and rerun allowed
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
                curr = is_producing.get((grade, line, d))
                if d > 0:
                    prev = is_producing.get((grade, line, d - 1))
                    if curr is not None and prev is not None:
                        model.AddBoolAnd([curr, prev.Not()]).OnlyEnforceIf(is_start)
                        model.AddBoolOr([curr.Not(), prev]).OnlyEnforceIf(is_start.Not())
                else:
                    if curr is not None:
                        model.Add(curr == 1).OnlyEnforceIf(is_start)
                        model.Add(is_start == 1).OnlyEnforceIf(curr)
                if d < num_days - 1:
                    nex = is_producing.get((grade, line, d + 1))
                    if curr is not None and nex is not None:
                        model.AddBoolAnd([curr, nex.Not()]).OnlyEnforceIf(is_end)
                        model.AddBoolOr([curr.Not(), nex]).OnlyEnforceIf(is_end.Not())
                else:
                    if curr is not None:
                        model.Add(curr == 1).OnlyEnforceIf(is_end)
                        model.Add(is_end == 1).OnlyEnforceIf(curr)

            # minimum-run enforcement where possible (avoid shutdown days)
            for d in range(num_days):
                is_start = is_start_vars[(grade, line, d)]
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
                            fut = is_producing.get((grade, line, d + k))
                            if fut is not None:
                                model.Add(fut == 1).OnlyEnforceIf(is_start)

            # max-run sliding window
            for d in range(max(0, num_days - max_run)):
                consecutive = []
                for k in range(max_run + 1):
                    if d + k < num_days:
                        if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                            break
                        v = is_producing.get((grade, line, d + k))
                        if v is not None:
                            consecutive.append(v)
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

    # Solver parameters
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit * 60.0
    solver.parameters.num_search_workers = 8
    solver.parameters.random_seed = 42

    with st.spinner("Solving... This may take a moment depending on model size."):
        start_time = time.time()
        status = solver.Solve(model)
        wall_time = time.time() - start_time

    # Build result structure
    result = {
        'status': solver.StatusName(status),
        'objective': (solver.ObjectiveValue() if status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None),
        'runtime': wall_time,
        'production': {},
        'inventory': {},
        'stockout': {},
        'is_producing': {}
    }

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

    # Transition counting
    transition_count = {l: 0 for l in lines}
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

    # Save to session
    st.session_state.last_result = result
    return result

# -----------------------------
# Wizard page implementations
# -----------------------------
def page_welcome():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h2>Welcome — Production Scheduler Assistant</h2>', unsafe_allow_html=True)
    st.markdown('<div class="muted">A guided flow will take you from uploading data to running optimization and downloading results. Click "Next" to begin.</div>', unsafe_allow_html=True)
    st.markdown('<hr>', unsafe_allow_html=True)
    st.markdown('<div style="display:flex;gap:12px;">', unsafe_allow_html=True)
    st.markdown('<div style="flex:1;"><div class="kpi"><div style="font-weight:700;">Preserve Solver</div><div class="muted">All OR-Tools logic kept intact</div></div></div>', unsafe_allow_html=True)
    st.markdown('<div style="flex:1;"><div class="kpi"><div style="font-weight:700;">Material UX</div><div class="muted">Wizard-style assistance for desktop</div></div></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    # Footer actions
    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("Next →", key="welcome_next"):
            nav_forward()

def page_upload():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 1 — Upload Excel Data</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Upload an .xlsx file containing Plant, Inventory and Demand sheets. Optional Transition_* sheets supported.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)

    st.markdown('<div class="file-zone">', unsafe_allow_html=True)
    uploaded = st.file_uploader("Drop or click to select Excel file", type=["xlsx"])
    if uploaded:
        try:
            uploaded.seek(0)
            raw = uploaded.read()
            st.session_state.uploaded_bytes = raw
            st.success("File uploaded into session.")
            # quick parse preview
            plant, inv, dem, transitions = try_read_sheets(raw)
            st.session_state.plant_df = plant
            st.session_state.inventory_df = inv
            st.session_state.demand_df = dem
            st.session_state.transition_dfs = transitions
        except Exception as e:
            st.error(f"Upload error: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    # Preview (compact)
    if st.session_state.plant_df is not None:
        st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
        st.markdown('<div class="md-card"><strong>Plant (preview)</strong></div>', unsafe_allow_html=True)
        st.dataframe(st.session_state.plant_df.iloc[:8, :], use_container_width=True)
    if st.session_state.inventory_df is not None:
        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        st.markdown('<div class="md-card"><strong>Inventory (preview)</strong></div>', unsafe_allow_html=True)
        st.dataframe(st.session_state.inventory_df.iloc[:8, :], use_container_width=True)
    if st.session_state.demand_df is not None:
        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        st.markdown('<div class="md-card"><strong>Demand (preview)</strong></div>', unsafe_allow_html=True)
        st.dataframe(st.session_state.demand_df.iloc[:8, :6], use_container_width=True)

    # navigation
    col1, col2, col3 = st.columns([1,1,1])
    with col1:
        if st.button("Back", key="upload_back"):
            nav_back()
    with col2:
        if st.button("Validate →", key="upload_validate"):
            nav_forward()
    with col3:
        if st.button("Skip Validate →", key="upload_skip_validate"):
            # allow skipping validation if user wants to proceed
            nav_forward()

def page_validate():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 2 — Validate Data</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">The assistant will run quick checks — missing sheets, date types, negative demand, reasonable capacities.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)

    issues = []
    if st.session_state.plant_df is None:
        issues.append("Plant sheet missing or unreadable.")
    if st.session_state.inventory_df is None:
        issues.append("Inventory sheet missing or unreadable.")
    if st.session_state.demand_df is None:
        issues.append("Demand sheet missing or unreadable.")

    # quick content checks
    try:
        if st.session_state.demand_df is not None:
            date_col = st.session_state.demand_df.columns[0]
            if not pd.api.types.is_datetime64_any_dtype(st.session_state.demand_df[date_col]):
                issues.append("Demand date column is not a recognized datetime column.")
            # negative demand
            for col in st.session_state.demand_df.columns[1:]:
                if (st.session_state.demand_df[col].fillna(0) < 0).any():
                    issues.append(f"Negative demand values detected in column '{col}'.")
    except Exception:
        issues.append("Demand data validation raised an exception — please inspect the sheet.")

    try:
        if st.session_state.plant_df is not None:
            if 'Capacity per day' not in st.session_state.plant_df.columns:
                issues.append("'Capacity per day' column missing in Plant sheet.")
    except Exception:
        issues.append("Plant data validation raised an exception — please inspect the sheet.")

    if issues:
        st.markdown('<div class="md-card"><strong>Issues found</strong></div>', unsafe_allow_html=True)
        for it in issues:
            st.error(it)
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("Back — Fix Upload", key="validate_back"):
                nav_back()
        with col2:
            if st.button("Proceed Anyway →", key="validate_forced_next"):
                nav_forward()
    else:
        st.success("No immediate issues found. Data looks good.")
        col1, col2 = st.columns([1,1])
        with col1:
            if st.button("Back", key="validate_back_ok"):
                nav_back()
        with col2:
            if st.button("Next →", key="validate_next_ok"):
                nav_forward()

def page_params():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 3 — Configure Parameters</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Batch parameter submission reduces accidental re-runs. Use presets for quick starts.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)

    with st.form("params_form"):
        col1, col2 = st.columns(2)
        with col1:
            tl = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=st.session_state.params['time_limit_min'])
            buf = st.number_input("Buffer days", min_value=0, max_value=30, value=st.session_state.params['buffer_days'])
            st.write("Presets")
            p1, p2, p3 = st.columns(3)
            if p1.button("Balanced"):
                tl, buf = 15, 3
            if p2.button("Min-Transitions"):
                tl, buf = 30, 2
            if p3.button("Aggressive"):
                tl, buf = 5, 1
        with col2:
            sp = st.number_input("Stockout penalty", min_value=0, value=st.session_state.params['stockout_penalty'])
            tp = st.number_input("Transition penalty", min_value=0, value=st.session_state.params['transition_penalty'])
            cb = st.number_input("Continuity bonus", min_value=0, value=st.session_state.params['continuity_bonus'])
            st.markdown('<div class="muted">Tip: increase transition penalty to discourage grade changes.</div>', unsafe_allow_html=True)
        submitted = st.form_submit_button("Apply & Continue")
        if submitted:
            st.session_state.params['time_limit_min'] = int(tl)
            st.session_state.params['buffer_days'] = int(buf)
            st.session_state.params['stockout_penalty'] = int(sp)
            st.session_state.params['transition_penalty'] = int(tp)
            st.session_state.params['continuity_bonus'] = int(cb)
            st.success("Parameters saved.")
            nav_forward()

    # footer nav
    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("Back", key="params_back"):
            nav_back()

def page_run():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 4 — Run Optimization</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Start the solver. You will be taken to Results when run completes.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    if st.session_state.uploaded_bytes is None:
        st.warning("No uploaded data found. Please go back to Upload step.")
        if st.button("Back to Upload"):
            go_to_step_key('upload')
        return

    # show params summary
    st.markdown('<div class="md-card"><strong>Parameters</strong></div>', unsafe_allow_html=True)
    st.write(st.session_state.params)

    # run controls
    if st.button("Run Optimization", key="run_now"):
        # execute solver function
        res = run_solver_from_session()
        if res:
            st.success("Run complete.")
            go_to_step_key('results')

    # navigation
    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("Back", key="run_back"):
            nav_back()

def page_results():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 5 — Results Dashboard</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Review key metrics, Gantt, inventory trends and tables. If desired, re-run with adjusted parameters.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    res = st.session_state.last_result
    if not res:
        st.info("No results available. Run the optimizer to produce results.")
        if st.button("Go to Run"):
            go_to_step_key('run')
        return

    # KPIs
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Objective", f"{res['objective']:,}" if res['objective'] is not None else "N/A")
    with col2:
        st.metric("Status", res.get('status', 'N/A'))
    with col3:
        st.metric("Runtime (s)", f"{res.get('runtime', 0):.1f}")
    with col4:
        st.metric("Transitions", f"{res.get('transitions', {}).get('total', 0)}")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    # Production Gantt-like view
    st.markdown('<div class="md-card"><strong>Production Schedule (per line)</strong></div>', unsafe_allow_html=True)
    for line, day_map in res['is_producing'].items():
        st.markdown(f"### 🏭 {line}")
        schedule = []
        cur = None
        start = None
        prev = None
        for date, grade in day_map.items():
            if grade != cur:
                if cur is not None:
                    schedule.append((cur, start, prev))
                cur = grade
                start = date
            prev = date
        if cur is not None:
            schedule.append((cur, start, prev))
        if not schedule:
            st.info(f"No schedule for {line}.")
            continue
        sched_df = pd.DataFrame([{"Grade": g, "Start": s, "End": e} for g,s,e in schedule])
        st.dataframe(sched_df, use_container_width=True)

    # Inventory charts
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
    st.markdown('<div class="md-card"><strong>Inventory Trends</strong></div>', unsafe_allow_html=True)
    inv = res.get('inventory', {})
    if inv:
        for grade, mapping in inv.items():
            x = list(mapping.keys())
            y = list(mapping.values())
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=x, y=y, mode='lines+markers', name=grade))
            fig.update_layout(title=f"Inventory - {grade}", height=320, margin=dict(t=40))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No inventory data to display.")

    # navigation
    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("Back", key="results_back"):
            nav_back()
    with col2:
        if st.button("Next →", key="results_next"):
            nav_forward()

def page_download():
    st.markdown('<div class="md-card">', unsafe_allow_html=True)
    st.markdown('<h3>Step 6 — Download Outputs</h3>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Export CSVs or Excel for production schedule and inventory. Restart the wizard when done.</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)

    res = st.session_state.last_result
    if not res:
        st.info("No results available to download.")
    else:
        # production CSV
        prod_rows = []
        for grade, mp in res['production'].items():
            for date, qty in mp.items():
                prod_rows.append({"Grade": grade, "Date": date, "Qty": qty})
        prod_df = pd.DataFrame(prod_rows)
        csv_prod = prod_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Production CSV", csv_prod, file_name="production_schedule.csv", mime="text/csv")

        # inventory CSV
        inv_rows = []
        for grade, mp in res['inventory'].items():
            for date, qty in mp.items():
                inv_rows.append({"Grade": grade, "Date": date, "Inventory": qty})
        inv_df = pd.DataFrame(inv_rows)
        csv_inv = inv_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Inventory CSV", csv_inv, file_name="inventory.csv", mime="text/csv")

    st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("Restart Wizard"):
            # clear session but keep uploaded bytes maybe? we'll reset everything
            st.session_state.step_index = 0
            st.session_state.uploaded_bytes = None
            st.session_state.plant_df = None
            st.session_state.inventory_df = None
            st.session_state.demand_df = None
            st.session_state.transition_dfs = {}
            st.session_state.last_result = None
            st.success("Wizard restarted.")
    with col2:
        if st.button("Back", key="download_back"):
            nav_back()

# -----------------------------
# Page router
# -----------------------------
st.markdown('<div class="wizard-shell">', unsafe_allow_html=True)
step_bar()
st.markdown('<div style="height:12px"></div>', unsafe_allow_html=True)
# container card
if st.session_state.step_index == 0:
    page_welcome()
elif st.session_state.step_index == 1:
    page_upload()
elif st.session_state.step_index == 2:
    page_validate()
elif st.session_state.step_index == 3:
    page_params()
elif st.session_state.step_index == 4:
    page_run()
elif st.session_state.step_index == 5:
    page_results()
elif st.session_state.step_index == 6:
    page_download()
else:
    page_welcome()
st.markdown('</div>', unsafe_allow_html=True)

# Bottom quick nav (always available)
footer_cols = st.columns([1,1,1,1])
with footer_cols[0]:
    if st.button("◀ Back", key="footer_back"):
        nav_back()
with footer_cols[1]:
    if st.button("Next ▶", key="footer_next"):
        nav_forward()
with footer_cols[2]:
    if st.button("Go to Results", key="footer_results"):
        go_to_step_key('results')
with footer_cols[3]:
    if st.button("Run Now (Quick)", key="footer_run"):
        if st.session_state.uploaded_bytes is None:
            st.warning("No uploaded data.")
        else:
            res = run_solver_from_session()
            if res:
                go_to_step_key('results')
