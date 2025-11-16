# app_rewrite.py
"""
Single-file Streamlit app:
- UI inspired by test.py (clean, step-by-step)
- Solver (CP-SAT) inspired by app.py but improved:
    * "No" in transition sheets => hard forbidden transitions
    * Transition penalty applies uniformly for any allowed transition
    * Continuity bonus removed
    * Min/Max run days strictly enforced
    * Rerun restriction enforced
    * Force start dates enforced
    * Shutdown periods block production
    * Inventory & stockout variables with penalties
- Plotly visuals kept (Gantt, inventory, etc.)
- NOT modular (single file) — per your request
"""

import io
import time
from datetime import datetime, timedelta
import math
from typing import Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

import streamlit as st
from ortools.sat.python import cp_model

# -------------------------
# UI Styling (test.py-like)
# -------------------------
st.set_page_config(page_title="Polymer Production Scheduler (Rebuilt)",
                   layout="wide")

st.markdown(
    """
    <style>
    /* Simplified test.py-like styling */
    .app-bar { background: linear-gradient(135deg,#667eea 0%,#764ba2 100%); color: white; padding: 1.2rem; border-radius: 8px; margin-bottom: 1rem;}
    .metric-card { padding: 1rem; border-radius: 8px; background: white; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
    .section-header { font-size: 1.1rem; font-weight: 700; margin: .5rem 0; }
    .info-box { padding: .5rem; border-radius: 6px; background: #e9f2ff; border-left: 4px solid #2196f3; margin: .5rem 0; }
    .alert-box { padding: .6rem; border-radius: 6px; background: #fff4e6; border-left: 4px solid #ff9800; margin: .5rem 0; }
    .success-box { padding: .6rem; border-radius: 6px; background: #e8f5e9; border-left: 4px solid #4caf50; margin: .5rem 0; }
    .divider { height: 1px; background: #e9e9e9; margin: 1rem 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="app-bar"><h2>🏭 Polymer Production Scheduler — Rebuilt UI + Solver</h2></div>', unsafe_allow_html=True)

# -------------------------
# Utility functions
# -------------------------
def read_excel_sheets(excel_bytes: io.BytesIO) -> Dict[str, pd.DataFrame]:
    """
    Safely read expected sheets. Return dict with DataFrames or None if missing.
    Expected sheets: Plant, Inventory, Demand, Transition_<PlantName> (optional)
    """
    excel_bytes.seek(0)
    xl = pd.ExcelFile(excel_bytes)
    sheets = {}
    for name in xl.sheet_names:
        try:
            sheets[name] = xl.parse(name)
        except Exception:
            sheets[name] = None
    return sheets


def get_possible_transition_sheet_names(plant_name: str) -> List[str]:
    return [
        f"Transition_{plant_name}",
        f"Transition_{plant_name.replace(' ', '_')}",
        f"Transition{plant_name.replace(' ', '')}",
    ]


def process_shutdown_dates(plant_df: pd.DataFrame, dates: List[datetime.date]) -> Dict[str, List[int]]:
    """
    Returns dict plant -> list of day indices that are shutdown
    """
    shutdown_periods = {}
    for _, row in plant_df.iterrows():
        plant = str(row['Plant']).strip()
        shutdown_days = []
        start = row.get('Shutdown Start Date', None)
        end = row.get('Shutdown End Date', None)
        if pd.notna(start) and pd.notna(end):
            try:
                start_dt = pd.to_datetime(start).date()
                end_dt = pd.to_datetime(end).date()
                for idx, d in enumerate(dates):
                    if start_dt <= d <= end_dt:
                        shutdown_days.append(idx)
            except Exception:
                shutdown_days = []
        shutdown_periods[plant] = shutdown_days
    return shutdown_periods


def load_transition_matrices(excel_bytes: io.BytesIO, plant_df: pd.DataFrame) -> Dict[str, Optional[pd.DataFrame]]:
    excel_bytes.seek(0)
    transition_dfs = {}
    for _, row in plant_df.iterrows():
        plant_name = str(row['Plant']).strip()
        df_found = None
        for sname in get_possible_transition_sheet_names(plant_name):
            try:
                excel_bytes.seek(0)
                df = pd.read_excel(excel_bytes, sheet_name=sname, index_col=0)
                # normalize index/columns to string
                df.index = df.index.astype(str).str.strip()
                df.columns = df.columns.astype(str).str.strip()
                df_found = df
                break
            except Exception:
                continue
        transition_dfs[plant_name] = df_found
    return transition_dfs


def safe_int(x, default=0):
    try:
        return int(x)
    except Exception:
        return default


# -------------------------
# Stepper UI - Upload
# -------------------------
if 'step' not in st.session_state:
    st.session_state.step = 1

col1, col2 = st.columns([3,1])

with col1:
    st.markdown("### 1) Upload Excel input")
    uploaded_file = st.file_uploader("Upload the Excel workbook (Plant, Inventory, Demand, Transition_...)", type=["xlsx"])
with col2:
    st.markdown("### Controls")
    if st.button("Reset App"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.experimental_rerun()

if uploaded_file:
    st.session_state.uploaded_file = uploaded_file
    st.session_state.step = 2

# -------------------------
# Step 2 - Preview & Configure
# -------------------------
if st.session_state.step >= 2 and 'uploaded_file' in st.session_state:
    excel_bytes = io.BytesIO(st.session_state.uploaded_file.read())
    sheets = read_excel_sheets(excel_bytes)

    # Validate and preview required sheets
    st.markdown('<div class="section-header">📄 Data Preview</div>', unsafe_allow_html=True)
    try:
        plant_df = sheets.get('Plant', None)
        inventory_df = sheets.get('Inventory', None)
        demand_df = sheets.get('Demand', None)
    except Exception as e:
        st.error(f"Error reading sheets: {e}")
        st.stop()

    if plant_df is None or inventory_df is None or demand_df is None:
        st.error("Missing required sheets. Please ensure 'Plant', 'Inventory', and 'Demand' sheets exist.")
        st.stop()

    # Show previews in three columns
    pcol1, pcol2, pcol3 = st.columns(3)
    with pcol1:
        st.subheader("🏭 Plant")
        st.dataframe(plant_df.head(50), use_container_width=True)
    with pcol2:
        st.subheader("📦 Inventory")
        st.dataframe(inventory_df.head(50), use_container_width=True)
    with pcol3:
        st.subheader("📊 Demand (first 50 rows)")
        st.dataframe(demand_df.head(50), use_container_width=True)

    # Basic parsing
    # Dates list
    try:
        raw_dates = pd.to_datetime(demand_df.iloc[:, 0]).dt.date.tolist()
    except Exception as e:
        st.error("Could not parse dates in Demand sheet first column. Ensure it's a proper date column.")
        st.stop()

    dates = sorted(list(dict.fromkeys(raw_dates)))  # keep order but unique
    if len(dates) == 0:
        st.error("No dates parsed from Demand sheet.")
        st.stop()

    # Buffer days option
    buffer_days = st.number_input("Buffer days (extra days after last demand, used for planning)", min_value=0, max_value=60, value=7)

    # Extend dates
    last_date = dates[-1]
    for i in range(1, buffer_days + 1):
        dates.append(last_date + timedelta(days=i))
    num_days = len(dates)
    formatted_dates = [d.strftime("%d-%b-%y") for d in dates]

    # Lines/plants
    lines = [str(x).strip() for x in plant_df['Plant'].tolist()]
    capacities = {str(row['Plant']).strip(): safe_int(row.get('Capacity per day', 0)) for _, row in plant_df.iterrows()}
    # default capacity if missing
    for l in lines:
        if capacities.get(l, 0) <= 0:
            capacities[l] = 1000  # fallback to a large default; you can change

    # Grades: use demand columns except first
    grades = [c for c in demand_df.columns.tolist() if c != demand_df.columns[0]]
    grades = [str(g).strip() for g in grades]

    if len(grades) == 0:
        st.error("No grades found in Demand sheet (no columns beyond date).")
        st.stop()

    # Build demand_data dict grade -> date -> demand
    demand_data: Dict[str, Dict[datetime.date, float]] = {}
    for g in grades:
        demand_data[g] = {}
    try:
        for idx in range(len(demand_df)):
            row_date = pd.to_datetime(demand_df.iloc[idx, 0]).date()
            for g in grades:
                val = demand_df.iloc[idx][g] if g in demand_df.columns else 0
                try:
                    demand_data[g][row_date] = float(val) if pd.notna(val) else 0.0
                except Exception:
                    demand_data[g][row_date] = 0.0
    except Exception as e:
        st.warning("Error while parsing demand values; falling back to zeros for problematic cells.")
        for g in grades:
            for d in dates:
                demand_data[g].setdefault(d, 0.0)

    # Fill buffer-day demand zeros
    for g in grades:
        for d in dates[-buffer_days:]:
            demand_data[g].setdefault(d, 0.0)

    # Inventory sheet parsing
    # Expect columns: Grade, Initial Inventory (MT), Min Inventory (safety), Min Closing Inventory, Max Inventory, Rerun Allowed (Yes/No), Lines (comma separated), Force Start Date
    inventory_by_grade = {}
    min_inventory = {}
    min_closing_inventory = {}
    max_inventory = {}
    initial_inventory = {}
    allowed_lines = {}  # grade -> list of lines permitted
    rerun_allowed = {}  # (grade,line) -> bool
    force_start = {}  # (grade,line) -> date or None

    for _, row in inventory_df.iterrows():
        try:
            grade_name = str(row['Grade Name']).strip()
        except Exception:
            continue
        initial_inventory[grade_name] = float(row.get('Initial Inventory', 0) if pd.notna(row.get('Initial Inventory', 0)) else 0.0)
        min_inventory[grade_name] = float(row.get('Min. Inventory', 0) if pd.notna(row.get('Min. Inventory', 0)) else 0.0)
        min_closing_inventory[grade_name] = float(row.get('Min. Closing Inventory', 0) if pd.notna(row.get('Min. Closing Inventory', 0)) else 0.0)
        max_inventory[grade_name] = float(row.get('Inventory', 1000000) if pd.notna(row.get('Inventory', 1000000)) else 1000000)

        # Allowed lines parsing: may be comma separated
        raw_lines = row.get('Lines', "")
        if pd.isna(raw_lines) or str(raw_lines).strip() == "":
            allowed = lines[:]  # if missing, assume all lines allowed (but user should specify)
        else:
            allowed = [l.strip() for l in str(raw_lines).split(',') if l.strip() in lines]
            if len(allowed) == 0:
                # fallback: if none match, assume all
                allowed = lines[:]
        allowed_lines[grade_name] = allowed

        # Rerun allowed default yes
        for ln in allowed:
            rr = row.get('Rerun Allowed', 'Yes')
            if isinstance(rr, str):
                rr_flag = rr.strip().lower() not in ['no', 'n', 'false', '0']
            else:
                rr_flag = bool(rr)
            rerun_allowed[(grade_name, ln)] = rr_flag

        # Force start
        raw_force = row.get('Force Start Date', None)
        if pd.notna(raw_force):
            try:
                fdate = pd.to_datetime(raw_force).date()
            except Exception:
                fdate = None
        else:
            fdate = None
        for ln in allowed:
            force_start[(grade_name, ln)] = fdate

    # If inventory sheet doesn't list a grade, create defaults
    for g in grades:
        if g not in initial_inventory:
            initial_inventory[g] = 0.0
            min_inventory[g] = 0.0
            min_closing_inventory[g] = 0.0
            max_inventory[g] = 1e9
            allowed_lines[g] = lines[:]
            for ln in lines:
                rerun_allowed[(g, ln)] = True
                force_start[(g, ln)] = None

    # Transition matrices
    transition_dfs = load_transition_matrices(excel_bytes, plant_df)
    # Convert to dict: transition_rules[plant][prev][next] -> True/False
    transition_rules: Dict[str, Dict[str, Dict[str, bool]]] = {}
    for plant, df in transition_dfs.items():
        if df is not None:
            # ensure rows and columns cover grades (may have extra)
            tr = {}
            for prev in df.index.astype(str).str.strip():
                tr[prev] = {}
                for nxt in df.columns.astype(str).str.strip():
                    val = str(df.loc[prev, nxt]).strip().lower()
                    allowed = (val == 'yes')
                    tr[prev][nxt] = allowed
            transition_rules[plant] = tr
        else:
            # If not present, treat as "all allowed"
            transition_rules[plant] = None

    # Shutdown processing
    shutdown_periods = process_shutdown_dates(plant_df, dates)

    # Solver configuration inputs
    st.markdown('<div class="section-header">⚙️ Solver configuration</div>', unsafe_allow_html=True)
    colA, colB, colC = st.columns(3)
    with colA:
        time_limit_min = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=5)
        transition_penalty = st.number_input("Transition penalty (per changeover)", min_value=0, value=10)
        stockout_penalty = st.number_input("Stockout penalty (per MT)", min_value=1, value=100)
    with colB:
        max_run_default = st.number_input("Default Max Run Days (if not specified per grade)", min_value=1, value=30)
        min_run_default = st.number_input("Default Min Run Days (if not specified per grade)", min_value=1, value=1)
    with colC:
        solver_workers = st.number_input("Solver threads", min_value=1, max_value=16, value=8)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Extra: allow a small checkbox for enforcing transitions strictly (should be ON by your rule)
    enforce_transition_hard = st.checkbox("Enforce 'No' transitions as HARD constraints (recommended)", value=True)
    remove_continuity_bonus = st.checkbox("Remove continuity bonus (recommended)", value=True, help="Continuity bonus redundant when transitions penalized.")

    # Button to run optimization
    if st.button("🎯 Run Optimization", type="primary"):
        progress = st.progress(0)
        status_box = st.empty()
        try:
            status_box.markdown('<div class="info-box">🔧 Building model...</div>', unsafe_allow_html=True)
            time.sleep(0.5)

            # -------------------------
            # Build CP-SAT model
            # -------------------------
            model = cp_model.CpModel()

            # Create variables only for allowed combos: grade-line pairs from allowed_lines
            is_producing = {}  # (grade,line,day) -> BoolVar
            production = {}    # (grade,line,day) -> IntVar (MT produced that day)
            for g in grades:
                for ln in allowed_lines[g]:
                    for d in range(num_days):
                        key = (g, ln, d)
                        is_producing[key] = model.NewBoolVar(f'is_prod_{g}_{ln}_{d}')
                        # production amount: either 0 or capacity (we keep integral MT units; can be scaled)
                        # allow partial production on buffer days
                        if d < num_days - buffer_days:
                            prod_var = model.NewIntVar(0, capacities[ln], f'prod_{g}_{ln}_{d}')
                            # link: if is_producing then production == capacity else 0
                            model.Add(prod_var == capacities[ln]).OnlyEnforceIf(is_producing[key])
                            model.Add(prod_var == 0).OnlyEnforceIf(is_producing[key].Not())
                        else:
                            prod_var = model.NewIntVar(0, capacities[ln], f'prod_{g}_{ln}_{d}')
                            model.Add(prod_var <= capacities[ln]).OnlyEnforceIf(is_producing[key])
                            model.Add(prod_var == 0).OnlyEnforceIf(is_producing[key].Not())
                        production[key] = prod_var

            # Per-line-per-day capacity: sum of productions on a line = capacity or 0 on shutdown
            for ln in lines:
                for d in range(num_days):
                    prod_vars = []
                    for g in grades:
                        if (g, ln, d) in production:
                            prod_vars.append(production[(g, ln, d)])
                    if ln in shutdown_periods and d in shutdown_periods[ln]:
                        # forced zero production
                        model.Add(sum(prod_vars) == 0)
                        # also enforce is_producing false
                        for g in grades:
                            if (g, ln, d) in is_producing:
                                model.Add(is_producing[(g, ln, d)] == 0)
                    else:
                        # On actual planning days we enforce capacity (exact fill); on buffer days we allow <=
                        if d < num_days - buffer_days:
                            model.Add(sum(prod_vars) <= capacities[ln])
                            # note: allow not fully utilized to satisfy inventory/stockout tradeoffs
                        else:
                            model.Add(sum(prod_vars) <= capacities[ln])

            # One grade at most per line per day
            for ln in lines:
                for d in range(num_days):
                    producing_bools = []
                    for g in grades:
                        if (g, ln, d) in is_producing:
                            producing_bools.append(is_producing[(g, ln, d)])
                    if producing_bools:
                        model.Add(sum(producing_bools) <= 1)

            # Inventory variables: inventory at start of day d (0..num_days)
            inventory_vars = {}
            for g in grades:
                for d in range(num_days + 1):
                    inventory_vars[(g, d)] = model.NewIntVar(0, int(max_inventory.get(g, 1000000)), f'inv_{g}_{d}')

            # Stockout variables for each day (how much unmet demand)
            stockout_vars = {}
            for g in grades:
                for d in range(num_days):
                    stockout_vars[(g, d)] = model.NewIntVar(0, 1000000, f'stockout_{g}_{d}')

            # Inventory balance constraints:
            # inventory[g,d+1] = inventory[g,d] + production_of_g_on_day_d - demand[g,d] + stockout[g,d] (stockout is positive when demand > supply)
            for g in grades:
                # initial inventory at day 0
                model.Add(inventory_vars[(g, 0)] == int(initial_inventory.get(g, 0)))
                for d in range(num_days):
                    # sum production of g across lines on day d
                    prod_sum_terms = []
                    for ln in allowed_lines[g]:
                        if (g, ln, d) in production:
                            prod_sum_terms.append(production[(g, ln, d)])
                    if prod_sum_terms:
                        prod_sum = sum(prod_sum_terms)
                    else:
                        # no plant can produce this grade
                        prod_sum = 0

                    demand_val = int(demand_data.get(g, {}).get(dates[d], 0.0))
                    # inventory_{d+1} = inventory_{d} + prod_sum - demand + stockout
                    # Rearranged: inventory_{d+1} - inventory_{d} - prod_sum + demand == stockout
                    # We'll express as linear constraint: inventory_{d+1} == inventory_{d} + prod_sum - (demand - stockout)
                    model.Add(inventory_vars[(g, d+1)] == inventory_vars[(g, d)] + prod_sum - (demand_val - stockout_vars[(g, d)]))

            # Min closing inventory: for user-specified days (end of actual horizon excluding buffer) -- enforce as soft via objective OR as hard? We'll enforce as soft via deficit vars
            closing_deficit_vars = {}
            for g in grades:
                closing_day_idx = num_days - buffer_days  # this is first buffer day index
                if closing_day_idx < 1:
                    closing_day_idx = num_days  # fallback
                closing_inventory_var = inventory_vars[(g, closing_day_idx)]
                min_cl = int(min_closing_inventory.get(g, 0))
                if min_cl > 0:
                    deficit = model.NewIntVar(0, 1000000, f'closing_deficit_{g}')
                    model.Add(deficit >= min_cl - closing_inventory_var)
                    model.Add(deficit >= 0)
                    closing_deficit_vars[g] = deficit

            # Min inventory daily (safety stock) as soft via deficits
            daily_deficit_vars = {}
            for g in grades:
                mi = int(min_inventory.get(g, 0))
                if mi > 0:
                    for d in range(num_days):
                        deficit = model.NewIntVar(0, 1000000, f'deficit_{g}_{d}')
                        model.Add(deficit >= mi - inventory_vars[(g, d+1)])
                        model.Add(deficit >= 0)
                        daily_deficit_vars[(g, d)] = deficit

            # Start variables to enforce min run days and rerun limits
            is_start = {}
            is_end = {}
            # start if day d is producing and day d-1 was not (or d==0)
            for g in grades:
                for ln in allowed_lines[g]:
                    for d in range(num_days):
                        key = (g, ln, d)
                        if key in is_producing:
                            s = model.NewBoolVar(f'start_{g}_{ln}_{d}')
                            is_start[key] = s
                            # if producing at d and not producing at d-1 => start
                            if d == 0:
                                # start when is_producing true
                                model.Add(s == is_producing[key])
                            else:
                                prev_key = (g, ln, d-1)
                                if prev_key in is_producing:
                                    # s = prod[d] AND NOT prod[d-1]
                                    model.AddBoolAnd([is_producing[key], is_producing[prev_key].Not()]).OnlyEnforceIf(s)
                                    model.AddBoolOr([is_producing[key].Not(), is_producing[prev_key]]).OnlyEnforceIf(s.Not())
                                else:
                                    model.Add(s == is_producing[key])
                        # end var
                        e = model.NewBoolVar(f'end_{g}_{ln}_{d}')
                        is_end[(g, ln, d)] = e
                        if key in is_producing:
                            if d == num_days - 1:
                                model.Add(e == is_producing[key])
                            else:
                                next_key = (g, ln, d+1)
                                if next_key in is_producing:
                                    model.AddBoolAnd([is_producing[key], is_producing[next_key].Not()]).OnlyEnforceIf(e)
                                    model.AddBoolOr([is_producing[key].Not(), is_producing[next_key]]).OnlyEnforceIf(e.Not())
                                else:
                                    model.Add(e == is_producing[key])

            # Min / Max run enforcement (min_run_default / max_run_default), accounting for shutdown break
            # For each start, we enforce that that run continues for at least min_run (unless shutdown blocks it)
            for g in grades:
                for ln in allowed_lines[g]:
                    min_run = int(inventory_df.loc[inventory_df['Grade Name'] == g, 'Min. Run Days'].iloc[0]) if (g in inventory_df['Grade Name'].astype(str).tolist() and 'Min. Run Days' in inventory_df.columns) else min_run_default
                    max_run = int(inventory_df.loc[inventory_df['Grade Name'] == g, 'Max. Run Days'].iloc[0]) if (g in inventory_df['Grade Name'].astype(str).tolist() and 'Max. Run Days' in inventory_df.columns) else max_run_default

                    for d in range(num_days):
                        key = (g, ln, d)
                        if key not in is_start:
                            continue
                        s = is_start[key]
                        # compute how many days are available consecutively starting at d (stop at shutdown)
                        max_possible = 0
                        max_days_range = []
                        for k in range(d, num_days):
                            if ln in shutdown_periods and k in shutdown_periods[ln]:
                                break
                            max_possible += 1
                            max_days_range.append(k)
                        # enforce min run only if enough days available
                        if max_possible >= min_run:
                            # for each offset 0..min_run-1, prod must be 1 when s=1
                            for offset in range(min_run):
                                dd = d + offset
                                if dd < num_days and (g, ln, dd) in is_producing:
                                    model.Add(is_producing[(g, ln, dd)] == 1).OnlyEnforceIf(s)
                        # enforce max run: sliding window ensure no run longer than max_run
                        if max_possible > max_run:
                            # for windows of length max_run+1, at most max_run days can be 1 (i.e. cannot have all max_run+1 true)
                            for start_w in range(d, d + max_possible - max_run):
                                window_vars = []
                                for kk in range(start_w, start_w + max_run + 1):
                                    if kk < num_days and (g, ln, kk) in is_producing:
                                        window_vars.append(is_producing[(g, ln, kk)])
                                if len(window_vars) == max_run + 1:
                                    model.Add(sum(window_vars) <= max_run)

            # Rerun prohibition (if rerun_allowed==False then allow at most one start for (grade,line))
            for g in grades:
                for ln in allowed_lines[g]:
                    if not rerun_allowed.get((g, ln), True):
                        starts = [is_start[(g, ln, d)] for d in range(num_days) if (g, ln, d) in is_start]
                        if starts:
                            model.Add(sum(starts) <= 1)

            # Force-start enforcement (if force-start exists for a grade-line, ensure that day production starts there)
            for (g, ln), fdate in force_start.items():
                if fdate is not None:
                    # find index
                    try:
                        idx = dates.index(fdate)
                    except ValueError:
                        # date not in planning horizon
                        idx = None
                    if idx is not None:
                        # enforce is_producing[(g,ln,idx)] == 1
                        if (g, ln, idx) in is_producing:
                            model.Add(is_producing[(g, ln, idx)] == 1)
                        else:
                            # no variable exists (line not allowed or missing) -> infeasible; we'll detect solver infeasibility later
                            pass

            # Transition rules enforcement:
            # For each line, day d -> d+1, if previous grade p and next grade n and transition_rules[line][p][n] == False (No), then forbid is_producing[p,line,d] AND is_producing[n,line,d+1]
            for ln in lines:
                tr = transition_rules.get(ln, None)
                for d in range(num_days - 1):
                    for p in grades:
                        prev_key = (p, ln, d)
                        if prev_key not in is_producing:
                            continue
                        for n in grades:
                            next_key = (n, ln, d+1)
                            if next_key not in is_producing:
                                continue
                            allowed = True
                            if tr is None:
                                allowed = True
                            else:
                                # if p not in tr or n not in tr[p], treat missing as allowed (but warn earlier)
                                allowed = bool(tr.get(p, {}).get(n, True))
                            if not allowed and enforce_transition_hard:
                                # forbid both = 1
                                model.Add(is_producing[prev_key] + is_producing[next_key] <= 1)

            # Transition counting & penalty:
            # We'll create transition_bool[ln,d] = 1 if production on both days AND grade changes
            transition_bool = {}
            transition_vars_list = []
            for ln in lines:
                for d in range(num_days - 1):
                    # Check if any production occurs on both days
                    day_d_prod_bools = [is_producing[(g, ln, d)] for g in grades if (g, ln, d) in is_producing]
                    day_d1_prod_bools = [is_producing[(g, ln, d+1)] for g in grades if (g, ln, d+1) in is_producing]
                    if not day_d_prod_bools or not day_d1_prod_bools:
                        continue
                    prod_day_d = model.NewBoolVar(f'prod_exists_{ln}_{d}')
                    prod_day_d1 = model.NewBoolVar(f'prod_exists_{ln}_{d+1}')
                    model.AddMaxEquality(prod_day_d, day_d_prod_bools)
                    model.AddMaxEquality(prod_day_d1, day_d1_prod_bools)

                    # Continuity detection: if same grade continues then not a transition
                    # Create continuity bools per grade and take max
                    continuity_bools = []
                    for g in grades:
                        if (g, ln, d) in is_producing and (g, ln, d+1) in is_producing:
                            same_bool = model.NewBoolVar(f'same_{g}_{ln}_{d}')
                            # same_bool == 1 iff both days produce g
                            model.AddBoolAnd([is_producing[(g, ln, d)], is_producing[(g, ln, d+1)]]).OnlyEnforceIf(same_bool)
                            model.AddBoolOr([is_producing[(g, ln, d)].Not(), is_producing[(g, ln, d+1)].Not()]).OnlyEnforceIf(same_bool.Not())
                            continuity_bools.append(same_bool)
                    if continuity_bools:
                        has_continuity = model.NewBoolVar(f'has_cont_{ln}_{d}')
                        model.AddMaxEquality(has_continuity, continuity_bools)
                    else:
                        has_continuity = model.NewBoolVar(f'has_cont_{ln}_{d}')
                        model.Add(has_continuity == 0)

                    trans = model.NewBoolVar(f'trans_{ln}_{d}')
                    # trans => production on both days AND no continuity
                    # trans => prod_day_d == 1, prod_day_d1 ==1, has_continuity == 0
                    # implement with reified constraints:
                    model.AddBoolAnd([prod_day_d, prod_day_d1, has_continuity.Not()]).OnlyEnforceIf(trans)
                    # if trans == 0 then at least one of (not both days producing OR continuity)
                    model.AddBoolOr([prod_day_d.Not(), prod_day_d1.Not(), has_continuity]).OnlyEnforceIf(trans.Not())

                    transition_bool[(ln, d)] = trans
                    transition_vars_list.append(trans)

            # Objective construction
            objective_terms = []

            # Stockout penalty (heavy)
            for g in grades:
                for d in range(num_days):
                    objective_terms.append(stockout_penalty * stockout_vars[(g, d)])

            # Transition penalty (applies to each transition var)
            for tvar in transition_vars_list:
                objective_terms.append(transition_penalty * tvar)

            # Closing deficit penalty (if desired)
            for g, dv in closing_deficit_vars.items():
                objective_terms.append(stockout_penalty * dv)  # reuse stockout_penalty as weight (user can tweak)

            # Daily min inventory deficits
            for dv in daily_deficit_vars.values():
                objective_terms.append(10 * dv)  # small weight vs stockout

            # Optional holding cost (small)
            for g in grades:
                for d in range(num_days):
                    objective_terms.append(1 * inventory_vars[(g, d)])

            model.Minimize(sum(objective_terms))

            progress.progress(30)
            status_box.markdown('<div class="info-box">⚡ Solving optimization problem...</div>', unsafe_allow_html=True)

            # -------------------------
            # Solver config & run
            # -------------------------
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = float(time_limit_min * 60)
            solver.parameters.num_search_workers = int(solver_workers)
            solver.parameters.random_seed = 42
            solver.parameters.log_search_progress = False
            solver.parameters.cp_model_presolve = True

            # Callback to capture intermediate solutions (minimal)
            class SimpleCallback(cp_model.CpSolverSolutionCallback):
                def __init__(self, production_vars, is_producing_vars, inventory_vars, stockout_vars, grades, lines, dates):
                    cp_model.CpSolverSolutionCallback.__init__(self)
                    self.best = None
                    self.production = production_vars
                    self.is_producing = is_producing_vars
                    self.inventory = inventory_vars
                    self.stockout = stockout_vars
                    self.grades = grades
                    self.lines = lines
                    self.dates = dates
                    self.solutions = []

                def on_solution_callback(self):
                    # capture objective and a compact representation
                    obj = self.ObjectiveValue()
                    t = time.time()
                    sol = {
                        'time': t,
                        'objective': obj,
                        'is_producing': {},
                        'production': {},
                        'inventory': {}
                    }
                    for ln in self.lines:
                        sol['is_producing'][ln] = {}
                        for d_idx in range(len(self.dates)):
                            sol['is_producing'][ln][self.dates[d_idx].strftime("%d-%b-%y")] = None
                            for g in self.grades:
                                k = (g, ln, d_idx)
                                if k in self.is_producing and self.Value(self.is_producing[k]) == 1:
                                    sol['is_producing'][ln][self.dates[d_idx].strftime("%d-%b-%y")] = g
                                    break
                    self.solutions.append(sol)

            cb = SimpleCallback(production, is_producing, inventory_vars, stockout_vars, grades, lines, dates)

            start = time.time()
            status = solver.SolveWithSolutionCallback(model, cb)
            solve_time = time.time() - start

            progress.progress(80)

            # -------------------------
            # Check solver status and extract solution
            # -------------------------
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                status_box.markdown('<div class="success-box">✅ Solution found</div>', unsafe_allow_html=True)
            else:
                status_box.markdown('<div class="alert-box">⚠️ No feasible solution found. Check constraints and inputs.</div>', unsafe_allow_html=True)

            progress.progress(90)

            # Use latest captured solution if any
            if hasattr(cb, 'solutions') and len(cb.solutions) > 0:
                result = cb.solutions[-1]
            else:
                # fallback: extract directly
                result = {
                    'objective': solver.ObjectiveValue(),
                    'is_producing': {},
                    'production': {},
                    'inventory': {}
                }
                for ln in lines:
                    result['is_producing'][ln] = {}
                    for d in range(num_days):
                        date_key = dates[d].strftime("%d-%b-%y")
                        result['is_producing'][ln][date_key] = None
                        for g in grades:
                            key = (g, ln, d)
                            if key in is_producing and solver.Value(is_producing[key]) == 1:
                                result['is_producing'][ln][date_key] = g
                                break

            # Build summary tables for UI
            production_totals = {g: {ln: 0 for ln in lines} for g in grades}
            stockout_totals = {g: 0 for g in grades}
            plant_totals = {ln: 0 for ln in lines}

            for g in grades:
                for ln in allowed_lines[g]:
                    for d in range(num_days):
                        key = (g, ln, d)
                        if key in production:
                            try:
                                val = solver.Value(production[key])
                            except Exception:
                                val = 0
                            production_totals[g][ln] += val
                            plant_totals[ln] += val
                for d in range(num_days):
                    try:
                        stockout_totals[g] += solver.Value(stockout_vars[(g, d)])
                    except Exception:
                        stockout_totals[g] += 0

            # Count transitions per line
            transition_count_per_line = {ln: 0 for ln in lines}
            total_transitions = 0
            for ln in lines:
                last_grade = None
                for d in range(num_days):
                    current = None
                    for g in grades:
                        if (g, ln, d) in is_producing:
                            if solver.Value(is_producing[(g, ln, d)]) == 1:
                                current = g
                                break
                    if current is not None:
                        if last_grade is not None and current != last_grade:
                            transition_count_per_line[ln] += 1
                            total_transitions += 1
                        last_grade = current

            progress.progress(95)

            # -------------------------
            # Render Results UI (Plotly charts similar to app.py)
            # -------------------------
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-header">📈 Results</div>', unsafe_allow_html=True)

            # Metrics
            colm1, colm2, colm3, colm4 = st.columns(4)
            with colm1:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.markdown(f"**Objective**: {solver.ObjectiveValue():.0f}")
                st.markdown('</div>', unsafe_allow_html=True)
            with colm2:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.markdown(f"**Total Transitions**: {total_transitions}")
                st.markdown('</div>', unsafe_allow_html=True)
            with colm3:
                total_stockout_mt = sum(stockout_totals.values())
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.markdown(f"**Total Stockout (MT)**: {total_stockout_mt:.0f}")
                st.markdown('</div>', unsafe_allow_html=True)
            with colm4:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.markdown(f"**Solve time (s)**: {solve_time:.1f}")
                st.markdown('</div>', unsafe_allow_html=True)

            # Tabs for visuals
            tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Summary", "📦 Inventory"])

            with tab1:
                st.markdown("#### Production Gantt by Line")
                sorted_grades = sorted(grades)
                base_colors = px.colors.qualitative.Vivid
                grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

                for ln in lines:
                    st.markdown(f"##### 🏭 {ln}")
                    gantt_data = []
                    for d in range(num_days):
                        date = dates[d]
                        for g in grades:
                            if (g, ln, d) in is_producing and solver.Value(is_producing[(g, ln, d)]) == 1:
                                gantt_data.append({
                                    "Grade": g,
                                    "Start": date,
                                    "Finish": date + timedelta(days=1),
                                    "Line": ln
                                })
                                break
                    if len(gantt_data) == 0:
                        st.info(f"No production scheduled for {ln}")
                        continue
                    gantt_df = pd.DataFrame(gantt_data)
                    fig = px.timeline(gantt_df, x_start="Start", x_end="Finish", y="Grade", color="Grade",
                                      color_discrete_map=grade_color_map)
                    fig.update_yaxes(autorange="reversed")
                    # add shutdown vrect if exists
                    if ln in shutdown_periods and shutdown_periods[ln]:
                        sd = shutdown_periods[ln]
                        if sd:
                            start_shutdown = dates[sd[0]]
                            end_shutdown = dates[sd[-1]] + timedelta(days=1)
                            fig.add_vrect(x0=start_shutdown, x1=end_shutdown, fillcolor="red", opacity=0.12, layer="below")
                    fig.update_layout(height=300, margin=dict(l=60, r=20, t=20, b=60))
                    st.plotly_chart(fig, use_container_width=True)

            with tab2:
                st.markdown("#### Summary Table")
                total_prod_data = []
                for g in grades:
                    row = {'Grade': g}
                    total = 0
                    for ln in lines:
                        val = production_totals[g].get(ln, 0)
                        row[ln] = int(val)
                        total += int(val)
                    row['Total Produced'] = total
                    row['Total Stockout'] = int(stockout_totals.get(g, 0))
                    total_prod_data.append(row)
                totals_row = {'Grade': 'TOTAL'}
                for ln in lines:
                    totals_row[ln] = plant_totals.get(ln, 0)
                totals_row['Total Produced'] = sum(plant_totals.values())
                totals_row['Total Stockout'] = sum(stockout_totals.values())
                total_prod_data.append(totals_row)
                st.dataframe(pd.DataFrame(total_prod_data), use_container_width=True)

                st.markdown("#### Transitions by Plant")
                tr_data = [{"Plant": ln, "Transitions": transition_count_per_line[ln]} for ln in lines]
                st.dataframe(pd.DataFrame(tr_data), use_container_width=True)

            with tab3:
                st.markdown("#### Inventory Trends")
                for g in grades:
                    inv_vals = []
                    for d in range(num_days):
                        try:
                            inv_vals.append(solver.Value(inventory_vars[(g, d)]))
                        except Exception:
                            inv_vals.append(0)
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=dates, y=inv_vals, mode="lines+markers", name=g))
                    fig.update_layout(height=250, margin=dict(l=40, r=10, t=30, b=40))
                    st.plotly_chart(fig, use_container_width=True)

            progress.progress(100)
            status_box.markdown(f'<div class="success-box">🎉 Done — objective {solver.ObjectiveValue():.0f}  |  time {solve_time:.1f}s</div>', unsafe_allow_html=True)

        except Exception as ex:
            st.error(f"Solver error: {ex}")
            import traceback as tb
            st.text(tb.format_exc())
            progress.progress(100)
