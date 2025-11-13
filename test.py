"""
app_full.py — Polymer Production Scheduler (Deterministic, Full-Capacity, Scaled MT)
Complete file from imports through visualizations.

Key properties:
- All numeric inputs are Metric Tonnes (MT); internally scaled by SCALE = 1000 (1 MT -> 1000 units).
- Production is forced to full plant capacity when a run is scheduled (and zero otherwise).
- Deterministic solver: fixed random seed and single search worker by default.
- No create_sample_workbook fallback — user must provide correct Excel template.
- Includes Streamlit UI and visualization (Plotly Gantt-like chart for schedule, inventory & stockout charts).
- Uses OR-Tools CP-SAT for scheduling and deterministic settings.

Usage:
- Run with `streamlit run app_full.py` in an environment containing: ortools, pandas, plotly, streamlit.
"""

import io
import math
import logging
from datetime import datetime, timedelta, date
from typing import Dict, List, Any, Tuple
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from ortools.sat.python import cp_model

# Streamlit (optional) — app has UI
try:
    import streamlit as st
except Exception:
    st = None

# ----------------------------
# Config & Logging
# ----------------------------
SCALE = 1000  # 1 MT -> 1000 internal units (integer)
DEFAULT_RANDOM_SEED = 123456789
DEFAULT_NUM_WORKERS = 1

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app_full")

# ----------------------------
# Helper functions
# ----------------------------
def to_int_mt(value: Any) -> int:
    """Convert numeric MT to integer units. NaN/invalid -> 0."""
    try:
        if value is None:
            return 0
        v = float(value)
        if math.isnan(v):
            return 0
        return int(round(v * SCALE))
    except Exception:
        return 0

def to_float_mt(value: int) -> float:
    """Convert integer units back to MT."""
    return float(value) / SCALE

def ensure_date(d) -> date:
    """Return python.date from various inputs or raise."""
    if isinstance(d, date):
        return d
    if isinstance(d, datetime):
        return d.date()
    try:
        return pd.to_datetime(d).date()
    except Exception as e:
        raise ValueError(f"Invalid date value: {d}") from e

# ----------------------------
# Problem data container
# ----------------------------
class ProblemData:
    def __init__(self):
        self.dates: List[date] = []
        self.grades: List[str] = []
        self.plants: List[str] = []
        self.capacities: Dict[str, int] = {}
        self.allowed_lines: Dict[str, List[str]] = {}
        self.initial_inventory: Dict[str, int] = {}
        self.min_closing_inventory: Dict[str, int] = {}
        self.min_run_days: Dict[str, int] = {}
        self.max_run_days: Dict[str, int] = {}
        self.demand: Dict[str, Dict[date, int]] = {}
        self.transition_rules: Dict[str, Dict[str, List[str]]] = {}
        self.shutdown_days: Dict[str, set] = {}
        self.penalty_stockout_per_mt: int = int(1000 * SCALE)
        self.penalty_transition: int = 1
        self.force_full_capacity: bool = True
        self.random_seed: int = DEFAULT_RANDOM_SEED
        self.num_workers: int = DEFAULT_NUM_WORKERS

# ----------------------------
# Model builder
# ----------------------------
def build_model(data: ProblemData) -> Tuple[cp_model.CpModel, Dict[str, Any]]:
    model = cp_model.CpModel()
    vars_map = {}

    grades = data.grades
    plants = data.plants
    dates = data.dates
    num_days = len(dates)

    # Bound estimation
    BIG = sum(data.capacities.values()) * num_days + sum(data.initial_inventory.values()) + 10 * SCALE

    production = {}
    is_producing = {}

    # Create production & bools: production equals capacity when running, else zero
    for g in grades:
        for p in data.allowed_lines.get(g, []):
            cap = int(max(0, data.capacities.get(p, 0)))
            for d in range(num_days):
                key = (g, p, d)
                prod_var = model.NewIntVar(0, cap, f'prod_{g}_{p}_{d}')
                production[key] = prod_var
                run_var = model.NewBoolVar(f'isprod_{g}_{p}_{d}')
                is_producing[key] = run_var
                # Enforce full-capacity when running, zero when not
                model.Add(prod_var == cap).OnlyEnforceIf(run_var)
                model.Add(prod_var == 0).OnlyEnforceIf(run_var.Not())
                # Shutdown handling
                if d in data.shutdown_days.get(p, set()):
                    model.Add(run_var == 0)

    vars_map['production'] = production
    vars_map['is_producing'] = is_producing

    # Inventory & stockout
    inventory = {}
    stockout = {}
    for g in grades:
        for d in range(num_days + 1):
            inv = model.NewIntVar(0, BIG, f'inv_{g}_{d}')
            inventory[(g, d)] = inv
    for g in grades:
        for d in range(num_days):
            so = model.NewIntVar(0, BIG, f'so_{g}_{d}')
            stockout[(g, d)] = so

    vars_map['inventory'] = inventory
    vars_map['stockout'] = stockout

    # Opening inventory
    for g in grades:
        inv0 = data.initial_inventory.get(g, 0)
        model.Add(inventory[(g, 0)] == inv0)

    # Inventory balance
    for g in grades:
        for d in range(num_days):
            demand_val = data.demand.get(g, {}).get(dates[d], 0)
            prod_list = []
            for p in data.allowed_lines.get(g, []):
                k = (g, p, d)
                if k in production:
                    prod_list.append(production[k])
            if prod_list:
                total_prod = sum(prod_list)
            else:
                total_prod = model.NewConstant(0)
            # inv_{d+1} + (demand - stockout) == inv_d + total_prod => as rearranged below
            model.Add(inventory[(g, d + 1)] + (demand_val - stockout[(g, d)]) == inventory[(g, d)] + total_prod)
            model.Add(stockout[(g, d)] <= demand_val)

    # Min closing inventory
    last = num_days
    for g in grades:
        min_cl = data.min_closing_inventory.get(g, 0)
        model.Add(inventory[(g, last)] >= min_cl)

    # Run-length constraints via sliding windows and start enforcement
    for g in grades:
        min_run = data.min_run_days.get(g, 1)
        max_run = data.max_run_days.get(g, num_days)
        for p in data.allowed_lines.get(g, []):
            bools = [is_producing[(g, p, d)] for d in range(num_days)]
            # max-run: any window of size max_run+1 has sum <= max_run
            if max_run < num_days:
                window = max_run + 1
                for s in range(0, num_days - window + 1):
                    model.Add(sum(bools[s:s + window]) <= max_run)
            # min-run via start indicator
            if min_run > 1:
                for d in range(num_days):
                    curr = bools[d]
                    if d == 0:
                        start_b = model.NewBoolVar(f'start_{g}_{p}_{d}')
                        model.Add(start_b == curr)
                    else:
                        prev = bools[d - 1]
                        start_b = model.NewBoolVar(f'start_{g}_{p}_{d}')
                        # start_b implies curr==1 and prev==0
                        model.Add(curr == 1).OnlyEnforceIf(start_b)
                        model.Add(prev == 0).OnlyEnforceIf(start_b)
                        # additionally, if curr==1 and prev==0 -> start_b = 1 (reified)
                        model.AddBoolAnd([curr, prev.Not()]).OnlyEnforceIf(start_b)
                    # If a start occurs, ensure the next min_run days are on (or forbid start near horizon end)
                    if d + min_run <= num_days:
                        span = bools[d:d + min_run]
                        model.Add(sum(span) >= min_run).OnlyEnforceIf(start_b)
                    else:
                        model.Add(start_b == 0)

    # Transition rules (hard disallow)
    for p in plants:
        rules = data.transition_rules.get(p, {})
        for d in range(num_days - 1):
            for prev_grade, allowed_next in rules.items():
                if allowed_next is None:
                    continue
                for curr_grade in [g for g in grades if p in data.allowed_lines.get(g, [])]:
                    if curr_grade not in allowed_next:
                        pk = (prev_grade, p, d)
                        nk = (curr_grade, p, d + 1)
                        if pk in is_producing and nk in is_producing:
                            model.Add(is_producing[pk] + is_producing[nk] <= 1)

    # Objective: minimize stockout weighted sum
    obj_terms = []
    for g in grades:
        for d in range(num_days):
            obj_terms.append(stockout[(g, d)] * data.penalty_stockout_per_mt)
    model.Minimize(sum(obj_terms))

    vars_map['grades'] = grades
    vars_map['plants'] = plants
    vars_map['dates'] = dates
    vars_map['num_days'] = num_days

    return model, vars_map

# ----------------------------
# Callback to capture solutions
# ----------------------------
class CaptureSolutionsCallback(cp_model.CpSolverSolutionCallback):
    def __init__(self, vars_map):
        super().__init__()
        self._vars = vars_map
        self.solutions = []

    def on_solution_callback(self):
        sol_snapshot = {'is_producing': {}, 'production': {}, 'inventory': {}, 'stockout': {}}
        prod = self._vars.get('production', {})
        isprod = self._vars.get('is_producing', {})
        inv = self._vars.get('inventory', {})
        so = self._vars.get('stockout', {})

        for k, v in isprod.items():
            try:
                sol_snapshot['is_producing'][k] = int(self.Value(v))
            except Exception:
                sol_snapshot['is_producing'][k] = 0
        for k, v in prod.items():
            try:
                sol_snapshot['production'][k] = int(self.Value(v))
            except Exception:
                sol_snapshot['production'][k] = 0
        for k, v in inv.items():
            try:
                sol_snapshot['inventory'][k] = int(self.Value(v))
            except Exception:
                sol_snapshot['inventory'][k] = 0
        for k, v in so.items():
            try:
                sol_snapshot['stockout'][k] = int(self.Value(v))
            except Exception:
                sol_snapshot['stockout'][k] = 0

        try:
            obj = self.ObjectiveValue()
        except Exception:
            obj = None
        self.solutions.append({'solution': sol_snapshot, 'objective': obj, 'time': self.WallTime()})

    def num_solutions(self):
        return len(self.solutions)

# ----------------------------
# Read Excel workbook into ProblemData
# ----------------------------
def read_input_workbook(xl_bytes: bytes) -> ProblemData:
    data = ProblemData()
    xl = pd.ExcelFile(io.BytesIO(xl_bytes))

    # Required sheets
    required = {'Plant', 'Inventory', 'Demand'}
    if not required.issubset(set(xl.sheet_names)):
        raise ValueError("Excel must contain sheets: 'Plant', 'Inventory', 'Demand'.")

    # Plant sheet
    plant_df = xl.parse('Plant')
    plant_df.columns = [c.strip() for c in plant_df.columns]
    plants = []
    capacities = {}
    shutdown_raw = {}
    for idx, row in plant_df.iterrows():
        plant = str(row.get('Plant', '')).strip()
        if not plant:
            continue
        plants.append(plant)
        cap = to_int_mt(row.get('Capacity per day', 0))
        capacities[plant] = cap
        s_raw = row.get('Shutdown Start Date', None)
        e_raw = row.get('Shutdown End Date', None)
        if pd.notna(s_raw) and pd.notna(e_raw):
            try:
                s = ensure_date(s_raw)
                e = ensure_date(e_raw)
                shutdown_raw[plant] = (s, e)
            except Exception:
                shutdown_raw[plant] = None
        else:
            shutdown_raw[plant] = None

    data.plants = plants
    data.capacities = capacities

    # Inventory sheet
    inv_df = xl.parse('Inventory')
    inv_df.columns = [c.strip() for c in inv_df.columns]
    grades = []
    allowed_lines = {}
    initial_inventory = {}
    min_closing_inventory = {}
    min_run_days = {}
    max_run_days = {}
    for idx, row in inv_df.iterrows():
        grade = str(row.get('Grade Name', '')).strip()
        if not grade:
            continue
        grades.append(grade)
        lines_field = str(row.get('Lines', '')).strip()
        lines = [s.strip() for s in lines_field.split(',')] if lines_field else []
        allowed_lines[grade] = [l for l in lines if l]
        initial_inventory[grade] = to_int_mt(row.get('Opening Inventory', 0))
        min_closing_inventory[grade] = to_int_mt(row.get('Min. Closing Inventory', 0))
        min_run_days[grade] = int(float(row.get('Min. Run Days', 1))) if row.get('Min. Run Days', '') != '' else 1
        max_run_days[grade] = int(float(row.get('Max. Run Days', max(1, min_run_days[grade]))))

    data.grades = grades
    data.allowed_lines = allowed_lines
    data.initial_inventory = initial_inventory
    data.min_closing_inventory = min_closing_inventory
    data.min_run_days = min_run_days
    data.max_run_days = max_run_days

    # Demand sheet
    demand_df = xl.parse('Demand')
    demand_df.columns = [c.strip() for c in demand_df.columns]
    if demand_df.shape[1] < 2:
        raise ValueError("Demand sheet must have Date column plus at least one grade column.")
    date_series = pd.to_datetime(demand_df.iloc[:, 0], errors='coerce')
    if date_series.isnull().any():
        raise ValueError("Invalid dates in Demand sheet first column.")
    dates = [ensure_date(d) for d in date_series.tolist()]
    data.dates = dates

    demand_grades = [c for c in demand_df.columns[1:]]
    for g in demand_grades:
        if g not in grades:
            # Add grade with defaults if demand present but no inventory row
            grades.append(g)
            data.initial_inventory[g] = 0
            data.min_closing_inventory[g] = 0
            data.allowed_lines[g] = []
            data.min_run_days[g] = 1
            data.max_run_days[g] = len(dates)
        series = {}
        for i, dt in enumerate(dates):
            raw = demand_df.iloc[i, demand_df.columns.get_loc(g)]
            series[dt] = to_int_mt(raw)
        data.demand[g] = series

    # Map shutdown dates to day indices
    shutdown_indices = {}
    for p in data.plants:
        tuple_ = shutdown_raw.get(p)
        if tuple_ is None:
            shutdown_indices[p] = set()
            continue
        s, e = tuple_
        idxs = set()
        for i, dt in enumerate(data.dates):
            if s <= dt <= e:
                idxs.add(i)
        shutdown_indices[p] = idxs
    data.shutdown_days = shutdown_indices

    # Set strong stockout penalty scaled
    data.penalty_stockout_per_mt = int(1000 * SCALE)
    data.penalty_transition = 1

    return data

# ----------------------------
# Solve wrapper
# ----------------------------
def solve_problem(data: ProblemData, time_limit_seconds: float = 60.0) -> Dict[str, Any]:
    model, vars_map = build_model(data)
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(time_limit_seconds)
    solver.parameters.random_seed = int(data.random_seed)
    solver.parameters.num_search_workers = int(data.num_workers)

    callback = CaptureSolutionsCallback(vars_map)
    result = solver.SolveWithSolutionCallback(model, callback)

    if callback.num_solutions() > 0:
        canonical = callback.solutions[-1]['solution']
        obj = callback.solutions[-1]['objective']
    else:
        canonical = {'is_producing': {}, 'production': {}, 'inventory': {}, 'stockout': {}}
        try:
            prod_map = vars_map.get('production', {})
            isprod_map = vars_map.get('is_producing', {})
            inv_map = vars_map.get('inventory', {})
            so_map = vars_map.get('stockout', {})
            for k, v in isprod_map.items():
                canonical['is_producing'][k] = int(solver.Value(v))
            for k, v in prod_map.items():
                canonical['production'][k] = int(solver.Value(v))
            for k, v in inv_map.items():
                canonical['inventory'][k] = int(solver.Value(v))
            for k, v in so_map.items():
                canonical['stockout'][k] = int(solver.Value(v))
            obj = solver.ObjectiveValue()
        except Exception:
            canonical = None
            obj = None

    status = {
    cp_model.OPTIMAL: "OPTIMAL",
    cp_model.FEASIBLE: "FEASIBLE",
    cp_model.UNKNOWN: "UNKNOWN",
    cp_model.INFEASIBLE: "INFEASIBLE",
    cp_model.MODEL_INVALID: "MODEL_INVALID"
}.get(result, f"STATUS_{result}")

    return {
        'status': status,
        'objective': obj,
        'solution': canonical,
        'solver_code': int(result),
        'seed': data.random_seed,
        'num_solutions_captured': callback.num_solutions()
    }

# ----------------------------
# Visualization helpers (Plotly)
# ----------------------------
def build_production_schedule_df(solution: Dict[str, Any], data: ProblemData) -> pd.DataFrame:
    rows = []
    prod = solution.get('production', {})
    isprod = solution.get('is_producing', {})
    for (g, p, d), val in prod.items():
        rows.append({
            'Grade': g,
            'Plant': p,
            'DayIndex': d,
            'Date': data.dates[d],
            'Production_MT': to_float_mt(val),
            'IsProducing': isprod.get((g, p, d), 0)
        })
    return pd.DataFrame(rows)

def gantt_schedule_plot(prod_df: pd.DataFrame, data: ProblemData) -> go.Figure:
    """
    Create a Gantt-like figure: each row is a plant, colored by grade across days.
    Because production is full-capacity only, we mark days as blocks.
    """
    if prod_df.empty:
        return go.Figure()
    # For Gantt, build segments per (Plant, Grade) with contiguous day runs
    items = []
    for p in data.plants:
        for g in data.grades:
            sub = prod_df[(prod_df['Plant'] == p) & (prod_df['Grade'] == g) & (prod_df['IsProducing'] == 1)].sort_values('DayIndex')
            if sub.empty:
                continue
            # find contiguous ranges
            idxs = sub['DayIndex'].tolist()
            start = idxs[0]
            prev = start
            for i in idxs[1:] + [None]:
                if i is None or i != prev + 1:
                    # emit block start..prev
                    items.append({'Plant': p, 'Grade': g, 'Start': data.dates[start], 'End': data.dates[prev] + timedelta(days=1)})
                    if i is not None:
                        start = i
                        prev = i
                else:
                    prev = i
    if not items:
        return go.Figure()
    df_items = pd.DataFrame(items)
    # Use plotly.timeline
    fig = px.timeline(df_items, x_start="Start", x_end="End", y="Plant", color="Grade", title="Production Schedule (Gantt)")
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=400)
    return fig

def inventory_time_series(solution: Dict[str, Any], data: ProblemData) -> go.Figure:
    inv = solution.get('inventory', {})
    rows = []
    for (g, d), val in inv.items():
        # d may be 0..num_days
        dt = data.dates[d] if d < len(data.dates) else (data.dates[-1] + timedelta(days=1))
        rows.append({'Grade': g, 'Date': dt, 'Inventory_MT': to_float_mt(val)})
    df = pd.DataFrame(rows)
    if df.empty:
        return go.Figure()
    fig = px.line(df, x='Date', y='Inventory_MT', color='Grade', markers=True, title='Inventory over Time')
    return fig

def stockout_bar_chart(solution: Dict[str, Any], data: ProblemData) -> go.Figure:
    so = solution.get('stockout', {})
    rows = []
    for (g, d), val in so.items():
        rows.append({'Grade': g, 'Date': data.dates[d], 'Stockout_MT': to_float_mt(val)})
    df = pd.DataFrame(rows)
    if df.empty:
        return go.Figure()
    fig = px.bar(df, x='Date', y='Stockout_MT', color='Grade', barmode='stack', title='Stockouts by Day')
    return fig

# ----------------------------
# Streamlit app
# ----------------------------
def run_streamlit_app():
    if st is None:
        print("Streamlit not available.")
        return
    st.set_page_config(layout="wide", page_title="Polymer Production Scheduler")
    st.title("Polymer Production Scheduler — Deterministic & Full Capacity")

    st.markdown("""
    **Important notes**
    - Inputs (capacity, demand, inventory) must be in Metric Tonnes (MT).
    - Internal scaling: 1 MT = 1000 internal units. Results are shown in MT.
    - Production is forced to full plant capacity when scheduled.
    - Deterministic solve: single worker and fixed seed by default (change in sidebar if needed).
    """)

    uploaded = st.file_uploader("Upload Excel file with sheets: Plant, Inventory, Demand", type=["xls", "xlsx"])
    if uploaded is None:
        st.info("Please upload the template Excel file (no fallback generation).")
        return

    try:
        data = read_input_workbook(uploaded.read())
    except Exception as e:
        st.error(f"Error reading workbook: {e}")
        return

    with st.sidebar:
        st.header("Solver options")
        seed = st.number_input("Random seed", value=data.random_seed, step=1)
        workers = st.number_input("Search workers", value=data.num_workers, step=1)
        time_limit = st.number_input("Time limit (sec)", value=60.0, step=10.0)
        if seed != data.random_seed or workers != data.num_workers:
            data.random_seed = int(seed)
            data.num_workers = int(workers)

    st.write(f"Planning horizon: {len(data.dates)} days")
    if st.button("Solve"):
        with st.spinner("Solving..."):
            res = solve_problem(data, time_limit_seconds=float(time_limit))
        st.success(f"Status: {res['status']} | Objective: {res['objective']} | Seed: {res['seed']}")
        sol = res.get('solution')
        if not sol:
            st.error("No feasible solution found.")
            return

        # Production DataFrame
        prod_df = build_production_schedule_df(sol, data)
        st.subheader("Production schedule (per plant/day)")
        if prod_df.empty:
            st.write("No production scheduled.")
        else:
            st.dataframe(prod_df.sort_values(['Plant', 'Date', 'Grade']))

        # Visualizations
        st.subheader("Visualizations")
        g_fig = gantt_schedule_plot(prod_df, data)
        if g_fig and g_fig.data:
            st.plotly_chart(g_fig, use_container_width=True)
        inv_fig = inventory_time_series(sol, data)
        if inv_fig and inv_fig.data:
            st.plotly_chart(inv_fig, use_container_width=True)
        so_fig = stockout_bar_chart(sol, data)
        if so_fig and so_fig.data:
            st.plotly_chart(so_fig, use_container_width=True)

        # Provide download of production table as CSV
        if not prod_df.empty:
            csv = prod_df.to_csv(index=False)
            st.download_button("Download production CSV", csv, file_name="production_schedule.csv", mime="text/csv")

# ----------------------------
# Small deterministic test (CLI)
# ----------------------------
def small_test_run():
    data = ProblemData()
    data.random_seed = DEFAULT_RANDOM_SEED
    data.num_workers = DEFAULT_NUM_WORKERS
    today = date.today()
    data.dates = [today + timedelta(days=i) for i in range(3)]
    data.plants = ['PlantA']
    data.capacities = {'PlantA': to_int_mt(10.0)}
    data.grades = ['G1', 'G2']
    data.allowed_lines = {'G1': ['PlantA'], 'G2': ['PlantA']}
    data.initial_inventory = {'G1': to_int_mt(2.0), 'G2': to_int_mt(0.0)}
    data.min_closing_inventory = {'G1': to_int_mt(0.0), 'G2': to_int_mt(0.0)}
    data.min_run_days = {'G1': 1, 'G2': 1}
    data.max_run_days = {'G1': 3, 'G2': 3}
    data.demand = {
        'G1': {data.dates[0]: to_int_mt(5.0), data.dates[1]: to_int_mt(0.0), data.dates[2]: to_int_mt(0.0)},
        'G2': {data.dates[0]: to_int_mt(0.0), data.dates[1]: to_int_mt(6.0), data.dates[2]: to_int_mt(0.0)}
    }
    data.shutdown_days = {'PlantA': set()}
    return solve_problem(data, time_limit_seconds=10.0)

# ----------------------------
# Main
# ----------------------------
if __name__ == "__main__":
    # If streamlit is running, use run_streamlit_app. Otherwise run small test.
    if st is not None:
        run_streamlit_app()
    else:
        print("Streamlit not available — running CLI deterministic test.")
        r1 = small_test_run()
        print("Run 1:", r1['status'], "obj:", r1['objective'])
        r2 = small_test_run()
        print("Run 2:", r2['status'], "obj:", r2['objective'])
        print("Solutions identical:", r1['solution'] == r2['solution'])
