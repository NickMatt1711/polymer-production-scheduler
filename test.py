
import io
import os
import math
import logging
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Dict, List, Tuple, Any

import pandas as pd

# OR-Tools CP-SAT
from ortools.sat.python import cp_model

# Streamlit for UI (optional when running as script)
try:
    import streamlit as st
except Exception:
    st = None  # If streamlit is not available, the module can still be used programmatically.

# ----------------------------
# Configuration & Constants
# ----------------------------

# Scale: convert Metric Tonnes (MT) to integer units for CP-SAT.
# 1 MT = 1000 units (kg). Using 1000 preserves 0.001 MT precision.
SCALE = 1000

# Deterministic solver settings
DEFAULT_RANDOM_SEED = 123456789
DEFAULT_NUM_WORKERS = 1  # 1 worker ensures reproducible search order in CP-SAT.

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("app_fixed")

# ----------------------------
# Helper functions
# ----------------------------

def to_int_mt(value: Any) -> int:
    """
    Convert input numeric value in Metric Tonnes (MT) to internal integer units (SCALE).
    Treat NaN or invalid values as 0.
    """
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
    """
    Convert internal integer units back to Metric Tonnes (MT).
    """
    return float(value) / SCALE

def ensure_date(d) -> date:
    """
    Convert a value to a datetime.date object (preserve input order & raise on invalid).
    """
    if isinstance(d, date):
        return d
    if isinstance(d, datetime):
        return d.date()
    try:
        return pd.to_datetime(d).date()
    except Exception as e:
        raise ValueError(f"Invalid date value: {d}") from e

# ----------------------------
# Data model structures
# ----------------------------

class ProblemData:
    """
    Container for data required to build the CP-SAT model.
    All numeric values are converted to integer units (SCALE).
    """
    def __init__(self):
        self.dates: List[date] = []
        self.grades: List[str] = []
        self.plants: List[str] = []
        # capacities[plant] -> int (scaled)
        self.capacities: Dict[str, int] = {}
        # allowed_lines_by_grade[grade] -> list of plants
        self.allowed_lines: Dict[str, List[str]] = {}
        # initial_inventory[grade] -> int (scaled)
        self.initial_inventory: Dict[str, int] = {}
        # min_closing_inventory[grade] -> int (scaled)
        self.min_closing_inventory: Dict[str, int] = {}
        # min_run_days[grade] and max_run_days[grade] (in days)
        self.min_run_days: Dict[str, int] = {}
        self.max_run_days: Dict[str, int] = {}
        # demand[grade][date] -> int (scaled)
        self.demand: Dict[str, Dict[date, int]] = {}
        # transition_rules[plant][prev_grade] = set(allowed_next_grades)  (if plant-specific transitions exist)
        self.transition_rules: Dict[str, Dict[str, List[str]]] = {}
        # shutdown_days[plant] -> set(day_index)
        self.shutdown_days: Dict[str, set] = {}
        # penalties (all in scaled units where applicable)
        self.penalty_stockout_per_mt: int = to_int_mt(1000)  # default big penalty; user may override
        self.penalty_transition: int = 1  # penalty per disallowed transition (as integer)
        # other options
        self.force_full_capacity = True  # user's requirement (must be True)
        self.random_seed = DEFAULT_RANDOM_SEED
        self.num_workers = DEFAULT_NUM_WORKERS

# ----------------------------
# Model builder
# ----------------------------

def build_model(data: ProblemData) -> Tuple[cp_model.CpModel, Dict[str, Any]]:
    """
    Build the CP-SAT model using the ProblemData inputs.

    Returns:
        model, vars_dict
    where vars_dict contains:
        - production[(grade, plant, day)] -> IntVar (scaled units)
        - is_producing[(grade, plant, day)] -> BoolVar
        - inventory[(grade, day_idx)] -> IntVar (scaled)
        - stockout[(grade, day_idx)] -> IntVar (scaled)
        - auxiliary indexing maps as needed
    """
    model = cp_model.CpModel()
    vars_dict = {}

    grades = data.grades
    plants = data.plants
    dates = data.dates
    num_days = len(dates)

    # Domains for inventory & stockout
    max_total_capacity = max(data.capacities.values()) if data.capacities else 0
    BIG = sum(data.capacities.values()) * num_days + sum(data.initial_inventory.values()) + 10 * SCALE

    # Create production & is_producing vars
    production = {}
    is_producing = {}
    for g in grades:
        for p in data.allowed_lines.get(g, []):
            cap = data.capacities.get(p, 0)
            # ensure capacity is non-negative int
            cap = int(max(0, cap))
            for d in range(num_days):
                key = (g, p, d)
                # production quantity variable (0..cap)
                prod_var = model.NewIntVar(0, cap, f'prod_{g}_{p}_{d}')
                production[key] = prod_var
                # boolean indicating whether this grade runs on this plant & day
                run_var = model.NewBoolVar(f'isprod_{g}_{p}_{d}')
                is_producing[key] = run_var

                # Enforce user's requirement: full-capacity production when running, zero when not running
                # production == cap if run_var == 1
                model.Add(prod_var == cap).OnlyEnforceIf(run_var)
                # production == 0 if run_var == 0
                model.Add(prod_var == 0).OnlyEnforceIf(run_var.Not())

                # Enforce shutdowns: if the plant is shutdown on this day, run_var == 0
                shutdown_set = data.shutdown_days.get(p, set())
                if d in shutdown_set:
                    model.Add(run_var == 0)

    vars_dict['production'] = production
    vars_dict['is_producing'] = is_producing

    # Inventory and stockout variables
    inventory = {}
    stockout = {}

    for g in grades:
        for d in range(num_days + 1):
            # inventory at day d (inventory[grade, day]) with day 0 = opening inventory
            inv_var = model.NewIntVar(0, BIG, f'inv_{g}_{d}')
            inventory[(g, d)] = inv_var
    for g in grades:
        for d in range(num_days):
            so_var = model.NewIntVar(0, BIG, f'stockout_{g}_{d}')
            stockout[(g, d)] = so_var

    vars_dict['inventory'] = inventory
    vars_dict['stockout'] = stockout

    # Initialize opening inventory (day 0)
    for g in grades:
        inv0 = data.initial_inventory.get(g, 0)
        model.Add(inventory[(g, 0)] == inv0)

    # Demand & inventory balance
    for g in grades:
        for d in range(num_days):
            demand_val = data.demand.get(g, {}).get(data.dates[d], 0)
            # Compute total production for grade g on day d across all plants
            prods = []
            for p in data.allowed_lines.get(g, []):
                key = (g, p, d)
                if key in production:
                    prods.append(production[key])
            if not prods:
                # No plant can produce this grade; production sum = 0 constant
                total_prod = model.NewConstant(0)
            else:
                total_prod = sum(prods)

            # Inventory balance:
            # inv_{d} + total_prod = demand - stockout + inv_{d+1}
            # Rearranged: inv_{d+1} = inv_{d} + total_prod - (demand - stockout)
            # All ints (scaled)
            model.Add(inventory[(g, d + 1)] + (demand_val - stockout[(g, d)]) == inventory[(g, d)] + total_prod)

            # stockout cannot exceed demand
            model.Add(stockout[(g, d)] <= demand_val)

    # Min closing inventory at planning horizon end
    last_day = num_days
    for g in grades:
        min_cl = data.min_closing_inventory.get(g, 0)
        model.Add(inventory[(g, last_day)] >= min_cl)

    # Run-length constraints using sliding-window encoding:
    # For each (grade, plant), the boolean sequence is is_producing[(g,p,d)] for d in 0..num_days-1
    # Enforce max_run_days: for every window of size (max_run_days + 1), sum <= max_run_days
    # Enforce min_run_days: whenever a start occurs (isprod[d]==1 and isprod[d-1]==0) we require
    # sum(isprod[d .. d+min_run_days-1]) >= min_run_days
    for g in grades:
        min_run = data.min_run_days.get(g, 1)
        max_run = data.max_run_days.get(g, num_days)
        for p in data.allowed_lines.get(g, []):
            # collect is_producing bools for this pair
            bools = [is_producing[(g, p, d)] for d in range(num_days)]
            # max run encoding (sliding window)
            if max_run < num_days:
                window = max_run + 1
                for start in range(0, num_days - window + 1):
                    model.Add(sum(bools[start:start + window]) <= max_run)
            # min run encoding via start indicator
            if min_run > 1:
                # build start indicators: start_d == 1 if bool[d]==1 and (d==0 or bool[d-1]==0)
                for d in range(num_days):
                    curr = bools[d]
                    if d == 0:
                        # if production at day 0 -> it is a start
                        start_bool = model.NewBoolVar(f'start_{g}_{p}_{d}')
                        model.Add(start_bool == curr)
                    else:
                        prev = bools[d - 1]
                        start_bool = model.NewBoolVar(f'start_{g}_{p}_{d}')
                        # start_bool >= curr - prev  and start_bool <= curr
                        # Implement equivalently with reified constraints:
                        model.Add(curr == 1).OnlyEnforceIf(start_bool)
                        model.AddBoolOr([curr.Not(), start_bool.Not()])
                        # The above two constraints ensure start_bool -> curr and (curr and prev==0) -> start_bool
                        # To enforce start_bool implies prev==0:
                        model.Add(prev == 0).OnlyEnforceIf(start_bool)

                    # If start_bool == 1 then ensure sum(curr..curr+min_run-1) >= min_run
                    if d + min_run <= num_days:
                        span = bools[d:d + min_run]
                        # sum(span) >= min_run * start_bool
                        model.Add(sum(span) >= min_run).OnlyEnforceIf(start_bool)
                    else:
                        # If a start occurs too late to satisfy min_run before horizon end -> forbid start
                        model.Add(start_bool == 0)

    # Transition rules (disallow certain grade sequences on same plant across adjacent days)
    # For each plant and each adjacent day pair, if prev_grade -> next_grade is disallowed, enforce prev + next <= 1
    for p in plants:
        rules_for_plant = data.transition_rules.get(p, {})
        for d in range(num_days - 1):
            for prev_grade, allowed_next in rules_for_plant.items():
                # if allowed_next is None => allow all
                if allowed_next is None:
                    continue
                # iterate candidate next grades (all grades that can be produced on p)
                for curr_grade in [g for g in grades if p in data.allowed_lines.get(g, [])]:
                    if curr_grade not in allowed_next:
                        prev_key = (prev_grade, p, d)
                        next_key = (curr_grade, p, d + 1)
                        if prev_key in is_producing and next_key in is_producing:
                            model.Add(is_producing[prev_key] + is_producing[next_key] <= 1)

    # Objective: minimize weighted sum of stockouts + transitions (if any)
    objective_terms = []

    # stockout penalty
    for g in grades:
        for d in range(num_days):
            objective_terms.append(stockout[(g, d)] * data.penalty_stockout_per_mt)

    # Optional: penalize transitions that are disallowed less severe - we already disallow them,
    # but if there are soft preferences they can be added here (not used now)

    model.Minimize(sum(objective_terms))

    # Pack variables for return
    vars_dict['grades'] = grades
    vars_dict['plants'] = plants
    vars_dict['dates'] = dates
    vars_dict['num_days'] = num_days

    return model, vars_dict

# ----------------------------
# Solution callback (capture intermediate solutions)
# ----------------------------

class CaptureSolutionsCallback(cp_model.CpSolverSolutionCallback):
    """
    Capture intermediate solutions produced by CP-SAT. We will use the last captured solution
    as the canonical solution for display to avoid mismatch with solver.Value() after search.
    """
    def __init__(self, variables_map: Dict[str, Any]):
        super().__init__()
        self._vars = variables_map
        self.solutions = []
        self._solution_count = 0

    def on_solution_callback(self):
        self._solution_count += 1
        # capture relevant vars: is_producing and production and inventory/stockout if present
        sol = {
            'is_producing': {},
            'production': {},
            'inventory': {},
            'stockout': {}
        }
        prod_map = self._vars.get('production', {})
        isprod_map = self._vars.get('is_producing', {})
        inv_map = self._vars.get('inventory', {})
        so_map = self._vars.get('stockout', {})

        # store is_producing
        for k, v in isprod_map.items():
            try:
                sol['is_producing'][k] = int(self.Value(v))
            except Exception:
                sol['is_producing'][k] = 0
        for k, v in prod_map.items():
            try:
                sol['production'][k] = int(self.Value(v))
            except Exception:
                sol['production'][k] = 0
        for k, v in inv_map.items():
            try:
                sol['inventory'][k] = int(self.Value(v))
            except Exception:
                sol['inventory'][k] = 0
        for k, v in so_map.items():
            try:
                sol['stockout'][k] = int(self.Value(v))
            except Exception:
                sol['stockout'][k] = 0

        # capture solution metadata (objective, time)
        try:
            obj = self.ObjectiveValue()
        except Exception:
            obj = None

        self.solutions.append({'solution': sol, 'objective': obj, 'time': self.WallTime()})

    def num_solutions(self):
        return len(self.solutions)

# ----------------------------
# Utility functions for reading Excel inputs (Streamlit UI integration)
# ----------------------------

def read_input_workbook(xl_bytes: bytes) -> ProblemData:
    """
    Parse Excel bytes and return ProblemData.
    Expects sheets: 'Plant', 'Inventory', 'Demand' at minimum.
    The function preserves input date row order and validates columns.
    Numeric fields are converted to internal integer units (SCALE).
    """
    data = ProblemData()

    xio = io.BytesIO(xl_bytes)
    xl = pd.ExcelFile(xio)

    if 'Plant' not in xl.sheet_names or 'Inventory' not in xl.sheet_names or 'Demand' not in xl.sheet_names:
        raise ValueError("Excel file must contain 'Plant', 'Inventory', and 'Demand' sheets with required columns.")

    # Read Plant sheet
    plant_df = xl.parse('Plant', dtype=str)
    plant_df = plant_df.fillna('')

    # Expected Plant columns: Plant, Capacity per day (MT), Lines (optional), Shutdown Start Date, Shutdown End Date
    # Normalize column names by lowercase & strip
    plant_df.columns = [c.strip() for c in plant_df.columns]
    # Build plants & capacities
    plants = []
    capacities = {}
    shutdown_days = {}
    for idx, row in plant_df.iterrows():
        plant_name = str(row.get('Plant', '')).strip()
        if not plant_name:
            continue
        plants.append(plant_name)
        cap = to_int_mt(row.get('Capacity per day', 0))
        capacities[plant_name] = cap
        # Shutdown dates may be blank; if present convert to date range
        start_raw = row.get('Shutdown Start Date', None)
        end_raw = row.get('Shutdown End Date', None)
        try:
            if pd.notna(start_raw) and pd.notna(end_raw):
                start_d = ensure_date(start_raw)
                end_d = ensure_date(end_raw)
                shutdown_days[plant_name] = (start_d, end_d)
        except Exception:
            shutdown_days[plant_name] = None

    data.plants = plants
    data.capacities = capacities

    # Read Inventory sheet
    inv_df = xl.parse('Inventory', dtype=str)
    inv_df.columns = [c.strip() for c in inv_df.columns]
    inv_df = inv_df.fillna('')

    grades = []
    allowed_lines = {}
    initial_inventory = {}
    min_closing_inventory = {}
    min_run = {}
    max_run = {}

    for idx, row in inv_df.iterrows():
        grade = str(row.get('Grade Name', '')).strip()
        if not grade:
            continue
        grades.append(grade)
        lines_field = str(row.get('Lines', '')).strip()
        # Lines can be comma-separated plant names; allow empty => no permitted lines
        lines = [s.strip() for s in lines_field.split(',')] if lines_field else []
        allowed_lines[grade] = [ln for ln in lines if ln]
        initial_inventory[grade] = to_int_mt(row.get('Opening Inventory', 0))
        min_closing_inventory[grade] = to_int_mt(row.get('Min. Closing Inventory', 0))
        min_run[grade] = int(float(row.get('Min. Run Days', 1))) if row.get('Min. Run Days', '') != '' else 1
        max_run[grade] = int(float(row.get('Max. Run Days', max(1, min_run.get(grade, 1)))))

    data.grades = grades
    data.allowed_lines = allowed_lines
    data.initial_inventory = initial_inventory
    data.min_closing_inventory = min_closing_inventory
    data.min_run_days = min_run
    data.max_run_days = max_run

    # Read Demand sheet
    demand_df = xl.parse('Demand')
    # preserve row order; first column must be date
    demand_df.columns = [c.strip() for c in demand_df.columns]
    if demand_df.shape[1] < 2:
        raise ValueError("Demand sheet must have at least two columns: Date and one grade demand column.")
    # Convert first column to dates
    demand_dates = pd.to_datetime(demand_df.iloc[:, 0], errors='coerce')
    if demand_dates.isnull().any():
        raise ValueError("Invalid date values in Demand sheet first column.")
    dates = [ensure_date(d) for d in demand_dates.tolist()]
    data.dates = dates

    # Remaining columns are assumed to be grades - match with inventory grades
    demand_grades = [c for c in demand_df.columns[1:]]
    # For each grade present in demand, map values into data.demand
    for g in demand_grades:
        if g not in grades:
            # Ignore demand columns for grades not declared in Inventory sheet but log a warning
            logger.warning(f"Grade '{g}' found in Demand but not in Inventory sheet. Adding it with zero inventory defaults.")
            # add default structures for this new grade
            grades.append(g)
            data.initial_inventory[g] = 0
            data.min_closing_inventory[g] = 0
            data.allowed_lines[g] = []
            data.min_run_days[g] = 1
            data.max_run_days[g] = len(dates)
        gvals = {}
        for idx, d in enumerate(dates):
            raw_val = demand_df.iloc[idx, demand_df.columns.get_loc(g)]
            gvals[d] = to_int_mt(raw_val)
        data.demand[g] = gvals

    # Process shutdown_days mapping to day indices
    # Map plant shutdown date ranges (if provided) to indices in data.dates
    shutdown_indices = {}
    for p in data.plants:
        sd_tuple = shutdown_days.get(p, None)
        if not sd_tuple:
            shutdown_indices[p] = set()
            continue
        start_d, end_d = sd_tuple
        idxs = set()
        for i, d in enumerate(data.dates):
            if start_d <= d <= end_d:
                idxs.add(i)
        shutdown_indices[p] = idxs

    data.shutdown_days = shutdown_indices

    # Default penalties: choose sensible scaled values if not provided
    # Here, we set stockout penalty high to prioritize avoiding stockouts.
    data.penalty_stockout_per_mt = int(1000 * SCALE)  # e.g., 1000 MT penalty per 1 MT stockout (very large)
    data.penalty_transition = 1

    return data

# ----------------------------
# Solve function
# ----------------------------

def solve_problem(data: ProblemData, time_limit_seconds: float = 60.0) -> Dict[str, Any]:
    """
    Solve the scheduling problem with CP-SAT deterministically.
    Returns a dictionary with results (using the last callback solution if any).
    """
    model, vars_map = build_model(data)

    solver = cp_model.CpSolver()
    # Deterministic parameters
    solver.parameters.max_time_in_seconds = float(time_limit_seconds)
    solver.parameters.random_seed = int(data.random_seed)
    solver.parameters.num_search_workers = int(data.num_workers)
    # reduce log volume but keep some info
    solver.parameters.log_search_progress = False

    # Capture intermediate solutions
    callback = CaptureSolutionsCallback(vars_map)
    # Call Solve with callback to capture intermediate solutions, but we will use the last captured solution
    result = solver.SolveWithSolutionCallback(model, callback)

    # Gather canonical solution: prefer last callback solution if any; else use solver values
    if callback.num_solutions() > 0:
        canonical = callback.solutions[-1]['solution']
        obj_val = callback.solutions[-1]['objective']
    else:
        # No callback solutions captured -> fallback to solver.Value reads for final solution
        canonical = {'is_producing': {}, 'production': {}, 'inventory': {}, 'stockout': {}}
        prod_map = vars_map.get('production', {})
        isprod_map = vars_map.get('is_producing', {})
        inv_map = vars_map.get('inventory', {})
        so_map = vars_map.get('stockout', {})
        try:
            for k, v in isprod_map.items():
                canonical['is_producing'][k] = int(solver.Value(v))
            for k, v in prod_map.items():
                canonical['production'][k] = int(solver.Value(v))
            for k, v in inv_map.items():
                canonical['inventory'][k] = int(solver.Value(v))
            for k, v in so_map.items():
                canonical['stockout'][k] = int(solver.Value(v))
            obj_val = solver.ObjectiveValue()
        except Exception:
            # If solver didn't return a feasible solution, return empty result and status code
            canonical = None
            obj_val = None

    status_str = {
        cp_model.OPTIMAL: "OPTIMAL",
        cp_model.FEASIBLE: "FEASIBLE",
        cp_model.UNKNOWN: "UNKNOWN",
        cp_model.INFEASIBLE: "INFEASIBLE",
        cp_model.MODEL_INVALID: "MODEL_INVALID",
        cp_model.NOT_SOLVED: "NOT_SOLVED"
    }.get(result, f"STATUS_{result}")

    results = {
        'status': status_str,
        'objective': obj_val,
        'solution': canonical,
        'solver_status_code': int(result),
        'solver_seed': data.random_seed,
        'num_solutions_captured': callback.num_solutions()
    }

    return results

# ----------------------------
# Small deterministic test harness
# ----------------------------

def small_test_run() -> Dict[str, Any]:
    """
    Build a small deterministic instance (2 grades, 1 plant, 3 days) to validate behavior.
    This can be used to check repeatability.
    """
    data = ProblemData()
    data.random_seed = DEFAULT_RANDOM_SEED
    data.num_workers = DEFAULT_NUM_WORKERS

    # Dates
    today = date.today()
    data.dates = [today + timedelta(days=i) for i in range(3)]
    data.num_days = 3

    # Plants & capacities (1 plant)
    data.plants = ['PlantA']
    # capacity 10 MT/day -> scaled
    data.capacities = {'PlantA': to_int_mt(10.0)}

    # Grades: G1, G2 both allowed on PlantA
    data.grades = ['G1', 'G2']
    data.allowed_lines = {'G1': ['PlantA'], 'G2': ['PlantA']}

    # Inventories
    data.initial_inventory = {'G1': to_int_mt(2.0), 'G2': to_int_mt(0.0)}
    data.min_closing_inventory = {'G1': to_int_mt(0.0), 'G2': to_int_mt(0.0)}

    # Run days constraints
    data.min_run_days = {'G1': 1, 'G2': 1}
    data.max_run_days = {'G1': 3, 'G2': 3}

    # Demand (MT)
    # Day0 Day1 Day2
    data.demand = {
        'G1': {data.dates[0]: to_int_mt(5.0), data.dates[1]: to_int_mt(0.0), data.dates[2]: to_int_mt(0.0)},
        'G2': {data.dates[0]: to_int_mt(0.0), data.dates[1]: to_int_mt(6.0), data.dates[2]: to_int_mt(0.0)},
    }

    # No shutdowns
    data.shutdown_days = {'PlantA': set()}

    # Penalties
    data.penalty_stockout_per_mt = int(1000 * SCALE)
    data.penalty_transition = 1

    # Force full capacity production is true by default in ProblemData
    data.force_full_capacity = True

    # Solve
    res = solve_problem(data, time_limit_seconds=10.0)
    return res

# ----------------------------
# Entry point for Streamlit UI (if used)
# ----------------------------

def run_streamlit_app():
    if st is None:
        print("Streamlit not available in this environment. Run the script with Streamlit in a proper env.")
        return

    st.set_page_config(layout="wide", page_title="Polymer Production Scheduler (Deterministic)")
    st.title("Polymer Production Scheduler — Deterministic, Scaled (MT)")

    st.markdown("""
    **Notes**
    - All numeric inputs (capacity, demand, inventory) are in Metric Tonnes (MT).
    - Internally the model scales MT -> integer units using SCALE = 1000 (1 MT -> 1000).
    - Production, when scheduled, will be exactly the plant capacity (user requirement).
    - The solver is deterministic (fixed random seed and single worker).
    """)

    uploaded = st.file_uploader("Upload Excel template (must contain 'Plant', 'Inventory', 'Demand' sheets)", type=['xls', 'xlsx'])
    if uploaded is None:
        st.info("Please upload the Excel file with Plant, Inventory and Demand sheets.")
        return

    try:
        data = read_input_workbook(uploaded.read())
    except Exception as e:
        st.error(f"Failed to read workbook: {e}")
        return

    # Parameters panel
    with st.sidebar:
        st.header("Solver Options")
        seed = st.number_input("Solver random seed", value=data.random_seed, step=1)
        workers = st.number_input("Search workers (set 1 for reproducible)", value=data.num_workers, step=1)
        time_limit = st.number_input("Time limit (seconds)", value=60.0, step=10.0)
        if seed != data.random_seed or workers != data.num_workers:
            data.random_seed = int(seed)
            data.num_workers = int(workers)

    st.write("Planning horizon:", len(data.dates), "days")
    if st.button("Solve"):
        with st.spinner("Solving..."):
            results = solve_problem(data, time_limit_seconds=float(time_limit))
        st.success(f"Solve status: {results['status']} — Objective: {results['objective']}")
        st.write("Solver seed:", results.get('solver_seed'))
        # Display summary tables: production schedule, inventory, stockout
        sol = results.get('solution')
        if sol is None:
            st.error("No feasible solution found.")
            return
        # Build dataframes for display
        prod_rows = []
        for (g, p, d), val in sol['production'].items():
            mt_val = to_float_mt(val)
            prod_rows.append({'Grade': g, 'Plant': p, 'Day': data.dates[d], 'Production_MT': mt_val, 'IsProducing': sol['is_producing'].get((g, p, d), 0)})
        if prod_rows:
            prod_df = pd.DataFrame(prod_rows)
            st.dataframe(prod_df)
        inv_rows = []
        for (g, day_idx), val in sol['inventory'].items():
            mt_val = to_float_mt(val)
            inv_rows.append({'Grade': g, 'DayIndex': day_idx, 'Day': data.dates[day_idx] if day_idx < len(data.dates) else 'End', 'Inventory_MT': mt_val})
        if inv_rows:
            inv_df = pd.DataFrame(inv_rows)
            st.dataframe(inv_df)
        so_rows = []
        for (g, d), val in sol['stockout'].items():
            mt_val = to_float_mt(val)
            so_rows.append({'Grade': g, 'Day': data.dates[d], 'Stockout_MT': mt_val})
        if so_rows:
            so_df = pd.DataFrame(so_rows)
            st.dataframe(so_df)

# ----------------------------
# If run as script, perform a small deterministic test and write results to stdout
# ----------------------------

if __name__ == '__main__':
    # Run small deterministic test to demonstrate repeatability
    print("Running small deterministic test instance...")
    r1 = small_test_run()
    print("First run status:", r1['status'], "objective:", r1['objective'], "seed:", r1.get('solver_seed'))
    # Run again to verify determinism with same seed
    r2 = small_test_run()
    print("Second run status:", r2['status'], "objective:", r2['objective'], "seed:", r2.get('solver_seed'))
    identical = (r1['solution'] == r2['solution'])
    print("Solutions identical across runs:", identical)
