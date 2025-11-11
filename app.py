# app.py (FULL script)
# Optimized polymer production scheduler with fixes:
# - Solver performance improvements (cached sums / less Python work inside loops)
# - Restored Total Stockout in production table
# - Inventory charts corrected to plot inventory variable values
# - Gridlines restored in Plotly visualizations
# - All displayed dates consistently formatted DD-MMM-YY

import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta, datetime, date
import numpy as np
import time
import io
from matplotlib import colormaps
import plotly.express as px
import plotly.graph_objects as go

# ----------------- Config -----------------
DATE_DISPLAY = "%d-%b-%y"      # DD-MMM-YY
LARGE_INT = 200_000
DEFAULT_NUM_WORKERS = 8

st.set_page_config(page_title="Polymer Production Scheduler", page_icon="🏭", layout="wide")

# ----------------- Utilities -----------------
def fmt(d):
    """Format date to DD-MMM-YY for UI display."""
    if d is None:
        return ""
    if isinstance(d, (datetime, pd.Timestamp)):
        d = d.date()
    if isinstance(d, date):
        return d.strftime(DATE_DISPLAY)
    try:
        return pd.to_datetime(d).date().strftime(DATE_DISPLAY)
    except Exception:
        return str(d)

def to_date(val):
    """Convert value to datetime.date or None."""
    if pd.isna(val) or val == "":
        return None
    if isinstance(val, date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None

def create_sample_workbook():
    """Create sample Excel bytes with template data (dates in workbook are datetimes)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        plant_df = pd.DataFrame({
            "Plant": ["Plant1", "Plant2"],
            "Capacity per day": [1500, 1000],
            "Material Running": ["Moulding", "BOPP"],
            "Expected Run Days": [1, 3],
            "Shutdown Start Date": [None, "15-Nov-25"],
            "Shutdown End Date": [None, "18-Nov-25"]
        })
        plant_df.to_excel(writer, sheet_name="Plant", index=False)

        inventory_df = pd.DataFrame({
            "Grade Name": ["BOPP", "Moulding", "Raffia", "TQPP", "Yarn"],
            "Opening Inventory": [500, 16000, 3000, 1700, 2500],
            "Min. Closing Inventory": [1000, 2000, 1000, 500, 500],
            "Min. Inventory": [500, 1000, 1000, 0, 0],
            "Max. Inventory": [20000, 20000, 20000, 6000, 6000],
            "Min. Run Days": [3, 1, 1, 2, 2],
            "Max. Run Days": [10, 10, 10, 8, 8],
            "Increment Days": ['', '', '', '', ''],
            "Force Start Date": ['', '', '', '', ''],
            "Lines": ['Plant1, Plant2', 'Plant1, Plant2', 'Plant1, Plant2', 'Plant1, Plant2', 'Plant2'],
            "Rerun Allowed": ['Yes', 'Yes', 'Yes', 'No', 'No']
        })
        inventory_df.to_excel(writer, sheet_name="Inventory", index=False)

        # Demand for November 2025 (30 days)
        current_year, current_month = 2025, 11
        dates = pd.date_range(start=f"{current_year}-{current_month:02d}-01", periods=30, freq="D")
        demand_df = pd.DataFrame({
            "Date": dates,
            "BOPP": [400]*len(dates),
            "Moulding": [400]*len(dates),
            "Raffia": [600]*len(dates),
            "TQPP": [300]*len(dates),
            "Yarn": [100]*len(dates)
        })
        demand_df.to_excel(writer, sheet_name="Demand", index=False)

        # Transition sheets
        trans1 = pd.DataFrame({
            "From": ["BOPP", "Moulding", "Raffia", "TQPP"],
            "BOPP": ["Yes", "No", "Yes", "No"],
            "Moulding": ["No", "Yes", "Yes", "Yes"],
            "Raffia": ["Yes", "Yes", "Yes", "Yes"],
            "TQPP": ["No", "Yes", "Yes", "Yes"]
        })
        trans1.to_excel(writer, sheet_name="Transition_Plant1", index=False)

        trans2 = pd.DataFrame({
            "From": ["BOPP", "Moulding", "Raffia", "TQPP", "Yarn"],
            "BOPP": ["Yes", "No", "Yes", "Yes", "No"],
            "Moulding": ["No", "Yes", "Yes", "Yes", "Yes"],
            "Raffia": ["Yes", "Yes", "Yes", "Yes", "No"],
            "TQPP": ["Yes", "Yes", "Yes", "Yes", "No"],
            "Yarn": ["No", "Yes", "No", "No", "Yes"]
        })
        trans2.to_excel(writer, sheet_name="Transition_Plant2", index=False)

    output.seek(0)
    return output

# ----------------- Helper: read sheets -----------------
def read_all_sheets(excel_bytes):
    xls = pd.ExcelFile(excel_bytes)
    sheets = {}
    for name in ["Plant", "Inventory", "Demand"]:
        if name in xls.sheet_names:
            sheets[name] = pd.read_excel(xls, sheet_name=name)
        else:
            sheets[name] = None
    transition_sheets = {s: pd.read_excel(xls, sheet_name=s, index_col=0) for s in xls.sheet_names if s.startswith("Transition")}
    return sheets, transition_sheets

# ----------------- Shutdown processing -----------------
def compute_shutdown_periods(plant_df, planning_dates):
    shutdown_periods = {}
    date_to_index = {d: i for i, d in enumerate(planning_dates)}
    for _, row in plant_df.iterrows():
        plant = row.get("Plant")
        start = to_date(row.get("Shutdown Start Date"))
        end = to_date(row.get("Shutdown End Date"))
        days = set()
        if start and end:
            if start > end:
                st.warning(f"⚠️ Shutdown start date after end date for {plant}. Ignoring shutdown.")
            else:
                for single in pd.date_range(start=start, end=end):
                    ddate = single.date()
                    if ddate in date_to_index:
                        days.add(date_to_index[ddate])
                if days:
                    st.info(f"🔧 Shutdown scheduled for {plant}: {fmt(start)} to {fmt(end)} ({len(days)} days)")
                else:
                    st.info(f"ℹ️ Shutdown period for {plant} is outside planning horizon")
        shutdown_periods[plant] = days
    return shutdown_periods

# ----------------- Solution callback (lightweight) -----------------
class SolutionCallback(cp_model.CpSolverSolutionCallback):
    def __init__(self, production_vars, inventory_vars, stockout_vars, is_producing, grades, lines, formatted_dates, num_days):
        super().__init__()
        self._production = production_vars
        self._inventory = inventory_vars
        self._stockout = stockout_vars
        self._is_prod = is_producing
        self._grades = grades
        self._lines = lines
        self._formatted_dates = formatted_dates
        self._num_days = num_days
        self.start_time = time.time()
        self.solution = None
        self.times = []

    def on_solution_callback(self):
        elapsed = time.time() - self.start_time
        self.times.append(elapsed)
        sol = {"time": elapsed, "production": {}, "inventory": {}, "stockout": {}, "is_producing": {}}
        # inventory -> record values
        for g in self._grades:
            sol["inventory"][g] = {}
            for d in range(self._num_days + 1):
                key = (g, d)
                if key in self._inventory:
                    sol["inventory"][g][self._formatted_dates[d] if d < self._num_days else "final"] = int(self.Value(self._inventory[key]))
        # production/stockout
        for (g, pl, d), var in self._production.items():
            val = int(self.Value(var))
            if val:
                sol["production"].setdefault(g, {}).setdefault(pl, {})[self._formatted_dates[d]] = val
        for (g, d), var in self._stockout.items():
            val = int(self.Value(var))
            if val:
                sol["stockout"].setdefault(g, {})[self._formatted_dates[d]] = val
        # is_producing to map line->date->grade
        for pl in self._lines:
            sol["is_producing"][pl] = {dt: None for dt in self._formatted_dates}
        for (g, pl, d), b in self._is_prod.items():
            if int(self.Value(b)) == 1:
                sol["is_producing"][pl][self._formatted_dates[d]] = g
        self.solution = sol

# ----------------- Streamlit UI -----------------
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #1f77b4; text-align: center; margin-bottom: 1rem; }
    .section-header { font-size: 1.25rem; color: #1f77b4; margin-top: 1.2rem; margin-bottom: 0.6rem; }
    .info-box { padding: 0.65rem; border-radius: 0.5rem; background-color: #d1ecf1; border: 1px solid #bee5eb; color: #0c5460; }
    .success-box { padding: 0.65rem; border-radius: 0.5rem; background-color: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏭 Polymer Production Scheduler (Optimized & Fixed)</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("📁 Data Input")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"], help="Upload an Excel file with Plant, Inventory, and Demand sheets")
    if uploaded_file:
        st.success("✅ File uploaded successfully!")
        st.header("⚙️ Optimization Parameters")
        time_limit_min = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=10)
        buffer_days = st.number_input("Buffer days", min_value=0, max_value=7, value=3)
        stockout_penalty = st.number_input("Stockout penalty", min_value=1, value=10)
        transition_penalty = st.number_input("Transition penalty", min_value=1, value=10)
        continuity_bonus = st.number_input("Continuity bonus", min_value=0, value=1)

if not uploaded_file:
    st.markdown("""
    <div class="info-box">
    <h3>Welcome!</h3>
    <p>Upload your file or download the sample template below. Dates are shown as <strong>DD-MMM-YY</strong>.</p>
    </div>
    """, unsafe_allow_html=True)
    sample = create_sample_workbook()
    st.download_button("📥 Download Sample Template", data=sample, file_name="polymer_production_template.xlsx")
    st.stop()

# ----------------- Main processing -----------------
try:
    uploaded_file.seek(0)
    sheets, transition_sheets = read_all_sheets(uploaded_file)
    plant_df = sheets.get("Plant")
    inventory_df = sheets.get("Inventory")
    demand_df = sheets.get("Demand")

    if plant_df is None or inventory_df is None or demand_df is None:
        st.error("Missing one of the required sheets: Plant / Inventory / Demand.")
        st.stop()

    # Preview data
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Plant Data")
        st.dataframe(plant_df, use_container_width=True)
    with col2:
        st.subheader("Inventory Data")
        st.dataframe(inventory_df, use_container_width=True)
    with col3:
        st.subheader("Demand Data (dates displayed DD-MMM-YY)")
        dp = demand_df.copy()
        date_col = dp.columns[0]
        try:
            dp[date_col] = pd.to_datetime(dp[date_col])
        except Exception:
            pass
        if pd.api.types.is_datetime64_any_dtype(dp[date_col]):
            dp[date_col] = dp[date_col].dt.strftime(DATE_DISPLAY)
        st.dataframe(dp, use_container_width=True)

    # Build structures
    lines = list(plant_df["Plant"].astype(str).tolist())
    capacities = {row["Plant"]: int(row["Capacity per day"]) for _, row in plant_df.iterrows()}

    date_col = demand_df.columns[0]
    try:
        demand_df[date_col] = pd.to_datetime(demand_df[date_col])
    except Exception:
        st.error("Unable to parse dates in Demand sheet. Ensure first column contains parseable dates.")
        st.stop()

    base_dates = sorted(demand_df[date_col].dt.date.unique().tolist())
    # Add buffer days
    last_date = base_dates[-1]
    for i in range(1, int(buffer_days) + 1):
        base_dates.append(last_date + timedelta(days=i))
    num_days = len(base_dates)
    formatted_dates = [d.strftime(DATE_DISPLAY) for d in base_dates]

    grades = [c for c in demand_df.columns if c != date_col]

    # Build demand data maps
    demand_map = {row[date_col].date(): row for _, row in demand_df.iterrows()}
    demand_data = {}
    for g in grades:
        gp = {}
        for d in base_dates[: len(demand_df)]:
            row = demand_map.get(d)
            if row is not None:
                v = row.get(g, 0)
                gp[d] = 0 if pd.isna(v) else int(v)
            else:
                gp[d] = 0
        # buffer days
        for d in base_dates[len(demand_df):]:
            gp[d] = 0
        demand_data[g] = gp

    # Inventory & grade/plant params
    initial_inventory = {}
    min_inventory = {}
    max_inventory = {}
    min_closing_inventory = {}
    min_run_days = {}
    max_run_days = {}
    force_start_date = {}
    allowed_lines = {g: [] for g in grades}
    rerun_allowed = {}
    grade_inv_seen = set()

    for _, row in inventory_df.iterrows():
        grade = row.get("Grade Name")
        if pd.isna(grade):
            continue
        grade = str(grade).strip()
        lines_value = row.get("Lines")
        if pd.notna(lines_value) and str(lines_value).strip() != "":
            plants_for_row = [x.strip() for x in str(lines_value).split(",")]
        else:
            plants_for_row = list(lines)
            st.warning(f"⚠️ Lines for grade '{grade}' are not specified; allowing all plants by default.")
        for pl in plants_for_row:
            if pl not in allowed_lines[grade]:
                allowed_lines[grade].append(pl)
        if grade not in grade_inv_seen:
            initial_inventory[grade] = int(row["Opening Inventory"]) if pd.notna(row.get("Opening Inventory")) else 0
            min_inventory[grade] = int(row["Min. Inventory"]) if pd.notna(row.get("Min. Inventory")) else 0
            max_inventory[grade] = int(row["Max. Inventory"]) if pd.notna(row.get("Max. Inventory")) else LARGE_INT
            min_closing_inventory[grade] = int(row["Min. Closing Inventory"]) if pd.notna(row.get("Min. Closing Inventory")) else 0
            grade_inv_seen.add(grade)
        for pl in plants_for_row:
            key = (grade, pl)
            min_run_days[key] = int(row["Min. Run Days"]) if pd.notna(row.get("Min. Run Days")) else 1
            max_run_days[key] = int(row["Max. Run Days"]) if pd.notna(row.get("Max. Run Days")) else num_days
            fsd = to_date(row.get("Force Start Date"))
            force_start_date[key] = fsd
            rerun_str = row.get("Rerun Allowed")
            if pd.notna(rerun_str) and isinstance(rerun_str, str) and rerun_str.strip().lower() == "no":
                rerun_allowed[key] = False
            else:
                rerun_allowed[key] = True

    # Material running
    material_running_info = {}
    for _, row in plant_df.iterrows():
        pl = row.get("Plant")
        mat = row.get("Material Running")
        exp = row.get("Expected Run Days")
        if pd.notna(mat) and pd.notna(exp):
            try:
                material_running_info[pl] = (str(mat).strip(), int(exp))
            except:
                st.warning(f"⚠️ Invalid material running data for {pl}")

    # Transitions
    transition_rules = {}
    for pl in lines:
        sheet_name = f"Transition_{pl}"
        df = None
        if sheet_name in transition_sheets:
            df = transition_sheets[sheet_name]
        else:
            # try approximate matches
            alt = [s for s in transition_sheets.keys() if s.lower().endswith(pl.lower())]
            if alt:
                df = transition_sheets[alt[0]]
        if df is not None:
            df = df.fillna("No")
            tmap = {}
            for prev in df.index.astype(str):
                allowed = [col for col in df.columns if str(df.loc[prev, col]).strip().lower() == "yes"]
                tmap[prev] = allowed
            transition_rules[pl] = tmap
            st.info(f"✅ Loaded transition matrix for {pl}")
        else:
            transition_rules[pl] = None
            st.info(f"ℹ️ No transition matrix for {pl}; allowing all transitions.")

    shutdown_periods = compute_shutdown_periods(plant_df, base_dates)

    st.markdown('<div class="section-header">🚀 Optimization</div>', unsafe_allow_html=True)

    if st.button("Run Production Optimization", type="primary"):
        prog = st.progress(0)
        status_box = st.empty()
        try:
            status_box.info("🔄 Building model...")
            prog.progress(10)

            model = cp_model.CpModel()

            # Pre-create inventory & stockout vars
            inventory_vars = {}
            stockout_vars = {}
            for g in grades:
                upper = max_inventory.get(g, LARGE_INT)
                for d in range(num_days + 1):
                    inventory_vars[(g, d)] = model.NewIntVar(0, upper, f"inv_{g}_{d}")
                for d in range(num_days):
                    # bound stockout by demand (tight bound)
                    dd = demand_data[g][base_dates[d]]
                    stockout_vars[(g, d)] = model.NewIntVar(0, dd if dd > 0 else 0, f'stock_{g}_{d}')

            # Create production & is_producing only for allowed combos
            production = {}
            is_producing = {}
            for g in grades:
                for pl in allowed_lines.get(g, []):
                    cap = capacities.get(pl, 0)
                    for d in range(num_days):
                        b = model.NewBoolVar(f"is_{g}_{pl}_{d}")
                        is_producing[(g, pl, d)] = b
                        p = model.NewIntVar(0, cap, f"prod_{g}_{pl}_{d}")
                        production[(g, pl, d)] = p
                        # link: p <= cap * b
                        model.Add(p <= cap * b)

            # Inventory initialization
            for g in grades:
                model.Add(inventory_vars[(g, 0)] == initial_inventory.get(g, 0))

            # Per-line at most one grade/day
            for pl in lines:
                for d in range(num_days):
                    vars_producing = [is_producing[(g, pl, d)] for g in grades if (g, pl, d) in is_producing]
                    if vars_producing:
                        model.Add(sum(vars_producing) <= 1)

            # Material running enforcement
            for pl, tup in material_running_info.items():
                material, exp_days = tup
                for d in range(min(num_days, exp_days)):
                    key = (material, pl, d)
                    if key in is_producing:
                        model.Add(is_producing[key] == 1)
                        # make sure others are zero
                        for g in grades:
                            if g != material:
                                k2 = (g, pl, d)
                                if k2 in is_producing:
                                    model.Add(is_producing[k2] == 0)

            # Apply shutdowns (no production)
            for pl, days in shutdown_periods.items():
                if not days:
                    continue
                for d in days:
                    for g in grades:
                        key = (g, pl, d)
                        if key in is_producing:
                            model.Add(is_producing[key] == 0)
                            model.Add(production[key] == 0)

            # Inventory flow + stockout constraints
            # To improve performance we cache lists of production vars per (grade,day)
            prod_cache = {}  # (g,d) -> list of production vars
            for g in grades:
                for d in range(num_days):
                    lst = [production[(g, pl, d)] for pl in allowed_lines.get(g, []) if (g, pl, d) in production]
                    prod_cache[(g, d)] = lst

            for g in grades:
                for d in range(num_days):
                    produced_vars = prod_cache.get((g, d), [])
                    if produced_vars:
                        produced_sum = sum(produced_vars)
                    else:
                        produced_sum = 0
                    demand_today = demand_data[g][base_dates[d]]
                    s_var = stockout_vars[(g, d)]
                    # stockout >= demand - (inventory + produced)
                    model.Add(s_var >= demand_today - (inventory_vars[(g, d)] + produced_sum))
                    # bound already handled by var domain
                    # inventory_{d+1} == inventory_d + produced - demand + stockout
                    model.Add(inventory_vars[(g, d + 1)] == inventory_vars[(g, d)] + produced_sum - demand_today + s_var)

            # Force full capacity usage for non-buffer non-shutdown days
            for pl in lines:
                cap = capacities.get(pl, 0)
                for d in range(num_days - int(buffer_days)):
                    if d in shutdown_periods.get(pl, set()):
                        continue
                    prod_vars = [production[(g, pl, d)] for g in grades if (g, pl, d) in production]
                    if prod_vars:
                        model.Add(sum(prod_vars) == cap)
                for d in range(num_days - int(buffer_days), num_days):
                    if d in shutdown_periods.get(pl, set()):
                        continue
                    prod_vars = [production[(g, pl, d)] for g in grades if (g, pl, d) in production]
                    if prod_vars:
                        model.Add(sum(prod_vars) <= cap)

            # Force start dates if present and within horizon
            for (g, pl), start_dt in force_start_date.items():
                if start_dt:
                    try:
                        idx = base_dates.index(start_dt)
                        key = (g, pl, idx)
                        if key in is_producing:
                            model.Add(is_producing[key] == 1)
                            st.info(f"✅ Enforced force start date for {g} on {pl} at {fmt(start_dt)}")
                        else:
                            st.warning(f"⚠️ Force start date {fmt(start_dt)} for {g} on {pl} cannot be enforced (combination not allowed)")
                    except ValueError:
                        st.warning(f"⚠️ Force start date {fmt(start_dt)} for {g} on {pl} not within planning horizon")

            # Min-run days: sliding-window with start variables to minimize added constraints
            start_vars = {}
            for (g, pl, d), bvar in list(is_producing.items()):
                s = model.NewBoolVar(f"start_{g}_{pl}_{d}")
                start_vars[(g, pl, d)] = s
                model.Add(s <= bvar)
                # If day>0 and prev exists define s => prev==0
                if (g, pl, d - 1) in is_producing:
                    prev = is_producing[(g, pl, d - 1)]
                    # s -> not prev
                    model.Add(prev + s <= 1)

            # Enforce windowed min-run where possible
            for (g, pl, d), s in list(start_vars.items()):
                min_run = min_run_days.get((g, pl), 1)
                if min_run <= 1:
                    continue
                available = []
                for k in range(min_run):
                    if d + k < num_days and (d + k) not in shutdown_periods.get(pl, set()):
                        if (g, pl, d + k) in is_producing:
                            available.append(is_producing[(g, pl, d + k)])
                        else:
                            available = []
                            break
                    else:
                        available = []
                        break
                if len(available) >= min_run:
                    model.Add(sum(available) >= min_run * s)

            # Transition & continuity terms (objective)
            objective_terms = []
            for pl in lines:
                for d in range(num_days - 1):
                    for g1 in grades:
                        if (g1, pl, d) not in is_producing:
                            continue
                        for g2 in grades:
                            if g2 == g1:
                                continue
                            if (g2, pl, d + 1) not in is_producing:
                                continue
                            # If transition disallowed -> forbid
                            if transition_rules.get(pl) and g1 in transition_rules[pl]:
                                allowed_next = transition_rules[pl][g1]
                                if g2 not in allowed_next:
                                    model.Add(is_producing[(g1, pl, d)] + is_producing[(g2, pl, d + 1)] <= 1)
                                    continue
                            tvar = model.NewBoolVar(f"trans_{pl}_{d}_{g1}_to_{g2}")
                            model.AddBoolAnd([is_producing[(g1, pl, d)], is_producing[(g2, pl, d + 1)]]).OnlyEnforceIf(tvar)
                            model.Add(tvar <= is_producing[(g1, pl, d)])
                            model.Add(tvar <= is_producing[(g2, pl, d + 1)])
                            objective_terms.append((transition_penalty, tvar))
                    # continuity bonus (continuing same grade)
                    for g in grades:
                        if (g, pl, d) in is_producing and (g, pl, d + 1) in is_producing:
                            cvar = model.NewBoolVar(f"cont_{pl}_{d}_{g}")
                            model.AddBoolAnd([is_producing[(g, pl, d)], is_producing[(g, pl, d + 1)]]).OnlyEnforceIf(cvar)
                            objective_terms.append((-continuity_bonus, cvar))

            # Rerun constraint (if not allowed, limit starts)
            for (g, pl) in [(g, pl) for g in grades for pl in allowed_lines.get(g, [])]:
                if not rerun_allowed.get((g, pl), True):
                    starts = [start_vars[(g, pl, d)] for d in range(num_days) if (g, pl, d) in start_vars]
                    if starts:
                        model.Add(sum(starts) <= 1)

            # Closing inventory soft constraint at last actual (non-buffer) day
            closing_day = num_days - int(buffer_days) - 1
            if closing_day < 0:
                closing_day = num_days - 1
            for g in grades:
                min_cl = min_closing_inventory.get(g, 0)
                if min_cl > 0:
                    deficit = model.NewIntVar(0, min_cl, f"closing_def_{g}")
                    model.Add(deficit >= min_cl - inventory_vars[(g, closing_day)])
                    objective_terms.append((2 * stockout_penalty, deficit))

            # Stockout penalties
            for g in grades:
                for d in range(num_days):
                    objective_terms.append((stockout_penalty, stockout_vars[(g, d)]))

            # Build objective linear expression
            linear_objs = []
            for coef, var in objective_terms:
                linear_objs.append(coef * var)
            model.Minimize(sum(linear_objs))

            prog.progress(60)
            status_box.info("⚡ Solving (OR-Tools CP-SAT)...")

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = float(time_limit_min) * 60.0
            solver.parameters.num_search_workers = DEFAULT_NUM_WORKERS

            callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, formatted_dates, num_days)
            start = time.time()
            res = solver.Solve(model, callback)
            elapsed = time.time() - start

            prog.progress(100)

            if res == cp_model.OPTIMAL:
                st.success("✅ Optimization completed optimally.")
            elif res == cp_model.FEASIBLE:
                st.success("✅ Feasible solution found (optimality not proven).")
            else:
                st.warning("⚠️ Solver finished without feasible/optimal solution.")

            if callback.solution is None:
                st.error("No solution recorded. Please check data / constraints.")
                st.stop()

            sol = callback.solution

            # Summary metrics and tables
            st.markdown('<div class="section-header">📈 Results Summary</div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                try:
                    st.metric("Objective Value", f"{solver.ObjectiveValue():,.0f}")
                except Exception:
                    st.metric("Objective Value", "NA")
            with c2:
                total_transitions = 0
                for pl in lines:
                    for d in range(num_days - 1):
                        t0 = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                        t1 = sol["is_producing"].get(pl, {}).get(formatted_dates[d + 1])
                        if t0 and t1 and t0 != t1:
                            total_transitions += 1
                st.metric("Total Transitions (approx)", total_transitions)
            with c3:
                total_stockout = sum(sum(v.values()) for v in sol.get("stockout", {}).values())
                st.metric("Total Stockouts (MT)", f"{int(total_stockout):,}")
            with c4:
                st.metric("Planning Horizon", f"{num_days} days")

            # Display shutdown info
            if any(len(s) for s in shutdown_periods.values()):
                st.subheader("🔧 Plant Shutdown Information")
                for pl, days in shutdown_periods.items():
                    if days:
                        start_idx, end_idx = min(days), max(days)
                        st.info(f"**{pl}**: {fmt(base_dates[start_idx])} to {fmt(base_dates[end_idx])} ({len(days)} days)")

            # Production totals with Total Stockout restored
            st.subheader("Total Production by Grade and Plant (MT)")
            production_totals = []
            plant_totals = {pl: 0 for pl in lines}
            stockout_totals = {g: 0 for g in grades}
            for g in grades:
                row = {"Grade": g}
                total_g = 0
                for pl in lines:
                    qty = 0
                    # sum production from solver values
                    for d in range(num_days):
                        key = (g, pl, d)
                        if key in production:
                            qty += int(solver.Value(production[key]))
                    row[pl] = qty
                    total_g += qty
                    plant_totals[pl] += qty
                # compute stockout totals from recorded solution
                s_total = sum(sol.get("stockout", {}).get(g, {}).values()) if sol.get("stockout", {}).get(g) else 0
                row["Total Produced"] = total_g
                row["Total Stockout"] = int(s_total)
                stockout_totals[g] = int(s_total)
                production_totals.append(row)
            totals_row = {"Grade": "Total"}
            for pl in lines:
                totals_row[pl] = plant_totals[pl]
            totals_row["Total Produced"] = sum(plant_totals.values())
            totals_row["Total Stockout"] = sum(stockout_totals.values())
            production_totals.append(totals_row)
            st.dataframe(pd.DataFrame(production_totals), use_container_width=True)

            # Production schedule by line (table)
            st.subheader("Production Schedule by Line")
            sorted_grades = sorted(grades)
            base_colors = px.colors.qualitative.Vivid
            grade_color_map = {g: base_colors[i % len(base_colors)] for i, g in enumerate(sorted_grades)}

            for pl in lines:
                st.markdown(f"### 🏭 {pl}")
                schedule = []
                cur = None
                start_day = None
                for d in range(num_days):
                    grade_today = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                    if grade_today != cur:
                        if cur is not None:
                            end_date = base_dates[d - 1]
                            schedule.append({
                                "Grade": cur,
                                "Start Date": fmt(start_day),
                                "End Date": fmt(end_date),
                                "Days": (end_date - start_day).days + 1
                            })
                        cur = grade_today
                        start_day = base_dates[d] if grade_today else None
                if cur:
                    end_date = base_dates[num_days - 1]
                    schedule.append({
                        "Grade": cur,
                        "Start Date": fmt(start_day),
                        "End Date": fmt(end_date),
                        "Days": (end_date - start_day).days + 1
                    })
                if not schedule:
                    st.info(f"No production data available for {pl}.")
                else:
                    df_sched = pd.DataFrame(schedule)
                    def color_grade(v):
                        if v in grade_color_map:
                            return f"background-color: {grade_color_map[v]}; color: white; font-weight:bold; text-align:center;"
                        return ""
                    st.dataframe(df_sched.style.applymap(color_grade, subset=["Grade"]), use_container_width=True)

            # Production visualization (Gantt-like)
            st.subheader("Production Visualization")
            for pl in lines:
                st.markdown(f"### Production Schedule - {pl}")
                gantt = []
                for d in range(num_days):
                    gt = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                    if gt:
                        gantt.append({"Grade": gt, "Start": base_dates[d], "Finish": base_dates[d] + timedelta(days=1), "Line": pl})
                if not gantt:
                    st.info(f"No production data available for {pl}.")
                    continue
                gdf = pd.DataFrame(gantt)
                fig = px.timeline(gdf, x_start="Start", x_end="Finish", y="Grade", color="Grade",
                                  color_discrete_map=grade_color_map, category_orders={"Grade": sorted_grades},
                                  title=f"Production Schedule - {pl}")
                # add shutdown shading
                if shutdown_periods.get(pl):
                    sdays = sorted(shutdown_periods[pl])
                    s0, s1 = base_dates[sdays[0]], base_dates[sdays[-1]] + timedelta(days=1)
                    fig.add_vrect(x0=s0, x1=s1, fillcolor="red", opacity=0.2, layer="below", line_width=0,
                                  annotation_text="Shutdown", annotation_position="top left")
                # restore gridlines and ticks
                fig.update_yaxes(autorange="reversed", title=None, showgrid=True, gridcolor="lightgray", gridwidth=1)
                fig.update_xaxes(title="Date", showgrid=True, gridcolor="lightgray", gridwidth=1, tickvals=base_dates, tickformat="%d-%b", dtick="D1")
                fig.update_layout(height=350, bargap=0.2, showlegend=True, legend_title_text="Grade",
                                  margin=dict(l=60, r=160, t=60, b=60), plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig, use_container_width=True)

            # Inventory plots: ensure actual inventory values are used and gridlines restored
            st.subheader("Inventory Levels")
            last_actual_day = num_days - int(buffer_days) - 1
            if last_actual_day < 0:
                last_actual_day = num_days - 1

            for g in sorted_grades:
                # inventory values d=0..num_days
                inv_vals = [int(solver.Value(inventory_vars[(g, d)])) for d in range(num_days + 1)]
                # Plot only first num_days (inventory at each day; we align x axis with base_dates + final point)
                xs = base_dates + [base_dates[-1] + timedelta(days=1)]  # optional final tick
                # But Plotly expects matching lengths; we will plot x as base_dates (num_days) and y as inv_vals[:num_days]
                fig = go.Figure()
                # show line for each day (inventory at start-of-day for day 0..num_days-1, and final inventory at end)
                fig.add_trace(go.Scatter(
                    x=base_dates,
                    y=[inv_vals[d] for d in range(num_days)],
                    mode="lines+markers",
                    name=g,
                    line=dict(color=grade_color_map[g] if g in grade_color_map else None, width=3),
                    marker=dict(size=6),
                    hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"
                ))
                # Add min/max lines
                fig.add_hline(y=min_inventory.get(g, 0), line=dict(color="red", width=2, dash="dash"),
                              annotation_text=f"Min: {min_inventory.get(g,0):,.0f}", annotation_position="top left",
                              annotation_font_color="red")
                max_inv_val = max_inventory.get(g, LARGE_INT)
                if max_inv_val < LARGE_INT:
                    fig.add_hline(y=max_inv_val, line=dict(color="green", width=2, dash="dash"),
                                  annotation_text=f"Max: {max_inv_val:,.0f}", annotation_position="bottom left",
                                  annotation_font_color="green")
                # Shutdown shading for plants that produce this grade
                shutdown_added = False
                for pl in allowed_lines.get(g, []):
                    if shutdown_periods.get(pl):
                        sdays = sorted(shutdown_periods[pl])
                        s0, s1 = base_dates[sdays[0]], base_dates[sdays[-1]] + timedelta(days=1)
                        fig.add_vrect(x0=s0, x1=s1, fillcolor="red", opacity=0.1, layer="below", line_width=0,
                                      annotation_text=f"Shutdown: {pl}" if not shutdown_added else "",
                                      annotation_position="top left")
                        shutdown_added = True
                # restore gridlines & layout
                fig.update_layout(title=f"Inventory Level - {g}",
                                  xaxis=dict(title="Date", tickvals=base_dates, tickformat="%d-%b", showgrid=True, gridcolor="lightgray", gridwidth=1),
                                  yaxis=dict(title="Inventory Volume (MT)", showgrid=True, gridcolor="lightgray"),
                                  plot_bgcolor="white", paper_bgcolor="white", margin=dict(l=60, r=80, t=60, b=60),
                                  height=420, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            st.error(f"Error during optimization: {e}")
            import traceback
            st.error(traceback.format_exc())

except Exception as e:
    st.error(f"Unexpected error: {e}")
    import traceback
    st.error(traceback.format_exc())
