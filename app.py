# app.py (refactored & optimized)
# Key goals: faster preprocessing, fewer solver vars, consistent DD-MMM-YY dates, unchanged visuals.
import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta, datetime, date
import matplotlib.pyplot as plt
import numpy as np
import time
import io
from matplotlib import colormaps
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile
import os

################################################################################
# ----------------------------- Configuration ---------------------------------
################################################################################
# Date display format required: DD-MMM-YY
DATE_DISPLAY = "%d-%b-%y"

# Solver worker config
DEFAULT_NUM_WORKERS = 8

# Large-but-finite bound for int variables (helps solver)
LARGE_INT = 200_000

# Streamlit page config
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

################################################################################
# --------------------------- Utility Functions -------------------------------
################################################################################
def fmt(d):
    """Format a date (datetime.date or datetime) to DD-MMM-YY string."""
    if d is None:
        return ""
    if isinstance(d, (datetime, pd.Timestamp)):
        d = d.date()
    if isinstance(d, date):
        return d.strftime(DATE_DISPLAY)
    # try parsing
    try:
        return pd.to_datetime(d).date().strftime(DATE_DISPLAY)
    except Exception:
        return str(d)

def to_date(val):
    """Convert strings, timestamps to datetime.date. Returns None if invalid."""
    if pd.isna(val) or val == "":
        return None
    if isinstance(val, date):
        return val
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None

def read_all_sheets(excel_bytes):
    """Read needed sheets once and return dict of DataFrames."""
    # Read all sheets into a dict for fast access, catching missing ones
    xls = pd.ExcelFile(excel_bytes)
    sheets = {}
    for name in ["Plant", "Inventory", "Demand"]:
        if name in xls.sheet_names:
            sheets[name] = pd.read_excel(xls, sheet_name=name)
        else:
            sheets[name] = None
    # Also read any Transition_* sheets
    transition_sheets = {s: pd.read_excel(xls, sheet_name=s, index_col=0) for s in xls.sheet_names if s.startswith("Transition")}
    return sheets, transition_sheets

def create_sample_workbook():
    """Returns BytesIO Excel sample with dates formatted as DD-MMM-YY (two-digit year)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Plant data (shutdowns using DD-MMM-YY display)
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

        # Demand: using November 2025 as template (dates are real datetimes)
        current_year, current_month = 2025, 11
        dates = pd.date_range(start=f"{current_year}-{current_month:02d}-01", periods=30, freq="D")
        demand = pd.DataFrame({
            "Date": dates,
            "BOPP": [400]*len(dates),
            "Moulding": [400]*len(dates),
            "Raffia": [600]*len(dates),
            "TQPP": [300]*len(dates),
            "Yarn": [100]*len(dates)
        })
        demand.to_excel(writer, sheet_name="Demand", index=False)

        # Simple transition matrices
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

        # Autofit & format date cells (best-effort)
        workbook = writer.book
        for sheet in workbook.worksheets:
            for col_cells in sheet.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    if cell.value is None:
                        continue
                    text = str(cell.value)
                    max_len = max(max_len, len(text))
                sheet.column_dimensions[col_letter].width = min((max_len + 2) * 1.2, 60)

    output.seek(0)
    return output

################################################################################
# -------------------------- Shutdown & Preprocessing -------------------------
################################################################################
def compute_shutdown_periods(plant_df, planning_dates):
    """
    Convert Shutdown Start/End in plant_df to dict of plant -> set(day_index)
    planning_dates: list of datetime.date
    Returns: dict {plant: set(indices)} and logs info to Streamlit
    """
    shutdown_periods = {}
    # Map date->index for fast lookup
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
                # include dates in planning horizon
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

################################################################################
# ----------------------------- Solver Callback --------------------------------
################################################################################
class SolutionCallback(cp_model.CpSolverSolutionCallback):
    """
    Light-weight solution recorder. Keeps only the last solution (to reduce memory)
    while recording times. Formats dates to DD-MMM-YY for UI display.
    """
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
        # Save a compact version of the solution (only non-zero production & stockouts)
        sol = {
            "time": elapsed,
            "production": {},
            "inventory": {},
            "stockout": {},
            "is_producing": {}
        }
        # is_producing: only record grade name per line/day
        for line in self._lines:
            sol["is_producing"][line] = {}
            for d in range(self._num_days):
                sol["is_producing"][line][self._formatted_dates[d]] = None
        for grade in self._grades:
            # inventory over days
            sol["inventory"][grade] = {}
            for d in range(self._num_days + 1):
                inv_key = (grade, d)
                if inv_key in self._inventory:
                    sol["inventory"][grade][self._formatted_dates[d] if d < self._num_days else "final"] = int(self.Value(self._inventory[inv_key]))
            # production & stockout
            for key, var in self._production.items():
                g, line, day = key
                if g != grade:
                    continue
                val = int(self.Value(var))
                if val:
                    sol["production"].setdefault(grade, {}).setdefault(line, {})[self._formatted_dates[day]] = val
            for d in range(self._num_days):
                s_key = (grade, d)
                if s_key in self._stockout:
                    sval = int(self.Value(self._stockout[s_key]))
                    if sval:
                        sol["stockout"].setdefault(grade, {})[self._formatted_dates[d]] = sval
        # fill is_producing from booleans (one pass)
        for (g, line, d), var in self._is_prod.items():
            if int(self.Value(var)) == 1:
                sol["is_producing"][line][self._formatted_dates[d]] = g
        self.solution = sol

################################################################################
# ------------------------------ Streamlit App --------------------------------
################################################################################
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #1f77b4; text-align: center; margin-bottom: 1rem; }
    .section-header { font-size: 1.25rem; color: #1f77b4; margin-top: 1.2rem; margin-bottom: 0.6rem; }
    .info-box { padding: 0.65rem; border-radius: 0.5rem; background-color: #d1ecf1; border: 1px solid #bee5eb; color: #0c5460; }
    .success-box { padding: 0.65rem; border-radius: 0.5rem; background-color: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🏭 Polymer Production Scheduler — Optimized</div>', unsafe_allow_html=True)

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

################################################################################
# If no upload -> show sample download and instructions
################################################################################
if not uploaded_file:
    st.markdown("""
    <div class="info-box">
    <h3>Welcome! Upload a file or download the sample template below.</h3>
    <ul>
      <li>Dates are displayed as <strong>DD-MMM-YY</strong> (e.g., 15-Nov-25).</li>
      <li>Sheets required: <code>Plant</code>, <code>Inventory</code>, <code>Demand</code>.</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)
    sample = create_sample_workbook()
    st.download_button("📥 Download Sample Template", data=sample, file_name="polymer_production_template.xlsx")
    with st.expander("📋 Required Excel Format"):
        st.markdown("""
        - Plant: Plant, Capacity per day, Material Running, Expected Run Days, Shutdown Start Date, Shutdown End Date
        - Inventory: Grade Name, Opening Inventory, Min. Inventory, Max. Inventory, Min. Run Days, Max. Run Days, Force Start Date, Lines, Rerun Allowed, Min. Closing Inventory
        - Demand: first col = Dates, subsequent cols = demand per Grade (col names match Grade Name)
        - Transition sheets (optional): name 'Transition_<PlantName>' with previous grade in index and next grades as columns (values 'Yes'/'No')
        """)
    st.stop()

################################################################################
# ------------------------------- Main Flow -----------------------------------
################################################################################
try:
    # Read all sheets once to memory
    uploaded_file.seek(0)
    sheets, transition_sheets = read_all_sheets(uploaded_file)

    plant_df = sheets.get("Plant")
    inventory_df = sheets.get("Inventory")
    demand_df = sheets.get("Demand")

    if plant_df is None:
        st.error("Plant sheet missing.")
        st.stop()
    if inventory_df is None:
        st.error("Inventory sheet missing.")
        st.stop()
    if demand_df is None:
        st.error("Demand sheet missing.")
        st.stop()

    # Show previews (format dates for demand display)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Plant Data")
        st.dataframe(plant_df, use_container_width=True)
    with col2:
        st.subheader("Inventory Data")
        st.dataframe(inventory_df, use_container_width=True)
    with col3:
        st.subheader("Demand Data (dates displayed DD-MMM-YY)")
        demand_preview = demand_df.copy()
        first_col = demand_preview.columns[0]
        # Attempt to coerce to datetime for display; if already ok, format
        if not pd.api.types.is_datetime64_any_dtype(demand_preview[first_col]):
            try:
                demand_preview[first_col] = pd.to_datetime(demand_preview[first_col])
            except Exception:
                pass
        if pd.api.types.is_datetime64_any_dtype(demand_preview[first_col]):
            demand_preview[first_col] = demand_preview[first_col].dt.strftime(DATE_DISPLAY)
        st.dataframe(demand_preview, use_container_width=True)

    # Build core structures: lines, capacities, grades, allowed_lines, inventory params
    lines = list(plant_df["Plant"].astype(str).tolist())
    capacities = {row["Plant"]: int(row["Capacity per day"]) for _, row in plant_df.iterrows()}

    date_col = demand_df.columns[0]
    # Make sure demand dates are parsed to datetime.date
    try:
        demand_df[date_col] = pd.to_datetime(demand_df[date_col])
    except Exception:
        # If parse fails, stop (can't build date horizon)
        st.error("Unable to parse dates in Demand sheet. Ensure first column contains parseable dates.")
        st.stop()

    # Planning horizon: unique dates in demand sorted
    base_dates = sorted(demand_df[date_col].dt.date.unique().tolist())
    # Append buffer days
    last_date = base_dates[-1]
    for i in range(1, int(buffer_days) + 1):
        base_dates.append(last_date + timedelta(days=i))
    num_days = len(base_dates)
    formatted_dates = [d.strftime(DATE_DISPLAY) for d in base_dates]

    # Grades from Demand header (exclude date column)
    grades = [c for c in demand_df.columns if c != date_col]

    # Build demand_data: dict grade -> {date: qty}
    demand_data = {}
    # Convert demand_df rows to mapping for speed
    demand_map = {row[date_col].date(): row for _, row in demand_df.iterrows()}
    for grade in grades:
        grade_map = {}
        for d in base_dates[: len(demand_df)]:  # original days
            row = demand_map.get(d)
            if row is not None:
                val = row.get(grade, 0)
                grade_map[d] = 0 if pd.isna(val) else int(val)
            else:
                grade_map[d] = 0
        # buffer days -> 0 demand
        for d in base_dates[len(demand_df):]:
            grade_map[d] = 0
        demand_data[grade] = grade_map

    # Inventory params (global per grade and per grade-plant)
    initial_inventory = {}
    min_inventory = {}
    max_inventory = {}
    min_closing_inventory = {}
    min_run_days = {}
    max_run_days = {}
    force_start_date = {}
    allowed_lines = {g: [] for g in grades}
    rerun_allowed = {}

    # Track which global inventory parameters for a grade already set
    grade_inv_seen = set()

    for _, row in inventory_df.iterrows():
        grade = row.get("Grade Name")
        if pd.isna(grade):
            continue
        grade = str(grade).strip()
        # Lines handling (if empty -> all lines)
        lines_value = row.get("Lines")
        if pd.notna(lines_value) and str(lines_value).strip() != "":
            plants_for_row = [x.strip() for x in str(lines_value).split(",")]
        else:
            plants_for_row = list(lines)
            st.warning(f"⚠️ Lines for grade '{grade}' are not specified; allowed on all plants by default.")

        for pl in plants_for_row:
            if pl not in allowed_lines[grade]:
                allowed_lines[grade].append(pl)

        # Global params
        if grade not in grade_inv_seen:
            initial_inventory[grade] = int(row["Opening Inventory"]) if pd.notna(row.get("Opening Inventory")) else 0
            min_inventory[grade] = int(row["Min. Inventory"]) if pd.notna(row.get("Min. Inventory")) else 0
            max_inventory[grade] = int(row["Max. Inventory"]) if pd.notna(row.get("Max. Inventory")) else LARGE_INT
            min_closing_inventory[grade] = int(row["Min. Closing Inventory"]) if pd.notna(row.get("Min. Closing Inventory")) else 0
            grade_inv_seen.add(grade)

        # Per grade-plant specifics
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

    # Material running info
    material_running_info = {}
    for _, row in plant_df.iterrows():
        pl = row.get("Plant")
        material = row.get("Material Running")
        exp_days = row.get("Expected Run Days")
        if pd.notna(material) and pd.notna(exp_days):
            try:
                material_running_info[pl] = (str(material).strip(), int(exp_days))
            except Exception:
                st.warning(f"⚠️ Invalid Material Running data for plant {pl}.")

    # Transition rules: build mapping for each plant
    transition_rules = {}
    for pl in lines:
        sheet_name = f"Transition_{pl}"
        # also handle underscores and spaces in names
        if sheet_name in transition_sheets:
            df = transition_sheets[sheet_name]
        else:
            # try alternatives
            alt = [s for s in transition_sheets.keys() if s.lower().endswith(pl.lower())]
            df = transition_sheets[alt[0]] if alt else None
        if df is not None:
            # ensure strings & lowercase
            trans_map = {}
            df = df.fillna("no")
            for prev in df.index.astype(str):
                allowed = [col for col in df.columns if str(df.loc[prev, col]).strip().lower() == "yes"]
                trans_map[prev] = allowed
            transition_rules[pl] = trans_map
            st.info(f"✅ Loaded transition matrix for {pl}")
        else:
            transition_rules[pl] = None
            st.info(f"ℹ️ No transition matrix for {pl}; allowing all transitions.")

    # Compute shutdowns
    shutdown_periods = compute_shutdown_periods(plant_df, base_dates)

    st.markdown('<div class="section-header">🚀 Optimization</div>', unsafe_allow_html=True)
    if st.button("Run Production Optimization", type="primary"):
        progress = st.progress(0)
        status = st.empty()

        try:
            status.info("🔄 Building and optimizing model...")
            progress.progress(10)

            model = cp_model.CpModel()

            # Only create variables for allowed (grade,line) combos
            is_producing = {}   # boolean var for (grade,line,day)
            production = {}     # int var for production quantity (grade,line,day) [0..capacity]
            inventory_vars = {} # (grade,day) inventory
            stockout_vars = {}  # (grade,day) stockout amount

            # Pre-create inventory and stockout vars per grade/day
            for g in grades:
                # inventory d=0..num_days
                for d in range(num_days + 1):
                    upper = max_inventory.get(g, LARGE_INT)
                    inventory_vars[(g, d)] = model.NewIntVar(0, upper, f"inv_{g}_{d}")
                for d in range(num_days):
                    # bound stockout by demand for that day (tight bound helps)
                    demand_today = demand_data[g][base_dates[d]]
                    stockout_vars[(g, d)] = model.NewIntVar(0, demand_today if demand_today > 0 else 0, f"stock_{g}_{d}")

            # Create production/is_producing variables only for allowed combos
            for g in grades:
                for pl in allowed_lines.get(g, []):
                    cap = capacities.get(pl, 0)
                    for d in range(num_days):
                        bvar = model.NewBoolVar(f"is_{g}_{pl}_{d}")
                        is_producing[(g, pl, d)] = bvar
                        # production quantity bounded by plant capacity
                        pvar = model.NewIntVar(0, cap, f"prod_{g}_{pl}_{d}")
                        production[(g, pl, d)] = pvar
                        # Link production to boolean: p <= cap * is_prod
                        model.Add(pvar <= cap * bvar)
                        # If day is not in buffer zone, we prefer full-capacity when producing (helps objective feasibility)
                        # but we will enforce full capacity per-line day later (sum of productions == cap)
                        # Keep per-prod linking minimal here.

            # 1) Inventory initialization
            for g in grades:
                model.Add(inventory_vars[(g, 0)] == initial_inventory.get(g, 0))

            # 2) Per-line at most one grade produced per day
            for pl in lines:
                for d in range(num_days):
                    vars_producing = [is_producing[(g, pl, d)] for g in grades if (g, pl, d) in is_producing]
                    if vars_producing:
                        model.Add(sum(vars_producing) <= 1)

            # 3) Force 'material running' days
            for pl, tup in material_running_info.items():
                material, exp_days = tup
                # Enforce first exp_days (or fewer if horizon smaller)
                for d in range(min(exp_days, num_days)):
                    key = (material, pl, d)
                    if key in is_producing:
                        model.Add(is_producing[key] == 1)
                        # ensure others are zero
                        for g in grades:
                            if g != material:
                                k2 = (g, pl, d)
                                if k2 in is_producing:
                                    model.Add(is_producing[k2] == 0)

            # 4) Shutdown restrictions: set is_producing and prod to 0 on shutdown days
            for pl, days in shutdown_periods.items():
                if not days:
                    continue
                for d in days:
                    for g in grades:
                        key = (g, pl, d)
                        if key in is_producing:
                            model.Add(is_producing[key] == 0)
                            model.Add(production[key] == 0)

            # 5) Inventory flow + stockout linear constraints (compact)
            for g in grades:
                for d in range(num_days):
                    # total produced today for grade g across allowed lines
                    produced_vars = [production[(g, pl, d)] for pl in allowed_lines.get(g, []) if (g, pl, d) in production]
                    produced_today = sum(produced_vars) if produced_vars else 0
                    demand_today = demand_data[g][base_dates[d]]
                    # stockout >= demand - (inventory + produced)
                    # inventory_next = inventory + produced_today - (demand - stockout)
                    model.Add(stockout_vars[(g, d)] >= demand_today - (inventory_vars[(g, d)] + produced_today))
                    # enforce stockout non-negative implicitly by var domain
                    # bound stockout by demand
                    model.Add(stockout_vars[(g, d)] <= demand_today)
                    # inventory update
                    # inventory_{d+1} == inventory_d + produced_today - (demand_today - stockout)
                    # rearrange: inventory_{d+1} == inventory_d + produced_today - demand_today + stockout
                    model.Add(inventory_vars[(g, d + 1)] == inventory_vars[(g, d)] + produced_today - demand_today + stockout_vars[(g, d)])
                    # upper bound on inventory enforced by var domain from creation

            # 6) Enforce full capacity usage on operational days (except buffer days & shutdown)
            for pl in lines:
                cap = capacities.get(pl, 0)
                for d in range(num_days - int(buffer_days)):
                    # skip shutdown days on this plant
                    if d in shutdown_periods.get(pl, set()):
                        continue
                    # sum production across grades == cap
                    prod_vars = [production[(g, pl, d)] for g in grades if (g, pl, d) in production]
                    if prod_vars:
                        model.Add(sum(prod_vars) == cap)
                # buffer days: <= capacity
                for d in range(num_days - int(buffer_days), num_days):
                    if d in shutdown_periods.get(pl, set()):
                        continue
                    prod_vars = [production[(g, pl, d)] for g in grades if (g, pl, d) in production]
                    if prod_vars:
                        model.Add(sum(prod_vars) <= cap)

            # 7) Force Start Date enforcement (if date in horizon)
            for (g, pl), start_dt in force_start_date.items():
                if start_dt:
                    try:
                        idx = base_dates.index(start_dt)
                        key = (g, pl, idx)
                        if key in is_producing:
                            model.Add(is_producing[key] == 1)
                            st.info(f"✅ Enforced force start date: {g} on {pl} at {fmt(start_dt)}")
                        else:
                            st.warning(f"⚠️ Force start date {fmt(start_dt)} for {g} on {pl} cannot be enforced (combo not allowed).")
                    except ValueError:
                        st.warning(f"⚠️ Force start date {fmt(start_dt)} for {g} on {pl} not present in planning horizon.")

            # 8) Min-Run enforcement using sliding windows (much fewer constraints)
            # Create start variables indicating run start at day d for (grade,pl)
            start_vars = {}
            for (g, pl, d), bvar in list(is_producing.items()):
                # only create start bool if (g,pl) exists
                s = model.NewBoolVar(f"start_{g}_{pl}_{d}")
                start_vars[(g, pl, d)] = s
                # s -> is_producing today
                model.Add(s <= bvar)
                # s -> (prev not producing OR d==0)
                if d > 0 and (g, pl, d - 1) in is_producing:
                    model.Add(s <= 1 - int(0))  # placeholder to avoid solver complaining; we'll add window constraints below
                # define s logically: s == is_producing & not prev_producing
                if d == 0:
                    # if day 0 and producing -> start
                    model.Add(bvar == 1).OnlyEnforceIf(s)
                    # ensure s implies bvar; already modeled with s <= bvar
                else:
                    prev = is_producing.get((g, pl, d - 1))
                    if prev is not None:
                        # s => is_producing today AND prev==0
                        model.AddBoolAnd([bvar, prev.Not()]).OnlyEnforceIf(s)
                        # If not s, then either not bvar or prev==1 (relaxed)
                        # No explicit exact bi-directional definition to save constraints

            # Now sliding window: For each potential start we enforce run length if possible (skip if shutdown interrupts)
            for (g, pl, d), s in list(start_vars.items()):
                min_run = min_run_days.get((g, pl), 1)
                if min_run <= 1:
                    continue
                # Check if there are min_run consecutive non-shutdown days from d
                available_days = 0
                indices = []
                for k in range(min_run):
                    if d + k < num_days and (d + k) not in shutdown_periods.get(pl, set()):
                        available_days += 1
                        indices.append(d + k)
                    else:
                        break
                if available_days >= min_run:
                    # sum(is_producing over window) >= min_run * s
                    window_vars = [is_producing[(g, pl, dd)] for dd in indices]
                    model.Add(sum(window_vars) >= min_run * s)

            # 9) Transition & continuity penalties: create only for relevant pairs
            objective_terms = []
            # stockout penalties added directly later to objective
            for pl in lines:
                for d in range(num_days - 1):
                    # For each ordered pair (g1,g2) such that both combos exist, create trans boolean
                    for g1 in grades:
                        if (g1, pl, d) not in is_producing:
                            continue
                        for g2 in grades:
                            if g2 == g1:
                                continue
                            if (g2, pl, d + 1) not in is_producing:
                                continue
                            # If transition_rules restrict, skip if not allowed
                            if transition_rules.get(pl) and g1 in transition_rules[pl]:
                                allowed_next = transition_rules[pl][g1]
                                if g2 not in allowed_next:
                                    # this combination is disallowed -> force not both
                                    model.Add(is_producing[(g1, pl, d)] + is_producing[(g2, pl, d + 1)] <= 1)
                                    continue
                            # Create transition boolean and penalize when g1 at d and g2 at d+1
                            tvar = model.NewBoolVar(f"trans_{pl}_{d}_{g1}_to_{g2}")
                            model.AddBoolAnd([is_producing[(g1, pl, d)], is_producing[(g2, pl, d + 1)]]).OnlyEnforceIf(tvar)
                            # if any producing var is false, tvar must be 0
                            model.Add(tvar <= is_producing[(g1, pl, d)])
                            model.Add(tvar <= is_producing[(g2, pl, d + 1)])
                            objective_terms.append((transition_penalty, tvar))
                # continuity (bonus) - reward continuing same grade between day d and d+1
                for d in range(num_days - 1):
                    for g in grades:
                        if (g, pl, d) in is_producing and (g, pl, d + 1) in is_producing:
                            cvar = model.NewBoolVar(f"cont_{pl}_{d}_{g}")
                            model.AddBoolAnd([is_producing[(g, pl, d)], is_producing[(g, pl, d + 1)]]).OnlyEnforceIf(cvar)
                            # negative bonus -> we will subtract from objective
                            objective_terms.append((-continuity_bonus, cvar))

            # 10) Rerun constraint simplified: avoid multiple starts if rerun not allowed
            for (g, pl) in [(g, pl) for g in grades for pl in allowed_lines.get(g, [])]:
                if not rerun_allowed.get((g, pl), True):
                    starts = [start_vars[(g, pl, d)] for d in range(num_days) if (g, pl, d) in start_vars]
                    if starts:
                        model.Add(sum(starts) <= 1)

            # 11) Closing inventory soft constraint as penalty (on the last non-buffer day)
            closing_day = num_days - int(buffer_days) - 1 if num_days - int(buffer_days) - 1 >= 0 else num_days - 1
            for g in grades:
                min_closing = min_closing_inventory.get(g, 0)
                if min_closing > 0:
                    deficit = model.NewIntVar(0, min_closing, f"closing_def_{g}")
                    # deficit >= min_closing - inventory_at_closing_day
                    model.Add(deficit >= min_closing - inventory_vars[(g, closing_day)])
                    # bound deficit
                    model.Add(deficit >= 0)
                    objective_terms.append((2 * stockout_penalty, deficit))

            # 12) Stockout penalties (sum over days)
            for g in grades:
                for d in range(num_days):
                    objective_terms.append((stockout_penalty, stockout_vars[(g, d)]))

            # Build linear objective
            linear_terms = []
            for coef, var in objective_terms:
                # coef may be negative -> allowed
                linear_terms.append(coef * var)
            model.Minimize(sum(linear_terms))

            progress.progress(60)
            status.info("⚡ Solving...")

            # Solve
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = float(time_limit_min) * 60.0
            # Allow 0 to let OR-Tools decide or set DEFAULT_NUM_WORKERS
            solver.parameters.num_search_workers = DEFAULT_NUM_WORKERS

            callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, formatted_dates, num_days)
            start_time = time.time()
            result_status = solver.Solve(model, callback)
            elapsed = time.time() - start_time

            progress.progress(100)

            # Interpret solver status
            if result_status == cp_model.OPTIMAL:
                st.success("✅ Optimization completed optimally.")
            elif result_status == cp_model.FEASIBLE:
                st.success("✅ Feasible solution found (optimality not proven).")
            else:
                st.warning("⚠️ Solver ended without producing a feasible/optimal solution.")

            # Show results if available
            if callback.solution is None:
                st.error("No solution recorded. Possibly infeasible model.")
                st.stop()

            sol = callback.solution

            # Summary metrics
            st.markdown('<div class="section-header">📈 Results Summary</div>', unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                try:
                    obj_val = solver.ObjectiveValue()
                    st.metric("Objective Value", f"{obj_val:,.0f}")
                except Exception:
                    st.metric("Objective Value", "NA")
            with col2:
                total_transitions = 0
                # count transitions from objective terms approximate by checking is_producing pattern
                for pl in lines:
                    for d in range(num_days - 1):
                        today = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                        nxt = sol["is_producing"].get(pl, {}).get(formatted_dates[d + 1])
                        if today and nxt and today != nxt:
                            total_transitions += 1
                st.metric("Total Transitions (approx)", total_transitions)
            with col3:
                total_stockout = sum(sum(v.values()) for v in sol.get("stockout", {}).values())
                st.metric("Total Stockouts (MT)", f"{int(total_stockout):,}")
            with col4:
                st.metric("Planning Horizon", f"{num_days} days")

            # Show Plant Shutdowns formatted
            if any(len(s) for s in shutdown_periods.values()):
                st.subheader("🔧 Plant Shutdowns")
                for pl, days in shutdown_periods.items():
                    if not days:
                        continue
                    start_idx, end_idx = min(days), max(days)
                    st.info(f"**{pl}**: {fmt(base_dates[start_idx])} to {fmt(base_dates[end_idx])} ({len(days)} days)")

            # Production totals table
            st.subheader("Total Production by Grade and Plant (MT)")
            prod_totals = []
            plant_totals = {pl: 0 for pl in lines}
            for g in grades:
                row = {"Grade": g}
                total_g = 0
                for pl in lines:
                    qty = 0
                    for d in range(num_days):
                        key = (g, pl, d)
                        if key in production:
                            qty += int(solver.Value(production[key]))
                    row[pl] = qty
                    total_g += qty
                    plant_totals[pl] += qty
                row["Total Produced"] = total_g
                prod_totals.append(row)
            totals = {"Grade": "Total"}
            for pl in lines:
                totals[pl] = plant_totals[pl]
            totals["Total Produced"] = sum(plant_totals.values())
            prod_totals.append(totals)
            st.dataframe(pd.DataFrame(prod_totals), use_container_width=True)

            # Schedule tables per line (sequence)
            st.subheader("Production Schedule by Line")
            sorted_grades = sorted(grades)
            base_colors = px.colors.qualitative.Vivid
            grade_color_map = {g: base_colors[i % len(base_colors)] for i, g in enumerate(sorted_grades)}

            for pl in lines:
                st.markdown(f"### 🏭 {pl}")
                schedule = []
                cur_grade = None
                start_day = None
                for d in range(num_days):
                    # find grade produced this day (if any)
                    grade_today = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                    if grade_today != cur_grade:
                        if cur_grade is not None:
                            end_day = base_dates[d - 1]
                            schedule.append({
                                "Grade": cur_grade,
                                "Start Date": fmt(start_day),
                                "End Date": fmt(end_day),
                                "Days": (end_day - start_day).days + 1
                            })
                        cur_grade = grade_today
                        start_day = base_dates[d] if grade_today else None
                if cur_grade:
                    end_day = base_dates[num_days - 1]
                    schedule.append({
                        "Grade": cur_grade,
                        "Start Date": fmt(start_day),
                        "End Date": fmt(end_day),
                        "Days": (end_day - start_day).days + 1
                    })
                if not schedule:
                    st.info(f"No production scheduled for {pl}.")
                else:
                    df_sched = pd.DataFrame(schedule)
                    # color grade column
                    def color_grade(v):
                        if v in grade_color_map:
                            return f"background-color: {grade_color_map[v]}; color: white; font-weight:bold"
                        return ""
                    st.dataframe(df_sched.style.applymap(color_grade, subset=["Grade"]), use_container_width=True)

            # Gantt-like visualization (per-line)
            st.subheader("Production Visualization")
            for pl in lines:
                st.markdown(f"### Production Schedule - {pl}")
                gantt = []
                for d in range(num_days):
                    grade_today = sol["is_producing"].get(pl, {}).get(formatted_dates[d])
                    if grade_today:
                        gantt.append({
                            "Grade": grade_today,
                            "Start": base_dates[d],
                            "Finish": base_dates[d] + timedelta(days=1),
                            "Line": pl
                        })
                if not gantt:
                    st.info(f"No production for {pl}.")
                    continue
                gdf = pd.DataFrame(gantt)
                fig = px.timeline(gdf, x_start="Start", x_end="Finish", y="Grade", color="Grade",
                                  color_discrete_map=grade_color_map, category_orders={"Grade": sorted_grades},
                                  title=f"Production Schedule - {pl}")
                # shutdown shading
                if shutdown_periods.get(pl):
                    sd = sorted(shutdown_periods[pl])
                    start_sd, end_sd = base_dates[sd[0]], base_dates[sd[-1]] + timedelta(days=1)
                    fig.add_vrect(x0=start_sd, x1=end_sd, fillcolor="red", opacity=0.2, layer="below",
                                  line_width=0, annotation_text="Shutdown", annotation_position="top left")
                fig.update_xaxes(tickvals=base_dates, tickformat="%d-%b")
                fig.update_layout(height=350)
                st.plotly_chart(fig, use_container_width=True)

            # Inventory Plots per grade
            st.subheader("Inventory Levels")
            last_day_actual = num_days - int(buffer_days) - 1 if num_days - int(buffer_days) - 1 >= 0 else num_days - 1
            for g in sorted_grades:
                inv_vals = [int(solver.Value(inventory_vars[(g, d)])) for d in range(num_days + 1)]
                xs = base_dates
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=xs, y=inv_vals[:-0], mode="lines+markers", name=g,
                                         line=dict(color=grade_color_map[g] if g in grade_color_map else None, width=3),
                                         hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"))
                # min/max lines
                fig.add_hline(y=min_inventory.get(g, 0), line=dict(color="red", dash="dash"), annotation_text=f"Min: {min_inventory.get(g,0)}")
                fig.add_hline(y=max_inventory.get(g, LARGE_INT), line=dict(color="green", dash="dash"), annotation_text=f"Max: {max_inventory.get(g, LARGE_INT)}")
                # shutdown zones for plants producing this grade
                added = False
                for pl in allowed_lines.get(g, []):
                    if shutdown_periods.get(pl):
                        sd = sorted(shutdown_periods[pl])
                        start_sd, end_sd = base_dates[sd[0]], base_dates[sd[-1]] + timedelta(days=1)
                        fig.add_vrect(x0=start_sd, x1=end_sd, fillcolor="red", opacity=0.1, layer="below", line_width=0,
                                      annotation_text=f"Shutdown: {pl}" if not added else "", annotation_position="top left")
                        added = True
                fig.update_layout(title=f"Inventory Level - {g}", xaxis=dict(tickvals=base_dates, tickformat="%d-%b"), height=420, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            st.error(f"Error running optimization: {e}")
            import traceback
            st.error(traceback.format_exc())

except Exception as e:
    st.error(f"General error: {e}")
    import traceback
    st.error(traceback.format_exc())
