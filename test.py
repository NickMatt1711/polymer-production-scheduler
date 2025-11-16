import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta
import time
import io
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import traceback

# ============================================================================
# PART 1: OPTIMIZATION ENGINE CLASS
# (Refactored logic from test(1).py)
# ============================================================================

class PolymerProductionModel:
    """
    Encapsulates the entire production scheduling optimization model.
    
    This class handles data preprocessing, model building, solving, and
    solution extraction, separating the optimization logic from the UI.
    """
    
    def __init__(self, plant_df, inventory_df, demand_df, params):
        self.plant_df = plant_df
        self.inventory_df = inventory_df
        self.demand_df = demand_df
        self.params = params
        
        # Solver components
        self.model = cp_model.CpModel()
        self.solver = cp_model.CpSolver()
        self.status = None
        self.callback = None
        
        # Data containers
        self.lines = []
        self.grades = []
        self.dates = []
        self.num_days = 0
        self.capacities = {}
        self.initial_inventory = {}
        self.min_inventory = {}
        self.max_inventory = {}
        self.min_closing_inventory = {}
        self.min_run_days = {}
        self.max_run_days = {}
        self.force_start_date = {}
        self.allowed_lines = {}
        self.rerun_allowed = {}
        self.material_running_info = {}
        self.demand_data = {}
        self.shutdown_periods = {}
        self.transition_rules = {}
        
        # Model variables
        self.is_producing = {}
        self.production = {}
        self.inventory_vars = {}
        self.stockout_vars = {}
        self.is_start_vars = {}

    def _preprocess_data(self):
        """
        Processes the raw DataFrames into structured data for the model.
        """
        # Extract plant data
        self.lines = list(self.plant_df['Plant'])
        self.capacities = {row['Plant']: row['Capacity per day'] for _, row in self.plant_df.iterrows()}
        
        # Extract grades from demand
        self.grades = [col for col in self.demand_df.columns if col != self.demand_df.columns[0]]
        self.allowed_lines = {grade: [] for grade in self.grades}
        
        grade_inventory_defined = set()
        
        # Process inventory sheet
        for _, row in self.inventory_df.iterrows():
            grade = row['Grade Name']
            if grade not in self.grades:
                continue # Skip grades not in demand sheet

            lines_value = row['Lines']
            plants_for_row = [x.strip() for x in str(lines_value).split(',')] if pd.notna(lines_value) and lines_value != '' else self.lines
            
            for plant in plants_for_row:
                if plant in self.lines and plant not in self.allowed_lines[grade]:
                    self.allowed_lines[grade].append(plant)
            
            if grade not in grade_inventory_defined:
                self.initial_inventory[grade] = row['Opening Inventory'] if pd.notna(row['Opening Inventory']) else 0
                self.min_inventory[grade] = row['Min. Inventory'] if pd.notna(row['Min. Inventory']) else 0
                self.max_inventory[grade] = row['Max. Inventory'] if pd.notna(row['Max. Inventory']) else 1_000_000_000
                self.min_closing_inventory[grade] = row['Min. Closing Inventory'] if pd.notna(row['Min. Closing Inventory']) else 0
                grade_inventory_defined.add(grade)
            
            for plant in plants_for_row:
                if plant not in self.lines: continue
                grade_plant_key = (grade, plant)
                self.min_run_days[grade_plant_key] = int(row['Min. Run Days']) if pd.notna(row['Min. Run Days']) else 1
                self.max_run_days[grade_plant_key] = int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999
                self.force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date() if pd.notna(row['Force Start Date']) else None
                rerun_val = str(row['Rerun Allowed']).strip().lower()
                self.rerun_allowed[grade_plant_key] = rerun_val not in ['no', 'n', 'false', '0'] if pd.notna(rerun_val) else True

        # Material running info
        for _, row in self.plant_df.iterrows():
            plant = row['Plant']
            material = str(row['Material Running']).strip()
            expected_days = row['Expected Run Days']
            if pd.notna(material) and pd.notna(expected_days) and material in self.grades:
                try:
                    self.material_running_info[plant] = (material, int(expected_days))
                except (ValueError, TypeError):
                    pass
        
        # Process demand data
        self.dates = sorted(list(set(self.demand_df.iloc[:, 0].dt.date.tolist())))
        last_date = self.dates[-1]
        buffer_days = self.params['buffer_days']
        
        for i in range(1, buffer_days + 1):
            self.dates.append(last_date + timedelta(days=i))
        self.num_days = len(self.dates)
        
        for grade in self.grades:
            self.demand_data[grade] = {self.demand_df.iloc[i, 0].date(): self.demand_df[grade].iloc[i] for i in range(len(self.demand_df)) if grade in self.demand_df.columns}
            for date in self.dates:
                if date not in self.demand_data[grade]:
                    self.demand_data[grade][date] = 0

        # Process shutdown periods
        self.shutdown_periods = process_shutdown_dates(self.plant_df, self.dates)
        
        # (Transition rules processing is omitted as it requires re-reading the Excel file)
        # (In a real app, you'd pass the excel_file bytes to the class)
        # For this demo, we assume no transition rules
        self.transition_rules = {}

    def _build_model(self):
        """
        Creates all CP-SAT variables, constraints, and the objective function.
        """
        buffer_days = self.params['buffer_days']

        # --- Variable Creation ---
        for grade in self.grades:
            for line in self.lines:
                for d in range(self.num_days):
                    key = (grade, line, d)
                    self.is_producing[key] = self.model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                    
                    if d < self.num_days - buffer_days:
                        prod_val = self.model.NewIntVar(0, self.capacities[line], f'production_{grade}_{line}_{d}')
                        self.model.Add(prod_val == self.capacities[line]).OnlyEnforceIf(self.is_producing[key])
                        self.model.Add(prod_val == 0).OnlyEnforceIf(self.is_producing[key].Not())
                    else:
                        prod_val = self.model.NewIntVar(0, self.capacities[line], f'production_{grade}_{line}_{d}')
                        self.model.Add(prod_val <= self.capacities[line] * self.is_producing[key])
                    
                    self.production[key] = prod_val
                    
                    if line not in self.allowed_lines[grade]:
                        self.model.Add(self.is_producing[key] == 0)

        for grade in self.grades:
            for d in range(self.num_days + 1):
                self.inventory_vars[(grade, d)] = self.model.NewIntVar(0, 1000000, f'inventory_{grade}_{d}')
        
        for grade in self.grades:
            for d in range(self.num_days):
                self.stockout_vars[(grade, d)] = self.model.NewIntVar(0, 1000000, f'stockout_{grade}_{d}')

        # --- Constraints ---
        
        # CATEGORY 1: PRODUCTION CAPACITY
        for line in self.lines:
            for d in range(self.num_days):
                prod_vars = [self.production[(g, line, d)] for g in self.grades]
                if line in self.shutdown_periods and d in self.shutdown_periods[line]:
                    self.model.Add(sum(prod_vars) == 0)
                elif d < self.num_days - buffer_days:
                    self.model.Add(sum(prod_vars) == self.capacities[line])
                else:
                    self.model.Add(sum(prod_vars) <= self.capacities[line])

        # CATEGORY 2: ONE GRADE PER LINE
        for line in self.lines:
            for d in range(self.num_days):
                self.model.Add(sum(self.is_producing[(g, line, d)] for g in self.grades) <= 1)

        # CATEGORY 3: MATERIAL RUNNING
        for plant, (material, expected_days) in self.material_running_info.items():
            for d in range(min(expected_days, self.num_days)):
                self.model.Add(self.is_producing[(material, plant, d)] == 1)
                for other_material in self.grades:
                    if other_material != material:
                        self.model.Add(self.is_producing[(other_material, plant, d)] == 0)

        # CATEGORY 4: INVENTORY BALANCE
        for grade in self.grades:
            self.model.Add(self.inventory_vars[(grade, 0)] == self.initial_inventory[grade])
            for d in range(self.num_days):
                produced = sum(self.production[(grade, line, d)] for line in self.lines)
                demand = self.demand_data[grade].get(self.dates[d], 0)
                
                self.model.Add(self.inventory_vars[(grade, d + 1)] == 
                               self.inventory_vars[(grade, d)] + produced - demand + self.stockout_vars[(grade, d)])
                self.model.Add(self.stockout_vars[(grade, d)] >= demand - self.inventory_vars[(grade, d)] - produced)
                self.model.Add(self.stockout_vars[(grade, d)] >= 0)
                self.model.Add(self.inventory_vars[(grade, d + 1)] >= 0)

        # CATEGORY 5: INVENTORY LIMITS
        for grade in self.grades:
            for d in range(1, self.num_days + 1):
                self.model.Add(self.inventory_vars[(grade, d)] <= self.max_inventory[grade])

        # CATEGORY 6: FORCE START DATE
        for grade_plant_key, start_date in self.force_start_date.items():
            if start_date:
                grade, plant = grade_plant_key
                try:
                    start_day_index = self.dates.index(start_date)
                    self.model.Add(self.is_producing[(grade, plant, start_day_index)] == 1)
                except ValueError:
                    pass

        # CATEGORY 7: RUN LENGTH
        for grade in self.grades:
            for line in self.allowed_lines[grade]:
                grade_plant_key = (grade, line)
                min_run = self.min_run_days.get(grade_plant_key, 1)
                max_run = self.max_run_days.get(grade_plant_key, 9999)
                
                for d in range(self.num_days):
                    is_start = self.model.NewBoolVar(f'start_{grade}_{line}_{d}')
                    self.is_start_vars[(grade, line, d)] = is_start
                    current_prod = self.is_producing[(grade, line, d)]
                    
                    if d > 0:
                        prev_prod = self.is_producing[(grade, line, d - 1)]
                        self.model.AddBoolAnd([current_prod, prev_prod.Not()]).OnlyEnforceIf(is_start)
                        self.model.AddBoolOr([current_prod.Not(), prev_prod]).OnlyEnforceIf(is_start.Not())
                    else:
                        self.model.Add(is_start == 1).OnlyEnforceIf(current_prod)
                
                # Min run
                for d in range(self.num_days):
                    is_start = self.is_start_vars[(grade, line, d)]
                    for k in range(min_run):
                        if d + k < self.num_days:
                            if not (line in self.shutdown_periods and (d + k) in self.shutdown_periods[line]):
                                self.model.Add(self.is_producing[(grade, line, d + k)] == 1).OnlyEnforceIf(is_start)
                
                # Max run
                for d in range(self.num_days - max_run):
                    consecutive_days = []
                    for k in range(max_run + 1):
                        if d + k < self.num_days:
                            if line in self.shutdown_periods and (d + k) in self.shutdown_periods[line]:
                                break # Stop if shutdown breaks run
                            consecutive_days.append(self.is_producing[(grade, line, d + k)])
                    if len(consecutive_days) == max_run + 1:
                        self.model.Add(sum(consecutive_days) <= max_run)

        # CATEGORY 8: TRANSITION RULES (Skipped - see _preprocess_data)

        # CATEGORY 9: RERUN ALLOWED
        for grade_plant_key, allowed in self.rerun_allowed.items():
            if not allowed:
                grade, line = grade_plant_key
                starts = [self.is_start_vars[(grade, line, d)] for d in range(self.num_days) 
                          if (grade, line, d) in self.is_start_vars]
                if starts:
                    self.model.Add(sum(starts) <= 1)

        # --- Objective Function ---
        objective_terms = []
        stockout_penalty = self.params['stockout_penalty']
        transition_penalty = self.params['transition_penalty']
        
        # TIER 1: CRITICAL - Stockouts
        for grade in self.grades:
            for d in range(self.num_days):
                objective_terms.append(stockout_penalty * self.stockout_vars[(grade, d)])
        
        # TIER 2: IMPORTANT - Minimum inventory violations
        for grade in self.grades:
            for d in range(self.num_days):
                if self.min_inventory[grade] > 0:
                    deficit = self.model.NewIntVar(0, 1000000, f'deficit_{grade}_{d}')
                    self.model.Add(deficit >= self.min_inventory[grade] - self.inventory_vars[(grade, d + 1)])
                    objective_terms.append(stockout_penalty * deficit) # Penalize at same high rate
        
        # TIER 3: IMPORTANT - Closing inventory targets
        for grade in self.grades:
            if self.min_closing_inventory[grade] > 0:
                closing_inv = self.inventory_vars[(grade, self.num_days - buffer_days)]
                closing_deficit = self.model.NewIntVar(0, 1000000, f'closing_deficit_{grade}')
                self.model.Add(closing_deficit >= self.min_closing_inventory[grade] - closing_inv)
                objective_terms.append(stockout_penalty * closing_deficit * 3) # Penalize even higher
        
        # TIER 4: OPERATIONAL - Transitions
        for line in self.lines:
            for d in range(self.num_days - 1):
                any_transition = self.model.NewBoolVar(f'transition_{line}_{d}')
                continuity_indicators = []
                for grade in self.grades:
                    same_grade = self.model.NewBoolVar(f'same_{grade}_{line}_{d}')
                    self.model.AddBoolAnd([self.is_producing[(grade, line, d)], 
                                           self.is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(same_grade)
                    continuity_indicators.append(same_grade)
                
                has_continuity = self.model.NewBoolVar(f'has_continuity_{line}_{d}')
                self.model.AddMaxEquality(has_continuity, continuity_indicators)
                
                prod_day_d = self.model.NewBoolVar(f'prod_{line}_{d}')
                prod_day_d_plus_1 = self.model.NewBoolVar(f'prod_{line}_{d+1}')
                self.model.AddMaxEquality(prod_day_d, [self.is_producing[(g, line, d)] for g in self.grades])
                self.model.AddMaxEquality(prod_day_d_plus_1, [self.is_producing[(g, line, d+1)] for g in self.grades])
                
                # Transition = (Prod Day T) AND (Prod Day T+1) AND (NOT Continuous)
                self.model.AddBoolAnd([prod_day_d, prod_day_d_plus_1, has_continuity.Not()]).OnlyEnforceIf(any_transition)
                self.model.AddBoolOr([prod_day_d.Not(), prod_day_d_plus_1.Not(), has_continuity]).OnlyEnforceIf(any_transition.Not())
                
                objective_terms.append(transition_penalty * any_transition)
        
        # TIER 5: EFFICIENCY - Inventory holding costs
        for grade in self.grades:
            for d in range(self.num_days):
                objective_terms.append(1 * self.inventory_vars[(grade, d)]) # Low cost
        
        self.model.Minimize(sum(objective_terms))


    def solve(self, time_limit_min):
        """
        Runs the full pipeline: preprocess, build, solve, and extract solution.
        """
        # --- 1. Preprocess Data ---
        try:
            self._preprocess_data()
        except Exception as e:
            print(f"Error during preprocessing: {e}")
            raise

        # --- 2. Build Model ---
        try:
            self._build_model()
        except Exception as e:
            print(f"Error during model building: {e}")
            raise

        # --- 3. Configure Solver ---
        self.solver.parameters.max_time_in_seconds = time_limit_min * 60.0
        self.solver.parameters.num_search_workers = 8
        self.solver.parameters.random_seed = 42
        self.solver.parameters.linearization_level = 2
        self.solver.parameters.cp_model_probing_level = 2
        self.solver.parameters.symmetry_level = 4
        self.solver.parameters.optimize_with_core = True
        self.solver.parameters.search_branching = cp_model.PORTFOLIO_SEARCH
        self.solver.parameters.log_search_progress = True
        
        class MinimalCallback(cp_model.CpSolverSolutionCallback):
            def __init__(self):
                super().__init__()
                self.objectives = []
                self.times = []
                self.start_time = time.time()
            
            def on_solution_callback(self):
                self.objectives.append(self.ObjectiveValue())
                self.times.append(time.time() - self.start_time)
        
        self.callback = MinimalCallback()

        # --- 4. Solve Model ---
        start_time = time.time()
        self.status = self.solver.Solve(self.model, self.callback)
        solve_time = time.time() - start_time

        # --- 5. Extract and Return Solution ---
        if self.status == cp_model.OPTIMAL or self.status == cp_model.FEASIBLE:
            solution_data = self._extract_solution(solve_time)
            return solution_data
        else:
            return None # No solution found

    def _extract_solution(self, solve_time):
        """
        Extracts data from the solver into a clean dictionary.
        """
        solution_data = {
            "status": self.solver.StatusName(self.status),
            "objective_value": self.solver.ObjectiveValue(),
            "solve_time": solve_time,
            "solution_production": {},
            "solution_inventory": {},
            "solution_stockout": {},
            "solution_schedule": {},
            "formatted_dates": [date.strftime('%d-%b-%y') for date in self.dates],
            "dates": self.dates,
            "grades": self.grades,
            "lines": self.lines,
            "num_days": self.num_days,
            "shutdown_periods": self.shutdown_periods,
            "min_inventory": self.min_inventory,
            "max_inventory": self.max_inventory,
            "allowed_lines": self.allowed_lines
        }

        # Use solver.Value() to extract values
        for grade in self.grades:
            solution_data["solution_production"][grade] = {}
            for d in range(self.num_days):
                date_key = solution_data["formatted_dates"][d]
                total_prod = sum(self.solver.Value(self.production[(grade, line, d)]) for line in self.lines)
                if total_prod > 0:
                    solution_data["solution_production"][grade][date_key] = total_prod
        
        for grade in self.grades:
            solution_data["solution_inventory"][grade] = {}
            for d in range(self.num_days + 1):
                val = self.solver.Value(self.inventory_vars[(grade, d)])
                if d < self.num_days:
                    solution_data["solution_inventory"][grade][solution_data["formatted_dates"][d]] = val
                else:
                    solution_data["solution_inventory"][grade]['final'] = val

        for grade in self.grades:
            solution_data["solution_stockout"][grade] = {}
            for d in range(self.num_days):
                val = self.solver.Value(self.stockout_vars[(grade, d)])
                if val > 0:
                    solution_data["solution_stockout"][grade][solution_data["formatted_dates"][d]] = val
        
        for line in self.lines:
            solution_data["solution_schedule"][line] = {}
            for d in range(self.num_days):
                solution_data["solution_schedule"][line][solution_data["formatted_dates"][d]] = None
                for grade in self.grades:
                    if self.solver.Value(self.is_producing[(grade, line, d)]) == 1:
                        solution_data["solution_schedule"][line][solution_data["formatted_dates"][d]] = grade
                        break

        # Calculate transitions
        total_transitions = 0
        transition_count_per_line = {line: 0 for line in self.lines}
        for line in self.lines:
            last_grade = None
            for d in range(self.num_days):
                current_grade = solution_data["solution_schedule"][line][solution_data["formatted_dates"][d]]
                if current_grade is not None:
                    if last_grade is not None and current_grade != last_grade:
                        transition_count_per_line[line] += 1
                        total_transitions += 1
                    last_grade = current_grade
        
        solution_data["total_transitions"] = total_transitions
        solution_data["transition_count_per_line"] = transition_count_per_line
        solution_data["total_stockouts"] = sum(sum(v.values()) for v in solution_data["solution_stockout"].values())

        return solution_data


# ============================================================================
# HELPER FUNCTIONS (for UI)
# ============================================================================

def get_sample_workbook():
    """Retrieve the sample workbook from the same directory"""
    try:
        current_dir = Path(__file__).parent
        sample_path = current_dir / "polymer_production_template.xlsx"
        
        if sample_path.exists():
            with open(sample_path, "rb") as f:
                return io.BytesIO(f.read())
        else:
            st.warning("Sample template file not found.")
            return None
    except Exception as e:
        st.warning(f"Could not load sample template: {e}")
        return None

def process_shutdown_dates(plant_df, dates):
    """Process shutdown dates for each plant (used in both UI and Model)"""
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
                
                shutdown_days = [d for d, date in enumerate(dates) if start_date <= date <= end_date]
                
                if shutdown_days:
                    shutdown_periods[plant] = shutdown_days
                    st.info(f"🔧 Shutdown scheduled for {plant}: {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')} ({len(shutdown_days)} days)")
                else:
                    shutdown_periods[plant] = []
                    
            except Exception as e:
                st.warning(f"⚠️ Invalid shutdown dates for {plant}: {e}")
                shutdown_periods[plant] = []
        else:
            shutdown_periods[plant] = []
    
    return shutdown_periods

# ============================================================================
# PART 2: STREAMLIT UI
# (This remains largely the same, but now calls the class)
# ============================================================================

# --- Page Config & Session State ---
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

if 'step' not in st.session_state:
    st.session_state.step = 1
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'optimization_params' not in st.session_state:
    st.session_state.optimization_params = {
        'buffer_days': 3,
        'time_limit_min': 10,
        'stockout_penalty': 1000,
        'transition_penalty': 100,
    }

# --- CSS (Unchanged) ---
st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none; }
    .main .block-container { padding-top: 3rem; padding-bottom: 3rem; max-width: 1200px; }
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }
    .app-bar { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-align: center; padding: 2rem 3rem; margin: -3rem -3rem 3rem -3rem; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); border-radius: 16px; }
    .app-bar h1 { margin: 0; font-size: 2rem; font-weight: 600; }
    .app-bar p { margin: 0.5rem 0 0 0; font-size: 1rem; opacity: 0.95; }
    .step-indicator { display: flex; justify-content: center; align-items: center; margin: 2rem 0 3rem 0; position: relative; }
    .step { display: flex; flex-direction: column; align-items: center; position: relative; flex: 1; max-width: 200px; }
    .step-circle { width: 48px; height: 48px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 1.125rem; transition: all 0.3s; z-index: 2; background: white; border: 3px solid #e0e0e0; color: #9e9e9e; }
    .step-circle.active { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-color: #667eea; color: white; box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4); transform: scale(1.1); }
    .step-circle.completed { background: #4caf50; border-color: #4caf50; color: white; }
    .step-label { margin-top: 0.75rem; font-size: 0.875rem; font-weight: 500; color: #757575; }
    .step-label.active { color: #667eea; font-weight: 600; }
    .step-label.completed { color: #4caf50; }
    .step-line { position: absolute; top: 24px; left: 50%; right: -50%; height: 3px; background: #e0e0e0; z-index: 1; }
    .step-line.completed { background: #4caf50; }
    .material-card { background: #F0F2FF; border-radius: 16px; margin-bottom: 1.5rem; padding: 2rem; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08); border: 1px solid rgba(0, 0, 0, 0.06); }
    .card-title { font-size: 1.25rem; font-weight: 600; text-align: center; color: #212121; margin: 0 0 1rem 0; }
    .stButton > button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; padding: 0.75rem 2rem; font-weight: 600; font-size: 1rem; box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3); transition: all 0.3s; text-transform: uppercase; }
    .stButton > button:hover { box-shadow: 0 4px 16px rgba(102, 126, 234, 0.4); transform: translateY(-2px); }
    .metric-card { background: white; border-radius: 12px; padding: 1.5rem; text-align: center; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08); border-left: 4px solid; height: 100%; }
    .metric-card.primary { border-left-color: #667eea; }
    .metric-card.success { border-left-color: #4caf50; }
    .metric-card.warning { border-left-color: #ff9800; }
    .metric-card.info { border-left-color: #2196f3; }
    .metric-label { font-size: 0.75rem; text-transform: uppercase; letter-spacing: 1px; color: #757575; font-weight: 600; margin-bottom: 0.5rem; }
    .metric-value { font-size: 2rem; font-weight: 700; color: #212121; }
    .metric-subtitle { font-size: 0.75rem; color: #9e9e9e; margin-top: 0.25rem; }
    .chip { display: inline-flex; align-items: center; padding: 0.375rem 0.875rem; border-radius: 16px; font-size: 0.8125rem; font-weight: 500; margin: 0.25rem; }
    .chip.success { background: #e8f5e9; color: #2e7d32; }
    .chip.warning { background: #fff3e0; color: #e65100; }
    .chip.info { background: #e3f2fd; color: #1565c0; }
    .stTabs [data-baseweb="tab-list"] { gap: 0.5rem; background: white; padding: 0.5rem; border-radius: 12px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08); }
    .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 0.75rem 1.5rem; font-weight: 600; background: transparent; border: none; color: #757575; flex: 1; }
    .stTabs [aria-selected="true"] { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3); }
    .stProgress > div > div > div > div { background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); }
    .alert-box { padding: 1rem 1.5rem; border-radius: 8px; margin: 1rem 0; border-left: 4px solid; }
    .alert-box.info { background: #e3f2fd; border-left-color: #2196f3; color: #1565c0; }
    .alert-box.success { background: #e8f5e9; border-left-color: #4caf50; color: #2e7d32; }
    .alert-box.warning { background: #fff3e0; border-left-color: #ff9800; color: #e65100; }
    .divider { height: 1px; background: linear-gradient(90deg, transparent, #e0e0e0, transparent); margin: 2rem 0; }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown("""
<div class="app-bar">
    <h1>🏭 Polymer Production Scheduler</h1>
    <p>Optimized Multi-Plant Production Planning</p>
</div>
""", unsafe_allow_html=True)

# --- Step Indicator ---
step_status = ['active' if st.session_state.step == 1 else 'completed',
               'active' if st.session_state.step == 2 else ('completed' if st.session_state.step > 2 else ''),
               'active' if st.session_state.step == 3 else '']
st.markdown(f"""
<div class="step-indicator">
    <div class="step"><div class="step-circle {step_status[0]}">{'✓' if st.session_state.step > 1 else '1'}</div><div class="step-label {step_status[0]}">Upload Data</div><div class="step-line {step_status[0] if st.session_state.step > 1 else ''}"></div></div>
    <div class="step"><div class="step-circle {step_status[1]}">{'✓' if st.session_state.step > 2 else '2'}</div><div class="step-label {step_status[1]}">Configure</div><div class="step-line {step_status[1] if st.session_state.step > 2 else ''}"></div></div>
    <div class="step"><div class="step-circle {step_status[2]}">3</div><div class="step-label {step_status[2]}">View Results</div></div>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# STEP 1: UPLOAD DATA (Unchanged)
# ============================================================================
if st.session_state.step == 1:
    col1, col2 = st.columns([4, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"], label_visibility="collapsed")
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("✅ File uploaded successfully!")
            time.sleep(0.5)
            st.session_state.step = 2
            st.rerun()
    with col2:
        sample_workbook = get_sample_workbook()
        if sample_workbook:
            st.download_button("📥 Download Template", sample_workbook, "polymer_production_template.xlsx", use_container_width=True)
    
    st.markdown("""
    <div class="material-card">
        <div class="card-title">📋 Quick Start Guide</div>
        <ol>
            <li>Download the Excel template</li>
            <li>Fill in your production data (Plant, Inventory, Demand sheets)</li>
            <li>Upload your completed file</li>
            <li>Configure optimization parameters</li>
            <li>Run optimization and analyze results</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

# ============================================================================
# STEP 2: PREVIEW & CONFIGURE (Unchanged)
# ============================================================================
elif st.session_state.step == 2:
    try:
        uploaded_file = st.session_state.uploaded_file
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        tab1, tab2, tab3 = st.tabs(["🏭 Plant Data", "📦 Inventory Data", "📊 Demand Data"])
        with tab1:
            plant_df = pd.read_excel(excel_file, sheet_name='Plant')
            st.dataframe(plant_df, use_container_width=True, height=300)
        with tab2:
            excel_file.seek(0)
            inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
            st.dataframe(inventory_df, use_container_width=True, height=300)
        with tab3:
            excel_file.seek(0)
            demand_df = pd.read_excel(excel_file, sheet_name='Demand')
            st.dataframe(demand_df, use_container_width=True, height=300)
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="material-card"><div class="card-title">⚙️ Optimization Parameters</div></div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.optimization_params['time_limit_min'] = st.number_input("⏱️ Time Limit (minutes)", min_value=1, value=st.session_state.optimization_params['time_limit_min'])
            st.session_state.optimization_params['buffer_days'] = st.number_input("📅 Planning Buffer (days)", min_value=0, value=st.session_state.optimization_params['buffer_days'])
        with col2:
            st.session_state.optimization_params['stockout_penalty'] = st.number_input("🎯 Stockout Penalty (per MT)", min_value=1, value=st.session_state.optimization_params['stockout_penalty'])
            st.session_state.optimization_params['transition_penalty'] = st.number_input("🔄 Transition Penalty (per changeover)", min_value=1, value=st.session_state.optimization_params['transition_penalty'])
        
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        
        col1, _, col3 = st.columns([1, 2, 1])
        with col1:
            if st.button("← Back to Upload", use_container_width=True):
                st.session_state.step = 1
                st.rerun()
        with col3:
            if st.button("Run Optimization →", type="primary", use_container_width=True):
                st.session_state.step = 3
                st.rerun()
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        if st.button("← Back to Upload"): st.session_state.step = 1; st.rerun()

# ============================================================================
# STEP 3: OPTIMIZATION & RESULTS (Modified to use the class)
# ============================================================================
elif st.session_state.step == 3:
    st.markdown('<div class="material-card"><div class="card-title">⚡ Running Optimization</div></div>', unsafe_allow_html=True)
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # --- 1. Load Data ---
        status_text.markdown('<div class="alert-box info">📊 Loading data...</div>', unsafe_allow_html=True)
        uploaded_file = st.session_state.uploaded_file
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        plant_df = pd.read_excel(excel_file, sheet_name='Plant')
        excel_file.seek(0)
        inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
        excel_file.seek(0)
        demand_df = pd.read_excel(excel_file, sheet_name='Demand')
        
        params = st.session_state.optimization_params
        progress_bar.progress(10)
        
        # --- 2. Initialize and Solve Model ---
        status_text.markdown('<div class="alert-box info">🔧 Building and solving model...</div>', unsafe_allow_html=True)
        
        # This is the key change: instantiating and running the model class
        model_runner = PolymerProductionModel(plant_df, inventory_df, demand_df, params)
        solution_data = model_runner.solve(params['time_limit_min'])
        
        progress_bar.progress(100)
        
        # --- 3. Process Results ---
        if solution_data:
            status_text.markdown(f'<div class="alert-box success">✅ {solution_data["status"]} solution found!</div>', unsafe_allow_html=True)
            time.sleep(0.5)
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # Metrics
            st.markdown('<div class="material-card"><div class="card-title">📊 Optimization Results</div></div>', unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="metric-card primary"><div class="metric-label">Objective Value</div><div class="metric-value">{solution_data["objective_value"]:,.0f}</div></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-card info"><div class="metric-label">Transitions</div><div class="metric-value">{solution_data["total_transitions"]}</div></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-card {"success" if solution_data["total_stockouts"] == 0 else "warning"}"><div class="metric-label">Stockouts (MT)</div><div class="metric-value">{solution_data["total_stockouts"]:,.0f}</div></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="metric-card info"><div class="metric-label">Solve Time</div><div class="metric-value">{solution_data["solve_time"]:.1f}s</div></div>', unsafe_allow_html=True)
            
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            
            # Tabbed results
            tab1, tab2, tab3 = st.tabs(["📅 Production Schedule", "📊 Summary Analytics", "📦 Inventory Trends"])
            
            # Unpack solution data for plotting
            lines = solution_data["lines"]
            grades = solution_data["grades"]
            dates = solution_data["dates"]
            num_days = solution_data["num_days"]
            shutdown_periods = solution_data["shutdown_periods"]
            
            sorted_grades = sorted(grades)
            base_colors = px.colors.qualitative.Vivid
            grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

            with tab1:
                for line in lines:
                    st.markdown(f"#### 🏭 {line}")
                    gantt_data = []
                    for d in range(num_days):
                        grade = solution_data["solution_schedule"][line][solution_data["formatted_dates"][d]]
                        if grade:
                            gantt_data.append({"Grade": grade, "Start": dates[d], "Finish": dates[d] + timedelta(days=1), "Line": line})
                    
                    if not gantt_data:
                        st.info(f"No production scheduled for {line}"); continue

                    fig = px.timeline(pd.DataFrame(gantt_data), x_start="Start", x_end="Finish", y="Grade", color="Grade",
                                      color_discrete_map=grade_color_map, category_orders={"Grade": sorted_grades})
                    if line in shutdown_periods and shutdown_periods[line]:
                        start_shutdown = dates[shutdown_periods[line][0]]
                        end_shutdown = dates[shutdown_periods[line][-1]] + timedelta(days=1)
                        fig.add_vrect(x0=start_shutdown, x1=end_shutdown, fillcolor="red", opacity=0.15, layer="below", line_width=0, annotation_text="Shutdown")
                    
                    fig.update_yaxes(autorange="reversed", title=None)
                    fig.update_xaxes(title="Date", tickformat="%d-%b")
                    fig.update_layout(height=350, showlegend=True, legend=dict(orientation="v", y=0.5, x=1.02), plot_bgcolor="white")
                    st.plotly_chart(fig, use_container_width=True)

            with tab2:
                prod_data, grade_totals, plant_totals, stockout_totals = [], {}, {l: 0 for l in lines}, {}
                for grade in grades:
                    grade_totals[grade] = sum(solution_data["solution_production"][grade].values())
                    stockout_totals[grade] = sum(solution_data["solution_stockout"][grade].values())
                    row = {'Grade': grade}
                    for line in lines:
                        line_prod = 0
                        for d in range(num_days):
                            if solution_data["solution_schedule"][line][solution_data["formatted_dates"][d]] == grade:
                                line_prod += plant_df.loc[plant_df['Plant'] == line, 'Capacity per day'].values[0]
                        row[line] = line_prod
                        plant_totals[line] += line_prod
                    row['Total Produced'] = grade_totals[grade]
                    row['Total Stockout'] = stockout_totals[grade]
                    prod_data.append(row)
                
                totals_row = {'Grade': 'TOTAL', **plant_totals, 'Total Produced': sum(plant_totals.values()), 'Total Stockout': sum(stockout_totals.values())}
                prod_data.append(totals_row)
                st.dataframe(pd.DataFrame(prod_data), use_container_width=True, hide_index=True)

            with tab3:
                for grade in sorted_grades:
                    inv_data = solution_data["solution_inventory"][grade]
                    inv_df = pd.DataFrame(list(inv_data.items()), columns=['Date', 'Inventory'])
                    inv_df = inv_df[inv_df['Date'] != 'final']
                    inv_df['Date'] = pd.to_datetime(inv_df['Date'], format='%d-%b-%y')
                    inv_df = inv_df.sort_values(by='Date')

                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=inv_df['Date'], y=inv_df['Inventory'], mode="lines+markers", name=grade, line=dict(color=grade_color_map[grade], width=3)))
                    
                    if solution_data["min_inventory"][grade] > 0:
                        fig.add_hline(y=solution_data["min_inventory"][grade], line=dict(color="#ef4444", width=2, dash="dash"), annotation_text=f"Min")
                    if solution_data["max_inventory"][grade] < 1000000000:
                        fig.add_hline(y=solution_data["max_inventory"][grade], line=dict(color="#10b981", width=2, dash="dash"), annotation_text=f"Max")

                    fig.update_layout(title=f"Inventory Level - {grade}", xaxis_title="Date", yaxis_title="Inventory (MT)", plot_bgcolor="white", height=420, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)

        else:
            # No solution found
            status_text.markdown('<div class="alert-box warning">⚠️ No Feasible Solution Found</div>', unsafe_allow_html=True)
            with st.expander("Troubleshooting Guide", expanded=True):
                st.markdown("""
                **Common Causes:**
                - **Constraint Conflicts:** Minimum run days may be too long for available windows, especially around shutdowns.
                - **Capacity Issues:** Total demand may exceed production capacity.
                - **Inventory Targets:** Minimum closing inventory targets may be unachievable.
                - **Transition Rules:** The transition matrix (if used) might be too restrictive.
                """)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        col1, col2, _ = st.columns([1, 1, 2])
        with col1:
            if st.button("🔄 New Optimization", use_container_width=True):
                st.session_state.step = 1; st.session_state.uploaded_file = None; st.rerun()
        with col2:
            if st.button("🔧 Adjust Parameters", use_container_width=True):
                st.session_state.step = 2; st.rerun()

    except Exception as e:
        st.markdown(f'<div class="alert-box warning"><strong>❌ Error During Optimization</strong><br>{str(e)}</div>', unsafe_allow_html=True)
        with st.expander("View Technical Details"):
            st.code(traceback.format_exc())
        if st.button("← Return to Start"):
            st.session_state.step = 1; st.session_state.uploaded_file = None; st.rerun()

# --- Footer ---
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown("<div style='text-align: center; color: #9e9e9e; font-size: 0.875rem;'>Polymer Production Scheduler • Refactored Architecture v4.0</div>", unsafe_allow_html=True)
