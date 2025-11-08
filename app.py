import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np
import time
import io
from matplotlib import colormaps
import matplotlib.colors as mcolors

# Set page configuration
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #1f77b4;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
    }
</style>
""", unsafe_allow_html=True)

# Main title
st.markdown('<div class="main-header">🏭 Polymer Production Scheduler</div>', unsafe_allow_html=True)

# Solution callback class
class SolutionCallback(cp_model.CpSolverSolutionCallback):
    def __init__(self, production, inventory, stockout, is_producing, grades, lines, dates, num_days):
        cp_model.CpSolverSolutionCallback.__init__(self)
        self.production = production
        self.inventory = inventory
        self.stockout = stockout
        self.is_producing = is_producing
        self.grades = grades
        self.lines = lines
        self.dates = dates
        self.num_days = num_days
        self.solutions = []
        self.solution_times = []
        self.start_time = time.time()

    def on_solution_callback(self):
        current_time = time.time() - self.start_time
        self.solution_times.append(current_time)
        current_obj = self.ObjectiveValue()

        # Store the current solution
        solution = {
            'objective': current_obj,
            'time': current_time,
            'production': {},
            'inventory': {},
            'stockout': {},
            'is_producing': {}
        }

        # Extract production values
        for grade in self.grades:
            solution['production'][grade] = {}
            for line in self.lines:
                for d in range(self.num_days):
                    key = (grade, line, d)
                    if key in self.production:
                        value = self.Value(self.production[key])
                        if value > 0:
                            date_key = self.dates[d]
                            if date_key not in solution['production'][grade]:
                                solution['production'][grade][date_key] = 0
                            solution['production'][grade][date_key] += value

        # Extract inventory values
        for grade in self.grades:
            solution['inventory'][grade] = {}
            for d in range(self.num_days + 1):
                key = (grade, d)
                if key in self.inventory:
                    if d < self.num_days:
                        solution['inventory'][grade][self.dates[d] if d > 0 else 'initial'] = self.Value(self.inventory[key])
                    else:
                        solution['inventory'][grade]['final'] = self.Value(self.inventory[key])

        # Extract stockout values
        for grade in self.grades:
            solution['stockout'][grade] = {}
            for d in range(self.num_days):
                key = (grade, d)
                if key in self.stockout:
                    value = self.Value(self.stockout[key])
                    if value > 0:
                        solution['stockout'][grade][self.dates[d]] = value

        # Extract production schedule
        for line in self.lines:
            solution['is_producing'][line] = {}
            for d in range(self.num_days):
                date_key = self.dates[d]
                solution['is_producing'][line][date_key] = None
                for grade in self.grades:
                    key = (grade, line, d)
                    if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                        solution['is_producing'][line][date_key] = grade
                        break

        self.solutions.append(solution)

        # Calculate transitions
        transition_count_per_line = {line: 0 for line in self.lines}
        total_transitions = 0

        for line in self.lines:
            last_grade = None
            for date in self.dates:
                current_grade = solution['is_producing'][line].get(date)
                if current_grade is not None and last_grade is not None:
                    if current_grade != last_grade:
                        transition_count_per_line[line] += 1
                        total_transitions += 1
                if current_grade is not None:
                    last_grade = current_grade

        solution['transitions'] = {
            'per_line': transition_count_per_line,
            'total': total_transitions
        }

    def num_solutions(self):
        return len(self.solutions)

# Sidebar for file upload and parameters
with st.sidebar:
    st.header("📁 Data Input")
    
    uploaded_file = st.file_uploader(
        "Upload Excel File", 
        type=["xlsx"],
        help="Upload an Excel file with Plant, Inventory, and Demand sheets"
    )
    
    if uploaded_file:
        st.success("✅ File uploaded successfully!")
        
        st.header("⚙️ Optimization Parameters")
        time_limit_min = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=5)
        buffer_days = st.number_input("Buffer days", min_value=0, max_value=7, value=3)
        stockout_penalty = st.number_input("Stockout penalty", min_value=1, value=10)
        transition_penalty = st.number_input("Transition penalty", min_value=1, value=10)
        continuity_bonus = st.number_input("Continuity bonus", min_value=0, value=1)

# Main content area
if uploaded_file:
    try:
        # Read the Excel file
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Read sheets
        plant_df = pd.read_excel(excel_file, sheet_name='Plant')
        excel_file.seek(0)
        inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
        excel_file.seek(0)
        demand_df = pd.read_excel(excel_file, sheet_name='Demand')
        
        # Show data preview
        st.markdown('<div class="section-header">📊 Data Preview</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("Plant Data")
            st.dataframe(plant_df, use_container_width=True)
        
        with col2:
            st.subheader("Inventory Data")
            st.dataframe(inventory_df, use_container_width=True)
        
        with col3:
            st.subheader("Demand Data")
            st.dataframe(demand_df, use_container_width=True)
        
        # Load transition matrices
        excel_file.seek(0)
        transition_dfs = {}
        for i in range(len(plant_df)):
            plant_name = plant_df['Plant'].iloc[i]
            sheet_name = f'Transition_{plant_name}'
            try:
                excel_file.seek(0)
                transition_dfs[plant_name] = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=0)
                st.info(f"✅ Loaded transition matrix for {plant_name}")
            except:
                st.info(f"ℹ️ No transition matrix found for {plant_name}. Assuming no transition constraints.")
                transition_dfs[plant_name] = None
        
        # Run optimization button
        st.markdown('<div class="section-header">🚀 Optimization</div>', unsafe_allow_html=True)
        
        if st.button("Run Production Optimization", type="primary", use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Data preprocessing (FIXED VERSION)
            status_text.markdown('<div class="info-box">🔄 Preprocessing data...</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            
            num_lines = len(plant_df)
            lines = list(plant_df['Plant'])
            capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
            grades = list(inventory_df['Grade Name'])
            initial_inventory = {row['Grade Name']: row['Opening Inventory'] for index, row in inventory_df.iterrows()}
            min_inventory = {row['Grade Name']: row['Min. Inventory'] for index, row in inventory_df.iterrows()}
            max_inventory = {row['Grade Name']: row['Max. Inventory'] for index, row in inventory_df.iterrows()}
            min_run_days = {row['Grade Name']: int(row['Min. Run Days']) if pd.notna(row['Min. Run Days']) else 1 for index, row in inventory_df.iterrows()}

            # SAFELY handle optional columns
            force_start_date = {}
            min_closing_inventory = {}

            for index, row in inventory_df.iterrows():
                grade = row['Grade Name']
                
                # Handle Force Start Date (optional column)
                if 'Force Start Date' in inventory_df.columns:
                    force_start_date[grade] = pd.to_datetime(row['Force Start Date']).date() if pd.notna(row['Force Start Date']) else None
                else:
                    force_start_date[grade] = None
                
                # Handle Min Closing Inventory (optional column)  
                if 'Min. Closing Inventory' in inventory_df.columns:
                    min_closing_inventory[grade] = row['Min. Closing Inventory'] if pd.notna(row['Min. Closing Inventory']) else 0
                else:
                    min_closing_inventory[grade] = 0

            allowed_lines = {
                row['Grade Name']: [x.strip() for x in str(row['Plant']).split(',')] if pd.notna(row['Plant']) else lines
                for index, row in inventory_df.iterrows()
            }

            rerun_allowed = {}
            for index, row in inventory_df.iterrows():
                rerun_val = row['Rerun Allowed']
                if isinstance(rerun_val, str) and rerun_val.strip().lower() == 'yes':
                    rerun_allowed[row['Grade Name']] = True
                else:
                    rerun_allowed[row['Grade Name']] = False

            max_run_days = {row['Grade Name']: int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999 for index, row in inventory_df.iterrows()}

            # Handle Material Running info (optional columns in Plant sheet)
            material_running_info = {}
            if 'Material Running' in plant_df.columns and 'Expected Run Days' in plant_df.columns:
                material_running_info = {
                    row['Plant']: (row['Material Running'], int(row['Expected Run Days']))
                    for index, row in plant_df.iterrows()
                    if pd.notna(row['Material Running']) and pd.notna(row['Expected Run Days'])
                }

            # Process demand data
            demand_data = {}
            dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            num_days = len(dates)
            last_date = dates[-1]
            for i in range(1, buffer_days + 1):
                dates.append(last_date + timedelta(days=i))
            num_days = len(dates)
            
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
            
            # Process transition rules
            transition_rules = {}
            for line, df in transition_dfs.items():
                if df is not None:
                    transition_rules[line] = {}
                    for prev_grade in df.index:
                        allowed_transitions = []
                        for current_grade in df.columns:
                            if str(df.loc[prev_grade, current_grade]).lower() == 'yes':
                                allowed_transitions.append(current_grade)
                        transition_rules[line][prev_grade] = allowed_transitions
                else:
                    transition_rules[line] = None
            
            progress_bar.progress(30)
            status_text.markdown('<div class="info-box">🔧 Building optimization model...</div>', unsafe_allow_html=True)
            
            # --- Create the CP-SAT Model ---
            model = cp_model.CpModel()

            # Decision Variables
            is_producing = {}
            production = {}
            for grade in grades:
                allowed = allowed_lines[grade]
                for line in allowed:
                    for d in range(num_days):
                        is_producing[(grade, line, d)] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                        if d < num_days - buffer_days:
                            production_value = capacities[line]
                        else:
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value <= capacities[line] * is_producing[(grade, line, d)])
                        production[(grade, line, d)] = production_value

            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')

            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')

            # Single Grade Production per Line per Day
            for line in lines:
                for d in range(num_days):
                    producing_vars = []
                    for grade in grades:
                        if line in allowed_lines[grade]:
                            producing_vars.append(is_producing[(grade, line, d)])
                    model.Add(sum(producing_vars) <= 1)

            # Handle Material Running (only if data exists)
            for plant, (material, expected_days) in material_running_info.items():
                for d in range(min(expected_days, num_days)):
                    model.Add(is_producing[(material, plant, d)] == 1)
                    for other_material in grades:
                        if other_material != material and plant in allowed_lines[other_material]:
                            model.Add(is_producing[(other_material, plant, d)] == 0)

            # --- Constraints ---
            objective = 0

            # Initial Inventory
            for grade in grades:
                model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])

            # Inventory Balance
            for grade in grades:
                for d in range(num_days):
                    produced_today = sum(production[(grade, line, d)] for line in allowed_lines[grade] if (grade, line, d) in production)
                    demand_today = demand_data[grade].get(dates[d], 0)

                    # Stockout and inventory update
                    stockout_var = stockout_vars[(grade, d)]
                    avail_today = model.NewIntVar(0, 100000, f'avail_{grade}_{d}')
                    model.Add(avail_today == inventory_vars[(grade, d)] + produced_today)

                    enough_supply = model.NewBoolVar(f'enough_supply_{grade}_{d}')
                    model.Add(avail_today >= demand_today).OnlyEnforceIf(enough_supply)
                    model.Add(avail_today < demand_today).OnlyEnforceIf(enough_supply.Not())

                    model.Add(stockout_var == 0).OnlyEnforceIf(enough_supply)
                    model.Add(stockout_var == demand_today - avail_today).OnlyEnforceIf(enough_supply.Not())

                    fulfilled = model.NewIntVar(0, 100000, f'fulfilled_{grade}_{d}')
                    model.Add(fulfilled == demand_today).OnlyEnforceIf(enough_supply)
                    model.Add(fulfilled == avail_today).OnlyEnforceIf(enough_supply.Not())

                    model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - fulfilled)

                    # Min inventory deficit penalty
                    if min_inventory[grade] > 0:
                        min_inv_value = int(min_inventory[grade])
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        below_min = model.NewBoolVar(f'below_min_{grade}_{d}')
                        model.Add(inventory_vars[(grade, d + 1)] < min_inv_value).OnlyEnforceIf(below_min)
                        model.Add(inventory_vars[(grade, d + 1)] >= min_inv_value).OnlyEnforceIf(below_min.Not())
                        model.Add(deficit == min_inv_value - inventory_vars[(grade, d + 1)]).OnlyEnforceIf(below_min)
                        model.Add(deficit == 0).OnlyEnforceIf(below_min.Not())
                        objective += stockout_penalty * deficit

            # Closing Inventory constraint
            for grade in grades:
                model.Add(inventory_vars[(grade, num_days - buffer_days)] >= int(min_closing_inventory[grade]))

            # Inventory Limits
            for grade in grades:
                for d in range(1, num_days + 1):
                    model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])

            # Line Capacity (Full Utilization)
            for line in lines:
                for d in range(num_days - buffer_days):
                    model.Add(sum(production[(grade, line, d)] for grade in grades if (grade, line, d) in production) == capacities[line])
                for d in range(num_days - buffer_days, num_days):
                    model.Add(sum(production[(grade, line, d)] for grade in grades if (grade, line, d) in production) <= capacities[line])

            # Force Start Date (only if specified)
            for grade in grades:
                if force_start_date[grade]:
                    try:
                        start_day_index = dates.index(force_start_date[grade])
                        force_production_constraints = []
                        for line in allowed_lines[grade]:
                            force_production_constraints.append(is_producing[(grade, line, start_day_index)])
                        if force_production_constraints:
                            model.AddBoolOr(force_production_constraints)
                        st.info(f"Force start date for grade '{grade}' set to day ({force_start_date[grade]})")
                    except ValueError:
                        st.warning(f"Force start date '{force_start_date[grade]}' for grade '{grade}' not found in demand dates.")

            # Minimum & Maximum Run Days
            is_start_vars = {}
            for grade in grades:
                for line in allowed_lines[grade]:
                    for d in range(num_days - min_run_days[grade] + 1):
                        is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                        is_start_vars[(grade, line, d)] = is_start
                        if d > 0:
                            model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d - 1)].Not()]).OnlyEnforceIf(is_start)
                            model.AddBoolOr([is_producing[(grade, line, d)].Not(), is_producing[(grade, line, d - 1)]]).OnlyEnforceIf(is_start.Not())
                        else:
                            model.Add(is_producing[(grade, line, d)] == 1).OnlyEnforceIf(is_start)
                            model.Add(is_start == 1).OnlyEnforceIf(is_producing[(grade, line, d)])

                        # Min Run Days
                        for k in range(1, min_run_days[grade]):
                            if d + k < num_days:
                                model.Add(is_producing[(grade, line, d + k)] == 1).OnlyEnforceIf(is_start)

                        # Max Run Days
                        if max_run_days[grade] < num_days and d + max_run_days[grade] < num_days:
                            model.Add(is_producing[(grade, line, d + max_run_days[grade])] == 0).OnlyEnforceIf(is_start)

            # Transition Rules
            for line in lines:
                if transition_rules.get(line):
                    for d in range(num_days - 1):
                        for prev_grade in grades:
                            if prev_grade in transition_rules[line]:
                                allowed_next = transition_rules[line][prev_grade]
                                for current_grade in grades:
                                    if current_grade != prev_grade and current_grade not in allowed_next:
                                        model.Add(is_producing[(prev_grade, line, d)] + is_producing[(current_grade, line, d + 1)] <= 1)

            # Rerun Allowed Constraints
            month_starts = {}
            month_ends = {}
            for d, date in enumerate(dates):
                key = (date.year, date.month)
                if key not in month_starts:
                    month_starts[key] = d
                month_ends[key] = d

            for grade in grades:
                if not rerun_allowed[grade]:
                    for line in allowed_lines[grade]:
                        for (year, month), start_day in month_starts.items():
                            end_day = month_ends[(year, month)]
                            produced_in_month = [is_producing[(grade, line, d)] for d in range(start_day, end_day + 1) if (grade, line, d) in is_producing]
                            if produced_in_month:
                                produced_at_all = model.NewBoolVar(f'produced_at_all_{grade}_{line}_{year}_{month}')
                                model.AddBoolOr(produced_in_month).OnlyEnforceIf(produced_at_all)
                                model.Add(sum(produced_in_month) == 0).OnlyEnforceIf(produced_at_all.Not())
                                days_producing = model.NewIntVar(0, end_day - start_day + 1, f'days_producing_{grade}_{line}_{year}_{month}')
                                model.Add(days_producing == sum(produced_in_month))
                                has_production = model.NewBoolVar(f'has_production_{grade}_{line}_{year}_{month}')
                                model.Add(days_producing >= min_run_days[grade]).OnlyEnforceIf(has_production)
                                model.Add(days_producing == 0).OnlyEnforceIf(has_production.Not())
                                model.AddImplication(produced_at_all, has_production)
                                starts_in_month = []
                                for d in range(start_day, end_day + 1):
                                    if (grade, line, d) in is_producing:
                                        if d > start_day:
                                            prev_d = d - 1
                                            is_start = model.NewBoolVar(f'is_start_{grade}_{line}_{d}_{year}_{month}')
                                            model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, prev_d)].Not()]).OnlyEnforceIf(is_start)
                                            model.AddBoolOr([is_producing[(grade, line, d)].Not(), is_producing[(grade, line, prev_d)]]).OnlyEnforceIf(is_start.Not())
                                            starts_in_month.append(is_start)
                                        else:
                                            is_start = model.NewBoolVar(f'is_start_{grade}_{line}_{d}_{year}_{month}')
                                            model.Add(is_producing[(grade, line, d)] == 1).OnlyEnforceIf(is_start)
                                            starts_in_month.append(is_start)
                                if starts_in_month:
                                    model.Add(sum(starts_in_month) <= 1).OnlyEnforceIf(produced_at_all)

            # Objective Function
            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

            # Transition Penalties and Continuity Bonus
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

                    # Continuity bonus for same grade
                    for grade in grades:
                        if line in allowed_lines[grade]:
                            continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                            model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                            objective += -continuity_bonus * continuity

            model.Minimize(objective)

            progress_bar.progress(50)
            status_text.markdown('<div class="info-box">⚡ Running optimization solver...</div>', unsafe_allow_html=True)

            # --- Solve the Model ---
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit_min * 60.0
            solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, num_days)

            # Solve with progress updates
            start_time = time.time()
            status = solver.Solve(model, solution_callback)
            
            progress_bar.progress(100)
            status_text.markdown('<div class="success-box">✅ Optimization completed successfully!</div>', unsafe_allow_html=True)

            # Display results
            st.markdown('<div class="section-header">📈 Results</div>', unsafe_allow_html=True)

            if solution_callback.num_solutions() > 0:
                best_solution = solution_callback.solutions[-1]

                # Key metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Objective Value", f"{best_solution['objective']:,.0f}")
                with col2:
                    st.metric("Total Transitions", best_solution['transitions']['total'])
                with col3:
                    total_stockouts = sum(sum(best_solution['stockout'][g].values()) for g in grades)
                    st.metric("Total Stockouts", f"{total_stockouts:,.0f}")
                with col4:
                    st.metric("Planning Horizon", f"{num_days} days")

                # Production schedule
                st.subheader("Production Schedule by Line")
                schedule_data = []
                for line in lines:
                    for date, grade in best_solution['is_producing'][line].items():
                        if grade:
                            schedule_data.append({'Line': line, 'Date': date, 'Grade': grade})
                
                if schedule_data:
                    schedule_df = pd.DataFrame(schedule_data)
                    st.dataframe(schedule_df, use_container_width=True)

                # Create visualization
                st.subheader("Production Visualization")
                
                # Create color map for grades
                cmap = colormaps.get_cmap('tab20')
                grade_colors = {}
                for idx, grade in enumerate(grades):
                    grade_colors[grade] = cmap(idx % 20)

                # Create production charts for each line
                for line in lines:
                    st.subheader(f"Production Chart - {line}")
                    
                    # Create dataframe for this line's production
                    line_data = []
                    for d in range(num_days):
                        date = dates[d]
                        for grade in grades:
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                line_data.append({
                                    'Date': date,
                                    'Grade': grade,
                                    'Production': solver.Value(production[(grade, line, d)])
                                })
                    
                    if line_data:
                        line_df = pd.DataFrame(line_data)
                        pivot_df = line_df.pivot_table(index='Date', columns='Grade', values='Production', aggfunc='sum').fillna(0)
                        
                        fig, ax = plt.subplots(figsize=(12, 6))
                        bottom = np.zeros(len(pivot_df))
                        
                        for grade in pivot_df.columns:
                            ax.bar(pivot_df.index, pivot_df[grade], bottom=bottom, label=grade, color=grade_colors[grade])
                            bottom += pivot_df[grade].values
                        
                        ax.set_title(f'Production Schedule - {line}')
                        ax.set_xlabel('Date')
                        ax.set_ylabel('Production Volume')
                        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                        plt.xticks(rotation=45)
                        plt.tight_layout()
                        st.pyplot(fig)

                # Download results
                st.markdown('<div class="section-header">📥 Download Results</div>', unsafe_allow_html=True)
                
                # Create Excel report
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Write summary sheet
                    summary_data = {
                        'Metric': ['Objective Value', 'Total Transitions', 'Total Stockouts', 'Planning Horizon', 'Time Limit (min)'],
                        'Value': [best_solution['objective'], best_solution['transitions']['total'], total_stockouts, f"{num_days} days", time_limit_min]
                    }
                    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Write production schedule
                    if schedule_data:
                        pd.DataFrame(schedule_data).to_excel(writer, sheet_name='Production Schedule', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Excel Report",
                    data=output,
                    file_name="production_schedule_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error("No solutions found during optimization. Please check your constraints and data.")
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.info("Please make sure your Excel file has the required sheets: 'Plant', 'Inventory', and 'Demand'")

else:
    # Welcome message when no file is uploaded
    st.markdown("""
    <div class="info-box">
    <h3>Welcome to the Polymer Production Scheduler! 🏭</h3>
    <p>This application helps optimize your polymer production schedule.</p>
    <p><strong>To get started:</strong></p>
    <ol>
        <li>Upload an Excel file with the required sheets (Plant, Inventory, Demand)</li>
        <li>Configure optimization parameters in the sidebar</li>
        <li>Run the optimization and view results</li>
        <li>Download the production schedule report</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>Polymer Production Scheduler • Built with Streamlit</div>",
    unsafe_allow_html=True
)
