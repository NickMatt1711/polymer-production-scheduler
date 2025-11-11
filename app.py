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
import base64
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Border, Side
import tempfile
import os
import plotly.express as px
import plotly.graph_objects as go

def create_sample_workbook():
    """Create a sample Excel workbook with the required format"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Plant sheet
        plant_data = {
            'Plant': ['Plant1', 'Plant2'],
            'Capacity per day': [1500, 1000],
            'Material Running': ['Moulding', 'BOPP'],
            'Expected Run Days': [1, 3]
        }
        plant_df = pd.DataFrame(plant_data)
        plant_df.to_excel(writer, sheet_name='Plant', index=False)
        
        # Inventory sheet
        inventory_data = {
            'Grade Name': ['BOPP', 'Moulding', 'Raffia', 'TQPP', 'Yarn'],
            'Opening Inventory': [500, 16000, 3000, 1700, 2500],
            'Min. Closing Inventory': [5000, 5000, 5000, 1500, 2000],
            'Min. Inventory': [500, 1000, 1000, 0, 0],
            'Max. Inventory': [20000, 20000, 20000, 6000, 6000],
            'Min. Run Days': [5, 1, 1, 3, 2],
            'Max. Run Days': [6, 6, 6, 5, 5],
            'Increment Days': ['', '', '', '', ''],
            'Force Start Date': ['', '', '', '', ''],
            'Lines': ['Plant1, Plant2', 'Plant1, Plant2', 'Plant1, Plant2', 'Plant1, Plant2', 'Plant2'],
            'Rerun Allowed': ['Yes', 'Yes', 'Yes', 'No', 'No']
        }
        inventory_df = pd.DataFrame(inventory_data)
        inventory_df.to_excel(writer, sheet_name='Inventory', index=False)
        
        # Demand sheet - generate dates for current month as proper Excel dates
        import calendar
        from datetime import datetime
        
        # Get current year and month
        now = datetime.now()
        current_year = now.year
        current_month = now.month
        
        # Get number of days in current month
        num_days_in_month = calendar.monthrange(current_year, current_month)[1]
        
        # Create date range for current month (1st to last day)
        dates = pd.date_range(
            start=f'{current_year}-{current_month:02d}-01',
            periods=num_days_in_month,
            freq='D'
        )
        
        demand_data = {
            'Date': dates,  # Keep as datetime objects for Excel
            'BOPP': [600] * num_days_in_month,
            'Moulding': [500] * num_days_in_month,
            'Raffia': [850] * num_days_in_month,
            'TQPP': [400] * num_days_in_month,
            'Yarn': [150] * num_days_in_month
        }
        demand_df = pd.DataFrame(demand_data)
        demand_df.to_excel(writer, sheet_name='Demand', index=False)
        
        # Transition matrices
        # Plant1 transition matrix
        Plant1_transition = {
            'From': ['BOPP', 'Moulding', 'Raffia', 'TQPP'],
            'BOPP': ['Yes', 'No', 'Yes', 'No'],
            'Moulding': ['No', 'Yes', 'Yes', 'Yes'],
            'Raffia': ['Yes', 'Yes', 'Yes', 'Yes'],
            'TQPP': ['No', 'Yes', 'Yes', 'Yes']
        }
        plant1_transition_df = pd.DataFrame(Plant1_transition)
        plant1_transition_df.to_excel(writer, sheet_name='Transition_Plant1', index=False)
        
        # Plant2 transition matrix
        Plant2_transition = {
            'From': ['BOPP', 'Moulding', 'Raffia', 'TQPP', 'Yarn'],
            'BOPP': ['Yes', 'No', 'Yes', 'Yes', 'No'],
            'Moulding': ['No', 'Yes', 'Yes', 'Yes', 'Yes'],
            'Raffia': ['Yes', 'Yes', 'Yes', 'Yes', 'No'],
            'TQPP': ['Yes', 'Yes', 'Yes', 'Yes', 'No'],
            'Yarn': ['No', 'Yes', 'No', 'No', 'Yes']
        }
        plant2_transition_df = pd.DataFrame(Plant2_transition)
        plant2_transition_df.to_excel(writer, sheet_name='Transition_Plant2', index=False)
        
        # Get workbook for formatting
        workbook = writer.book
        
        # Apply date formatting FIRST
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            
            # Apply date formatting to the Date column in Demand sheet
            if sheet_name == 'Demand':
                # Format the Date column (column A)
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                    for cell in row:
                        cell.number_format = 'DD-MMM-YY'  # Excel date format
        
        # Function to autofit columns - UPDATED to handle formatted dates
        def autofit_columns(ws):
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                # Special handling for date columns in Demand sheet
                if ws.title == 'Demand' and column_letter == 'A':
                    # For date column, use the formatted length (DD-MMM-YY = 9 characters)
                    max_length = 9
                else:
                    # For other columns, calculate max length as before
                    for cell in column:
                        try:
                            if cell.value:
                                # If it's a date cell with formatting, use the formatted length
                                if hasattr(cell, 'number_format') and cell.number_format:
                                    if 'MMM' in cell.number_format or 'mmm' in cell.number_format:
                                        max_length = max(max_length, 9)  # DD-MMM-YY format
                                    else:
                                        max_length = max(max_length, len(str(cell.value)))
                                else:
                                    max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column_letter].width = min(adjusted_width, 50)  # Cap at 50
        
        # Apply autofit to all sheets AFTER formatting
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            autofit_columns(ws)
    
    output.seek(0)
    return output


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
    .stProgress > div > div > div > div {
        background-color: #1f77b4;
    }
</style>
""", unsafe_allow_html=True)

# Main title
st.markdown('<div class="main-header">🏭 Polymer Production Scheduler</div>', unsafe_allow_html=True)

# Solution callback class
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
                            date_key = self.formatted_dates[d]
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
                        solution['inventory'][grade][self.formatted_dates[d] if d > 0 else 'initial'] = self.Value(self.inventory[key])
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
                        solution['stockout'][grade][self.formatted_dates[d]] = value
        
        # Extract production schedule
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
        time_limit_min = st.number_input(
            "Time limit (minutes)",
            min_value=1,
            max_value=120,
            value=10,
            help="Maximum time to run the optimization"
        )
        
        buffer_days = st.number_input(
            "Buffer days",
            min_value=0,
            max_value=7,
            value=3,
            help="Additional days for planning buffer"
        )
        
        stockout_penalty = st.number_input(
            "Stockout penalty",
            min_value=1,
            value=10,
            help="Penalty weight for stockouts in objective function"
        )
        
        transition_penalty = st.number_input(
            "Transition penalty", 
            min_value=1,
            value=10,
            help="Penalty weight for production line transitions"
        )
        
        continuity_bonus = st.number_input(
            "Continuity bonus",
            min_value=0,
            value=1,
            help="Bonus for continuing the same grade (negative penalty)"
        )

# Main content area
if uploaded_file:
    try:
        # Load data - reset the file pointer first
        uploaded_file.seek(0)
        excel_file = io.BytesIO(uploaded_file.read())
        
        # Show data preview
        st.markdown('<div class="section-header">📊 Data Preview</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            try:
                plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                st.subheader("Plant Data")
                st.dataframe(plant_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Plant sheet: {e}")
                st.stop()
        
        with col2:
            try:
                excel_file.seek(0)
                inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                st.subheader("Inventory Data")
                st.dataframe(inventory_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Inventory sheet: {e}")
                st.stop()
        
        with col3:
            try:
                excel_file.seek(0)
                demand_df = pd.read_excel(excel_file, sheet_name='Demand')
                
                # Create a copy for display with formatted dates
                demand_display_df = demand_df.copy()
                
                # Format the date column to dd-mmm-yy
                date_column = demand_display_df.columns[0]
                if pd.api.types.is_datetime64_any_dtype(demand_display_df[date_column]):
                    demand_display_df[date_column] = demand_display_df[date_column].dt.strftime('%d-%b-%y')
                
                st.subheader("Demand Data")
                st.dataframe(demand_display_df, use_container_width=True)
            except Exception as e:
                st.error(f"Error reading Demand sheet: {e}")
                st.stop()
        
        # Reset file pointer for transition matrices
        excel_file.seek(0)
        
        # Load transition matrices - FIXED VERSION
        transition_dfs = {}
        for i in range(len(plant_df)):
            plant_name = plant_df['Plant'].iloc[i]
            
            # Try multiple possible sheet name formats
            possible_sheet_names = [
                f'Transition_{plant_name}',           # Transition_Plant1
                f'Transition_{plant_name.replace(" ", "_")}',  # Transition_Plant1
                f'Transition{plant_name.replace(" ", "")}',    # TransitionPlant1
            ]
            
            transition_df_found = None
            for sheet_name in possible_sheet_names:
                try:
                    excel_file.seek(0)
                    transition_df_found = pd.read_excel(excel_file, sheet_name=sheet_name, index_col=0)
                    st.info(f"✅ Loaded transition matrix for {plant_name} from sheet '{sheet_name}'")
                    break
                except:
                    continue
            
            if transition_df_found is not None:
                transition_dfs[plant_name] = transition_df_found
            else:
                st.info(f"ℹ️ No transition matrix found for {plant_name}. Assuming no transition constraints.")
                transition_dfs[plant_name] = None
        
        # Run optimization button
        st.markdown('<div class="section-header">🚀 Optimization</div>', unsafe_allow_html=True)
        
        if st.button("Run Production Optimization", type="primary", use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            results_placeholder = st.empty()
            
            # Initialize session state for solutions
            if 'solutions' not in st.session_state:
                st.session_state.solutions = []
            if 'best_solution' not in st.session_state:
                st.session_state.best_solution = None
            
            # Data preprocessing
            status_text.markdown('<div class="info-box">🔄 Preprocessing data...</div>', unsafe_allow_html=True)
            progress_bar.progress(10)
            
            try:
                # Data preprocessing with corrected column names and default values
                num_lines = len(plant_df)
                lines = list(plant_df['Plant'])
                capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
                
                # --- Modified Inventory Preprocessing: Allow duplicate grades per plant ---
                inventory_records = []
                
                for _, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    lines_list = [x.strip() for x in str(row['Lines']).split(',')] if pd.notna(row['Lines']) else lines
                
                    for line in lines_list:
                        record = {
                            'Grade': grade,
                            'Line': line,
                            'Opening Inventory': row.get('Opening Inventory', 0) or 0,
                            'Min. Inventory': row.get('Min. Inventory', 0) or 0,
                            'Max. Inventory': row.get('Max. Inventory', 1e9) or 1e9,
                            'Min. Closing Inventory': row.get('Min. Closing Inventory', 0) or 0,
                            'Min. Run Days': int(row.get('Min. Run Days', 1) or 1),
                            'Max. Run Days': int(row.get('Max. Run Days', 9999) or 9999),
                            'Force Start Date': pd.to_datetime(row['Force Start Date']).date()
                                                if pd.notna(row['Force Start Date']) else None,
                            'Rerun Allowed': str(row.get('Rerun Allowed', 'Yes')).strip().lower() == 'yes'
                        }
                        inventory_records.append(record)
                
                # Convert into lookup dictionaries indexed by (grade, line)
                initial_inventory = {(r['Grade'], r['Line']): r['Opening Inventory'] for r in inventory_records}
                min_inventory = {(r['Grade'], r['Line']): r['Min. Inventory'] for r in inventory_records}
                max_inventory = {(r['Grade'], r['Line']): r['Max. Inventory'] for r in inventory_records}
                min_closing_inventory = {(r['Grade'], r['Line']): r['Min. Closing Inventory'] for r in inventory_records}
                min_run_days = {(r['Grade'], r['Line']): r['Min. Run Days'] for r in inventory_records}
                max_run_days = {(r['Grade'], r['Line']): r['Max. Run Days'] for r in inventory_records}
                force_start_date = {(r['Grade'], r['Line']): r['Force Start Date'] for r in inventory_records}
                rerun_allowed = {(r['Grade'], r['Line']): r['Rerun Allowed'] for r in inventory_records}
                
                # Get unique grade list as before
                grades = sorted(list(set([r['Grade'] for r in inventory_records])))
                
                # Allowed lines mapping
                allowed_lines = {}
                for grade in grades:
                    allowed_lines[grade] = [r['Line'] for r in inventory_records if r['Grade'] == grade]
                
                          
                # Material running info with error handling
                material_running_info = {}
                for index, row in plant_df.iterrows():
                    plant = row['Plant']
                    material = row['Material Running']
                    expected_days = row['Expected Run Days']
                    
                    if pd.notna(material) and pd.notna(expected_days):
                        try:
                            material_running_info[plant] = (str(material).strip(), int(expected_days))
                        except (ValueError, TypeError):
                            st.warning(f"⚠️ Invalid Material Running or Expected Run Days for plant '{plant}', ignoring")
                    elif pd.notna(material) or pd.notna(expected_days):
                        st.warning(f"⚠️ Incomplete Material Running info for plant '{plant}', ignoring both fields")
                
            except Exception as e:
                st.error(f"Error in data preprocessing: {str(e)}")
                import traceback
                st.error(f"Traceback: {traceback.format_exc()}")
                st.stop()
            
            # Process demand data
            demand_data = {}
            dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            num_days = len(dates)
            last_date = dates[-1]
            for i in range(1, buffer_days + 1):
                dates.append(last_date + timedelta(days=i))
            num_days = len(dates)
            
            # Convert dates to DD-MMM-YY format for display
            formatted_dates = [date.strftime('%d-%b-%y') for date in dates]
            
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
            
            # Decision Variables - WITH COMPREHENSIVE SAFETY CHECKS
            is_producing = {}
            production = {}
            
            # Helper function to safely check if a grade-line combination exists
            def is_allowed_combination(grade, line):
                return line in allowed_lines.get(grade, [])
            
            # Create variables only for allowed combinations
            for grade in grades:
                for line in allowed_lines[grade]:  # This ensures we only create variables for allowed lines
                    for d in range(num_days):
                        key = (grade, line, d)
                        is_producing[key] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                        
                        # Create production variable based on whether it's a buffer day or not
                        if d < num_days - buffer_days:
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value == capacities[line]).OnlyEnforceIf(is_producing[key])
                            model.Add(production_value == 0).OnlyEnforceIf(is_producing[key].Not())
                        else:
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value <= capacities[line] * is_producing[key])
                        
                        production[key] = production_value
            
            # Safety wrapper for accessing production variables
            def get_production_var(grade, line, d):
                key = (grade, line, d)
                if key not in production:
                    return 0  # Return zero if the combination doesn't exist
                return production[key]
            
            # Safety wrapper for accessing is_producing variables  
            def get_is_producing_var(grade, line, d):
                key = (grade, line, d)
                if key not in is_producing:
                    return None  # Return None if the combination doesn't exist
                return is_producing[key]
            
            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')
            
            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')
            
            # Single Grade Production per Line per Day - WITH SAFETY CHECKS
            for line in lines:
                for d in range(num_days):
                    producing_vars = []
                    for grade in grades:
                        # Only include variables for allowed combinations
                        if is_allowed_combination(grade, line):
                            var = get_is_producing_var(grade, line, d)
                            if var is not None:
                                producing_vars.append(var)
                    if producing_vars:  # Only add constraint if there are variables
                        model.Add(sum(producing_vars) <= 1)
            
            # Handle Material Running - WITH SAFETY CHECKS
            for plant, (material, expected_days) in material_running_info.items():
                for d in range(min(expected_days, num_days)):
                    # Only add constraint if this combination is allowed
                    if is_allowed_combination(material, plant):
                        model.Add(get_is_producing_var(material, plant, d) == 1)
                        for other_material in grades:
                            if other_material != material and is_allowed_combination(other_material, plant):
                                model.Add(get_is_producing_var(other_material, plant, d) == 0)
            
            # --- Constraints ---
            objective = 0
            
            # Initial Inventory
            for grade in grades:
                model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])
            
            # Inventory Balance - WITH SAFETY CHECKS
            for grade in grades:
                for d in range(num_days):
                    # Use safe production variable access
                    produced_today = sum(
                        get_production_var(grade, line, d) 
                        for line in allowed_lines[grade]  # Only sum over allowed lines
                    )
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
            
            # Line Capacity (Full Utilization) - WITH SAFETY CHECKS
            for line in lines:
                for d in range(num_days - buffer_days):
                    # Only sum production for grades that are allowed on this line
                    production_vars = [
                        get_production_var(grade, line, d) 
                        for grade in grades 
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:  # Only add constraint if there are variables
                        model.Add(sum(production_vars) == capacities[line])
                
                for d in range(num_days - buffer_days, num_days):
                    production_vars = [
                        get_production_var(grade, line, d) 
                        for grade in grades 
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:  # Only add constraint if there are variables
                        model.Add(sum(production_vars) <= capacities[line])
            
            # Force Start Date - WITH SAFETY CHECKS
            for grade in grades:
                if force_start_date[grade]:
                    try:
                        start_day_index = dates.index(force_start_date[grade])
                        force_production_constraints = []
                        for line in allowed_lines[grade]:  # Only consider allowed lines
                            var = get_is_producing_var(grade, line, start_day_index)
                            if var is not None:
                                force_production_constraints.append(var)
                        if force_production_constraints:
                            model.AddBoolOr(force_production_constraints)
                        st.info(f"Force start date for grade '{grade}' set to day ({force_start_date[grade]})")
                    except ValueError:
                        st.warning(f"Force start date '{force_start_date[grade]}' for grade '{grade}' not found in demand dates.")
            
            # Minimum & Maximum Run Days - WITH SAFETY CHECKS
            is_start_vars = {}
            for grade in grades:
                for line in allowed_lines[grade]:  # Only consider allowed lines
                    for d in range(num_days - min_run_days[grade] + 1):
                        is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                        is_start_vars[(grade, line, d)] = is_start
                        
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
            
                        # Min Run Days
                        for k in range(1, min_run_days[grade]):
                            if d + k < num_days:
                                future_prod = get_is_producing_var(grade, line, d + k)
                                if future_prod is not None:
                                    model.Add(future_prod == 1).OnlyEnforceIf(is_start)
            
                        # Max Run Days
                        if max_run_days[grade] < num_days and d + max_run_days[grade] < num_days:
                            future_prod = get_is_producing_var(grade, line, d + max_run_days[grade])
                            if future_prod is not None:
                                model.Add(future_prod == 0).OnlyEnforceIf(is_start)
            
            # Transition Rules - WITH SAFETY CHECKS
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
            solution_callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, dates, formatted_dates, num_days)

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
                    st.metric("Total Stockouts", f"{total_stockouts:,.0f} MT")
                with col4:
                    st.metric("Planning Horizon", f"{num_days} days")

                # Create color map for grades
                cmap = colormaps.get_cmap('tab20')
                grade_colors = {}
                for idx, grade in enumerate(grades):
                    grade_colors[grade] = cmap(idx % 20)
                
                # Total Production Quantity Table
                st.subheader("Total Production by Grade and Plant (MT)")
                
                # Calculate total production by grade and plant
                production_totals = {}
                grade_totals = {}
                plant_totals = {line: 0 for line in lines}
                
                for grade in grades:
                    production_totals[grade] = {}
                    grade_totals[grade] = 0
                    for line in lines:
                        total_prod = 0
                        for d in range(num_days):
                            key = (grade, line, d)
                            if key in production:
                                total_prod += solver.Value(production[key])
                        production_totals[grade][line] = total_prod
                        grade_totals[grade] += total_prod
                        plant_totals[line] += total_prod
                
                # Create DataFrame
                total_prod_data = []
                for grade in grades:
                    row = {'Grade': grade}
                    for line in lines:
                        row[line] = production_totals[grade][line]
                    row['Total'] = grade_totals[grade]
                    total_prod_data.append(row)
                
                # Add plant totals row
                totals_row = {'Grade': 'Total'}
                for line in lines:
                    totals_row[line] = plant_totals[line]
                totals_row['Total'] = sum(plant_totals.values())
                total_prod_data.append(totals_row)
                
                total_prod_df = pd.DataFrame(total_prod_data)
                
                # Display the table
                st.dataframe(total_prod_df, use_container_width=True)
                
                # Production schedule with color coding
                st.subheader("Production Schedule by Line")
                
                sorted_grades = sorted(grades)
                base_colors = px.colors.qualitative.Vivid
                grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}
                
                for line in lines:
                    st.markdown(f"### 🏭 {line}")
                
                    schedule_data = []
                    current_grade = None
                    start_day = None
                
                    # Iterate day-by-day to detect grade transitions
                    for d in range(num_days):
                        date = dates[d]
                        grade_today = None
                
                        for grade in sorted_grades:
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                grade_today = grade
                                break
                
                        # Detect a new production run or a change in grade
                        if grade_today != current_grade:
                            if current_grade is not None:
                                # End the previous run
                                end_date = dates[d - 1]
                                duration = (end_date - start_day).days + 1
                                schedule_data.append({
                                    "Grade": current_grade,
                                    "Start Date": start_day.strftime("%d-%b-%y"),
                                    "End Date": end_date.strftime("%d-%b-%y"),
                                    "Days": duration
                                })
                            # Start a new run
                            current_grade = grade_today
                            start_day = date
                
                    # Add the final run if still active
                    if current_grade is not None:
                        end_date = dates[-1]
                        duration = (end_date - start_day).days + 1
                        schedule_data.append({
                            "Grade": current_grade,
                            "Start Date": start_day.strftime("%d-%b-%y"),
                            "End Date": end_date.strftime("%d-%b-%y"),
                            "Days": duration
                        })
                
                    if not schedule_data:
                        st.info(f"No production data available for {line}.")
                        continue
                
                    schedule_df = pd.DataFrame(schedule_data)
                
                    # ✅ Apply grade color styling
                    def color_grade(val):
                        if val in grade_color_map:
                            color = grade_color_map[val]
                            return f'background-color: {color}; color: white; font-weight: bold; text-align: center;'
                        return ''
                
                    # ✅ Render table in Streamlit with color coding
                    styled_df = schedule_df.style.applymap(color_grade, subset=['Grade'])
                    st.dataframe(styled_df, use_container_width=True)

                # Create visualization
                st.subheader("Production Visualization")
                
                for line in lines:
                    st.markdown(f"### Production Schedule - {line}")
                
                    gantt_data = []
                    for d in range(num_days):
                        date = dates[d]
                        for grade in sorted_grades:  # use sorted order for consistency
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                gantt_data.append({
                                    "Grade": grade,
                                    "Start": date,
                                    "Finish": date + timedelta(days=1),
                                    "Line": line
                                })
                
                    if not gantt_data:
                        st.info(f"No production data available for {line}.")
                        continue
                
                    gantt_df = pd.DataFrame(gantt_data)
                
                    # ✅ Build Gantt chart using consistent color map
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
                
                    # ✅ Axes and grid formatting
                    fig.update_yaxes(
                        autorange="reversed",
                        title=None,
                        showgrid=True,
                        gridcolor="lightgray",
                        gridwidth=1
                    )
                
                    fig.update_xaxes(
                        title="Date",
                        showgrid=True,
                        gridcolor="lightgray",
                        gridwidth=1,
                        tickvals=dates,
                        tickformat="%d-%b",
                        dtick="D1"
                    )
                
                    # ✅ Layout styling with right-aligned legend
                    fig.update_layout(
                        height=350,
                        bargap=0.2,
                        showlegend=True,
                        legend_title_text="Grade",
                        legend=dict(
                            traceorder="normal",  # keeps order as per sorted_grades
                            orientation="v",
                            yanchor="middle",
                            y=0.5,
                            xanchor="left",
                            x=1.02,  # place legend just outside right edge
                            bgcolor="rgba(255,255,255,0)",
                            bordercolor="lightgray",
                            borderwidth=0
                        ),
                        xaxis=dict(showline=True, showticklabels=True),
                        yaxis=dict(showline=True),
                        margin=dict(l=60, r=160, t=60, b=60),  # extra right margin for legend
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        font=dict(size=12),
                    )
                
                    st.plotly_chart(fig, use_container_width=True)

                # Create inventory charts with data labels
                st.subheader("Inventory Levels")
                
                last_actual_day = num_days - buffer_days - 1

                for grade in sorted_grades:
                    inventory_values = [solver.Value(inventory_vars[(grade, d)]) for d in range(num_days)]
                
                    # ✅ Key points
                    start_val = inventory_values[0]
                    end_val = inventory_values[last_actual_day]
                    highest_val = max(inventory_values[: last_actual_day + 1])
                    lowest_val = min(inventory_values[: last_actual_day + 1])
                
                    start_x = dates[0]
                    end_x = dates[last_actual_day]
                    highest_x = dates[inventory_values.index(highest_val)]
                    lowest_x = dates[inventory_values.index(lowest_val)]
                
                    fig = go.Figure()
                
                    # ✅ Inventory line
                    fig.add_trace(go.Scatter(
                        x=dates,
                        y=inventory_values,
                        mode="lines+markers",
                        name=grade,
                        line=dict(color=grade_color_map[grade], width=3),
                        marker=dict(size=6),
                        hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"
                    ))
                
                    # ✅ Min/Max inventory lines (with numeric labels)
                    fig.add_hline(
                        y=min_inventory[grade],
                        line=dict(color="red", width=2, dash="dash"),
                        annotation_text=f"Min: {min_inventory[grade]:,.0f}",
                        annotation_position="top left",
                        annotation_font_color="red"
                    )
                    fig.add_hline(
                        y=max_inventory[grade],
                        line=dict(color="green", width=2, dash="dash"),
                        annotation_text=f"Max: {max_inventory[grade]:,.0f}",
                        annotation_position="bottom left",
                        annotation_font_color="green"
                    )
                
                    # ✅ Inline annotation labels (Start / End / High / Low)
                    annotations = [
                        dict(
                            x=start_x, y=start_val,
                            text=f"Start: {start_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=-40, ay=30,
                            font=dict(color="black", size=11),
                            bgcolor="white", bordercolor="gray"
                        ),
                        dict(
                            x=end_x, y=end_val,
                            text=f"End: {end_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=40, ay=30,
                            font=dict(color="black", size=11),
                            bgcolor="white", bordercolor="gray"
                        ),
                        dict(
                            x=highest_x, y=highest_val,
                            text=f"High: {highest_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=0, ay=-40,
                            font=dict(color="darkgreen", size=11),
                            bgcolor="white", bordercolor="gray"
                        ),
                        dict(
                            x=lowest_x, y=lowest_val,
                            text=f"Low: {lowest_val:.0f}",
                            showarrow=True, arrowhead=2,
                            ax=0, ay=40,
                            font=dict(color="firebrick", size=11),
                            bgcolor="white", bordercolor="gray"
                        )
                    ]
                
                    # ✅ Layout configuration
                    fig.update_layout(
                        title=f"Inventory Level - {grade}",
                        xaxis=dict(
                            title="Date",
                            showgrid=True,
                            gridcolor="lightgray",
                            tickvals=dates,
                            tickformat="%d-%b",
                            dtick="D1"
                        ),
                        yaxis=dict(
                            title="Inventory Volume (MT)",
                            showgrid=True,
                            gridcolor="lightgray"
                        ),
                        annotations=annotations,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(l=60, r=80, t=60, b=60),
                        font=dict(size=12),
                        height=420,
                        showlegend=False
                    )
                
                    st.plotly_chart(fig, use_container_width=True)


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
    <p>This application helps optimize your polymer production schedule by:</p>
    <ul>
        <li>📊 Analyzing plant capacities and inventory constraints</li>
        <li>⚡ Optimizing production sequences to minimize transitions</li>
        <li>📈 Balancing inventory levels and meeting demand</li>
        <li>💾 Generating detailed production schedules and reports</li>
    </ul>
    <p><strong>To get started:</strong></p>
    <ol>
        <li>Upload an Excel file with the required sheets (Plant, Inventory, Demand)</li>
        <li>Configure optimization parameters in the sidebar</li>
        <li>Run the optimization and view results</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
    
    # Create sample workbook
    sample_workbook = create_sample_workbook()
    
    # Download sample file section
    st.markdown("---")
    st.markdown('<div class="section-header">📥 Get Started with Sample Template</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        **Download our sample template file to get started quickly:**
        - Includes all required sheets with proper formatting
        - Contains sample data that you can modify
        - Ready-to-use structure for the optimization
        """)
    
    with col2:
        st.download_button(
            label="📥 Download Sample Template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # Sample file format guide
    with st.expander("📋 Required Excel File Format Details"):
        st.markdown("""
        Your Excel file should contain the following sheets with these exact column headers:
        
        **1. Plant Sheet**
        - `Plant`: Plant names (e.g., Plant1, Plant2)
        - `Capacity per day`: Daily production capacity
        - `Material Running`: Currently running material (optional)
        - `Expected Run Days`: Expected run days (optional)
        
        **2. Inventory Sheet**
        - `Grade Name`: Material grades
        - `Opening Inventory`: Starting inventory levels
        - `Min. Inventory`: Minimum inventory requirements
        - `Max. Inventory`: Maximum inventory capacity
        - `Min. Run Days`: Minimum consecutive run days
        - `Max. Run Days`: Maximum consecutive run days
        - `Force Start Date`: Mandatory start dates (optional)
        - `Lines`: Allowed production lines (comma-separated)
        - `Rerun Allowed`: Whether rerun is allowed (Yes/No)
        - `Min. Closing Inventory`: Minimum closing inventory
        
        **3. Demand Sheet**
        - First column: Dates
        - Subsequent columns: Demand for each grade (column names should match grade names)
        
        **4. Transition Sheets (optional)**
        - Name format: `Transition_[PlantName]` (e.g., `Transition_Plant1`, `Transition_Plant2`)
        - Rows: Previous grade
        - Columns: Next grade  
        - Values: "yes" for allowed transitions
        """)

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>Polymer Production Scheduler • Built with Streamlit</div>",
    unsafe_allow_html=True
)
