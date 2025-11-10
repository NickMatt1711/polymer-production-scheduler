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
        
        # Autofit columns and apply date formatting for all sheets
        workbook = writer.book
        
        # Function to autofit columns
        def autofit_columns(ws):
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column_letter].width = min(adjusted_width, 50)  # Cap at 50
        
        # Apply autofit to all sheets and date formatting to Demand sheet
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            autofit_columns(ws)
            
            # Apply date formatting to the Date column in Demand sheet
            if sheet_name == 'Demand':
                # Find the Date column (assuming it's the first column)
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                    for cell in row:
                        cell.number_format = 'DD-MMM-YY'  # Excel date format
        
        # Also format the Force Start Date column in Inventory sheet if it contains dates
        inventory_ws = writer.sheets['Inventory']
        # Find the Force Start Date column (you might need to adjust the column index)
        # Let's find it by header name
        force_start_col = None
        for col in range(1, inventory_ws.max_column + 1):
            if inventory_ws.cell(row=1, column=col).value == 'Force Start Date':
                force_start_col = col
                break
        
        if force_start_col:
            for row in range(2, inventory_ws.max_row + 1):
                cell = inventory_ws.cell(row=row, column=force_start_col)
                if cell.value and str(cell.value).strip():  # If cell has a value
                    try:
                        # Convert to datetime and apply format
                        date_val = pd.to_datetime(cell.value)
                        cell.value = date_val
                        cell.number_format = 'DD-MMM-YY'
                    except:
                        # If it's not a valid date, leave as is
                        pass
    
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
                st.subheader("Demand Data")
                st.dataframe(demand_df, use_container_width=True)
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
                grades = list(inventory_df['Grade Name'])
                
                # Enhanced inventory processing with default values
                initial_inventory = {}
                min_inventory = {}
                max_inventory = {}
                min_closing_inventory = {}
                min_run_days = {}
                max_run_days = {}
                force_start_date = {}
                allowed_lines = {}
                rerun_allowed = {}
                
                for index, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    
                    # Opening Inventory - default 0
                    if pd.notna(row['Opening Inventory']):
                        initial_inventory[grade] = row['Opening Inventory']
                    else:
                        initial_inventory[grade] = 0
                        st.warning(f"⚠️ Opening Inventory for grade '{grade}' is missing, using default value: 0")
                    
                    # Min. Inventory - default 0
                    if pd.notna(row['Min. Inventory']):
                        min_inventory[grade] = row['Min. Inventory']
                    else:
                        min_inventory[grade] = 0
                        st.warning(f"⚠️ Min. Inventory for grade '{grade}' is missing, using default value: 0")
                    
                    # Max. Inventory - default very large number (limitless)
                    if pd.notna(row['Max. Inventory']):
                        max_inventory[grade] = row['Max. Inventory']
                    else:
                        max_inventory[grade] = 1000000000  # Very large number as "limitless"
                        st.warning(f"⚠️ Max. Inventory for grade '{grade}' is missing, using default value: 1000000000 (limitless)")
                    
                    # Min. Closing Inventory - default 0
                    if pd.notna(row['Min. Closing Inventory']):
                        min_closing_inventory[grade] = row['Min. Closing Inventory']
                    else:
                        min_closing_inventory[grade] = 0
                        st.warning(f"⚠️ Min. Closing Inventory for grade '{grade}' is missing, using default value: 0")
                    
                    # Min. Run Days - default 1
                    if pd.notna(row['Min. Run Days']):
                        min_run_days[grade] = int(row['Min. Run Days'])
                    else:
                        min_run_days[grade] = 1
                        st.warning(f"⚠️ Min. Run Days for grade '{grade}' is missing, using default value: 1")
                    
                    # Max. Run Days - default 9999 (very large)
                    if pd.notna(row['Max. Run Days']):
                        max_run_days[grade] = int(row['Max. Run Days'])
                    else:
                        max_run_days[grade] = 9999
                        st.warning(f"⚠️ Max. Run Days for grade '{grade}' is missing, using default value: 9999")
                    
                    # Force Start Date
                    if pd.notna(row['Force Start Date']):
                        try:
                            force_start_date[grade] = pd.to_datetime(row['Force Start Date']).date()
                        except:
                            force_start_date[grade] = None
                            st.warning(f"⚠️ Force Start Date for grade '{grade}' is invalid, ignoring")
                    else:
                        force_start_date[grade] = None
                    
                    # Lines - handle None values properly
                    lines_value = row['Lines']
                    if pd.notna(lines_value) and lines_value != '':
                        # Split by comma and strip whitespace
                        allowed_lines[grade] = [x.strip() for x in str(lines_value).split(',')]
                    else:
                        # If no lines specified, allow all lines
                        allowed_lines[grade] = lines
                        st.warning(f"⚠️ Lines for grade '{grade}' are not specified, allowing all lines: {lines}")
                    
                    # Rerun Allowed - default 'Yes'
                    rerun_val = row['Rerun Allowed']
                    if pd.notna(rerun_val) and isinstance(rerun_val, str) and rerun_val.strip().lower() == 'no':
                        rerun_allowed[grade] = False
                    else:
                        rerun_allowed[grade] = True
                        if pd.isna(rerun_val) or (isinstance(rerun_val, str) and rerun_val.strip().lower() != 'yes'):
                            st.warning(f"⚠️ Rerun Allowed for grade '{grade}' is missing or invalid, using default value: Yes")
            
                # FIXED: Changed from 'Plant' to 'Lines' column and handle None values properly
                allowed_lines = {}
                for index, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    lines_value = row['Lines']
                    if pd.notna(lines_value) and lines_value != '':
                        # Split by comma and strip whitespace
                        allowed_lines[grade] = [x.strip() for x in str(lines_value).split(',')]
                    else:
                        # If no lines specified, allow all lines
                        allowed_lines[grade] = lines
            
                rerun_allowed = {}
                for index, row in inventory_df.iterrows():
                    rerun_val = row['Rerun Allowed']
                    if isinstance(rerun_val, str) and rerun_val.strip().lower() == 'yes':
                        rerun_allowed[row['Grade Name']] = True
                    else:
                        rerun_allowed[row['Grade Name']] = False
            
                max_run_days = {row['Grade Name']: int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999 for index, row in inventory_df.iterrows()}
                min_closing_inventory = {row['Grade Name']: row['Min. Closing Inventory'] if pd.notna(row['Min. Closing Inventory']) else 0 for index, row in inventory_df.iterrows()}
            
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
                
                schedule_data = []
                for line in lines:
                    for date, grade in best_solution['is_producing'][line].items():
                        if grade:
                            schedule_data.append({'Line': line, 'Date': date, 'Grade': grade})
                
                if schedule_data:
                    schedule_df = pd.DataFrame(schedule_data)
                    
                    # Create a color mapping function
                    def color_grade(val):
                        if val in grade_colors:
                            color = mcolors.to_hex(grade_colors[val])
                            return f'background-color: {color}; color: white; font-weight: bold;'
                        return ''
                    
                    # Apply styling
                    styled_schedule = schedule_df.style.applymap(color_grade, subset=['Grade'])
                    st.dataframe(styled_schedule, use_container_width=True)

                # Create visualization
                st.subheader("Production Visualization")
                
                # Create production charts for each line
                for line in lines:
                    st.subheader(f"Production Chart - {line}")
                    
                    # Create dataframe for this line's production
                    line_data = []
                    for d in range(num_days):
                        date = formatted_dates[d]  # Use formatted date
                        for grade in grades:
                            if (grade, line, d) in is_producing and solver.Value(is_producing[(grade, line, d)]) == 1:
                                line_data.append({
                                    'Date': date,
                                    'Day': d + 1,  # Add day number starting from 1
                                    'Grade': grade,
                                    'Production': solver.Value(production[(grade, line, d)])
                                })
                    
                    if line_data:
                        line_df = pd.DataFrame(line_data)
                        # Use Day instead of Date for the pivot
                        pivot_df = line_df.pivot_table(index='Day', columns='Grade', values='Production', aggfunc='sum').fillna(0)
                        
                        fig, ax = plt.subplots(figsize=(12, 6))
                        bottom = np.zeros(len(pivot_df))
                        
                        for grade in pivot_df.columns:
                            ax.bar(pivot_df.index, pivot_df[grade], bottom=bottom, label=grade, color=grade_colors[grade])
                            bottom += pivot_df[grade].values
                        
                        ax.set_title(f'Production Schedule - {line}')
                        ax.set_xlabel('Day')
                        ax.set_ylabel('Production Volume (MT)')
                        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                        
                        # Set x-axis to show all day numbers
                        ax.set_xticks(range(1, num_days + 1))
                        ax.set_xticklabels(range(1, num_days + 1))
                        
                        plt.tight_layout()
                        st.pyplot(fig)

                # Create inventory charts with data labels
                st.subheader("Inventory Levels")
                
                for grade in grades: 
                    inventory_values = []
                    for d in range(num_days):
                        inventory_values.append(solver.Value(inventory_vars[(grade, d)]))
                    
                    fig, ax = plt.subplots(figsize=(12, 4))
                    day_numbers = list(range(1, num_days + 1))
                    line = ax.plot(day_numbers, inventory_values, marker='o', label=grade, color=grade_colors[grade], linewidth=2, markersize=6)
                    
                    # Add data labels
                    for i, (day, inv) in enumerate(zip(day_numbers, inventory_values)):
                        ax.annotate(f'{inv:.0f}', 
                                   (day, inv), 
                                   textcoords="offset points", 
                                   xytext=(0,10), 
                                   ha='center', 
                                   fontsize=8,
                                   bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor="gray", alpha=0.7))
                    
                    ax.axhline(y=min_inventory[grade], color='red', linestyle='--', label='Min Inventory')
                    ax.axhline(y=max_inventory[grade], color='green', linestyle='--', label='Max Inventory')
                    ax.set_title(f'Inventory Level - {grade}')
                    ax.set_xlabel('Day')
                    ax.set_ylabel('Inventory Volume (MT)')
                    ax.legend()
                    ax.grid(True, alpha=0.3)
                    plt.xticks(day_numbers)  # Show all day numbers
                    plt.tight_layout()
                    st.pyplot(fig)


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
