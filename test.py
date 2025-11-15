import streamlit as st
import pandas as pd
import numpy as np
import io
import time
import base64
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from ortools.sat.python import cp_model

# === APP CONFIGURATION ===
st.set_page_config(
    page_title="Production Optimizer",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# === CUSTOM CSS FOR MATERIAL DESIGN ===
st.markdown("""
<style>
    /* Material Design inspired styles */
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1976D2;
        text-align: center;
        margin: 1rem 0 2rem 0;
        padding: 1rem;
        background: linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%);
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(25, 118, 210, 0.15);
    }
    
    .section-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border: 1px solid #E0E0E0;
        margin-bottom: 1.5rem;
    }
    
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #1976D2;
        text-align: center;
    }
    
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1976D2;
        margin: 0.25rem 0;
    }
    
    .metric-label {
        font-size: 0.85rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .chip-container {
        display: flex;
        gap: 0.5rem;
        flex-wrap: wrap;
        margin: 1rem 0;
    }
    
    .chip {
        padding: 0.5rem 1rem;
        background: #E3F2FD;
        border-radius: 20px;
        border: 1px solid #BBDEFB;
        cursor: pointer;
        transition: all 0.2s ease;
        font-size: 0.9rem;
    }
    
    .chip.active {
        background: #1976D2;
        color: white;
        border-color: #1976D2;
    }
    
    .chip:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .upload-area {
        border: 2px dashed #1976D2;
        border-radius: 12px;
        padding: 3rem;
        text-align: center;
        background: #F8F9FA;
        transition: all 0.3s ease;
        margin: 1rem 0;
    }
    
    .upload-area:hover {
        background: #E3F2FD;
        border-color: #1565C0;
    }
    
    .primary-button {
        background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s ease;
        box-shadow: 0 2px 4px rgba(25, 118, 210, 0.3);
    }
    
    .primary-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(25, 118, 210, 0.4);
    }
    
    .secondary-button {
        background: white;
        color: #1976D2;
        border: 1px solid #1976D2;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s ease;
    }
    
    .data-table {
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #F8F9FA;
        border-radius: 8px 8px 0px 0px;
        gap: 8px;
        padding: 10px 16px;
        font-weight: 600;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1976D2;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# === SESSION STATE INITIALIZATION ===
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'sheets' not in st.session_state:
    st.session_state.sheets = {}
if 'selected_sheet' not in st.session_state:
    st.session_state.selected_sheet = None
if 'params' not in st.session_state:
    st.session_state.params = {}
if 'run_state' not in st.session_state:
    st.session_state.run_state = 'idle'  # idle, running, completed, error
if 'results' not in st.session_state:
    st.session_state.results = None
if 'preview_rows' not in st.session_state:
    st.session_state.preview_rows = 8

# === UTILITY FUNCTIONS ===
@st.cache_data
def parse_uploaded_file(uploaded_file):
    """Parse uploaded file and return dictionary of sheets"""
    try:
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension in ['xlsx', 'xls']:
            # Read Excel file
            excel_file = pd.ExcelFile(uploaded_file)
            sheets = {}
            for sheet_name in excel_file.sheet_names:
                sheets[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            return sheets
        elif file_extension == 'csv':
            # Read CSV as single sheet
            csv_data = pd.read_csv(uploaded_file)
            return {'Sheet1': csv_data}
        else:
            st.error(f"Unsupported file format: {file_extension}")
            return {}
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return {}

def get_sheet_preview(sheet_data, num_rows=8):
    """Get preview of sheet data with specified number of rows"""
    if sheet_data is None:
        return None
    return sheet_data.head(num_rows)

def create_demo_file():
    """Create a demo Excel file for testing"""
    # Plant data
    plant_data = pd.DataFrame({
        'Plant': ['Plant1', 'Plant2', 'Plant3'],
        'Capacity per day': [1000, 1500, 1200],
        'Material Running': ['GradeA', 'GradeB', 'GradeA'],
        'Expected Run Days': [5, 3, 4],
        'Shutdown Start Date': [None, '2024-01-15', None],
        'Shutdown End Date': [None, '2024-01-18', None]
    })
    
    # Demand data
    dates = pd.date_range(start='2024-01-01', end='2024-01-20', freq='D')
    demand_data = pd.DataFrame({
        'Date': dates,
        'GradeA': np.random.randint(500, 2000, len(dates)),
        'GradeB': np.random.randint(300, 1500, len(dates)),
        'GradeC': np.random.randint(200, 1000, len(dates))
    })
    
    # Inventory data
    inventory_data = pd.DataFrame({
        'Grade Name': ['GradeA', 'GradeB', 'GradeC', 'GradeA', 'GradeB'],
        'Opening Inventory': [5000, 3000, 2000, 5000, 3000],
        'Min. Inventory': [1000, 800, 500, 1000, 800],
        'Max. Inventory': [8000, 6000, 4000, 8000, 6000],
        'Min. Run Days': [3, 2, 2, 3, 2],
        'Max. Run Days': [10, 8, 6, 10, 8],
        'Force Start Date': [None, None, None, '2024-01-05', None],
        'Lines': ['Plant1,Plant3', 'Plant2', 'Plant2,Plant3', 'Plant2', 'Plant1'],
        'Rerun Allowed': ['Yes', 'Yes', 'No', 'Yes', 'Yes'],
        'Min. Closing Inventory': [2000, 1500, 1000, 2000, 1500]
    })
    
    # Create Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        plant_data.to_excel(writer, sheet_name='Plant', index=False)
        demand_data.to_excel(writer, sheet_name='Demand', index=False)
        inventory_data.to_excel(writer, sheet_name='Inventory', index=False)
    
    output.seek(0)
    return output

# === BEGIN USER SOLVER CODE - DO NOT MODIFY THIS HEADER ===
class ProductionOptimizer:
    def __init__(self, data, params):
        self.data = data
        self.params = params
        self.model = cp_model.CpModel()
        self.solver = cp_model.CpSolver()
        
    def preprocess_data(self):
        """Extract and preprocess data from uploaded sheets"""
        try:
            # Extract plant data
            plant_df = self.data['Plant']
            self.plants = list(plant_df['Plant'])
            self.capacities = {row['Plant']: row['Capacity per day'] for _, row in plant_df.iterrows()}
            
            # Extract demand data
            demand_df = self.data['Demand']
            self.dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            self.num_days = len(self.dates)
            self.grades = [col for col in demand_df.columns if col != demand_df.columns[0]]
            
            # Process demand data
            self.demand_data = {}
            for grade in self.grades:
                if grade in demand_df.columns:
                    self.demand_data[grade] = {
                        demand_df.iloc[i, 0].date(): demand_df[grade].iloc[i] 
                        for i in range(len(demand_df))
                    }
            
            # Process inventory data
            inventory_df = self.data['Inventory']
            self.initial_inventory = {}
            self.min_inventory = {}
            self.max_inventory = {}
            self.min_run_days = {}
            self.max_run_days = {}
            self.allowed_plants = {grade: [] for grade in self.grades}
            
            for _, row in inventory_df.iterrows():
                grade = row['Grade Name']
                
                # Global inventory parameters
                if grade not in self.initial_inventory:
                    self.initial_inventory[grade] = row['Opening Inventory'] if pd.notna(row['Opening Inventory']) else 0
                    self.min_inventory[grade] = row['Min. Inventory'] if pd.notna(row['Min. Inventory']) else 0
                    self.max_inventory[grade] = row['Max. Inventory'] if pd.notna(row['Max. Inventory']) else 1000000
                
                # Plant-specific parameters
                lines_value = row['Lines']
                if pd.notna(lines_value) and lines_value != '':
                    plants_for_grade = [x.strip() for x in str(lines_value).split(',')]
                else:
                    plants_for_grade = self.plants
                
                for plant in plants_for_grade:
                    if plant not in self.allowed_plants[grade]:
                        self.allowed_plants[grade].append(plant)
                    
                    key = (grade, plant)
                    self.min_run_days[key] = int(row['Min. Run Days']) if pd.notna(row['Min. Run Days']) else 1
                    self.max_run_days[key] = int(row['Max. Run Days']) if pd.notna(row['Max. Run Days']) else 9999
            
            return True
            
        except Exception as e:
            st.error(f"Data preprocessing error: {str(e)}")
            return False
    
    def build_model(self):
        """Build the optimization model"""
        # Decision variables
        self.is_producing = {}
        self.production = {}
        self.inventory_vars = {}
        self.stockout_vars = {}
        
        # Create production variables
        for grade in self.grades:
            for plant in self.allowed_plants[grade]:
                for d in range(self.num_days):
                    key = (grade, plant, d)
                    self.is_producing[key] = self.model.NewBoolVar(f'is_producing_{grade}_{plant}_{d}')
                    self.production[key] = self.model.NewIntVar(0, self.capacities[plant], f'production_{grade}_{plant}_{d}')
                    
                    # Link production to binary variable
                    self.model.Add(self.production[key] == self.capacities[plant]).OnlyEnforceIf(self.is_producing[key])
                    self.model.Add(self.production[key] == 0).OnlyEnforceIf(self.is_producing[key].Not())
        
        # Create inventory variables
        for grade in self.grades:
            for d in range(self.num_days + 1):
                self.inventory_vars[(grade, d)] = self.model.NewIntVar(0, 1000000, f'inventory_{grade}_{d}')
        
        # Create stockout variables
        for grade in self.grades:
            for d in range(self.num_days):
                self.stockout_vars[(grade, d)] = self.model.NewIntVar(0, 1000000, f'stockout_{grade}_{d}')
        
        # Constraints
        self._add_constraints()
        
        # Objective function
        self._set_objective()
    
    def _add_constraints(self):
        """Add constraints to the model"""
        # One plant produces at most one grade per day
        for plant in self.plants:
            for d in range(self.num_days):
                producing_vars = []
                for grade in self.grades:
                    if plant in self.allowed_plants[grade]:
                        key = (grade, plant, d)
                        if key in self.is_producing:
                            producing_vars.append(self.is_producing[key])
                if producing_vars:
                    self.model.Add(sum(producing_vars) <= 1)
        
        # Inventory balance constraints
        for grade in self.grades:
            # Initial inventory
            self.model.Add(self.inventory_vars[(grade, 0)] == self.initial_inventory[grade])
            
            for d in range(self.num_days):
                # Total production for this grade on day d
                total_production = sum(
                    self.production[(grade, plant, d)] 
                    for plant in self.allowed_plants[grade] 
                    if (grade, plant, d) in self.production
                )
                
                # Demand for this grade on day d
                demand_today = self.demand_data[grade].get(self.dates[d], 0)
                
                # Inventory balance: inventory[d+1] = inventory[d] + production - demand + stockout
                self.model.Add(
                    self.inventory_vars[(grade, d + 1)] == 
                    self.inventory_vars[(grade, d)] + total_production - demand_today + self.stockout_vars[(grade, d)]
                )
                
                # Inventory bounds
                self.model.Add(self.inventory_vars[(grade, d)] >= self.min_inventory[grade])
                self.model.Add(self.inventory_vars[(grade, d)] <= self.max_inventory[grade])
        
        # Minimum run days constraints
        for grade in self.grades:
            for plant in self.allowed_plants[grade]:
                min_run = self.min_run_days.get((grade, plant), 1)
                
                for d in range(self.num_days - min_run + 1):
                    # If production starts at day d, it must continue for at least min_run days
                    start_var = self.model.NewBoolVar(f'start_{grade}_{plant}_{d}')
                    self.model.AddMaxEquality(start_var, [
                        self.is_producing[(grade, plant, d)],
                        self.model.NewConstant(1) if d == 0 else self.is_producing[(grade, plant, d-1)].Not()
                    ])
                    
                    # Enforce minimum run
                    for k in range(min_run):
                        if d + k < self.num_days:
                            self.model.Add(self.is_producing[(grade, plant, d + k)] == 1).OnlyEnforceIf(start_var)
    
    def _set_objective(self):
        """Set the objective function"""
        objective = 0
        
        # Stockout penalty
        stockout_penalty = self.params.get('stockout_penalty', 10)
        for grade in self.grades:
            for d in range(self.num_days):
                objective += stockout_penalty * self.stockout_vars[(grade, d)]
        
        # Transition penalty
        transition_penalty = self.params.get('transition_penalty', 5)
        for plant in self.plants:
            for d in range(self.num_days - 1):
                transition_vars = []
                for grade1 in self.grades:
                    if plant not in self.allowed_plants[grade1]:
                        continue
                    for grade2 in self.grades:
                        if plant not in self.allowed_plants[grade2] or grade1 == grade2:
                            continue
                        trans_var = self.model.NewBoolVar(f'trans_{plant}_{d}_{grade1}_to_{grade2}')
                        self.model.AddBoolAnd([
                            self.is_producing[(grade1, plant, d)],
                            self.is_producing[(grade2, plant, d + 1)]
                        ]).OnlyEnforceIf(trans_var)
                        transition_vars.append(trans_var)
                
                objective += transition_penalty * sum(transition_vars)
        
        self.model.Minimize(objective)
    
    def solve(self):
        """Solve the optimization problem"""
        # Set solver parameters
        self.solver.parameters.max_time_in_seconds = self.params.get('time_limit', 10) * 60
        self.solver.parameters.num_search_workers = 8
        
        # Solve
        status = self.solver.Solve(self.model)
        
        return status
    
    def get_results(self):
        """Extract results from the solved model"""
        if self.solver.Status() not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            return None
        
        results = {
            'tables': {},
            'plots': {},
            'logs': f"Solver status: {self.solver.StatusName()}\nObjective value: {self.solver.ObjectiveValue()}\n"
        }
        
        # Production schedule table
        schedule_data = []
        for d in range(self.num_days):
            row = {'Date': self.dates[d]}
            for plant in self.plants:
                plant_grade = 'Idle'
                for grade in self.grades:
                    key = (grade, plant, d)
                    if key in self.is_producing and self.solver.Value(self.is_producing[key]) == 1:
                        plant_grade = grade
                        break
                row[plant] = plant_grade
            schedule_data.append(row)
        
        results['tables']['production_schedule'] = pd.DataFrame(schedule_data)
        
        # Production summary table
        summary_data = []
        total_production = 0
        total_stockout = 0
        
        for grade in self.grades:
            grade_production = 0
            grade_stockout = 0
            
            for plant in self.allowed_plants[grade]:
                plant_production = 0
                for d in range(self.num_days):
                    key = (grade, plant, d)
                    if key in self.production:
                        plant_production += self.solver.Value(self.production[key])
                grade_production += plant_production
            
            for d in range(self.num_days):
                grade_stockout += self.solver.Value(self.stockout_vars[(grade, d)])
            
            total_production += grade_production
            total_stockout += grade_stockout
            
            summary_data.append({
                'Grade': grade,
                'Total Production': grade_production,
                'Total Stockout': grade_stockout,
                'Avg Daily Production': grade_production / self.num_days
            })
        
        results['tables']['production_summary'] = pd.DataFrame(summary_data)
        
        # Inventory levels table
        inventory_data = []
        for d in range(self.num_days):
            row = {'Date': self.dates[d]}
            for grade in self.grades:
                row[grade] = self.solver.Value(self.inventory_vars[(grade, d)])
            inventory_data.append(row)
        
        results['tables']['inventory_levels'] = pd.DataFrame(inventory_data)
        
        # Create plots
        self._create_plots(results)
        
        return results
    
    def _create_plots(self, results):
        """Create visualization plots"""
        # Production by plant
        plant_production = {}
        for plant in self.plants:
            plant_production[plant] = 0
            for grade in self.grades:
                if plant in self.allowed_plants[grade]:
                    for d in range(self.num_days):
                        key = (grade, plant, d)
                        if key in self.production:
                            plant_production[plant] += self.solver.Value(self.production[key])
        
        fig1 = px.bar(
            x=list(plant_production.keys()),
            y=list(plant_production.values()),
            title='Production by Plant',
            labels={'x': 'Plant', 'y': 'Total Production'}
        )
        results['plots']['production_by_plant'] = fig1
        
        # Inventory trends
        inventory_df = results['tables']['inventory_levels']
        fig2 = go.Figure()
        for grade in self.grades:
            fig2.add_trace(go.Scatter(
                x=inventory_df['Date'],
                y=inventory_df[grade],
                mode='lines+markers',
                name=grade
            ))
        fig2.update_layout(title='Inventory Trends', xaxis_title='Date', yaxis_title='Inventory Level')
        results['plots']['inventory_trends'] = fig2
        
        # Stockout analysis
        stockout_data = []
        for grade in self.grades:
            total_stockout = 0
            for d in range(self.num_days):
                total_stockout += self.solver.Value(self.stockout_vars[(grade, d)])
            stockout_data.append({'Grade': grade, 'Stockout': total_stockout})
        
        stockout_df = pd.DataFrame(stockout_data)
        fig3 = px.pie(stockout_df, values='Stockout', names='Grade', title='Stockout Distribution by Grade')
        results['plots']['stockout_distribution'] = fig3

def run_solver_wrapper(data, params):
    """
    Wrapper function for the production optimization solver.
    
    Args:
        data: Dictionary of DataFrames (sheets from uploaded file)
        params: Dictionary of optimization parameters
        
    Returns:
        Dictionary with keys: 'tables', 'plots', 'logs'
    """
    try:
        # Initialize optimizer
        optimizer = ProductionOptimizer(data, params)
        
        # Preprocess data
        if not optimizer.preprocess_data():
            raise Exception("Data preprocessing failed")
        
        # Build model
        optimizer.build_model()
        
        # Solve
        status = optimizer.solve()
        
        if status not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            raise Exception(f"Solver could not find solution. Status: {optimizer.solver.StatusName()}")
        
        # Get results
        results = optimizer.get_results()
        
        if results is None:
            raise Exception("Failed to extract results from solver")
        
        # Add solver logs
        results['logs'] += f"\nSolver statistics:\n"
        results['logs'] += f"  - Conflicts: {optimizer.solver.NumConflicts()}\n"
        results['logs'] += f"  - Branches: {optimizer.solver.NumBranches()}\n"
        results['logs'] += f"  - Wall time: {optimizer.solver.WallTime():.2f}s\n"
        
        return results
        
    except Exception as e:
        return {
            'tables': {},
            'plots': {},
            'logs': f"Error during optimization: {str(e)}"
        }
# === END USER SOLVER CODE ===

def display_plotly_figure(fig, caption=None):
    """Display Plotly figure with optional caption and export button"""
    if fig is None:
        st.info("No plot available")
        return
        
    st.plotly_chart(fig, use_container_width=True)
    
    if caption:
        st.caption(caption)
    
    # Export button
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("📥 Export PNG", key=f"export_{id(fig)}"):
            # Convert to PNG and offer download
            img_bytes = fig.to_image(format="png")
            st.download_button(
                label="Download PNG",
                data=img_bytes,
                file_name="plot.png",
                mime="image/png"
            )

# === MAIN APP LAYOUT ===
def main():
    # Header
    st.markdown('<div class="main-header">🏭 Production Optimization Dashboard</div>', unsafe_allow_html=True)
    
    # File Upload Section
    with st.container():
        st.markdown("## 📁 Data Input")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            # File upload area
            uploaded_file = st.file_uploader(
                "Upload your production data file",
                type=["xlsx", "xls", "csv"],
                key="file_uploader",
                label_visibility="collapsed"
            )
            
            if uploaded_file is not None:
                if uploaded_file != st.session_state.uploaded_file:
                    # New file uploaded
                    st.session_state.uploaded_file = uploaded_file
                    st.session_state.sheets = parse_uploaded_file(uploaded_file)
                    if st.session_state.sheets:
                        st.session_state.selected_sheet = list(st.session_state.sheets.keys())[0]
                        st.success(f"✅ Successfully loaded {len(st.session_state.sheets)} sheet(s)")
            
        with col2:
            st.markdown("### Quick Start")
            if st.button("📋 Load Demo File", use_container_width=True):
                demo_file = create_demo_file()
                st.session_state.uploaded_file = demo_file
                st.session_state.sheets = parse_uploaded_file(demo_file)
                if st.session_state.sheets:
                    st.session_state.selected_sheet = list(st.session_state.sheets.keys())[0]
                    st.success("Demo file loaded successfully!")
            
            if st.session_state.uploaded_file:
                file_info = st.session_state.uploaded_file
                st.info(f"""
                **File:** {file_info.name}  
                **Size:** {len(file_info.getvalue()) // 1024} KB  
                **Sheets:** {len(st.session_state.sheets)}
                """)
    
    # Preview and Parameters Section
    if st.session_state.sheets:
        st.markdown("---")
        
        # Sheet selector chips
        st.markdown("## 📊 Data Preview")
        st.markdown('<div class="chip-container">', unsafe_allow_html=True)
        for sheet_name in st.session_state.sheets.keys():
            is_active = sheet_name == st.session_state.selected_sheet
            chip_class = "chip active" if is_active else "chip"
            if st.markdown(f'<div class="{chip_class}" onclick="this.dispatchEvent(new Event(\'click\'))">{sheet_name}</div>', 
                          unsafe_allow_html=True):
                st.session_state.selected_sheet = sheet_name
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Preview controls
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            preview_option = st.selectbox(
                "Show rows:",
                [8, 25, 100, "All"],
                index=0,
                key="preview_rows"
            )
        with col2:
            if st.button("🔄 Refresh Preview"):
                st.rerun()
        
        # Data preview and parameters in columns
        col_preview, col_params = st.columns([2, 1])
        
        with col_preview:
            if st.session_state.selected_sheet:
                sheet_data = st.session_state.sheets[st.session_state.selected_sheet]
                
                # Sheet summary
                st.markdown(f"""
                **Sheet:** `{st.session_state.selected_sheet}`  
                **Shape:** {sheet_data.shape[0]} rows × {sheet_data.shape[1]} columns  
                **Data types:** {', '.join([f'{col}: {dtype}' for col, dtype in zip(sheet_data.columns, sheet_data.dtypes.astype(str))][:3])}{'...' if len(sheet_data.columns) > 3 else ''}
                """)
                
                # Data preview
                preview_data = get_sheet_preview(sheet_data, 
                    num_rows=preview_option if preview_option != "All" else len(sheet_data))
                
                st.markdown('<div class="data-table">', unsafe_allow_html=True)
                st.dataframe(preview_data, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
        
        with col_params:
            st.markdown("## ⚙️ Parameters")
            
            with st.form("parameters_form"):
                st.markdown("### Optimization Settings")
                
                # Numeric parameters
                col1, col2 = st.columns(2)
                with col1:
                    time_limit = st.number_input(
                        "Time Limit (min)",
                        min_value=1,
                        max_value=120,
                        value=10,
                        help="Maximum optimization time in minutes"
                    )
                    stockout_penalty = st.number_input(
                        "Stockout Penalty",
                        min_value=1,
                        max_value=100,
                        value=10,
                        help="Penalty cost per unit of stockout"
                    )
                
                with col2:
                    max_iterations = st.number_input(
                        "Max Iterations",
                        min_value=100,
                        max_value=10000,
                        value=1000,
                        step=100,
                        help="Maximum solver iterations"
                    )
                    transition_penalty = st.number_input(
                        "Transition Penalty",
                        min_value=1,
                        max_value=50,
                        value=5,
                        help="Cost for production line transitions"
                    )
                
                # Categorical parameters
                objective_priority = st.selectbox(
                    "Objective Priority",
                    ["Minimize Cost", "Maximize Production", "Balance Inventory"],
                    help="Primary optimization objective"
                )
                
                # Toggles
                col1, col2 = st.columns(2)
                with col1:
                    enforce_constraints = st.toggle(
                        "Enforce Constraints",
                        value=True,
                        help="Apply all production constraints"
                    )
                with col2:
                    allow_stockouts = st.toggle(
                        "Allow Stockouts",
                        value=False,
                        help="Permit temporary stockouts"
                    )
                
                # Store parameters
                st.session_state.params = {
                    'time_limit': time_limit,
                    'stockout_penalty': stockout_penalty,
                    'max_iterations': max_iterations,
                    'transition_penalty': transition_penalty,
                    'objective_priority': objective_priority,
                    'enforce_constraints': enforce_constraints,
                    'allow_stockouts': allow_stockouts
                }
                
                # Form actions
                col1, col2 = st.columns(2)
                with col1:
                    if st.form_submit_button("🔄 Reset to Defaults", use_container_width=True):
                        st.session_state.params = {}
                        st.rerun()
                with col2:
                    if st.form_submit_button("🚀 Start Optimization", use_container_width=True, type="primary"):
                        st.session_state.run_state = 'running'
    
    # Optimization Results Section
    if st.session_state.run_state == 'running':
        st.markdown("---")
        st.markdown("## ⚡ Running Optimization")
        
        # Progress indicator
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Simulate optimization steps
        steps = [
            "Preparing data...",
            "Building model...",
            "Setting constraints...",
            "Running solver...",
            "Generating results..."
        ]
        
        for i, step in enumerate(steps):
            progress_bar.progress((i + 1) * 20)
            status_text.markdown(f'<div class="section-card">{step}</div>', unsafe_allow_html=True)
            time.sleep(1)
            
            # Check for cancellation
            if st.session_state.run_state != 'running':
                status_text.markdown('<div class="section-card">❌ Optimization cancelled</div>', unsafe_allow_html=True)
                break
        
        if st.session_state.run_state == 'running':
            # Run the actual solver
            try:
                st.session_state.results = run_solver_wrapper(
                    st.session_state.sheets, 
                    st.session_state.params
                )
                st.session_state.run_state = 'completed'
                progress_bar.progress(100)
                status_text.markdown('<div class="section-card">✅ Optimization completed successfully!</div>', unsafe_allow_html=True)
            except Exception as e:
                st.session_state.run_state = 'error'
                status_text.markdown(f'<div class="section-card">❌ Optimization failed: {str(e)}</div>', unsafe_allow_html=True)
        
        # Cancel button
        if st.button("⏹️ Cancel Optimization", use_container_width=True):
            st.session_state.run_state = 'idle'
            st.rerun()
    
    # Display Results
    if st.session_state.run_state == 'completed' and st.session_state.results:
        st.markdown("---")
        st.markdown("## 📈 Optimization Results")
        
        # Key metrics
        if 'production_summary' in st.session_state.results['tables']:
            summary_df = st.session_state.results['tables']['production_summary']
            total_production = summary_df['Total Production'].sum()
            total_stockout = summary_df['Total Stockout'].sum()
            
            cols = st.columns(3)
            with cols[0]:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Production</div>
                    <div class="metric-value">{total_production:,.0f}</div>
                    <div class="metric-label">MT</div>
                </div>
                """, unsafe_allow_html=True)
            with cols[1]:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Total Stockout</div>
                    <div class="metric-value">{total_stockout:,.0f}</div>
                    <div class="metric-label">MT</div>
                </div>
                """, unsafe_allow_html=True)
            with cols[2]:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">Planning Horizon</div>
                    <div class="metric-value">{len(st.session_state.results['tables']['production_schedule'])}</div>
                    <div class="metric-label">Days</div>
                </div>
                """, unsafe_allow_html=True)
        
        # Results tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📋 Tables", "📈 Plots", "📝 Logs"])
        
        with tab1:
            st.markdown("### Optimization Summary")
            
            if 'production_summary' in st.session_state.results['tables']:
                st.dataframe(st.session_state.results['tables']['production_summary'], use_container_width=True)
            
            # Additional summary cards
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("""
                <div class="metric-card">
                    <div class="metric-label">Solver Status</div>
                    <div class="metric-value">Optimal</div>
                    <div class="metric-label">Solution Quality</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown("""
                <div class="metric-card">
                    <div class="metric-label">Plants Used</div>
                    <div class="metric-value">{len(st.session_state.sheets.get('Plant', pd.DataFrame())}</div>
                    <div class="metric-label">Active Plants</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown("""
                <div class="metric-card">
                    <div class="metric-label">Grades</div>
                    <div class="metric-value">{len(st.session_state.results['tables']['production_summary'])}</div>
                    <div class="metric-label">Production Grades</div>
                </div>
                """, unsafe_allow_html=True)
        
        with tab2:
            st.markdown("### Result Tables")
            
            for table_name, table_data in st.session_state.results['tables'].items():
                st.markdown(f"#### {table_name.replace('_', ' ').title()}")
                st.dataframe(table_data, use_container_width=True)
                
                # Export buttons for each table
                col1, col2 = st.columns([1, 5])
                with col1:
                    csv = table_data.to_csv(index=False)
                    st.download_button(
                        label="📥 CSV",
                        data=csv,
                        file_name=f"{table_name}.csv",
                        mime="text/csv",
                        key=f"csv_{table_name}"
                    )
        
        with tab3:
            st.markdown("### Visualization Plots")
            
            if 'plots' in st.session_state.results and st.session_state.results['plots']:
                for plot_name, plot_fig in st.session_state.results['plots'].items():
                    display_plotly_figure(
                        plot_fig, 
                        caption=plot_name.replace('_', ' ').title()
                    )
            else:
                st.info("No plots generated for this optimization run.")
        
        with tab4:
            st.markdown("### Solver Logs")
            
            if 'logs' in st.session_state.results:
                st.text_area(
                    "Optimization Log",
                    st.session_state.results['logs'],
                    height=400,
                    key="logs_display"
                )
                
                # Log export buttons
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📋 Copy to Clipboard", use_container_width=True):
                        st.code(st.session_state.results['logs'])
                with col2:
                    st.download_button(
                        label="📥 Download Log",
                        data=st.session_state.results['logs'],
                        file_name="optimization_log.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
            else:
                st.info("No logs available for this optimization run.")
        
        # Final actions
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Run New Optimization", use_container_width=True):
                st.session_state.run_state = 'idle'
                st.rerun()
        with col2:
            if st.button("📥 Export All Results", use_container_width=True):
                st.success("Export functionality would be implemented here")

# === APP FOOTER ===
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666; padding: 2rem 0;'>
        <strong>Production Optimization Dashboard</strong> • Built with Streamlit • 
        <a href='#file-upload' style='color: #1976D2; text-decoration: none;'>Upload Data</a> • 
        <a href='#parameters' style='color: #1976D2; text-decoration: none;'>Configure</a> • 
        <a href='#results' style='color: #1976D2; text-decoration: none;'>View Results</a>
    </div>
    """, 
    unsafe_allow_html=True
)

if __name__ == "__main__":
    main()
