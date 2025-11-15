# app.py
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
from pathlib import Path

# -------------------------
# Helper functions (kept intact)
# -------------------------
def get_sample_workbook():
    """Retrieve the sample workbook from the same directory as app.py"""
    try:
        current_dir = Path(__file__).parent
        sample_path = current_dir / "polymer_production_template.xlsx"
        if sample_path.exists():
            with open(sample_path, "rb") as f:
                return io.BytesIO(f.read())
        else:
            # If absent, create a minimal bytes object so download button still works
            return io.BytesIO(b"")
    except Exception as e:
        st.warning(f"Could not load sample template: {e}. Using empty template.")
        return io.BytesIO(b"")

def process_shutdown_dates(plant_df, dates):
    """Process shutdown dates for each plant"""
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
                shutdown_days = []
                for d, date in enumerate(dates):
                    if start_date <= date <= end_date:
                        shutdown_days.append(d)
                if shutdown_days:
                    shutdown_periods[plant] = shutdown_days
                    st.info(f"🔧 Shutdown scheduled for {plant}: {start_date.strftime('%d-%b-%y')} to {end_date.strftime('%d-%b-%y')} ({len(shutdown_days)} days)")
                else:
                    shutdown_periods[plant] = []
                    st.info(f"ℹ️ Shutdown period for {plant} is outside planning horizon")
            except Exception as e:
                st.warning(f"⚠️ Invalid shutdown dates for {plant}: {e}")
                shutdown_periods[plant] = []
        else:
            shutdown_periods[plant] = []
    return shutdown_periods

# -------------------------
# Page config and session state
# -------------------------
st.set_page_config(
    page_title="Polymer Production Scheduler",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Session state defaults (for smoother transitions)
if 'current_step' not in st.session_state:
    st.session_state.current_step = 0
if 'solutions' not in st.session_state:
    st.session_state.solutions = []
if 'best_solution' not in st.session_state:
    st.session_state.best_solution = None

# -------------------------
# Custom CSS - Light Material + Glassmorphism
# -------------------------
st.markdown(
    """
    <style>
    :root{
      --bg:#ffffff;
      --card:#ffffff;
      --muted:#6b7280;
      --accent1: #4f46e5; /* deep indigo */
      --accent2: #7c3aed; /* purple */
      --glass: rgba(255,255,255,0.55);
      --glass-border: rgba(124,58,237,0.12);
      --deep: #1e293b;
    }
    html, body, [data-testid="stAppViewContainer"]{
      background: linear-gradient(180deg, #fbfdff 0%, #ffffff 40%);
      color: var(--deep);
    }
    /* Top header */
    .app-header {
      display:flex; align-items:center; justify-content:space-between;
      padding: 18px 22px; margin-bottom: 20px;
      background: linear-gradient(135deg, rgba(79,70,229,0.95), rgba(124,58,237,0.95));
      color: white; border-radius: 14px;
      box-shadow: 0 6px 18px rgba(79,70,229,0.15);
    }
    .app-title {font-size:22px; font-weight:700; letter-spacing:0.2px;}
    .app-sub {font-size:13px; opacity:0.95;}
    /* Main white glass card */
    .glass-card {
      background: linear-gradient(180deg, rgba(255,255,255,0.6), rgba(255,255,255,0.48));
      border: 1px solid var(--glass-border);
      backdrop-filter: blur(6px) saturate(120%);
      border-radius: 12px;
      padding: 18px;
      box-shadow: 0 8px 24px rgba(16,24,40,0.05);
    }
    /* Fluent buttons */
    .stButton>button, .primary-btn {
      background: linear-gradient(135deg, var(--accent1), var(--accent2));
      color: white; border-radius: 10px; padding: 8px 14px; font-weight:700;
      border: none; box-shadow: 0 6px 14px rgba(124,58,237,0.12);
    }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 10px 22px rgba(79,70,229,0.14); }
    /* Input panels */
    .panel-title { font-size:16px; font-weight:700; margin-bottom:8px; color:var(--deep); }
    .muted { color: var(--muted); font-size:13px; }
    /* Tabs look */
    .stTabs [data-baseweb="tab-list"] {
      background: transparent; gap: 8px; margin-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] {
      background: rgba(255,255,255,0.9); border: 1px solid #eef2ff; box-shadow: none;
      border-radius: 8px; padding: 8px 14px; font-weight:600; color:var(--deep);
    }
    .stTabs [aria-selected="true"] {
      background: linear-gradient(90deg, rgba(79,70,229,0.08), rgba(124,58,237,0.06));
      border: 1px solid rgba(79,70,229,0.18);
      color: var(--deep);
      box-shadow: 0 6px 18px rgba(79,70,229,0.04);
    }
    /* Dataframe styling */
    .dataframe { border-radius:8px; box-shadow: 0 6px 20px rgba(16,24,40,0.03); }
    /* Metric cards */
    .metric {
      background: linear-gradient(180deg, rgba(124,58,237,0.12), rgba(79,70,229,0.06));
      border-radius: 10px; padding: 12px; text-align:left;
      border: 1px solid rgba(124,58,237,0.06);
    }
    .metric .val { font-weight:800; font-size:20px; color:var(--deep); }
    .metric .lbl { color:var(--muted); font-size:12px; margin-top:4px; }
    /* small helpers */
    .note { font-size:13px; color:var(--muted); }
    .warning { background:#fff7ed; border:1px solid #ffedd5; padding:8px; border-radius:8px; color:#92400e; }
    /* compact progress */
    .progress-inline { display:flex; align-items:center; gap:12px; }
    .kbd { background:#f3f4f6; padding:4px 8px; border-radius:6px; font-weight:600; font-size:12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------
# App header (no sidebar)
# -------------------------
st.markdown(
    f"""
    <div class="app-header">
      <div>
        <div class="app-title">🏭 Polymer Production Scheduler</div>
        <div class="app-sub">Multi-Plant Optimization • Shutdown-aware scheduling • Material Design (light)</div>
      </div>
      <div style="text-align:right;">
        <div style="font-size:12px;color:rgba(255,255,255,0.9);">Streamlined: Upload → Parameters → Optimize → Results</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -------------------------
# Top-level tabs guiding the progressive workflow
# -------------------------
tabs = st.tabs(["1️⃣ Upload & Validate", "2️⃣ Parameters & Optimize", "3️⃣ Results"])
tab_upload, tab_params, tab_results = tabs

# ---------- UPLOAD & VALIDATE ----------
with tab_upload:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">📥 Upload your Excel file</div>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Upload the Excel file containing sheets: <strong>Plant</strong>, <strong>Inventory</strong>, <strong>Demand</strong>. Use the sample template if unclear.</div>', unsafe_allow_html=True)
    col1, col2 = st.columns([3,1])
    with col1:
        uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], key="uploader", help="Drag & drop the Excel file here.")
    with col2:
        sample_workbook = get_sample_workbook()
        st.download_button(
            "📥 Sample template",
            data=sample_workbook,
            file_name="polymer_production_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.markdown("<hr/>", unsafe_allow_html=True)

    # Show progressive validation / file preview only when uploaded
    if uploaded_file:
        # Keep a copy of bytes for multiple reads
        uploaded_file.seek(0)
        excel_bytes = io.BytesIO(uploaded_file.read())

        st.markdown('<div class="panel-title" style="margin-top:6px;">🔍 Data Preview & Quality Checks</div>', unsafe_allow_html=True)
        # Tabbed cards for Plant / Inventory / Demand
        preview_tabs = st.tabs(["🏭 Plant", "📦 Inventory", "📊 Demand", "⚠️ Validation Report"])
        t_plant, t_inv, t_dem, t_val = preview_tabs

        # Common function to render dataframes with safe exceptions:
        def safe_read(sheet):
            try:
                excel_bytes.seek(0)
                df = pd.read_excel(excel_bytes, sheet_name=sheet)
                return df
            except Exception as e:
                return None

        plant_df = safe_read('Plant')
        inventory_df = safe_read('Inventory')
        demand_df = safe_read('Demand')

        # Plant tab
        with t_plant:
            if plant_df is None:
                st.error("Plant sheet not found or could not be read. Ensure a sheet named 'Plant' exists.")
            else:
                # Sanitize date columns for display without modifying underlying df used later
                display_df = plant_df.copy()
                for col in display_df.columns:
                    if pd.api.types.is_datetime64_any_dtype(display_df[col]):
                        display_df[col] = display_df[col].dt.strftime('%d-%b-%y')
                st.dataframe(display_df, use_container_width=True)
                st.markdown('<div class="note">Tip: Shutdown columns will be auto-detected and visualized.</div>', unsafe_allow_html=True)

        # Inventory tab
        with t_inv:
            if inventory_df is None:
                st.error("Inventory sheet not found or could not be read. Ensure a sheet named 'Inventory' exists.")
            else:
                display_df = inventory_df.copy()
                for col in display_df.columns:
                    if pd.api.types.is_datetime64_any_dtype(display_df[col]):
                        display_df[col] = display_df[col].dt.strftime('%d-%b-%y')
                st.dataframe(display_df, use_container_width=True)
                st.markdown('<div class="note">Tip: ' + ('Missing "Lines" values will default to all plants.' if 'Lines' in inventory_df.columns else 'Make sure "Lines" column exists for per-plant settings.') + '</div>', unsafe_allow_html=True)

        # Demand tab
        with t_dem:
            if demand_df is None:
                st.error("Demand sheet not found or could not be read. Ensure a sheet named 'Demand' exists.")
            else:
                display_df = demand_df.copy()
                # Format first column if it's datetime
                first_col = display_df.columns[0]
                if pd.api.types.is_datetime64_any_dtype(display_df[first_col]):
                    display_df[first_col] = display_df[first_col].dt.strftime('%d-%b-%y')
                st.dataframe(display_df, use_container_width=True)
                st.markdown('<div class="note">Tip: First column should be Dates. Other columns are grade demands (names must match Inventory "Grade Name").</div>', unsafe_allow_html=True)

        # Validation report: auto-detect typical issues
        with t_val:
            issues = []
            if plant_df is None:
                issues.append("Missing Plant sheet.")
            else:
                required_plant_cols = ['Plant', 'Capacity per day']
                for c in required_plant_cols:
                    if c not in plant_df.columns:
                        issues.append(f"Plant sheet missing column: {c}")

            if inventory_df is None:
                issues.append("Missing Inventory sheet.")
            else:
                if 'Grade Name' not in inventory_df.columns:
                    issues.append("Inventory sheet missing column: Grade Name")

            if demand_df is None:
                issues.append("Missing Demand sheet.")
            else:
                # check date column type
                date_col = demand_df.columns[0] if demand_df is not None else None
                if date_col is not None:
                    if not pd.api.types.is_datetime64_any_dtype(demand_df[date_col]):
                        issues.append("Demand first column doesn't look like dates. Convert to date format in Excel.")

            if issues:
                st.markdown('<div class="warning"><strong>Validation Issues Found</strong><ul style="margin:6px 0 0 18px;">' + ''.join([f"<li>{i}</li>" for i in issues]) + '</ul></div>', unsafe_allow_html=True)
            else:
                st.success("✅ Basic validation passed. Sheets and columns look good.")
                # Show compact shutdown visualization (compact list)
                if plant_df is not None and 'Shutdown Start Date' in plant_df.columns and 'Shutdown End Date' in plant_df.columns:
                    shutdowns = []
                    for _, r in plant_df.iterrows():
                        if pd.notna(r.get('Shutdown Start Date')) and pd.notna(r.get('Shutdown End Date')):
                            try:
                                a = pd.to_datetime(r['Shutdown Start Date']).date()
                                b = pd.to_datetime(r['Shutdown End Date']).date()
                                shutdowns.append((r['Plant'], a.strftime('%d-%b-%y'), b.strftime('%d-%b-%y')))
                            except:
                                continue
                    if shutdowns:
                        st.markdown("**Shutdowns detected:**")
                        st.table(pd.DataFrame(shutdowns, columns=["Plant", "Start", "End"]))
                    else:
                        st.info("No shutdowns detected in Plant sheet.")

    else:
        st.info("Upload an Excel file to preview and validate data. Use the sample template if you need an example.")

    st.markdown('</div>', unsafe_allow_html=True)

# ---------- PARAMETERS & OPTIMIZE ----------
with tab_params:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">⚙️ Parameters & Run Optimization</div>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Configure solver time limit, buffer days and objective weights. Most parameters are progressive and validated where possible.</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)

    # Layout: left column for core params, right column for advanced & quick actions
    left, right = st.columns([2, 1])
    with left:
        with st.expander("🔧 Basic Parameters (recommended defaults)", expanded=True):
            time_limit_min = st.number_input("Time limit (minutes)", min_value=1, max_value=120, value=10, help="Max time to run solver", key="time_limit")
            buffer_days = st.number_input("Buffer days", min_value=0, max_value=14, value=3, help="Days appended to planning horizon", key="buffer_days")
        with st.expander("🎯 Objective weights", expanded=False):
            stockout_penalty = st.number_input("Stockout penalty", min_value=1, value=10, help="Penalty weight for stockouts", key="stockout_penalty")
            transition_penalty = st.number_input("Transition penalty", min_value=1, value=10, help="Penalty weight for transitions", key="transition_penalty")
            continuity_bonus = st.number_input("Continuity bonus", min_value=0, value=1, help="Bonus for continuing same grade", key="continuity_bonus")

        st.markdown("<div style='margin-top:10px; display:flex; gap:10px;'>", unsafe_allow_html=True)
        run_col1, run_col2 = st.columns([1,1])
        with run_col1:
            run_button = st.button("🎯 Run Optimization", use_container_width=True, key="run_opt")
        with run_col2:
            quick_reset = st.button("↺ Reset Session", use_container_width=True, key="reset_session")

    with right:
        st.markdown("<div class='panel-title'>Status</div>", unsafe_allow_html=True)
        # Compact status box
        st.markdown('<div class="metric"><div class="val">Step: ' + str(st.session_state.current_step) + '</div><div class="lbl">Current workflow step</div></div>', unsafe_allow_html=True)
        st.markdown("<br/>", unsafe_allow_html=True)
        st.markdown("<div class='panel-title'>Shortcuts</div>", unsafe_allow_html=True)
        st.markdown("<div class='note'>Keyboard: <span class='kbd'>Shift + Enter</span> to run when focus is on inputs. Hover hints on charts and tables.</div>", unsafe_allow_html=True)
        st.markdown("<br/>", unsafe_allow_html=True)
        st.markdown("<div class='panel-title'>Tips</div>", unsafe_allow_html=True)
        st.markdown("<ul class='muted'><li>Use the Upload tab to inspect shutdowns before running.</li><li>Reduce time limit for quick test runs.</li></ul>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # Handle reset
    if quick_reset:
        st.session_state.current_step = 0
        st.session_state.solutions = []
        st.session_state.best_solution = None
        st.rerun()

    # Run optimization when requested
    if run_button:
        if 'uploader' in st.session_state and st.session_state.uploader is not None:
            uploaded_file = st.session_state.uploader
        else:
            uploaded_file = None

        # if no file uploaded in this session try to read the widget's file
        try:
            # Prefer reading widget directly (works immediately in the same run)
            uploaded_file = st.session_state.get("uploader", None)
        except Exception:
            uploaded_file = None

        # As a fallback attempt to read `uploaded_file` from the earlier upload widget in the upload tab
        if uploaded_file is None and 'uploader' in st.session_state:
            uploaded_file = st.session_state.uploader

        # local variable: if upload element exists in this run, we also have it in the upload tab
        widget_file = st.session_state.get("uploader", None)
        if widget_file is not None:
            uploaded_file = widget_file

        # If still None, try re-reading the widget (some Streamlit versions keep it)
        if uploaded_file is None:
            uploaded_file = st.file_uploader("Upload Excel (.xlsx) to run optimization", type=["xlsx"])

        if uploaded_file is None:
            st.error("Please upload an Excel file in the Upload tab before running optimization.")
        else:
            # Set session step
            st.session_state.current_step = 2

            # Keep bytes copy
            uploaded_file.seek(0)
            excel_file = io.BytesIO(uploaded_file.read())

            # Show a slim progress & status area
            progress_bar = st.progress(0)
            status_col = st.empty()
            results_area = st.empty()

            status_col.markdown('<div class="note">📄 Preprocessing data...</div>', unsafe_allow_html=True)
            time.sleep(0.5)
            progress_bar.progress(5)

            try:
                # === Begin: Data preprocessing & model building ===
                # This block is intentionally preserved from original app logic with identical variable names and constraints.
                # It only moves into the new UI container.
                excel_file.seek(0)
                plant_df = pd.read_excel(excel_file, sheet_name='Plant')
                excel_file.seek(0)
                inventory_df = pd.read_excel(excel_file, sheet_name='Inventory')
                excel_file.seek(0)
                demand_df = pd.read_excel(excel_file, sheet_name='Demand')

                # Basic derived info
                num_lines = len(plant_df)
                lines = list(plant_df['Plant'])
                capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
                grades = [col for col in demand_df.columns if col != demand_df.columns[0]]

                # Prepare inventory dictionaries (same logic)
                initial_inventory = {}
                min_inventory = {}
                max_inventory = {}
                min_closing_inventory = {}
                min_run_days = {}
                max_run_days = {}
                force_start_date = {}
                allowed_lines = {grade: [] for grade in grades}
                rerun_allowed = {}
                grade_inventory_defined = set()

                for index, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    lines_value = row.get('Lines', None)
                    if pd.notna(lines_value) and lines_value != '':
                        plants_for_row = [x.strip() for x in str(lines_value).split(',')]
                    else:
                        plants_for_row = lines
                        # Warning shown earlier in old app; keep it as info here
                        st.warning(f"⚠️ Lines for grade '{grade}' (row {index}) are not specified; allowing all lines.")

                    for plant in plants_for_row:
                        if plant not in allowed_lines.get(grade, []):
                            allowed_lines[grade].append(plant)

                    if grade not in grade_inventory_defined:
                        initial_inventory[grade] = row['Opening Inventory'] if pd.notna(row.get('Opening Inventory')) else 0
                        min_inventory[grade] = row['Min. Inventory'] if pd.notna(row.get('Min. Inventory')) else 0
                        max_inventory[grade] = row['Max. Inventory'] if pd.notna(row.get('Max. Inventory')) else 1000000000
                        min_closing_inventory[grade] = row['Min. Closing Inventory'] if pd.notna(row.get('Min. Closing Inventory')) else 0
                        grade_inventory_defined.add(grade)

                    for plant in plants_for_row:
                        grade_plant_key = (grade, plant)
                        min_run_days[grade_plant_key] = int(row['Min. Run Days']) if pd.notna(row.get('Min. Run Days')) else 1
                        max_run_days[grade_plant_key] = int(row['Max. Run Days']) if pd.notna(row.get('Max. Run Days')) else 9999
                        if pd.notna(row.get('Force Start Date')):
                            try:
                                force_start_date[grade_plant_key] = pd.to_datetime(row['Force Start Date']).date()
                            except:
                                force_start_date[grade_plant_key] = None
                                st.warning(f"⚠️ Invalid Force Start Date for grade '{grade}' on plant '{plant}'")
                        else:
                            force_start_date[grade_plant_key] = None
                        rerun_val = row.get('Rerun Allowed', None)
                        if pd.notna(rerun_val):
                            val_str = str(rerun_val).strip().lower()
                            if val_str in ['no', 'n', 'false', '0']:
                                rerun_allowed[grade_plant_key] = False
                            else:
                                rerun_allowed[grade_plant_key] = True
                        else:
                            rerun_allowed[grade_plant_key] = True

                material_running_info = {}
                for index, row in plant_df.iterrows():
                    plant = row['Plant']
                    material = row.get('Material Running')
                    expected_days = row.get('Expected Run Days')
                    if pd.notna(material) and pd.notna(expected_days):
                        try:
                            material_running_info[plant] = (str(material).strip(), int(expected_days))
                        except:
                            st.warning(f"⚠️ Invalid Material Running or Expected Run Days for plant '{plant}'")

            except Exception as e:
                st.error(f"Error in preprocessing: {e}")
                import traceback
                st.error(traceback.format_exc())
                progress_bar.progress(0)
                st.stop()

            progress_bar.progress(20)
            status_col.markdown('<div class="note">🔧 Building optimization model...</div>', unsafe_allow_html=True)
            time.sleep(0.6)

            # Prepare demand dates and horizon (identical logic)
            demand_dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
            num_days = len(demand_dates)
            last_date = demand_dates[-1]
            for i in range(1, buffer_days + 1):
                demand_dates.append(last_date + timedelta(days=i))
            num_days = len(demand_dates)
            formatted_dates = [date.strftime('%d-%b-%y') for date in demand_dates]

            demand_data = {}
            for grade in grades:
                if grade in demand_df.columns:
                    demand_data[grade] = {demand_df.iloc[i, 0].date(): demand_df[grade].iloc[i] for i in range(len(demand_df))}
                else:
                    demand_data[grade] = {date: 0 for date in demand_dates}
            for grade in grades:
                for date in demand_dates[-buffer_days:]:
                    if date not in demand_data[grade]:
                        demand_data[grade][date] = 0

            shutdown_periods = process_shutdown_dates(plant_df, demand_dates)

            # Load transition sheets (same mechanism)
            transition_dfs = {}
            for i in range(len(plant_df)):
                plant_name = plant_df['Plant'].iloc[i]
                possible_sheet_names = [
                    f'Transition_{plant_name}',
                    f'Transition_{plant_name.replace(" ", "_")}',
                    f'Transition{plant_name.replace(" ", "")}',
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
                    transition_dfs[plant_name] = None
                    st.info(f"ℹ️ No transition matrix found for {plant_name}. Assuming no transition constraints.")

            progress_bar.progress(35)
            status_col.markdown('<div class="note">⚡ Formulating CP-SAT model...</div>', unsafe_allow_html=True)
            time.sleep(0.6)

            # ----------------
            # Build model (logic kept identical to original)
            # ----------------
            model = cp_model.CpModel()
            is_producing = {}
            production = {}

            def is_allowed_combination(grade, line):
                return line in allowed_lines.get(grade, [])

            for grade in grades:
                for line in allowed_lines[grade]:
                    for d in range(num_days):
                        key = (grade, line, d)
                        is_producing[key] = model.NewBoolVar(f'is_producing_{grade}_{line}_{d}')
                        if d < num_days - buffer_days:
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value == capacities[line]).OnlyEnforceIf(is_producing[key])
                            model.Add(production_value == 0).OnlyEnforceIf(is_producing[key].Not())
                        else:
                            production_value = model.NewIntVar(0, capacities[line], f'production_{grade}_{line}_{d}')
                            model.Add(production_value <= capacities[line] * is_producing[key])
                        production[key] = production_value

            def get_production_var(grade, line, d):
                key = (grade, line, d)
                if key not in production:
                    return 0
                return production[key]

            def get_is_producing_var(grade, line, d):
                key = (grade, line, d)
                if key not in is_producing:
                    return None
                return is_producing[key]

            # SHUTDOWN constraints
            for line in lines:
                if line in shutdown_periods and shutdown_periods[line]:
                    for d in shutdown_periods[line]:
                        for grade in grades:
                            if is_allowed_combination(grade, line):
                                key = (grade, line, d)
                                if key in is_producing:
                                    model.Add(is_producing[key] == 0)
                                    model.Add(production[key] == 0)

            # shutdown demand check
            shutdown_demand = {}
            for grade in grades:
                shutdown_demand[grade] = 0
                for line in allowed_lines[grade]:
                    if line in shutdown_periods:
                        for d in shutdown_periods[line]:
                            shutdown_demand[grade] += demand_data[grade].get(demand_dates[d], 0)
            for grade, total_shutdown_demand in shutdown_demand.items():
                if total_shutdown_demand > initial_inventory[grade]:
                    st.warning(f"⚠️ Grade '{grade}': Shutdown periods require {total_shutdown_demand} MT from inventory (current: {initial_inventory[grade]} MT).")

            inventory_vars = {}
            for grade in grades:
                for d in range(num_days + 1):
                    inventory_vars[(grade, d)] = model.NewIntVar(0, 100000, f'inventory_{grade}_{d}')

            stockout_vars = {}
            for grade in grades:
                for d in range(num_days):
                    stockout_vars[(grade, d)] = model.NewIntVar(0, 100000, f'stockout_{grade}_{d}')

            for line in lines:
                for d in range(num_days):
                    producing_vars = []
                    for grade in grades:
                        if is_allowed_combination(grade, line):
                            var = get_is_producing_var(grade, line, d)
                            if var is not None:
                                producing_vars.append(var)
                    if producing_vars:
                        model.Add(sum(producing_vars) <= 1)

            for plant, (material, expected_days) in material_running_info.items():
                for d in range(min(expected_days, num_days)):
                    if is_allowed_combination(material, plant):
                        model.Add(get_is_producing_var(material, plant, d) == 1)
                        for other_material in grades:
                            if other_material != material and is_allowed_combination(other_material, plant):
                                model.Add(get_is_producing_var(other_material, plant, d) == 0)

            objective = 0

            # inventory balance & stockout corrected
            for grade in grades:
                model.Add(inventory_vars[(grade, 0)] == initial_inventory[grade])

            for grade in grades:
                for d in range(num_days):
                    produced_today = sum(
                        get_production_var(grade, line, d)
                        for line in allowed_lines[grade]
                    )
                    demand_today = demand_data[grade].get(demand_dates[d], 0)
                    supplied = model.NewIntVar(0, 100000, f'supplied_{grade}_{d}')
                    model.Add(supplied <= inventory_vars[(grade, d)] + produced_today)
                    model.Add(supplied <= demand_today)
                    model.Add(stockout_vars[(grade, d)] == demand_today - supplied)
                    model.Add(inventory_vars[(grade, d + 1)] == inventory_vars[(grade, d)] + produced_today - supplied)
                    model.Add(inventory_vars[(grade, d + 1)] >= 0)

            for grade in grades:
                for d in range(num_days):
                    if min_inventory[grade] > 0:
                        min_inv_value = int(min_inventory[grade])
                        inventory_tomorrow = inventory_vars[(grade, d + 1)]
                        deficit = model.NewIntVar(0, 100000, f'deficit_{grade}_{d}')
                        model.Add(deficit >= min_inv_value - inventory_tomorrow)
                        model.Add(deficit >= 0)
                        objective += stockout_penalty * deficit

            for grade in grades:
                closing_inventory = inventory_vars[(grade, num_days - buffer_days)]
                min_closing = min_closing_inventory[grade]
                if min_closing > 0:
                    closing_deficit = model.NewIntVar(0, 100000, f'closing_deficit_{grade}')
                    model.Add(closing_deficit >= min_closing - closing_inventory)
                    model.Add(closing_deficit >= 0)
                    objective += stockout_penalty * closing_deficit * 3

            for grade in grades:
                for d in range(1, num_days + 1):
                    model.Add(inventory_vars[(grade, d)] <= max_inventory[grade])

            for line in lines:
                for d in range(num_days - buffer_days):
                    if line in shutdown_periods and d in shutdown_periods[line]:
                        continue
                    production_vars = [
                        get_production_var(grade, line, d)
                        for grade in grades
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:
                        model.Add(sum(production_vars) == capacities[line])

                for d in range(num_days - buffer_days, num_days):
                    production_vars = [
                        get_production_var(grade, line, d)
                        for grade in grades
                        if is_allowed_combination(grade, line)
                    ]
                    if production_vars:
                        model.Add(sum(production_vars) <= capacities[line])

            for grade_plant_key, start_date in force_start_date.items():
                if start_date:
                    grade, plant = grade_plant_key
                    try:
                        start_day_index = demand_dates.index(start_date)
                        var = get_is_producing_var(grade, plant, start_day_index)
                        if var is not None:
                            model.Add(var == 1)
                            st.info(f"✅ Enforced force start date for grade '{grade}' on plant '{plant}' at day {start_date.strftime('%d-%b-%y')}")
                        else:
                            st.warning(f"⚠️ Cannot enforce force start date for grade '{grade}' on plant '{plant}' - combination not allowed")
                    except ValueError:
                        st.warning(f"⚠️ Force start date '{start_date.strftime('%d-%b-%y')}' for grade '{grade}' on plant '{plant}' not found in demand dates")

            # Minimum & Maximum Run Days (same approach)
            is_start_vars = {}
            run_end_vars = {}
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    min_run = min_run_days.get(grade_plant_key, 1)
                    max_run = max_run_days.get(grade_plant_key, 9999)
                    for d in range(num_days):
                        is_start = model.NewBoolVar(f'start_{grade}_{line}_{d}')
                        is_start_vars[(grade, line, d)] = is_start
                        is_end = model.NewBoolVar(f'end_{grade}_{line}_{d}')
                        run_end_vars[(grade, line, d)] = is_end
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
                        if d < num_days - 1:
                            next_prod = get_is_producing_var(grade, line, d + 1)
                            if current_prod is not None and next_prod is not None:
                                model.AddBoolAnd([current_prod, next_prod.Not()]).OnlyEnforceIf(is_end)
                                model.AddBoolOr([current_prod.Not(), next_prod]).OnlyEnforceIf(is_end.Not())
                        else:
                            if current_prod is not None:
                                model.Add(current_prod == 1).OnlyEnforceIf(is_end)
                                model.Add(is_end == 1).OnlyEnforceIf(current_prod)

                    # MIN run enforcement if consecutive non-shutdown days exist
                    for d in range(num_days):
                        is_start = is_start_vars[(grade, line, d)]
                        max_possible_run = 0
                        for k in range(min_run):
                            if d + k < num_days:
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    break
                                max_possible_run += 1
                        if max_possible_run >= min_run:
                            for k in range(min_run):
                                if d + k < num_days:
                                    if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                        continue
                                    future_prod = get_is_producing_var(grade, line, d + k)
                                    if future_prod is not None:
                                        model.Add(future_prod == 1).OnlyEnforceIf(is_start)

                    # MAX run (sliding window)
                    for d in range(num_days - max_run):
                        consecutive_days = []
                        for k in range(max_run + 1):
                            if d + k < num_days:
                                if line in shutdown_periods and (d + k) in shutdown_periods[line]:
                                    break
                                prod_var = get_is_producing_var(grade, line, d + k)
                                if prod_var is not None:
                                    consecutive_days.append(prod_var)
                        if len(consecutive_days) == max_run + 1:
                            model.Add(sum(consecutive_days) <= max_run)

            # Transition rules
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

            # Rerun not allowed
            for grade in grades:
                for line in allowed_lines[grade]:
                    grade_plant_key = (grade, line)
                    if not rerun_allowed.get(grade_plant_key, True):
                        starts = [is_start_vars[(grade, line, d)] for d in range(num_days) if (grade, line, d) in is_start_vars]
                        if starts:
                            model.Add(sum(starts) <= 1)

            # Stockout penalties & transitions (objective)
            for grade in grades:
                for d in range(num_days):
                    objective += stockout_penalty * stockout_vars[(grade, d)]

            for line in lines:
                for d in range(num_days - 1):
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
                            objective += transition_penalty * trans_var

                    for grade in grades:
                        if line in allowed_lines[grade]:
                            continuity = model.NewBoolVar(f'continuity_{line}_{d}_{grade}')
                            model.AddBoolAnd([is_producing[(grade, line, d)], is_producing[(grade, line, d + 1)]]).OnlyEnforceIf(continuity)
                            objective += -continuity_bonus * continuity

            model.Minimize(objective)

            progress_bar.progress(60)
            status_col.markdown('<div class="note">🧠 Solving with OR-Tools CP-SAT (this may take a few moments)...</div>', unsafe_allow_html=True)

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = time_limit_min * 60.0
            solver.parameters.num_search_workers = 8
            solver.parameters.random_seed = 42
            solver.parameters.log_search_progress = True

            # Callback class (kept identical)
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
                    solution = {
                        'objective': current_obj,
                        'time': current_time,
                        'production': {},
                        'inventory': {},
                        'stockout': {},
                        'is_producing': {}
                    }
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
                    for grade in self.grades:
                        solution['inventory'][grade] = {}
                        for d in range(self.num_days + 1):
                            key = (grade, d)
                            if key in self.inventory:
                                if d < self.num_days:
                                    solution['inventory'][grade][self.formatted_dates[d] if d > 0 else 'initial'] = self.Value(self.inventory[key])
                                else:
                                    solution['inventory'][grade]['final'] = self.Value(self.inventory[key])
                    for grade in self.grades:
                        solution['stockout'][grade] = {}
                        for d in range(self.num_days):
                            key = (grade, d)
                            if key in self.stockout:
                                value = self.Value(self.stockout[key])
                                if value > 0:
                                    solution['stockout'][grade][self.formatted_dates[d]] = value
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
                    # Count transitions
                    transition_count_per_line = {line: 0 for line in self.lines}
                    total_transitions = 0
                    for line in self.lines:
                        last_grade = None
                        for d in range(self.num_days):
                            current_grade = None
                            for grade in self.grades:
                                key = (grade, line, d)
                                if key in self.is_producing and self.Value(self.is_producing[key]) == 1:
                                    current_grade = grade
                                    break
                            if current_grade is not None:
                                if last_grade is not None and current_grade != last_grade:
                                    transition_count_per_line[line] += 1
                                    total_transitions += 1
                                last_grade = current_grade
                    solution['transitions'] = {
                        'per_line': transition_count_per_line,
                        'total': total_transitions
                    }
                    self.solutions.append(solution)

                def num_solutions(self):
                    return len(self.solutions)

            callback = SolutionCallback(production, inventory_vars, stockout_vars, is_producing, grades, lines, demand_dates, formatted_dates, num_days)

            start_time = time.time()
            status = solver.Solve(model, callback)
            elapsed = time.time() - start_time

            progress_bar.progress(100)
            # Status messaging (kept similar)
            if status == cp_model.OPTIMAL:
                status_col.markdown('<div class="note">✅ Optimization completed optimally.</div>', unsafe_allow_html=True)
            elif status == cp_model.FEASIBLE:
                status_col.markdown('<div class="note">✅ Found feasible solution (not proven optimal).</div>', unsafe_allow_html=True)
            else:
                status_col.markdown('<div class="warning">⚠️ Solver ended without a feasible solution.</div>', unsafe_allow_html=True)

            # Save results to session
            st.session_state.solutions = callback.solutions
            if callback.num_solutions() > 0:
                st.session_state.best_solution = callback.solutions[-1]
            else:
                st.session_state.best_solution = None

            # Move user to Results tab programmatically by setting step and rendering results (they can manually click too)
            st.session_state.current_step = 3
            st.success(f"Solver finished in {elapsed:.1f}s. Go to the Results tab to view outputs.")
            st.rerun()

# ---------- RESULTS ----------
with tab_results:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">📈 Results Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="muted">Summary metrics, production Gantt per line and inventory charts are presented below. All Plotly visuals are preserved from the original logic.</div>', unsafe_allow_html=True)
    st.markdown("<br/>", unsafe_allow_html=True)

    # If no solution in session, show helpful guidance
    if st.session_state.best_solution is None:
        st.info("No solution available. Run optimization from the Parameters tab after uploading a valid Excel file.")
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        best_solution = st.session_state.best_solution

        # Metrics row
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f"<div class='metric'><div class='val'>{best_solution['objective']:,}</div><div class='lbl'>Objective Value (lower better)</div></div>", unsafe_allow_html=True)
        with m2:
            st.markdown(f"<div class='metric'><div class='val'>{best_solution['transitions']['total']}</div><div class='lbl'>Total Transitions</div></div>", unsafe_allow_html=True)
        with m3:
            total_stockouts = sum(sum(best_solution['stockout'][g].values()) if isinstance(best_solution['stockout'][g], dict) else 0 for g in best_solution['stockout'])
            st.markdown(f"<div class='metric'><div class='val'>{total_stockouts:,} MT</div><div class='lbl'>Total Stockouts</div></div>", unsafe_allow_html=True)
        with m4:
            # derive planning horizon if earlier available in session - fallback
            horizon_days = None
            # Not always stored, attempt to derive from a solution inventory dict
            try:
                some_grade = next(iter(best_solution['inventory']))
                horizon_days = len(best_solution['inventory'][some_grade]) - 1
            except Exception:
                horizon_days = "N/A"
            st.markdown(f"<div class='metric'><div class='val'>{horizon_days}</div><div class='lbl'>Planning Horizon (days)</div></div>", unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)

        # Tabs: schedule, inventory, summary (keeps plotly code intact)
        res_tab1, res_tab2, res_tab3 = st.tabs(["📅 Schedule", "📦 Inventory", "📊 Summary"])

        # Rebuild some local state to run the exact same plotting code:
        # We attempt to re-read the last uploaded file to reconstruct plant/grades/capacities/dates used.
        # If not available, try to use variables stored earlier (best-effort).
        uploaded_file_bytes = None
        try:
            # attempt to read from the upload widget's state if present
            if 'uploader' in st.session_state and st.session_state.uploader is not None:
                uf = st.session_state.uploader
                uf.seek(0)
                uploaded_file_bytes = io.BytesIO(uf.read())
            else:
                # try to read the uploader widget directly (some versions keep it available)
                uf2 = st.session_state.get("uploader", None)
                if uf2:
                    uf2.seek(0)
                    uploaded_file_bytes = io.BytesIO(uf2.read())
        except Exception:
            uploaded_file_bytes = None

        # Fallback: if we cannot reconstruct all original inputs, we will reuse the best_solution structure to render schedule tables and charts.
        try:
            if uploaded_file_bytes:
                uploaded_file_bytes.seek(0)
                plant_df = pd.read_excel(uploaded_file_bytes, sheet_name='Plant')
                uploaded_file_bytes.seek(0)
                demand_df = pd.read_excel(uploaded_file_bytes, sheet_name='Demand')
                uploaded_file_bytes.seek(0)
                inventory_df = pd.read_excel(uploaded_file_bytes, sheet_name='Inventory')

                # Recreate supporting structures
                lines = list(plant_df['Plant'])
                capacities = {row['Plant']: row['Capacity per day'] for index, row in plant_df.iterrows()}
                grades = [col for col in demand_df.columns if col != demand_df.columns[0]]
                demand_dates = sorted(list(set(demand_df.iloc[:, 0].dt.date.tolist())))
                # buffer days: attempt to infer from session or default 3
                buffer_days = st.session_state.get("buffer_days", 3)
                last_date = demand_dates[-1]
                for i in range(1, buffer_days + 1):
                    demand_dates.append(last_date + timedelta(days=i))
                num_days = len(demand_dates)
                formatted_dates = [d.strftime('%d-%b-%y') for d in demand_dates]

                # Allowed lines mapping (from inventory)
                allowed_lines = {grade: [] for grade in grades}
                for index, row in inventory_df.iterrows():
                    grade = row['Grade Name']
                    lines_value = row.get('Lines', None)
                    if pd.notna(lines_value) and lines_value != '':
                        plants_for_row = [x.strip() for x in str(lines_value).split(',')]
                    else:
                        plants_for_row = lines
                    for pl in plants_for_row:
                        if pl not in allowed_lines[grade]:
                            allowed_lines[grade].append(pl)

                # Recreate shutdown periods if possible
                shutdown_periods = process_shutdown_dates(plant_df, demand_dates)
            else:
                # If we cannot re-read, attempt to infer from best_solution keys
                # Build lines and dates from best_solution content
                # Lines:
                lines = list(best_solution['is_producing'].keys())
                # Dates:
                # get first grade dict to find dates
                some_line = lines[0]
                dates_list = list(best_solution['is_producing'][some_line].keys())
                formatted_dates = dates_list
                # try to parse to datetime.date if possible
                try:
                    demand_dates = [pd.to_datetime(d).date() for d in dates_list]
                except:
                    demand_dates = [pd.to_datetime(d, format='%d-%b-%y').date() for d in dates_list]
                num_days = len(demand_dates)
                # Grades:
                grades = list(best_solution['production'].keys())
                # capacities unknown; set placeholder mapping if needed
                capacities = {line: 0 for line in lines}
                allowed_lines = {g: lines for g in grades}
                shutdown_periods = {}
        except Exception as e:
            st.warning(f"Could not reconstruct metadata for plotting from upload. Some charts may be approximate. ({e})")
            # minimal fallback
            lines = list(best_solution['is_producing'].keys())
            dates_list = list(next(iter(best_solution['is_producing'].values())).keys())
            formatted_dates = dates_list
            try:
                demand_dates = [pd.to_datetime(d).date() for d in dates_list]
            except:
                demand_dates = [pd.to_datetime(d, format='%d-%b-%y').date() for d in dates_list]
            num_days = len(demand_dates)
            grades = list(best_solution['production'].keys())
            capacities = {line: 0 for line in lines}
            allowed_lines = {g: lines for g in grades}
            shutdown_periods = {}

        # ---------- Schedule tab (kept Plotly render logic)
        with res_tab1:
            st.subheader("Production Schedule (Gantt & Line-wise)")
            # Use same plotting approach as original: timeline per line
            sorted_grades = sorted(grades) if grades else []
            base_colors = px.colors.qualitative.Vivid
            grade_color_map = {grade: base_colors[i % len(base_colors)] for i, grade in enumerate(sorted_grades)}

            for line in lines:
                st.markdown(f"### 🏭 {line}")
                # Build gantt_data from best_solution is_producing mapping
                gantt_data = []
                # If we have best_solution['is_producing'][line] keyed by date strings:
                try:
                    for date_str, grade in best_solution['is_producing'][line].items():
                        # date string to date object
                        try:
                            date_obj = pd.to_datetime(date_str).date()
                        except:
                            date_obj = pd.to_datetime(date_str, format='%d-%b-%y').date()
                        if grade:
                            gantt_data.append({"Grade": grade, "Start": date_obj, "Finish": date_obj + timedelta(days=1), "Line": line})
                except Exception:
                    # fallback: use production entries
                    for grade in best_solution['production']:
                        for dstr in best_solution['production'][grade].keys():
                            try:
                                d_obj = pd.to_datetime(dstr).date()
                            except:
                                d_obj = pd.to_datetime(dstr, format='%d-%b-%y').date()
                            if best_solution['production'][grade].get(dstr, 0) > 0:
                                gantt_data.append({"Grade": grade, "Start": d_obj, "Finish": d_obj + timedelta(days=1), "Line": line})

                if not gantt_data:
                    st.info(f"No production data available for {line}.")
                    continue

                gantt_df = pd.DataFrame(gantt_data)
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

                # shutdown visualization if available
                if line in shutdown_periods and shutdown_periods[line]:
                    shutdown_days = shutdown_periods[line]
                    try:
                        start_shutdown = demand_dates[shutdown_days[0]]
                        end_shutdown = demand_dates[shutdown_days[-1]] + timedelta(days=1)
                        fig.add_vrect(
                            x0=start_shutdown, x1=end_shutdown,
                            fillcolor="red", opacity=0.15, layer="below", line_width=0,
                            annotation_text="Shutdown", annotation_position="top left", annotation_font_color="red"
                        )
                    except Exception:
                        pass

                fig.update_yaxes(autorange="reversed", title=None, showgrid=True, gridcolor="lightgray", gridwidth=1)
                fig.update_xaxes(title="Date", showgrid=True, gridcolor="lightgray", gridwidth=1, tickvals=demand_dates if 'demand_dates' in locals() else None, tickformat="%d-%b", dtick="D1")
                fig.update_layout(height=340, bargap=0.15, margin=dict(l=60, r=160, t=60, b=60), plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("Line-wise Compressed Schedule Table")
            for line in lines:
                st.markdown(f"**{line}**")
                schedule_data = []
                current_grade = None
                start_day = None
                for i, dstr in enumerate(formatted_dates):
                    grade_today = None
                    try:
                        grade_today = best_solution['is_producing'][line].get(dstr, None)
                    except:
                        # try different key access pattern
                        grade_today = None
                    if grade_today != current_grade:
                        if current_grade is not None:
                            end_date = demand_dates[i - 1]
                            duration = (end_date - start_day).days + 1
                            schedule_data.append({"Grade": current_grade, "Start Date": start_day.strftime("%d-%b-%y"), "End Date": end_date.strftime("%d-%b-%y"), "Days": duration})
                        current_grade = grade_today
                        start_day = demand_dates[i]
                if current_grade is not None and start_day is not None:
                    end_date = demand_dates[-1]
                    duration = (end_date - start_day).days + 1
                    schedule_data.append({"Grade": current_grade, "Start Date": start_day.strftime("%d-%b-%y"), "End Date": end_date.strftime("%d-%b-%y"), "Days": duration})
                if schedule_data:
                    st.dataframe(pd.DataFrame(schedule_data), use_container_width=True)
                else:
                    st.info("No schedule segments for this line.")

        # ---------- Inventory tab (kept Plotly inventory plots)
        with res_tab2:
            st.subheader("Inventory Levels (per Grade)")
            try:
                for grade in sorted(grades):
                    # try to extract inventory time series from best_solution
                    inv_dict = best_solution['inventory'].get(grade, {})
                    # Build list of values aligned with formatted_dates
                    inv_vals = []
                    for d in formatted_dates:
                        if d in inv_dict:
                            inv_vals.append(inv_dict[d])
                        else:
                            # fallback to initial / zeros
                            if d == formatted_dates[0]:
                                inv_vals.append(inv_dict.get('initial', 0))
                            else:
                                inv_vals.append(inv_vals[-1] if inv_vals else 0)
                    # Plotly line chart
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=demand_dates, y=inv_vals, mode="lines+markers", name=grade, line=dict(width=3), marker=dict(size=6), hovertemplate="Date: %{x|%d-%b-%y}<br>Inventory: %{y:.0f} MT<extra></extra>"))
                    # add min/max lines if we can infer from inventory_df
                    if 'inventory_df' in locals() and not inventory_df.empty:
                        try:
                            min_inv_val = int(inventory_df[inventory_df['Grade Name']==grade]['Min. Inventory'].iloc[0])
                            max_inv_val = int(inventory_df[inventory_df['Grade Name']==grade]['Max. Inventory'].iloc[0])
                            fig.add_hline(y=min_inv_val, line=dict(color="red", width=2, dash="dash"), annotation_text=f"Min: {min_inv_val}", annotation_position="top left")
                            fig.add_hline(y=max_inv_val, line=dict(color="green", width=2, dash="dash"), annotation_text=f"Max: {max_inv_val}", annotation_position="bottom left")
                        except Exception:
                            pass
                    fig.update_layout(title=f"Inventory Level - {grade}", xaxis=dict(title="Date", tickformat="%d-%b"), yaxis=dict(title="Inventory (MT)"), height=420, plot_bgcolor="white", paper_bgcolor="white")
                    st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.warning(f"Could not render inventory charts: {e}")

        # ---------- Summary tab (production totals table)
        with res_tab3:
            st.subheader("Production Summary")
            try:
                # Build totals similar to original
                production_totals = {}
                grade_totals = {}
                plant_totals = {line: 0 for line in lines}
                stockout_totals = {}

                for grade in grades:
                    production_totals[grade] = {}
                    grade_totals[grade] = 0
                    stockout_totals[grade] = 0
                    for line in lines:
                        total_prod = 0
                        # find in best_solution's production map if available (it stores per-date numbers)
                        for dstr, val in best_solution['production'].get(grade, {}).items():
                            total_prod += val if isinstance(val, (int, float)) else 0
                        # assign same total to all lines if we have no breakdown (best-effort)
                        production_totals[grade][line] = total_prod if len(lines)==1 else 0
                        grade_totals[grade] += production_totals[grade][line]
                        plant_totals[line] += production_totals[grade][line]
                    # stockout totals
                    for dstr, v in best_solution['stockout'].get(grade, {}).items():
                        stockout_totals[grade] += v

                total_prod_data = []
                for grade in grades:
                    row = {'Grade': grade}
                    for line in lines:
                        row[line] = production_totals[grade].get(line, 0)
                    row['Total Produced'] = grade_totals[grade]
                    row['Total Stockout'] = stockout_totals.get(grade, 0)
                    total_prod_data.append(row)
                totals_row = {'Grade': 'Total'}
                totals_row.update({line: plant_totals[line] for line in lines})
                totals_row['Total Produced'] = sum(plant_totals.values())
                totals_row['Total Stockout'] = sum(stockout_totals.values())
                total_prod_data.append(totals_row)
                total_prod_df = pd.DataFrame(total_prod_data)
                st.dataframe(total_prod_df, use_container_width=True)
            except Exception as e:
                st.warning(f"Could not build production summary: {e}")

        st.markdown("</div>", unsafe_allow_html=True)

# Footer
st.markdown("<hr/>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center;color:var(--muted);font-size:13px;padding:8px 0;'>Polymer Production Scheduler • Material Design (Light) • Preserves solver & visualization logic</div>", unsafe_allow_html=True)
