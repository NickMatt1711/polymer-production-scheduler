#!/usr/bin/env python3
"""
test.py

Optimized CP-SAT production scheduler adapted to polymer_production_template.xlsx.

Key behavior (as requested):
- Option A capacity logic (one grade per plant per day; capacity is per-plant-per-day).
- MaterialRunning + Expected Run Days: enforced at horizon start (hard), skipping shutdown days.
- Transition matrices: "Yes" means allowed, "No" means forbidden (hard). Allowed transitions that change grade count as transitions and are penalized.
- Min Inventory (daily) and Min Closing Inventory (final day): SOFT constraints — shortages are penalized.
- Grade-as-integer encoding (one IntVar per plant-day), variables only created on feasible days.
- Enforces inventory balance, demand satisfaction (with stockout variables), production linking, min/max run, and run continuity constraints.
- Output: schedule CSV and inventory CSV.

Usage:
    python test.py --input polymer_production_template.xlsx --time_limit 5 --output_dir ./out
"""

import os
import sys
import argparse
import time
import math
from collections import defaultdict

import pandas as pd
import numpy as np
from ortools.sat.python import cp_model

# -------------------------
# Helper: read Excel sheets
# -------------------------
def read_input_excel(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    xls = pd.read_excel(path, sheet_name=None)
    # Normalize sheet names to lower-case keys
    sheets = {name.strip().lower(): df.copy() for name, df in xls.items()}
    return sheets

def parse_date(cell):
    if pd.isna(cell):
        return None
    if isinstance(cell, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(cell).date()
    try:
        return pd.to_datetime(cell).date()
    except Exception:
        raise ValueError(f"Cannot parse date: {cell}")

def parse_allowed_lines(cell):
    if pd.isna(cell):
        return None
    if isinstance(cell, (list, tuple, set)):
        return list(cell)
    s = str(cell).strip()
    if s == '':
        return None
    parts = [p.strip() for p in s.replace(';', ',').split(',') if p.strip()]
    return parts if parts else None

# -------------------------
# Main builder & solver
# -------------------------
def build_and_solve(input_path, time_limit_minutes=10, output_dir=None, debug=False):
    sheets = read_input_excel(input_path)

    # Lowercase keys map
    def get_sheet(name_variants):
        for v in name_variants:
            if v.lower() in sheets:
                return sheets[v.lower()]
        return None

    # Required sheets
    grades_df = get_sheet(['Grades'])
    lines_df = get_sheet(['Lines', 'Plants', 'Plant'])
    demand_df = get_sheet(['Demand'])
    if grades_df is None or lines_df is None or demand_df is None:
        raise ValueError("Input Excel must contain 'Grades', 'Lines' (or 'Plants'), and 'Demand' sheets.")

    shutdowns_df = get_sheet(['Shutdowns']) or pd.DataFrame(columns=['Line', 'StartDate', 'EndDate'])
    params_df = get_sheet(['Params']) or pd.DataFrame(columns=['Key', 'Value'])

    # Normalize columns (strip)
    grades_df.columns = [c.strip() for c in grades_df.columns]
    lines_df.columns = [c.strip() for c in lines_df.columns]
    demand_df.columns = [c.strip() for c in demand_df.columns]
    shutdowns_df.columns = [c.strip() for c in shutdowns_df.columns]
    params_df.columns = [c.strip() for c in params_df.columns]

    # -------------------------
    # Parse Grades sheet
    # Expected columns (case-insensitive):
    #   Grade, MinRun (optional), MaxRun (optional), AllowedLines (optional),
    #   InitialInventory (optional), MinInventory (optional), MinClosingInventory (optional)
    # -------------------------
    grades_df = grades_df.rename(columns={c: c.strip() for c in grades_df.columns})
    # find grade column
    grade_col = None
    for c in grades_df.columns:
        if c.lower() == 'grade':
            grade_col = c
            break
    if grade_col is None:
        raise ValueError("Grades sheet must contain a 'Grade' column.")
    # Build grade list and properties
    grade_list = []
    grades_info = {}
    for _, row in grades_df.iterrows():
        g = row.get(grade_col)
        if pd.isna(g):
            continue
        g = str(g).strip()
        if g == '':
            continue
        grade_list.append(g)
        # allowed lines
        allowed = None
        for opt in ['AllowedLines', 'Allowed Lines', 'Allowed', 'Lines']:
            if opt in grades_df.columns:
                allowed = parse_allowed_lines(row.get(opt))
                break
        # min/max run
        min_run = None
        max_run = None
        for opt in ['MinRun', 'Min Run', 'Min_Run']:
            if opt in grades_df.columns:
                v = row.get(opt)
                if not pd.isna(v):
                    min_run = int(v)
                break
        for opt in ['MaxRun', 'Max Run', 'Max_Run']:
            if opt in grades_df.columns:
                v = row.get(opt)
                if not pd.isna(v):
                    max_run = int(v)
                break
        # inventories
        init_inv = 0
        for opt in ['InitialInventory', 'OpeningInventory', 'Opening Inventory']:
            if opt in grades_df.columns:
                v = row.get(opt)
                if not pd.isna(v):
                    init_inv = int(v)
                break
        min_inv = 0
        for opt in ['MinInventory', 'Min Inventory', 'Min_Inventory']:
            if opt in grades_df.columns:
                v = row.get(opt)
                if not pd.isna(v):
                    min_inv = int(v)
                break
        min_closing = 0
        for opt in ['MinClosingInventory', 'Min Closing Inventory', 'Min_Closing_Inventory']:
            if opt in grades_df.columns:
                v = row.get(opt)
                if not pd.isna(v):
                    min_closing = int(v)
                break
        grades_info[g] = {
            'allowed_lines': allowed,  # None means allowed anywhere (resolved later)
            'min_run': min_run if min_run is not None else 1,
            'max_run': max_run if max_run is not None else 9999,
            'initial_inventory': init_inv,
            'min_inventory': min_inv,
            'min_closing_inventory': min_closing
        }

    if not grade_list:
        raise ValueError("Grades sheet contains no grades.")

    # -------------------------
    # Parse Lines/Plants sheet
    # Expected columns: Line (or Plant), Capacity per day (or Capacity)
    # Also: MaterialRunning (initial grade) and ExpectedRunDays for start-of-horizon enforced runs
    # -------------------------
    # find name/line column and capacity column
    name_col = None
    cap_col = None
    for c in lines_df.columns:
        cl = c.lower()
        if cl in ('line', 'plant', 'plantname', 'plant name'):
            name_col = c
        if cl in ('capacity', 'capacity per day', 'capacity_per_day'):
            cap_col = c
    if name_col is None or cap_col is None:
        # try heuristics
        for c in lines_df.columns:
            if name_col is None and any(k in c.lower() for k in ['line', 'plant']):
                name_col = c
            if cap_col is None and 'cap' in c.lower():
                cap_col = c
    if name_col is None or cap_col is None:
        raise ValueError("Lines/Plants sheet must contain a line/plant name column and a capacity column (e.g., 'Capacity' or 'Capacity per day').")

    lines = []
    line_capacity = {}
    initial_material = {}     # line -> grade name or None
    expected_run_days_start = {}  # line -> integer (expected run days at start)
    for _, row in lines_df.iterrows():
        name = row.get(name_col)
        if pd.isna(name):
            continue
        name = str(name).strip()
        if name == '':
            continue
        lines.append(name)
        cap = row.get(cap_col)
        cap_val = 0 if pd.isna(cap) else int(cap)
        line_capacity[name] = cap_val
        # MaterialRunning column (initial material)
        mat = None
        for opt in ['MaterialRunning', 'Material Running', 'Material']:
            if opt in lines_df.columns:
                if not pd.isna(row.get(opt)):
                    mat = str(row.get(opt)).strip()
                break
        initial_material[name] = mat
        # ExpectedRunDays column
        erd = 0
        for opt in ['Expected Run Days', 'ExpectedRunDays', 'Expected_Run_Days']:
            if opt in lines_df.columns:
                if not pd.isna(row.get(opt)):
                    erd = int(row.get(opt))
                break
        expected_run_days_start[name] = erd

    if not lines:
        raise ValueError("No lines/plants defined in Lines sheet.")

    # -------------------------
    # Parse Demand sheet
    # Two common formats:
    # 1) Wide: Date, <Grade1>, <Grade2>, ...  <-- your template uses wide format
    # 2) Long: Date, Grade, Demand
    # We'll detect wide format and extract per-grade demands
    # -------------------------
    demand_df = demand_df.copy()
    # find date column
    date_col = None
    for c in demand_df.columns:
        if c.lower() in ('date', 'day'):
            date_col = c
            break
    if date_col is None:
        # try first column
        date_col = demand_df.columns[0]
    # parse dates and build sorted horizon
    demand_df[date_col] = demand_df[date_col].apply(parse_date)
    date_values = sorted(set([d for d in demand_df[date_col].tolist() if d is not None]))
    if not date_values:
        raise ValueError("Demand sheet contains no dates.")
    day_index = {d: i for i, d in enumerate(date_values)}
    T = len(date_values)

    # build demand_by_grade_day
    demand_by_grade_day = {g: [0] * T for g in grade_list}
    # detect if wide: columns include grade names
    wide_mode = False
    grade_cols_present = [c for c in demand_df.columns if c in grade_list]
    if grade_cols_present:
        wide_mode = True
        for _, row in demand_df.iterrows():
            d = row[date_col]
            if pd.isna(d):
                continue
            d = parse_date(d)
            if d not in day_index:
                continue
            idx = day_index[d]
            for g in grade_list:
                if g in demand_df.columns:
                    v = row.get(g)
                    if pd.isna(v):
                        continue
                    demand_by_grade_day[g][idx] = int(v)
    else:
        # long mode: expect columns Grade and Demand (or similar)
        grade_col_d = None
        qty_col = None
        for c in demand_df.columns:
            if c.lower() == 'grade':
                grade_col_d = c
            if c.lower() in ('demand', 'qty', 'quantity'):
                qty_col = c
        if grade_col_d is None or qty_col is None:
            raise ValueError("Demand sheet not in wide format and missing Grade/Demand columns.")
        for _, row in demand_df.iterrows():
            d = parse_date(row[date_col])
            if d not in day_index:
                continue
            idx = day_index[d]
            g = row.get(grade_col_d)
            if pd.isna(g):
                continue
            g = str(g).strip()
            if g not in grade_list:
                # ignore unknown grade rows
                continue
            q = 0 if pd.isna(row.get(qty_col)) else int(row.get(qty_col))
            demand_by_grade_day[g][idx] = q

    # -------------------------
    # Parse Shutdowns
    # expected columns: Line, StartDate, EndDate (inclusive)
    # -------------------------
    shutdowns = defaultdict(list)  # line -> list of (start_idx, end_idx)
    if shutdowns_df is not None and not shutdowns_df.empty:
        # find cols
        col_line = None
        col_start = None
        col_end = None
        for c in shutdowns_df.columns:
            if c.lower() in ('line', 'plant'):
                col_line = c
            if c.lower() in ('startdate', 'start_date', 'start'):
                col_start = c
            if c.lower() in ('enddate', 'end_date', 'end'):
                col_end = c
        if col_line and col_start and col_end:
            for _, row in shutdowns_df.iterrows():
                line = row.get(col_line)
                if pd.isna(line):
                    continue
                line = str(line).strip()
                if line not in lines:
                    continue
                s = parse_date(row.get(col_start))
                e = parse_date(row.get(col_end))
                if s is None or e is None:
                    continue
                # clamp to horizon
                if e < date_values[0] or s > date_values[-1]:
                    continue
                s_clamped = max(s, date_values[0])
                e_clamped = min(e, date_values[-1])
                shutdowns[line].append((day_index[s_clamped], day_index[e_clamped]))

    # -------------------------
    # Parse per-plant transition matrices (optional)
    # Sheet names often like 'Transition_Plant1' or 'Transition_<line>'
    # Each such sheet should have first column as 'From' and subsequent columns as grades with Yes/No
    # -------------------------
    # Build a dict: transitions_allowed[line][from_gid][to_gid] -> True/False
    transitions_allowed = {}
    for line in lines:
        transitions_allowed[line] = [[True] * len(grade_list) for _ in range(len(grade_list))]  # default allow all
    # search sheets starting with 'transition'
    for sheet_name, df in sheets.items():
        if not sheet_name.lower().startswith('transition'):
            continue
        # determine which line this sheet refers to: if sheet name like transition_plant1 try to map
        # fallback: attempt to find a column header 'From' and row headers with grade names
        # derive line name from sheet name suffix if it matches a line
        line_ref = None
        parts = sheet_name.split('_')
        if len(parts) >= 2:
            candidate = '_'.join(parts[1:]).strip()
            # try exact match or case-insensitive
            for l in lines:
                if candidate.lower() == l.lower():
                    line_ref = l
                    break
        # if still None, try scanning sheet for 'From' column and infer nothing (we'll skip mapping unless unambiguous)
        if line_ref is None:
            # try any sheet that contains a word matching a line name
            for l in lines:
                if l.lower() in sheet_name.lower():
                    line_ref = l
                    break
        if line_ref is None:
            # skip ambiguous transition sheets
            continue
        df = df.copy()
        df.columns = [c.strip() for c in df.columns]
        # expecting first column is 'From' or grade names in rows and columns header are to-grades
        # Attempt long or wide parse:
        # if df has a column 'From' and columns matching grade_list, use them
        header_cols = [c for c in df.columns]
        # find row index mapping by reading first column values
        from_col = None
        for c in header_cols:
            if c.lower() in ('from',):
                from_col = c
                break
        if from_col is None:
            # try first column
            from_col = header_cols[0]
        # to-grade columns are header names excluding from_col that match grade_list
        to_cols = [c for c in header_cols if c != from_col and c in grade_list]
        if not to_cols:
            # maybe the sheet is transposed; skip
            continue
        # Build map
        allowed = [[False] * len(grade_list) for _ in range(len(grade_list))]
        for _, row in df.iterrows():
            frm = row.get(from_col)
            if pd.isna(frm):
                continue
            frm = str(frm).strip()
            if frm not in grade_list:
                continue
            i = grade_list.index(frm)
            for j, to_grade in enumerate(grade_list):
                if to_grade in df.columns:
                    val = row.get(to_grade)
                    if pd.isna(val):
                        # default to True (safe)
                        allowed[i][j] = True
                    else:
                        sval = str(val).strip().lower()
                        if sval in ('yes', 'y', '1', 'true', 'allow', 'allowed'):
                            allowed[i][j] = True
                        else:
                            allowed[i][j] = False
                else:
                    # column missing, default allow
                    allowed[i][j] = True
        transitions_allowed[line_ref] = allowed

    # -------------------------
    # Params with defaults
    # -------------------------
    default_params = {
        'stockout_cost': 1000.0,
        'transition_cost': 100.0,
        'production_cost': 0.0,
        'time_limit_minutes': time_limit_minutes if time_limit_minutes is not None else 10.0,
        'num_search_workers': None,
        'random_seed': 0,
        'min_inventory_penalty': 50.0,      # per unit shortfall per day
        'min_closing_penalty': 200.0,       # per unit shortfall on final day
        'write_outputs': True,
        'output_prefix': 'schedule_output'
    }
    params = dict(default_params)
    if params_df is not None and not params_df.empty:
        for _, row in params_df.iterrows():
            key = row.iloc[0]
            val = row.iloc[1] if len(row) > 1 else None
            if pd.isna(key):
                continue
            key = str(key).strip()
            if key in params:
                if val is None or (isinstance(val, float) and math.isnan(val)):
                    continue
                try:
                    if isinstance(params[key], bool):
                        params[key] = bool(val)
                    elif isinstance(params[key], int):
                        params[key] = int(val)
                    elif isinstance(params[key], float):
                        params[key] = float(val)
                    else:
                        params[key] = val
                except Exception:
                    # leave as default if parsing fails
                    continue
            else:
                # allow adding penalty params like min_inventory_penalty etc.
                try:
                    params[key] = float(val)
                except Exception:
                    params[key] = val

    # Override time limit if function arg provided
    if time_limit_minutes is not None:
        params['time_limit_minutes'] = float(time_limit_minutes)

    # -------------------------
    # Feasible days per line (exclude shutdown days)
    # -------------------------
    feasible_days_per_line = {}
    continuous_blocks = {}
    for line in lines:
        sd_flags = [0] * T
        for (s, e) in shutdowns.get(line, []):
            for d in range(s, e + 1):
                if 0 <= d < T:
                    sd_flags[d] = 1
        feasible = [d for d in range(T) if sd_flags[d] == 0]
        feasible_days_per_line[line] = feasible
        # continuous blocks
        blocks = []
        if feasible:
            start = feasible[0]
            prev = feasible[0]
            for d in feasible[1:]:
                if d == prev + 1:
                    prev = d
                else:
                    blocks.append((start, prev))
                    start = d
                    prev = d
            blocks.append((start, prev))
        continuous_blocks[line] = blocks

    # -------------------------
    # Max possible production per grade per day (UB) used to bound vars
    # -------------------------
    max_possible_prod_by_grade_day = {}
    for g in grade_list:
        arr = [0] * T
        for d in range(T):
            s = 0
            for line in lines:
                allowed_lines = grades_info[g]['allowed_lines']
                if allowed_lines is None or line in allowed_lines:
                    if d in feasible_days_per_line[line]:
                        s += line_capacity[line]
            arr[d] = s
        max_possible_prod_by_grade_day[g] = arr

    # -------------------------
    # Build CP-SAT model
    # -------------------------
    model = cp_model.CpModel()
    grade_to_idx = {g: i for i, g in enumerate(grade_list)}
    idx_to_grade = {i: g for g, i in grade_to_idx.items()}
    G = len(grade_list)
    IDLE = -1  # special idle value for line_grade

    # Variables:
    # line_grade[(line,d)] -> IntVar in domain {IDLE} U allowed grade indices for that line/day
    line_grade = {}
    # For convenience: production var per (line,grade_idx,d) if that assignment possible
    production_lgd = {}
    # produced total per (grade_idx,d)
    produced_by_grade_day = {}

    # Create line_grade only for feasible days. If day not feasible (shutdown) we won't create variable.
    for line in lines:
        for d in feasible_days_per_line[line]:
            # Determine allowed grade indices at this line-day (respecting grade-level allowed_lines if present)
            allowed_gids = []
            for g in grade_list:
                allowed_lines = grades_info[g]['allowed_lines']
                if allowed_lines is None or line in allowed_lines:
                    allowed_gids.append(grade_to_idx[g])
            if not allowed_gids:
                domain_vals = [IDLE]
            else:
                domain_vals = [IDLE] + allowed_gids
            # Create IntVar with min..max domain and then forbid non-members
            vmin = min(domain_vals)
            vmax = max(domain_vals)
            var = model.NewIntVar(vmin, vmax, f'linegrade_{line}_{d}')
            line_grade[(line, d)] = var
            # forbid values not in domain
            forbidden_vals = [val for val in range(vmin, vmax + 1) if val not in domain_vals]
            for fv in forbidden_vals:
                model.Add(var != fv)
            # For each allowed grade, create equality bool and production var
            for gid in allowed_gids:
                eq_b = model.NewBoolVar(f'lg_eq_{line}_{d}_g{gid}')
                # reify equality
                model.Add(var == gid).OnlyEnforceIf(eq_b)
                model.Add(var != gid).OnlyEnforceIf(eq_b.Not())
                # production variable for this line-grade-day
                pvar = model.NewIntVar(0, line_capacity[line], f'prod_{line}_g{gid}_{d}')
                production_lgd[(line, gid, d)] = pvar
                # if eq_b false then production zero
                model.Add(pvar == 0).OnlyEnforceIf(eq_b.Not())
                # if eq_b true then production <= capacity
                model.Add(pvar <= line_capacity[line]).OnlyEnforceIf(eq_b)

    # Aggregate produced_by_grade_day as sum over lines
    for g in grade_list:
        gid = grade_to_idx[g]
        for d in range(T):
            parts = []
            for line in lines:
                key = (line, gid, d)
                if key in production_lgd:
                    parts.append(production_lgd[key])
            if parts:
                agg_ub = sum(line_capacity[l] for l in lines)
                agg = model.NewIntVar(0, agg_ub, f'prod_agg_g{gid}_{d}')
                model.Add(agg == sum(parts))
                produced_by_grade_day[(gid, d)] = agg
            else:
                produced_by_grade_day[(gid, d)] = 0  # constant zero

    # Inventory / supply / stockout / shortage variables
    inventory = {}
    supplied = {}
    stockout = {}
    shortage_min_inv = {}   # daily min-inventory shortage (soft)
    shortage_min_closing = {}  # final-day closing shortage (soft)
    max_inventory_est = {}

    for g in grade_list:
        gid = grade_to_idx[g]
        init_inv = grades_info[g]['initial_inventory']
        # upper bound estimate on inventory: initial + sum possible production
        max_future = sum(max_possible_prod_by_grade_day[g])
        ub = init_inv + max_future
        max_inventory_est[gid] = ub
        for d in range(T):
            inv = model.NewIntVar(0, ub, f'inv_g{gid}_{d}')
            inventory[(gid, d)] = inv
            demand_q = demand_by_grade_day[g][d]
            sup_ub = min(ub, demand_q + max_possible_prod_by_grade_day[g][d])
            sup = model.NewIntVar(0, sup_ub, f'sup_g{gid}_{d}')
            supplied[(gid, d)] = sup
            st = model.NewIntVar(0, demand_q, f'stockout_g{gid}_{d}')
            stockout[(gid, d)] = st
            # daily min-inventory shortage (soft)
            min_inv = grades_info[g]['min_inventory']
            # maximum possible shortage is min_inv (if inventory >=0)
            sh_ub = max(0, min_inv)
            sh = model.NewIntVar(0, sh_ub, f'shmin_g{gid}_{d}')
            shortage_min_inv[(gid, d)] = sh

        # final day closing shortage var
        min_cl = grades_info[g]['min_closing_inventory']
        final_sh_ub = max(0, min_cl)
        final_sh = model.NewIntVar(0, final_sh_ub, f'shclose_g{gid}_final')
        shortage_min_closing[gid] = final_sh

    # Inventory balance
    for g in grade_list:
        gid = grade_to_idx[g]
        init_inv = grades_info[g]['initial_inventory']
        for d in range(T):
            prod_term = produced_by_grade_day[(gid, d)]
            sup = supplied[(gid, d)]
            inv = inventory[(gid, d)]
            demand_q = demand_by_grade_day[g][d]
            if d == 0:
                model.Add(inv == init_inv + (prod_term if isinstance(prod_term, int) or isinstance(prod_term, cp_model.IntVar) else prod_term) - sup)
            else:
                model.Add(inv == inventory[(gid, d - 1)] + (prod_term if isinstance(prod_term, int) or isinstance(prod_term, cp_model.IntVar) else prod_term) - sup)
            # bounds on supply and stockout
            model.Add(sup <= demand_q)
            model.Add(stockout[(gid, d)] == demand_q - sup)
            # supply cannot exceed available
            if d == 0:
                avail_expr = init_inv + (prod_term if isinstance(prod_term, int) or isinstance(prod_term, cp_model.IntVar) else prod_term)
            else:
                avail_expr = inventory[(gid, d - 1)] + (prod_term if isinstance(prod_term, int) or isinstance(prod_term, cp_model.IntVar) else prod_term)
            model.Add(sup <= avail_expr)
            # shortage min inventory: sh >= min_inventory - inventory
            min_inv = grades_info[g]['min_inventory']
            model.Add(shortage_min_inv[(gid, d)] >= min_inv - inv)
            model.Add(shortage_min_inv[(gid, d)] >= 0)
            # also ensure shortage cannot exceed min_inv
            model.Add(shortage_min_inv[(gid, d)] <= max(0, min_inv))

        # final day closing shortage: min_closing - inv_final
        final_inv = inventory[(gid, T - 1)]
        min_cl = grades_info[g]['min_closing_inventory']
        model.Add(shortage_min_closing[gid] >= min_cl - final_inv)
        model.Add(shortage_min_closing[gid] >= 0)
        model.Add(shortage_min_closing[gid] <= max(0, min_cl))

    # -------------------------
    # Enforce MaterialRunning initial runs (hard) using expected_run_days_start.
    # The interpretation: the plant is producing MaterialRunning at start of period for ExpectedRunDays available days.
    # If MaterialRunning is missing or expected_run_days == 0, skip.
    # If the specified material is not allowed at that plant, we do NOT force it (we'll warn).
    # -------------------------
    for line in lines:
        mat = initial_material.get(line)
        erd = expected_run_days_start.get(line, 0) or 0
        if mat is None or erd <= 0:
            continue
        if mat not in grade_list:
            print(f"Warning: MaterialRunning '{mat}' for line {line} not in Grades list; skipping initial run constraint.")
            continue
        gid = grade_to_idx[mat]
        # gather earliest erd feasible days (skip shutdowns)
        feasible = feasible_days_per_line[line]
        if not feasible:
            print(f"Warning: Line {line} has no feasible days in horizon; cannot enforce initial MaterialRunning.")
            continue
        # assign first erd feasible days or until feasible exhausted
        enforced_days = feasible[:erd]
        # check if mat allowed on this line (grade-level AllowedLines)
        allowed_lines_for_mat = grades_info[mat]['allowed_lines']
        if allowed_lines_for_mat is not None and line not in allowed_lines_for_mat:
            print(f"Warning: initial MaterialRunning {mat} not allowed on line {line} by grade AllowedLines; skipping enforcement.")
            continue
        # enforce equality for each enforced day: line_grade[(line,d)] == gid
        for d in enforced_days:
            if (line, d) in line_grade:
                model.Add(line_grade[(line, d)] == gid)
            else:
                # if variable doesn't exist (shouldn't happen), warn
                print(f"Warning: expected to enforce initial run on line {line} day {d} but this day is infeasible (shutdown).")

    # -------------------------
    # Forbidden transitions (hard) and transition detection
    # For each line and adjacent feasible days (d, d+1), if transition from a->b is "No" in transitions_allowed, forbid that assignment pair.
    # Also create trans_var (Bool) indicating a transition (grade changes between day d and d+1) when both days feasible and both not idle.
    # -------------------------
    trans_vars = {}  # (line,d) -> Bool
    for line in lines:
        allowed_mat = transitions_allowed.get(line, [[True]*G for _ in range(G)])
        for d in range(T - 1):
            if d in feasible_days_per_line[line] and (d + 1) in feasible_days_per_line[line]:
                v1 = line_grade[(line, d)]
                v2 = line_grade[(line, d + 1)]
                # forbid specific pairs where allowed_mat[from][to] == False
                forbidden_pairs = []
                for i in range(G):
                    for j in range(G):
                        if not allowed_mat[i][j]:
                            forbidden_pairs.append([i, j])
                # Also forbid transitions that involve IDLE? If allowed_mat doesn't define idle transitions, allowed them.
                if forbidden_pairs:
                    # AddForbiddenAssignments expects list of lists
                    try:
                        model.AddForbiddenAssignments([v1, v2], forbidden_pairs)
                    except Exception:
                        # fallback to pairwise reified forbids (slower)
                        for (i, j) in forbidden_pairs:
                            b1 = model.NewBoolVar(f'v1_eq_{line}_{d}_{i}')
                            b2 = model.NewBoolVar(f'v2_eq_{line}_{d+1}_{j}')
                            model.Add(v1 == i).OnlyEnforceIf(b1)
                            model.Add(v1 != i).OnlyEnforceIf(b1.Not())
                            model.Add(v2 == j).OnlyEnforceIf(b2)
                            model.Add(v2 != j).OnlyEnforceIf(b2.Not())
                            # forbid both true
                            model.AddBoolAnd([b1, b2]).OnlyEnforceIf(model.NewBoolVar(f'forbid_tmp_{line}_{d}_{i}_{j}')).OnlyEnforceIf(False)
                # create transition indicator: trans = 1 when both not idle and v1 != v2
                # We'll create eq bool then trans = 1 - eq_same when both not idle; but treat idle transitions as not transition
                eq = model.NewBoolVar(f'eq_same_{line}_{d}')
                model.Add(v1 == v2).OnlyEnforceIf(eq)
                model.Add(v1 != v2).OnlyEnforceIf(eq.Not())
                # both_producing bool: True if v1 != IDLE and v2 != IDLE
                prod1 = model.NewBoolVar(f'prod1_{line}_{d}')
                prod2 = model.NewBoolVar(f'prod2_{line}_{d+1}')
                model.Add(v1 != IDLE).OnlyEnforceIf(prod1)
                model.Add(v1 == IDLE).OnlyEnforceIf(prod1.Not())
                model.Add(v2 != IDLE).OnlyEnforceIf(prod2)
                model.Add(v2 == IDLE).OnlyEnforceIf(prod2.Not())
                both_prod = model.NewBoolVar(f'bothprod_{line}_{d}')
                # both_prod <= prod1, both_prod <= prod2, both_prod >= prod1 + prod2 -1
                model.Add(both_prod <= prod1)
                model.Add(both_prod <= prod2)
                model.Add(both_prod >= prod1 + prod2 - 1)
                tvar = model.NewBoolVar(f'trans_{line}_{d}')
                trans_vars[(line, d)] = tvar
                # trans == 1 iff (both_prod == 1) and (eq == 0)
                # So trans <= 1 - eq, trans <= both_prod, trans >= both_prod - eq
                model.Add(tvar <= (1)).OnlyEnforceIf()
                model.Add(tvar <= 1 - (eq))  # uses bool arithmetic
                model.Add(tvar <= both_prod)
                model.Add(tvar >= both_prod - eq)
            else:
                # skip pairs where one day not feasible
                pass

    # -------------------------
    # Enforce max_run via sliding window inside continuous blocks (avoid runs crossing shutdowns)
    # For each line, each grade with max_run < block_length, forbid any window of size max_run+1 being all equal to that grade.
    # -------------------------
    for line in lines:
        blocks = continuous_blocks[line]
        for (start, end) in blocks:
            block_len = end - start + 1
            for g in grade_list:
                gid = grade_to_idx[g]
                maxr = grades_info[g]['max_run']
                if maxr is None or maxr >= block_len:
                    continue
                w = maxr + 1
                for s in range(start, end - w + 2):
                    # For window days s..s+w-1, forbid all equal to gid simultaneously
                    vars_window = [line_grade[(line, d)] for d in range(s, s + w)]
                    # Add forbidden assignment: all equal to gid
                    forb = [[gid] * len(vars_window)]
                    try:
                        model.AddForbiddenAssignments(vars_window, forb)
                    except Exception:
                        # fallback: create boolean conjunction and forbid it
                        eq_bools = []
                        for d in range(s, s + w):
                            b = model.NewBoolVar(f'win_eq_{line}_{g}_{d}')
                            model.Add(line_grade[(line, d)] == gid).OnlyEnforceIf(b)
                            model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(b.Not())
                            eq_bools.append(b)
                        # sum(eq_bools) <= w-1
                        model.Add(sum(eq_bools) <= w - 1)

    # -------------------------
    # Enforce min_run (Expected Run Days or MinRun) as minimum length of a run.
    # We'll use the approach: if a run starts at day d for grade g on line, then the next min_run days must be that grade.
    # We'll build start indicators per (line,d,g) via reified equality between line_grade and grade + previous not equal.
    # -------------------------
    is_start = {}  # (line,d,gid) -> Bool
    for line in lines:
        for (start, end) in continuous_blocks[line]:
            for d in range(start, end + 1):
                # if day not feasible skip
                if (line, d) not in line_grade:
                    continue
                for g in grade_list:
                    gid = grade_to_idx[g]
                    # Check if grade allowed on this line at day d
                    allowed_lines = grades_info[g]['allowed_lines']
                    if allowed_lines is not None and line not in allowed_lines:
                        continue
                    # create eq_b for (line,d)==gid
                    eq_b = model.NewBoolVar(f'start_eq_{line}_{d}_g{gid}')
                    model.Add(line_grade[(line, d)] == gid).OnlyEnforceIf(eq_b)
                    model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(eq_b.Not())
                    # prev exists?
                    if d - 1 in feasible_days_per_line[line]:
                        # create prev_eq bool for (line,d-1)==gid
                        prev_eq = model.NewBoolVar(f'prev_eq_{line}_{d-1}_g{gid}')
                        model.Add(line_grade[(line, d - 1)] == gid).OnlyEnforceIf(prev_eq)
                        model.Add(line_grade[(line, d - 1)] != gid).OnlyEnforceIf(prev_eq.Not())
                        s_var = model.NewBoolVar(f'is_start_{line}_{d}_g{gid}')
                        is_start[(line, d, gid)] = s_var
                        # s_var = eq_b AND (not prev_eq)
                        model.Add(s_var <= eq_b)
                        model.Add(s_var <= prev_eq.Not())
                        model.Add(s_var >= eq_b - prev_eq)
                    else:
                        # beginning of block: start iff eq_b
                        s_var = model.NewBoolVar(f'is_start_{line}_{d}_g{gid}')
                        is_start[(line, d, gid)] = s_var
                        model.Add(s_var == eq_b)

    # Now min-run enforcement
    for line in lines:
        for (start, end) in continuous_blocks[line]:
            for d in range(start, end + 1):
                for g in grade_list:
                    gid = grade_to_idx[g]
                    key = (line, d, gid)
                    if key not in is_start:
                        continue
                    min_run = grades_info[g]['min_run']
                    if min_run is None or min_run <= 1:
                        continue
                    # if start at d, ensure that for k in 0..min_run-1, line_grade == gid
                    # but must check block boundaries
                    if d + min_run - 1 > end:
                        # cannot start this grade here => forbid is_start true
                        model.Add(is_start[key] == 0)
                        continue
                    # create implication: is_start -> sum_{k=0..min_run-1} eq(line, d+k) >= min_run
                    # Build eq bools for each window day
                    eqs = []
                    for k in range(min_run):
                        dd = d + k
                        b = model.NewBoolVar(f'minrun_eq_{line}_{dd}_g{gid}')
                        model.Add(line_grade[(line, dd)] == gid).OnlyEnforceIf(b)
                        model.Add(line_grade[(line, dd)] != gid).OnlyEnforceIf(b.Not())
                        eqs.append(b)
                    # sum(eqs) >= min_run * is_start
                    model.Add(sum(eqs) >= min_run * is_start[key])

    # -------------------------
    # Objective: minimize weighted sum of:
    #   - stockout (units unmet) * stockout_cost
    #   - transition count * transition_cost
    #   - production cost (optional) * total_production
    #   - min_inventory shortages * min_inventory_penalty
    #   - min_closing shortages * min_closing_penalty
    # -------------------------
    obj_terms = []

    # stockout
    stockout_cost = float(params.get('stockout_cost', default_params['stockout_cost']))
    stockout_vars = []
    for g in grade_list:
        gid = grade_to_idx[g]
        for d in range(T):
            stockout_vars.append(stockout[(gid, d)])
    if stockout_vars:
        obj_terms.append((stockout_cost, stockout_vars))

    # transitions
    transition_cost = float(params.get('transition_cost', default_params['transition_cost']))
    trans_var_list = [v for v in trans_vars.values()]
    if trans_var_list:
        obj_terms.append((transition_cost, trans_var_list))

    # production cost
    production_cost = float(params.get('production_cost', default_params['production_cost']))
    prod_vars_list = [p for p in production_lgd.values()]
    if production_cost != 0 and prod_vars_list:
        obj_terms.append((production_cost, prod_vars_list))

    # min inventory shortage penalties
    min_inventory_penalty = float(params.get('min_inventory_penalty', default_params['min_inventory_penalty']))
    mininv_vars = [v for v in shortage_min_inv.values()]
    if mininv_vars:
        obj_terms.append((min_inventory_penalty, mininv_vars))

    # min closing shortage penalties
    min_closing_penalty = float(params.get('min_closing_penalty', default_params['min_closing_penalty']))
    minclose_vars = [v for v in shortage_min_closing.values()]
    if minclose_vars:
        obj_terms.append((min_closing_penalty, minclose_vars))

    # Build linear objective scaled to integers
    # flatten terms
    terms_flat = []
    for (coeff, varlist) in obj_terms:
        for v in varlist:
            terms_flat.append((float(coeff), v))
    if terms_flat:
        # scale to integer coefficients
        max_decimals = 0
        for (c, _) in terms_flat:
            s = str(float(c))
            if '.' in s:
                decimals = len(s.split('.')[-1].rstrip('0'))
                max_decimals = max(max_decimals, decimals)
        scale = 10 ** min(max_decimals, 6)
        scaled_terms = []
        for (c, v) in terms_flat:
            scaled_terms.append((int(round(c * scale)), v))
        model.Minimize(sum(coef * var for (coef, var) in scaled_terms))
    else:
        model.Minimize(0)

    # -------------------------
    # Solver parameters & profiling
    # -------------------------
    solver = cp_model.CpSolver()
    tlim = float(params.get('time_limit_minutes', default_params['time_limit_minutes']))
    if tlim is not None and tlim > 0:
        solver.parameters.max_time_in_seconds = float(tlim * 60.0)
    import multiprocessing
    workers = params.get('num_search_workers', None)
    if workers is None:
        workers = max(1, min(8, multiprocessing.cpu_count()))
    solver.parameters.num_search_workers = int(workers)
    seed = int(params.get('random_seed', default_params['random_seed']))
    solver.parameters.random_seed = seed
    solver.parameters.log_search_progress = False
    solver.parameters.cp_model_presolve = True

    # model proto stats
    proto = model.Proto()
    num_vars = len(proto.variables)
    num_constraints = len(proto.constraints)
    print("MODEL STATS:")
    print(f"  Grades: {len(grade_list)}, Lines: {len(lines)}, Days: {T}")
    print(f"  Variables (proto): {num_vars}, Constraints (proto): {num_constraints}")
    print(f"  Feasible line-days sum: {sum(len(feasible_days_per_line[l]) for l in lines)}")
    print(f"  Time limit (min): {tlim}, workers: {workers}, seed: {seed}")
    sys.stdout.flush()

    # -------------------------
    # Solve
    # -------------------------
    start_time = time.time()
    result = solver.Solve(model)
    elapsed = time.time() - start_time
    status = solver.StatusName(result)
    print(f"SOLVE STATUS: {status} (time {elapsed:.2f}s)")

    # Extract objective (remember scaling)
    if 'scale' in locals():
        try:
            raw_obj = solver.ObjectiveValue()
            obj_val = raw_obj / float(scale)
        except Exception:
            obj_val = None
    else:
        try:
            obj_val = solver.ObjectiveValue()
        except Exception:
            obj_val = None
    if obj_val is not None:
        print(f"  Objective (approx): {obj_val}")

    # If solution found, extract schedule and inventory
    schedule_rows = []
    inv_rows = []
    if result in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for d_idx, d_date in enumerate(date_values):
            for line in lines:
                if d_idx not in feasible_days_per_line[line]:
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Shutdown',
                        'Grade': None,
                        'Produced': 0
                    })
                    continue
                if (line, d_idx) not in line_grade:
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Idle',
                        'Grade': None,
                        'Produced': 0
                    })
                    continue
                val = solver.Value(line_grade[(line, d_idx)])
                if val == IDLE:
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Idle',
                        'Grade': None,
                        'Produced': 0
                    })
                else:
                    grade_name = idx_to_grade[val]
                    qty = 0
                    key = (line, val, d_idx)
                    if key in production_lgd:
                        qty = solver.Value(production_lgd[key])
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Producing',
                        'Grade': grade_name,
                        'Produced': int(qty)
                    })
            # inventory rows per grade
            for g in grade_list:
                gid = grade_to_idx[g]
                inv = solver.Value(inventory[(gid, d_idx)])
                sup = solver.Value(supplied[(gid, d_idx)])
                st = solver.Value(stockout[(gid, d_idx)])
                prod = 0
                pbd = produced_by_grade_day[(gid, d_idx)]
                if isinstance(pbd, int):
                    prod = pbd
                else:
                    prod = solver.Value(pbd)
                shmin = solver.Value(shortage_min_inv[(gid, d_idx)])
                inv_rows.append({
                    'Date': d_date,
                    'Grade': g,
                    'Produced': int(prod),
                    'Supplied': int(sup),
                    'Stockout': int(st),
                    'Inventory': int(inv),
                    'Shortage_MinInventory': int(shmin)
                })
        # final shortages
        final_shortages = {g: solver.Value(shortage_min_closing[grade_to_idx[g]]) for g in grade_list}
    else:
        print("No feasible solution found; exiting with outputs as None.")
        final_shortages = {g: None for g in grade_list}

    schedule_df = pd.DataFrame(schedule_rows) if schedule_rows else pd.DataFrame()
    inv_df = pd.DataFrame(inv_rows) if inv_rows else pd.DataFrame()

    # Print summary
    if not inv_df.empty:
        total_stockout = inv_df['Stockout'].sum()
        total_produced = inv_df['Produced'].sum()
    else:
        total_stockout = 0
        total_produced = 0
    total_trans = sum(solver.Value(v) for v in trans_var_list) if 'trans_var_list' in locals() and trans_var_list else 0
    print(f"  Total produced: {total_produced}, Total stockout: {total_stockout}, Transitions: {total_trans}")
    print("  Final-day closing shortages per grade:")
    for g in grade_list:
        print(f"    {g}: {final_shortages.get(g)}")

    # Write outputs
    if params.get('write_outputs', True):
        out_pref = params.get('output_prefix', default_params['output_prefix'])
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            schedule_path = os.path.join(output_dir, f'{out_pref}_schedule.csv')
            inv_path = os.path.join(output_dir, f'{out_pref}_inventory.csv')
        else:
            schedule_path = f'{out_pref}_schedule.csv'
            inv_path = f'{out_pref}_inventory.csv'
        if not schedule_df.empty:
            schedule_df.to_csv(schedule_path, index=False)
            print(f"  Wrote schedule to: {schedule_path}")
        if not inv_df.empty:
            inv_df.to_csv(inv_path, index=False)
            print(f"  Wrote inventory to: {inv_path}")

    return {
        'status': status,
        'objective': obj_val,
        'schedule_df': schedule_df,
        'inventory_df': inv_df,
        'solve_time_seconds': elapsed
    }

# -------------------------
# Command-line entry
# -------------------------
def main():
    parser = argparse.ArgumentParser(description="Optimized scheduler (test.py)")
    parser.add_argument('--input', '-i', type=str, required=True, help='Input Excel file (template)')
    parser.add_argument('--time_limit', '-t', type=float, default=10.0, help='Time limit in minutes')
    parser.add_argument('--output_dir', '-o', type=str, default=None, help='Directory to write CSV outputs')
    parser.add_argument('--debug', action='store_true', help='Enable debug prints')
    args = parser.parse_args()
    res = build_and_solve(args.input, time_limit_minutes=args.time_limit, output_dir=args.output_dir, debug=args.debug)
    if res.get('schedule_df') is not None and not res['schedule_df'].empty:
        print("Sample schedule (first 10 rows):")
        print(res['schedule_df'].head(10).to_string(index=False))
    else:
        print("No schedule produced.")

if __name__ == '__main__':
    main()
