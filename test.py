#!/usr/bin/env python3
"""
optimized_scheduler.py

A CP-SAT based production scheduler optimized for performance and correctness.

Key features:
- Reads input from an Excel workbook (sheets: Grades, Lines, Demand, Shutdowns, Params).
- Creates variables only on feasible (non-shutdown) days.
- Encodes the grade produced on a (line,day) as a single IntVar (grade index or -1 for idle)
  to avoid quadratic transition variables.
- Links production to grade assignment with reified booleans.
- Enforces min-run and max-run constraints inside continuous blocks (periods without shutdown).
- Tightens variable domains to speed up propagation.
- Provides solver parameter tuning and model stats for profiling.
- Objective: minimize stockout + transition penalties + production cost (configurable).
"""

import argparse
import os
import sys
import math
import time
from collections import defaultdict, OrderedDict

import pandas as pd
import numpy as np
from ortools.sat.python import cp_model

# ----- Helper: robust excel reading with basic validation -----
def read_input_excel(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")

    # Read all sheets into dict keyed by lowercased sheet name
    xls = pd.read_excel(path, sheet_name=None)
    sheets = {name.lower(): df for name, df in xls.items()}

    # Grades sheet
    if 'grades' not in sheets:
        raise ValueError("Excel must contain a 'Grades' sheet with at least a 'Grade' column.")
    grades_df = sheets['grades'].copy()
    if 'grade' not in [c.lower() for c in grades_df.columns]:
        raise ValueError("'Grades' sheet must have a 'Grade' column.")
    grades_df.columns = [c.strip() for c in grades_df.columns]
    grades_df.rename(columns={c: c.strip() for c in grades_df.columns}, inplace=True)

    # Lines sheet
    if 'lines' not in sheets:
        raise ValueError("Excel must contain a 'Lines' sheet with at least 'Line' and 'Capacity' columns.")
    lines_df = sheets['lines'].copy()
    lines_df.columns = [c.strip() for c in lines_df.columns]
    if 'Line' not in lines_df.columns and 'line' not in lines_df.columns:
        # normalize
        lines_df.rename(columns={c: c.strip().title(): c for c in lines_df.columns}, inplace=True)
    # unify names
    lines_df.columns = [c.strip() for c in lines_df.columns]

    # Demand sheet
    if 'demand' not in sheets:
        raise ValueError("Excel must contain a 'Demand' sheet with 'Date', 'Grade', 'Demand' columns.")
    demand_df = sheets['demand'].copy()
    demand_df.columns = [c.strip() for c in demand_df.columns]

    # optional shutdowns
    shutdowns_df = sheets.get('shutdowns', pd.DataFrame(columns=['Line', 'StartDate', 'EndDate']))
    shutdowns_df = shutdowns_df.copy()
    shutdowns_df.columns = [c.strip() for c in shutdowns_df.columns]

    # optional params
    params_df = sheets.get('params', pd.DataFrame(columns=['Key', 'Value']))
    params_df = params_df.copy()
    params_df.columns = [c.strip() for c in params_df.columns]

    return grades_df, lines_df, demand_df, shutdowns_df, params_df


# ----- Utility parsing functions -----
def parse_allowed_lines(cell):
    if pd.isna(cell):
        return None
    if isinstance(cell, (list, tuple, set)):
        return list(cell)
    s = str(cell).strip()
    if s == '':
        return None
    # comma or semicolon separated
    parts = [p.strip() for p in s.replace(';', ',').split(',') if p.strip() != '']
    return parts if parts else None

def parse_date(x):
    if pd.isna(x):
        return None
    if isinstance(x, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(x).date()
    try:
        return pd.to_datetime(x).date()
    except Exception:
        raise ValueError(f"Cannot parse date: {x}")

# ----- Main model builder and solver function -----
def build_and_solve(input_path, time_limit_minutes=None, output_dir=None, debug=False, params_override=None):
    # Load excel
    grades_df, lines_df, demand_df, shutdowns_df, params_df = read_input_excel(input_path)

    # Normalize columns to expected names (case-insensitive)
    def col_get(df, name_variants, default=None):
        for v in name_variants:
            for c in df.columns:
                if c.lower() == v.lower():
                    return df[c]
        return default

    # Grades processing
    if 'Grade' in grades_df.columns:
        grade_col = 'Grade'
    else:
        # find column ignoring case
        cols_lower = {c.lower(): c for c in grades_df.columns}
        grade_col = cols_lower.get('grade')
        if grade_col is None:
            raise ValueError("Grades sheet needs 'Grade' column")
    grades_df = grades_df.rename(columns={grade_col: 'Grade'})

    # Extract grade list and attributes
    grade_list = []
    grades_info = {}
    for _, row in grades_df.iterrows():
        g = str(row['Grade']).strip()
        if g == '' or pd.isna(g):
            continue
        grade_list.append(g)

        allowed_lines = None
        # try common column names
        for c in ['AllowedLines', 'Allowed Lines', 'Allowed', 'Lines']:
            if c in grades_df.columns:
                allowed_lines = parse_allowed_lines(row.get(c))
                break
        # min/max run
        min_run = row.get('MinRun', None) if 'MinRun' in grades_df.columns else None
        max_run = row.get('MaxRun', None) if 'MaxRun' in grades_df.columns else None
        if pd.isna(min_run):
            min_run = None
        else:
            min_run = int(min_run)
        if pd.isna(max_run):
            max_run = None
        else:
            max_run = int(max_run)
        initial_inventory = int(row.get('InitialInventory', 0)) if 'InitialInventory' in grades_df.columns and not pd.isna(row.get('InitialInventory')) else 0
        min_closing = int(row.get('MinClosingInventory', 0)) if 'MinClosingInventory' in grades_df.columns and not pd.isna(row.get('MinClosingInventory')) else 0

        grades_info[g] = {
            'allowed_lines': allowed_lines,  # may be None and interpreted later
            'min_run': min_run if min_run is not None else 1,
            'max_run': max_run if max_run is not None else 9999,
            'initial_inventory': initial_inventory,
            'min_closing_inventory': min_closing,
        }

    if not grade_list:
        raise ValueError("No grades found in Grades sheet.")

    # Lines processing
    # find necessary columns
    lines_df.columns = [c.strip() for c in lines_df.columns]
    line_name_col = None
    cap_col = None
    for c in lines_df.columns:
        if c.lower() == 'line':
            line_name_col = c
        if c.lower() == 'capacity':
            cap_col = c
    if line_name_col is None or cap_col is None:
        raise ValueError("Lines sheet must contain 'Line' and 'Capacity' columns.")
    lines = []
    line_capacity = {}
    for _, row in lines_df.iterrows():
        line_name = str(row[line_name_col]).strip()
        if line_name == '' or pd.isna(line_name):
            continue
        cap_val = int(row[cap_col]) if not pd.isna(row[cap_col]) else 0
        lines.append(line_name)
        line_capacity[line_name] = int(cap_val)

    if not lines:
        raise ValueError("No production lines found in Lines sheet.")

    # Demand processing: expect rows with Date, Grade, Demand
    demand_df.columns = [c.strip() for c in demand_df.columns]
    # find columns
    date_col = None
    grade_col_d = None
    demand_col = None
    for c in demand_df.columns:
        if c.lower() in ('date', 'day'):
            date_col = c
        elif c.lower() == 'grade':
            grade_col_d = c
        elif c.lower() in ('demand', 'qty', 'quantity'):
            demand_col = c
    if date_col is None or grade_col_d is None or demand_col is None:
        raise ValueError("Demand sheet must have 'Date', 'Grade', and 'Demand' columns.")
    # Parse demand rows -> mapping (date,grade)->demand
    demand_df = demand_df.dropna(subset=[date_col, grade_col_d])  # demand may be NaN -> zero
    demand_df[date_col] = demand_df[date_col].apply(parse_date)
    demand_map = {}
    date_set = set()
    for _, row in demand_df.iterrows():
        d = row[date_col]
        g = str(row[grade_col_d]).strip()
        q = 0 if pd.isna(row.get(demand_col)) else int(row.get(demand_col))
        date_set.add(d)
        demand_map.setdefault(d, {})[g] = q

    if not date_set:
        raise ValueError("Demand sheet contains no dates.")
    sorted_dates = sorted(date_set)
    day_index = {d: idx for idx, d in enumerate(sorted_dates)}
    num_days = len(sorted_dates)

    # shutdowns processing
    shutdowns = defaultdict(list)  # mapping line -> list of (start_day_idx, end_day_idx)
    if shutdowns_df is not None and not shutdowns_df.empty:
        shutdowns_df.columns = [c.strip() for c in shutdowns_df.columns]
        # find columns
        col_line = None
        col_start = None
        col_end = None
        for c in shutdowns_df.columns:
            if c.lower() in ('line',):
                col_line = c
            if c.lower() in ('startdate', 'start_date', 'start'):
                col_start = c
            if c.lower() in ('enddate', 'end_date', 'end'):
                col_end = c
        if col_line is None or col_start is None or col_end is None:
            # If sheet present but columns missing, ignore gracefully
            shutdowns_df = pd.DataFrame(columns=['Line', 'StartDate', 'EndDate'])
        else:
            for _, row in shutdowns_df.iterrows():
                line = str(row[col_line]).strip()
                if line == '' or pd.isna(line):
                    continue
                s = parse_date(row[col_start])
                e = parse_date(row[col_end])
                if s is None or e is None:
                    continue
                # clamp to horizon
                # ignore shutdowns entirely outside horizon
                if e < sorted_dates[0] or s > sorted_dates[-1]:
                    continue
                s_clamped = max(s, sorted_dates[0])
                e_clamped = min(e, sorted_dates[-1])
                start_idx = day_index[s_clamped]
                end_idx = day_index[e_clamped]
                shutdowns[line].append((start_idx, end_idx))

    # Params - default
    default_params = {
        'stockout_cost': 1000.0,     # large penalty per unit of unmet demand (tune in excel params)
        'transition_cost': 100.0,    # penalty per line transition (tune)
        'production_cost': 0.0,      # per-unit production cost (optional)
        'time_limit_minutes': 10.0,
        'num_search_workers': None,  # if None, set to cpu count or 4
        'random_seed': 0,
        'write_outputs': True,
        'output_prefix': 'schedule_output'
    }
    # load params from params_df
    params = dict(default_params)
    if params_df is not None and not params_df.empty:
        for _, row in params_df.iterrows():
            if len(row) < 2:
                continue
            key = str(row.iloc[0]).strip()
            val = row.iloc[1]
            if pd.isna(key):
                continue
            key = key.strip()
            if key in ('stockout_cost', 'transition_cost', 'production_cost'):
                params[key] = float(val)
            elif key in ('time_limit_minutes',):
                params[key] = float(val)
            elif key in ('num_search_workers', 'random_seed'):
                if pd.isna(val):
                    params[key] = None
                else:
                    params[key] = int(val)
            elif key in ('write_outputs',):
                params[key] = bool(val)
            else:
                # allow writing custom params (ignored if unknown)
                params[key] = val

    # CLI overrides
    if params_override is not None:
        params.update(params_override)

    if time_limit_minutes is not None:
        params['time_limit_minutes'] = float(time_limit_minutes)

    # Prepare allowed lines per grade: if missing, assume all lines allowed
    allowed_on_line = {}
    for g in grade_list:
        al = grades_info[g]['allowed_lines']
        if al is None:
            allowed_on_line[g] = list(lines)  # all lines
        else:
            # sanitize list and ensure lines exist
            al_clean = [a for a in al if a in lines]
            if not al_clean:
                # fallback: allow all
                allowed_on_line[g] = list(lines)
            else:
                allowed_on_line[g] = al_clean

    # Build continuous blocks per line (periods without shutdown)
    feasible_days_per_line = {}
    continuous_blocks = {}
    for line in lines:
        sd = [0] * num_days  # 0 -> available, 1 -> shutdown
        for (sidx, eidx) in shutdowns.get(line, []):
            for d in range(sidx, eidx + 1):
                if 0 <= d < num_days:
                    sd[d] = 1
        feasible = [d for d in range(num_days) if sd[d] == 0]
        feasible_days_per_line[line] = feasible
        # find continuous blocks of available days
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

    # Precompute per-grade daily demand list (default 0 if missing)
    demand_by_grade_day = defaultdict(lambda: [0] * num_days)
    for d in sorted_dates:
        for g in grade_list:
            q = demand_map.get(d, {}).get(g, 0)
            demand_by_grade_day[g][day_index[d]] = int(q)

    # Safety checks
    # Ensure capacities are non-negative ints
    for l in lines:
        if l not in line_capacity:
            raise ValueError(f"Line {l} missing capacity in Lines sheet.")
        line_capacity[l] = int(line_capacity[l])

    # Map grade to index
    grade_to_idx = {g: i for i, g in enumerate(grade_list)}
    idx_to_grade = {i: g for g, i in grade_to_idx.items()}
    G = len(grade_list)
    L = len(lines)
    T = num_days

    # Compute an upper bound on possible production of a grade on a day (sum capacities of allowed lines that are available that day)
    max_possible_prod_by_grade_day = {}
    for g in grade_list:
        arr = [0] * T
        for d in range(T):
            s = 0
            for line in allowed_on_line[g]:
                if d in feasible_days_per_line[line]:
                    s += line_capacity[line]
            arr[d] = s
        max_possible_prod_by_grade_day[g] = arr

    # Setup model
    model = cp_model.CpModel()

    # Idle index for line_grade variable
    IDLE = -1

    # Variables:
    # line_grade[(line,d)] -> IntVar in {IDLE} U {0..G-1} but domain customized to allowed grades for that line on that day
    line_grade = {}
    # is_prod[(line,d)] -> BoolVar: whether line produces any grade on that day (not idle)
    is_prod = {}
    # production_line_grade[(line,grade_idx,d)] -> IntVar production of that grade on that line and day (only created if grade allowed and day feasible)
    production_lgd = {}
    # produced_by_grade_day[(grade,d)] aggregated across lines
    produced_by_grade_day = {}

    for line in lines:
        for d in feasible_days_per_line[line]:
            # domain for this line-day: allowed grade indices only
            allowed_grades_idx = []
            for g in grade_list:
                if line in allowed_on_line[g]:
                    allowed_grades_idx.append(grade_to_idx[g])
            # if none allowed, domain = {IDLE}
            if not allowed_grades_idx:
                dom = [IDLE]
            else:
                dom = [IDLE] + allowed_grades_idx
            # create IntVar with compact domain using linearization of domain
            # cp-sat doesn't have direct finite set domain in NewIntVar; we create min..max domain but we'll constrain illegal values out
            min_dom = min(dom)
            max_dom = max(dom)
            var = model.NewIntVar(min_dom, max_dom, f'linegrade_{line}_{d}')
            line_grade[(line, d)] = var
            # Now forbid values not in dom
            forbidden = [v for v in range(min_dom, max_dom + 1) if v not in dom]
            for fv in forbidden:
                model.Add(var != fv)

            # create is_prod bool and link it: is_prod iff line_grade != IDLE
            b = model.NewBoolVar(f'is_prod_{line}_{d}')
            is_prod[(line, d)] = b
            # reify: var == IDLE <=> b == 0
            model.Add(var == IDLE).OnlyEnforceIf(b.Not())
            model.Add(var != IDLE).OnlyEnforceIf(b)

            # For each allowed grade create production var and link to equality boolean
            for g in grade_list:
                gid = grade_to_idx[g]
                if line not in allowed_on_line[g]:
                    continue
                # create equality bool eq_g that means line_grade == gid
                eq_b = model.NewBoolVar(f'lg_eq_{line}_{d}_g{gid}')
                # reified equality
                model.Add(var == gid).OnlyEnforceIf(eq_b)
                model.Add(var != gid).OnlyEnforceIf(eq_b.Not())
                # production for this (line,g,d) - bounded by capacity
                prod_var = model.NewIntVar(0, line_capacity[line], f'prod_{line}_g{gid}_{d}')
                production_lgd[(line, gid, d)] = prod_var
                # link prod_var <= capacity * eq_b
                model.Add(prod_var <= line_capacity[line]).OnlyEnforceIf(eq_b)
                # if eq_b false then production zero
                model.Add(prod_var == 0).OnlyEnforceIf(eq_b.Not())
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
                agg = model.NewIntVar(0, sum(line_capacity[l] for l in lines), f'prod_agg_g{gid}_{d}')
                model.Add(agg == sum(parts))
                produced_by_grade_day[(gid, d)] = agg
            else:
                # no capacity for this grade on any line this day
                produced_by_grade_day[(gid, d)] = 0  # use integer 0 (not var) for simplicity later

    # Inventory & supplied & stockout variables per grade per day
    inventory = {}
    supplied = {}
    stockout = {}
    max_inventory_est = {}
    for g in grade_list:
        gid = grade_to_idx[g]
        # compute a simple upper bound for inventory: starting inventory + sum of possible production across horizon
        max_possible_future = sum(max_possible_prod_by_grade_day[g])
        max_inv = grades_info[g]['initial_inventory'] + max_possible_future
        max_inventory_est[gid] = max_inv
        for d in range(T):
            inv_var = model.NewIntVar(0, max_inv, f'inv_g{gid}_{d}')
            inventory[(gid, d)] = inv_var
            # demand for that day
            demand_q = demand_by_grade_day[g][d]
            sup_max = min(max_inv, demand_q + max_possible_prod_by_grade_day[g][d])
            sup_var = model.NewIntVar(0, sup_max, f'sup_g{gid}_{d}')
            supplied[(gid, d)] = sup_var
            stock_var = model.NewIntVar(0, demand_q, f'stockout_g{gid}_{d}')
            stockout[(gid, d)] = stock_var

    # Inventory balance constraints
    for g in grade_list:
        gid = grade_to_idx[g]
        init_inv = grades_info[g]['initial_inventory']
        for d in range(T):
            prod_term = produced_by_grade_day[(gid, d)]
            if isinstance(prod_term, int):
                prod_term_expr = prod_term
            else:
                prod_term_expr = prod_term
            sup_var = supplied[(gid, d)]
            inv_var = inventory[(gid, d)]
            demand_q = demand_by_grade_day[g][d]
            # inv(t) = inv(t-1) + produced(t) - supplied(t)
            if d == 0:
                model.Add(inv_var == init_inv + prod_term_expr - sup_var)
            else:
                model.Add(inv_var == inventory[(gid, d - 1)] + prod_term_expr - sup_var)
            # supplied constraints: 0 <= supplied <= demand and <= inv + production
            # We also add lower linking: supplied >= demand - stockout
            model.Add(sup_var <= demand_q)
            model.Add(sup_var >= demand_q - stockout[(gid, d)])
            # supply cannot exceed what is available that day
            # available = inv_prev + produced_today
            if d == 0:
                available_expr = init_inv + prod_term_expr
            else:
                available_expr = inventory[(gid, d - 1)] + prod_term_expr
            # supplied <= available
            # available_expr may be an IntVar + IntVar; cp-sat supports <= with var
            model.Add(sup_var <= available_expr)

    # Stockout definition: stockout = demand - supplied (already partially enforced by bounds)
    for g in grade_list:
        gid = grade_to_idx[g]
        for d in range(T):
            demand_q = demand_by_grade_day[g][d]
            model.Add(stockout[(gid, d)] == demand_q - supplied[(gid, d)])

    # Transition booleans: define eq_same[(line,d)] saying line_grade[line,d] == line_grade[line,d+1]
    # and trans[(line,d)] = 1 if they differ (i.e., a transition occurred between day d and d+1).
    eq_same = {}
    trans = {}
    for line in lines:
        for d in range(T - 1):
            # only define if both days feasible, else treat as no transition (line idle or shutdown)
            if d in feasible_days_per_line[line] and (d + 1) in feasible_days_per_line[line]:
                b_eq = model.NewBoolVar(f'eq_same_{line}_{d}')
                eq_same[(line, d)] = b_eq
                # reified equality between two IntVars:
                model.Add(line_grade[(line, d)] == line_grade[(line, d + 1)]).OnlyEnforceIf(b_eq)
                model.Add(line_grade[(line, d)] != line_grade[(line, d + 1)]).OnlyEnforceIf(b_eq.Not())
                tvar = model.NewBoolVar(f'trans_{line}_{d}')
                trans[(line, d)] = tvar
                # enforce tvar = not eq_same
                # eq_same + trans == 1
                model.Add(b_eq + tvar == 1)
            else:
                # do not define transition var when one of the days is shutdown
                pass

    # Enforce max_run constraint: for each line and each continuous block, forbid runs longer than max_run by sliding-window
    # For each window of length (max_run + 1) within a block, sum(is_prod in window) <= max_run
    for line in lines:
        for (start, end) in continuous_blocks[line]:
            length = end - start + 1
            # For each grade, retrieve per-day is_prod? is_prod exists per line-day
            for g in grade_list:
                gid = grade_to_idx[g]
                # max_run for this grade
                maxr = grades_info[g]['max_run']
                if maxr is None or maxr >= length:
                    continue
                # sliding windows of size maxr+1
                wsize = maxr + 1
                for s in range(start, end - wsize + 2):
                    window_days = list(range(s, s + wsize))
                    # For this grade, compute sum of eq booleans (line produces this grade on day)
                    bools = []
                    for d in window_days:
                        key = (line, gid, d)
                        # production_lgd present => we created an eq bool earlier indirectly via linking prod==0 or not
                        # We'll create a dedicated eq var for (line,d,grade) again for safety by checking if a var exists
                        # But we earlier created reified eq booleans named 'lg_eq_{line}_{d}_g{gid}'
                        # To find existing variable names is tricky; instead create new bool and connect via reified equality:
                        eq_b = model.NewBoolVar(f'window_eq_{line}_g{gid}_{d}')
                        model.Add(line_grade[(line, d)] == gid).OnlyEnforceIf(eq_b)
                        model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(eq_b.Not())
                        bools.append(eq_b)
                    model.Add(sum(bools) <= maxr)

    # Enforce min_run: for each possible start day s inside a continuous block, if is_start then the next min_run days must be produced
    # is_start = is_prod[cur] AND NOT is_prod[prev]
    is_start = {}
    for line in lines:
        for (start, end) in continuous_blocks[line]:
            for d in range(start, end + 1):
                if d not in feasible_days_per_line[line]:
                    continue
                cur = is_prod[(line, d)]
                prev_exists = (d - 1 in feasible_days_per_line[line])
                if prev_exists:
                    prev = is_prod[(line, d - 1)]
                    s_var = model.NewBoolVar(f'is_start_{line}_{d}')
                    is_start[(line, d)] = s_var
                    # s = cur AND (not prev) implemented via linear constraints:
                    model.Add(s_var <= cur)
                    model.Add(s_var <= prev.Not())
                    # s_var >= cur - prev  => ensures exact equivalence
                    # But prev.Not() is not an IntVar; we rewrite prev.Not() as 1 - prev
                    # Using integer linearization: s_var >= cur - prev
                    # CP-SAT booleans can be used in arithmetic: cur and prev are BoolVar
                    model.Add(s_var >= cur - prev)
                else:
                    # start of block -> is_start == cur
                    s_var = model.NewBoolVar(f'is_start_{line}_{d}')
                    is_start[(line, d)] = s_var
                    model.Add(s_var == cur)

    # For min_run: if is_start then sum is_prod for next min_run days >= min_run
    for line in lines:
        for (start, end) in continuous_blocks[line]:
            for d in range(start, end + 1):
                s_var = is_start.get((line, d), None)
                if s_var is None:
                    continue
                # for each grade that has a min_run > 1, ensure if the run starts with that grade then the next min_run days produce the same grade
                # To avoid per-grade branching here (heavy), we'll apply a simpler but conservative min_run:
                # If any run starts at day d on a line, then the run must last at least the min_run of whichever grade is selected on that day.
                # We implement a constraint: for every grade g allowed on this line that has min_run m_g,
                # create an implication: if is_start AND (line_grade == g) then sum_{k=0..m_g-1} is_prod(line, d+k) >= m_g
                # This is precise but requires reified equality; we already created those reified eq booleans earlier for each (line,g,d).
                for g in grade_list:
                    gid = grade_to_idx[g]
                    m_g = grades_info[g]['min_run']
                    if m_g is None or m_g <= 1:
                        continue
                    # check if enough days remain in the block
                    if d + m_g - 1 > end:
                        # cannot start this grade here because run would exceed block; forbid combination
                        # forbid: not (is_start && line_grade==g)
                        # i.e., enforce: if is_start then line_grade != g OR is_start == 0
                        # We implement: model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(is_start[(line,d)])
                        model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(is_start[(line, d)])
                        continue
                    # create eq boolean for (line,d) == gid
                    eq_b = model.NewBoolVar(f'minrun_eq_{line}_{d}_g{gid}')
                    model.Add(line_grade[(line, d)] == gid).OnlyEnforceIf(eq_b)
                    model.Add(line_grade[(line, d)] != gid).OnlyEnforceIf(eq_b.Not())
                    # now create condition: if is_start AND eq_b then sum is_prod over next m_g days >= m_g
                    # produce a helper bool that is conjunction of is_start and eq_b. There's no direct And to make a bool var, but we can enforce:
                    conj = model.NewBoolVar(f'minrun_conj_{line}_{d}_g{gid}')
                    # conj <= is_start, conj <= eq_b, conj >= is_start + eq_b -1
                    model.Add(conj <= is_start[(line, d)])
                    model.Add(conj <= eq_b)
                    model.Add(conj >= is_start[(line, d)] + eq_b - 1)
                    # sum is_prod over next m_g days
                    sum_vars = []
                    for k in range(m_g):
                        dd = d + k
                        sum_vars.append(is_prod[(line, dd)])
                    # if conj is true => sum >= m_g
                    # we'll linearize via: sum_vars >= m_g * conj
                    model.Add(sum(sum_vars) >= m_g * conj)

    # Objective: minimize stockout_cost * total_stockout + transition_cost * total_transitions + production_cost * total_production
    stockout_cost = float(params.get('stockout_cost', default_params['stockout_cost']))
    transition_cost = float(params.get('transition_cost', default_params['transition_cost']))
    production_cost = float(params.get('production_cost', default_params['production_cost']))

    objective_terms = []

    # stockout terms
    total_stockout_vars = []
    for g in grade_list:
        gid = grade_to_idx[g]
        for d in range(T):
            total_stockout_vars.append(stockout[(gid, d)])
    if total_stockout_vars:
        # sum is linear expression
        objective_terms.append((stockout_cost, total_stockout_vars))

    # transition terms
    total_trans_vars = []
    for key, tvar in trans.items():
        total_trans_vars.append(tvar)
    if total_trans_vars:
        objective_terms.append((transition_cost, total_trans_vars))

    # production cost (sum across production_lgd)
    total_production_vars = []
    for key, pvar in production_lgd.items():
        total_production_vars.append(pvar)
    if total_production_vars and production_cost != 0:
        objective_terms.append((production_cost, total_production_vars))

    # Set linear objective: sum(coeff * var)
    if objective_terms:
        # flatten
        linear_terms = []
        for coeff, var_list in objective_terms:
            for v in var_list:
                linear_terms.append((coeff, v))
        # CP-SAT allows using linear expressions for minimization via Add(sum(...) * coeff)
        # Build a single linear expression: sum(coeff * var)
        total_obj = None
        # We'll create a linear expression by summing weighted terms
        linear_expr = sum([int(coeff) * v if isinstance(coeff, (int, np.integer)) else coeff * v for (coeff, v) in linear_terms])
        # However, cp_model expects coefficients to be integers when building objective; using floats is allowed but can be scaled.
        # To be safe scale all coefficients to integers by a factor
        # Find scaling factor to remove decimals up to 3 decimal places
        coeffs = [coeff for (coeff, _) in linear_terms]
        # compute scale
        max_decimals = 0
        for c in coeffs:
            s = str(float(c))
            if '.' in s:
                decimals = len(s.split('.')[-1].rstrip('0'))
                max_decimals = max(max_decimals, decimals)
        scale = 10 ** min(max_decimals, 6)
        scaled_terms = []
        for (coeff, v) in linear_terms:
            scaled_terms.append((int(round(coeff * scale)), v))
        # Set objective
        model.Minimize(sum(coef * var for (coef, var) in scaled_terms))
    else:
        # nothing to minimize: set a dummy objective (minimize 0)
        model.Minimize(0)

    # Solver parameters & stats
    solver = cp_model.CpSolver()
    # time limit
    tlimit = float(params.get('time_limit_minutes', default_params['time_limit_minutes']))
    if tlimit is not None and tlimit > 0:
        solver.parameters.max_time_in_seconds = float(tlimit * 60.0)
    # workers
    import multiprocessing
    workers = params.get('num_search_workers', None)
    if workers is None:
        workers = max(1, min(8, multiprocessing.cpu_count()))
    solver.parameters.num_search_workers = int(workers)
    # random seed
    seed = params.get('random_seed', 0)
    solver.parameters.random_seed = int(seed)
    # disable verbose search logging by default (fast)
    solver.parameters.log_search_progress = False
    # ensure presolve/probing enabled
    solver.parameters.cp_model_presolve = True
    solver.parameters.maximize = False

    # Print model statistics for profiling
    model_proto = model.Proto()
    num_variables = len(model_proto.variables)
    num_bools = sum(1 for v in model_proto.variables if v.domain == [0, 1])
    num_constraints = len(model_proto.constraints)
    print("MODEL STATS BEFORE SOLVE:")
    print(f"  Grades: {G}, Lines: {L}, Days: {T}")
    print(f"  Variables (proto): {num_variables} (booleans approx {num_bools})")
    print(f"  Constraints (proto): {num_constraints}")
    print(f"  Feasible line-days (sum): {sum(len(feasible_days_per_line[l]) for l in lines)}")
    print(f"  Time limit (min): {tlimit}, workers: {workers}, seed: {seed}")
    sys.stdout.flush()

    # Solve
    start_time = time.time()
    result = solver.Solve(model)
    elapsed = time.time() - start_time

    status = solver.StatusName(result)
    print(f"SOLVE STATUS: {status} (time {elapsed:.2f}s)")
    # Objective value scaled, if we scaled coefficients we need to divide by scale for printing. We used 'scale' if set above
    try:
        obj_val = solver.ObjectiveValue()
        # If we scaled the objective, divide by scale
        if 'scale' in locals() and scale > 1:
            obj_val_display = obj_val / float(scale)
        else:
            obj_val_display = obj_val
        print(f"  Objective (scaled): {obj_val_display}")
    except Exception:
        obj_val_display = None

    # Extract solution if solution found
    if result in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        # Build schedules
        schedule_rows = []
        inv_rows = []
        for d_idx, d_date in enumerate(sorted_dates):
            for line in lines:
                if d_idx not in feasible_days_per_line[line]:
                    # shutdown day
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Shutdown',
                        'Grade': None,
                        'Produced': 0
                    })
                    continue
                lg_val = solver.Value(line_grade[(line, d_idx)])
                if lg_val == IDLE:
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Idle',
                        'Grade': None,
                        'Produced': 0
                    })
                else:
                    grade_name = idx_to_grade[lg_val]
                    # produced quantity aggregated across production_lgd for this (line,grade,d)
                    prod_qty = 0
                    key = (line, lg_val, d_idx)
                    if key in production_lgd:
                        prod_qty = solver.Value(production_lgd[key])
                    schedule_rows.append({
                        'Date': d_date,
                        'Line': line,
                        'Status': 'Producing',
                        'Grade': grade_name,
                        'Produced': int(prod_qty)
                    })

            # inventory and supply/stockout
            for g in grade_list:
                gid = grade_to_idx[g]
                inv_val = solver.Value(inventory[(gid, d_idx)])
                sup_val = solver.Value(supplied[(gid, d_idx)])
                stv = solver.Value(stockout[(gid, d_idx)])
                prod_val = 0
                # produced_by_grade_day may be int or var
                pbd = produced_by_grade_day[(gid, d_idx)]
                if isinstance(pbd, int):
                    prod_val = pbd
                else:
                    prod_val = solver.Value(pbd)
                inv_rows.append({
                    'Date': d_date,
                    'Grade': g,
                    'Produced': int(prod_val),
                    'Supplied': int(sup_val),
                    'Stockout': int(stv),
                    'Inventory': int(inv_val)
                })

        schedule_df = pd.DataFrame(schedule_rows)
        inv_df = pd.DataFrame(inv_rows)

        # Print small summary
        total_stockout = inv_df['Stockout'].sum() if not inv_df.empty else 0
        total_produced = inv_df['Produced'].sum() if not inv_df.empty else 0
        total_transitions = sum(solver.Value(t) for t in trans.values()) if trans else 0
        print(f"  Total produced: {total_produced}, Total stockout: {total_stockout}, Transitions: {total_transitions}")
        # Save outputs if requested
        if params.get('write_outputs', True):
            out_pref = params.get('output_prefix', default_params['output_prefix'])
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                schedule_path = os.path.join(output_dir, f'{out_pref}_schedule.csv')
                inv_path = os.path.join(output_dir, f'{out_pref}_inventory.csv')
            else:
                schedule_path = f'{out_pref}_schedule.csv'
                inv_path = f'{out_pref}_inventory.csv'
            schedule_df.to_csv(schedule_path, index=False)
            inv_df.to_csv(inv_path, index=False)
            print(f"  Wrote schedule to: {schedule_path}")
            print(f"  Wrote inventory to: {inv_path}")
        return {
            'status': status,
            'objective': obj_val_display,
            'schedule_df': schedule_df,
            'inventory_df': inv_df,
            'solve_time_seconds': elapsed
        }
    else:
        print("No feasible solution found.")
        return {
            'status': status,
            'objective': None,
            'schedule_df': None,
            'inventory_df': None,
            'solve_time_seconds': elapsed
        }


# ----- CLI entrypoint -----
def main():
    parser = argparse.ArgumentParser(description="Optimized CP-SAT production scheduler")
    parser.add_argument('--input', '-i', type=str, default='input.xlsx', help='Input Excel file path')
    parser.add_argument('--time_limit', '-t', type=float, default=None, help='Time limit in minutes (overrides Params sheet)')
    parser.add_argument('--output_dir', '-o', type=str, default=None, help='Directory to write outputs (CSVs).')
    parser.add_argument('--debug', action='store_true', help='Debug mode (keeps logging verbose).')
    args = parser.parse_args()
    res = build_and_solve(args.input, time_limit_minutes=args.time_limit, output_dir=args.output_dir, debug=args.debug)
    if res.get('schedule_df') is not None:
        print("Sample schedule (first 10 rows):")
        print(res['schedule_df'].head(10).to_string(index=False))
    else:
        print("No schedule produced.")


if __name__ == '__main__':
    main()
