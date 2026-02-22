from ortools.sat.python import cp_model
import pandas as pd

def solve_roster(employees_df, positions, constraints, col_map, shifts, avail_overrides=None, pref_weights=None, calc_potentials=False):
    """
    Main solver function with updated Double Shift logic.
    Double Morning (DM): 07:00-19:00 (Covers M + First half A)
    Double Night (DN): 19:00-07:00 (Covers Second half A + N)
    """
    model = cp_model.CpModel()
    
    # --- 1. Data Parsing ---
    emp_list = []
    
    # Safe helper for double shift column
    note_col = col_map.get('note')
    has_double_col = note_col and note_col not in ['None', 'ללא']
    
    for idx, row in employees_df.iterrows():
        name = row[col_map['name']]
        pos = row[col_map['pos']] 
        
        is_double = True # Default to allowed if no column specifies otherwise
        
        # Re-evaluating: The user wants to ignore the column IN THE APP. 
        # So col_map['note'] will be None.
        
        # However, to be safe:
        if has_double_col and note_col in employees_df.columns:
             val = str(row[note_col]).lower()
             if 'double' in val or 'כפולה' in val or 'כן' in val:
                 is_double = True
             else:
                 is_double = False
        
        # Availability parse
        avail = {}
        avail_from_override = {}  # tracks which days came from manual override
        override_df = avail_overrides.get(idx) if avail_overrides else None
        
        for s_col in shifts:
            day_avail = []
            
            # Check for override first
            clean_day_label = str(s_col).strip()
            
            # Defensive check
            is_valid_df = override_df is not None and hasattr(override_df, 'columns')
            if is_valid_df and clean_day_label in override_df.columns:
                # Use overridden values — these are MANUAL and must NOT be overridden by auto_doubles
                # Structure: 0=M, 1=A, 2=N, 3=DoubleM, 4=DoubleN
                try:
                    vals = override_df[clean_day_label].tolist()
                    if len(vals) >= 3:
                        if vals[0]: day_avail.append('M')
                        if vals[1]: day_avail.append('A')
                        if vals[2]: day_avail.append('N')
                    
                    # Double capability: read EXACTLY what the user set — never auto-override
                    if len(vals) >= 5:
                        if vals[3]: day_avail.append('Can_DM')
                        if vals[4]: day_avail.append('Can_DN')
                    
                    avail_from_override[s_col] = True  # mark as manually set
                except Exception:
                    avail_from_override[s_col] = False
            elif s_col in employees_df.columns:
                # Use original file data (Standard M/A/N)
                cell = str(row[s_col]).lower()
                if 'בוקר' in cell or 'morning' in cell: day_avail.append('M')
                if 'צהריים' in cell or 'afternoon' in cell: day_avail.append('A')
                if 'לילה' in cell or 'night' in cell: day_avail.append('N')
                
                # Global fallback for double shifts (only when NOT manually overridden)
                if is_double:
                    day_avail.append('Can_DM')
                    day_avail.append('Can_DN')
                avail_from_override[s_col] = False

            avail[s_col] = day_avail
            
        emp_list.append({
            'id': idx,
            'name': name,
            'pos': str(pos), 
            'avail': avail,
            'avail_from_override': avail_from_override  # which days were manually set
        })
    
    # Handle case with no employees
    if not emp_list:
        return {'status': 'No Employees Found', 'roster': None}

    # --- 2. Variables & Constraints ---
    SHIFTS = ['M', 'A', 'N', 'DM', 'DN']
    assignments = {} 
    
    # Objective tracking
    all_double_shifts = []
    slacks = []

    for p_idx, pos_data in enumerate(positions):
        pos_name = pos_data['name']
        req_m = pos_data['guards_morning']
        req_a = pos_data['guards_afternoon']
        req_n = pos_data['guards_night']
        
        for d in shifts:
            # vars for this pos/day by shift type
            pos_day_vars = {'M': [], 'A': [], 'N': [], 'DM': [], 'DN': []}

            for e in emp_list:
                # Qualification check (Flexible containment)
                emp_pos_str = str(e['pos']) if e['pos'] else ""
                # Normalize employee roles: split, strip, lower
                emp_roles = [r.strip().lower() for r in emp_pos_str.split(',') if r.strip()]
                
                # Check if this position is in their allowed roles
                # We normalize the required position name as well
                norm_pos_name = pos_name.strip().lower()
                
                is_qualified = False
                if 'all' in emp_roles:
                    is_qualified = True
                else:
                    # Flexible Match:
                    # 1. Exact token match
                    if norm_pos_name in emp_roles:
                        is_qualified = True
                    # 2. Substring match (e.g. "Security" in "Head of Security")
                    # We check if the expected Position Name is contained in any of the employee's roles
                    # OR if any of the employee's roles is contained in the Position Name.
                    elif any(norm_pos_name in r or r in norm_pos_name for r in emp_roles):
                        is_qualified = True
                
                if not is_qualified:
                    continue
                
                day_av = e['avail'].get(d, [])
                
                # Single Shifts
                if 'M' in day_av:
                    v = model.NewBoolVar(f"x_{e['id']}_{p_idx}_{d}_M")
                    assignments[(e['id'], p_idx, d, 'M')] = v
                    pos_day_vars['M'].append(v)
                
                if 'A' in day_av:
                    v = model.NewBoolVar(f"x_{e['id']}_{p_idx}_{d}_A")
                    assignments[(e['id'], p_idx, d, 'A')] = v
                    pos_day_vars['A'].append(v)
                
                if 'N' in day_av:
                    v = model.NewBoolVar(f"x_{e['id']}_{p_idx}_{d}_N")
                    assignments[(e['id'], p_idx, d, 'N')] = v
                    pos_day_vars['N'].append(v)
                
                # Double Shifts
                # Rule: auto_doubles only kicks in when the day was NOT manually overridden by the user.
                # If the user manually set Can_DM/Can_DN to False in the table → we MUST respect that.
                day_was_overridden = e.get('avail_from_override', {}).get(d, False)

                can_do_dm = 'Can_DM' in day_av
                if not day_was_overridden and constraints.get('auto_doubles', False) and 'M' in day_av:
                    # No manual override for this day → safe to apply global auto_doubles rule
                    can_do_dm = True

                if constraints['allow_double'] and can_do_dm:
                    v = model.NewBoolVar(f"x_{e['id']}_{p_idx}_{d}_DM")
                    assignments[(e['id'], p_idx, d, 'DM')] = v
                    pos_day_vars['DM'].append(v)
                    all_double_shifts.append(v)
                
                can_do_dn = 'Can_DN' in day_av
                if not day_was_overridden and constraints.get('auto_doubles', False) and 'N' in day_av:
                    # No manual override for this day → safe to apply global auto_doubles rule
                    can_do_dn = True

                if constraints['allow_double'] and can_do_dn:
                    v = model.NewBoolVar(f"x_{e['id']}_{p_idx}_{d}_DN")
                    assignments[(e['id'], p_idx, d, 'DN')] = v
                    pos_day_vars['DN'].append(v)
                    all_double_shifts.append(v)
            
            # Coverage Constraints with Slack (Priority-weighted Soft Constraints)
            # Penalty is scaled by position priority and shift priority.
            # Priority 1 (most important) -> highest penalty -> solver fills it first.
            # Formula: base * pos_weight * shift_weight
            # pos_weight  = 1/pos_priority  (priority 1 -> weight=1.0, priority 5 -> weight=0.2)
            # shift_weight = 1/shift_priority
            BASE_PENALTY = 10000
            pos_priority = pos_data.get('priority', 5)  # 1=most important
            pm_priority  = pos_data.get('priority_morning', 1)
            pa_priority  = pos_data.get('priority_afternoon', 1)
            pn_priority  = pos_data.get('priority_night', 1)

            def _w(pos_p, shift_p):
                """Convert 1-based priority to integer penalty weight.
                Lower number = more important = higher penalty.
                We invert: weight = BASE / (pos_p * shift_p)
                Then round to int for CP-SAT (integer coefficients required)."""
                return max(1, int(BASE_PENALTY / (pos_p * shift_p)))

            # Morning
            if req_m > 0:
                slack_m = model.NewIntVar(0, req_m, f"slack_{p_idx}_{d}_M")
                model.Add(sum(pos_day_vars['M'] + pos_day_vars['DM']) + slack_m == req_m)
                slacks.append((f"{d}|{pos_name}|בוקר", slack_m, _w(pos_priority, pm_priority)))
            
            # Afternoon
            if req_a > 0:
                slack_a1 = model.NewIntVar(0, req_a, f"slack_{p_idx}_{d}_A1")
                model.Add(sum(pos_day_vars['A'] + pos_day_vars['DM']) + slack_a1 == req_a)
                slacks.append((f"{d}|{pos_name}|צהריים (15:00-19:00)", slack_a1, _w(pos_priority, pa_priority)))

                slack_a2 = model.NewIntVar(0, req_a, f"slack_{p_idx}_{d}_A2")
                model.Add(sum(pos_day_vars['A'] + pos_day_vars['DN']) + slack_a2 == req_a)
                slacks.append((f"{d}|{pos_name}|צהריים (19:00-23:00)", slack_a2, _w(pos_priority, pa_priority)))

            # Night
            if req_n > 0:
                slack_n = model.NewIntVar(0, req_n, f"slack_{p_idx}_{d}_N")
                model.Add(sum(pos_day_vars['N'] + pos_day_vars['DN']) + slack_n == req_n)
                slacks.append((f"{d}|{pos_name}|לילה", slack_n, _w(pos_priority, pn_priority)))

    # --- 3. Employee Global Constraints ---
    for e in emp_list:
        for i, d in enumerate(shifts):
            # Gather assignments for (e, d) across all positions/shifts
            day_vars = []
            for p_idx in range(len(positions)):
                for s in SHIFTS:
                    if (e['id'], p_idx, d, s) in assignments:
                        day_vars.append(assignments[(e['id'], p_idx, d, s)])
            
            if not day_vars: continue

            # Max 1 shift per day (No Overlap)
            model.Add(sum(day_vars) <= 1)
            
            # No Back-to-Back (Night -> Next Morning)
            if constraints['no_back_to_back'] and i < len(shifts) - 1:
                d_next = shifts[i+1]
                
                # Night today (N or DN)
                night_today = []
                for p_idx in range(len(positions)):
                    if (e['id'], p_idx, d, 'N') in assignments: night_today.append(assignments[(e['id'], p_idx, d, 'N')])
                    if (e['id'], p_idx, d, 'DN') in assignments: night_today.append(assignments[(e['id'], p_idx, d, 'DN')])
                
                # Morning next (M or DM)
                morning_next = []
                for p_idx in range(len(positions)):
                    if (e['id'], p_idx, d_next, 'M') in assignments: morning_next.append(assignments[(e['id'], p_idx, d_next, 'M')])
                    if (e['id'], p_idx, d_next, 'DM') in assignments: morning_next.append(assignments[(e['id'], p_idx, d_next, 'DM')])
                
                if night_today and morning_next:
                    model.Add(sum(night_today) + sum(morning_next) <= 1)

    # --- Objective: Minimize Weighted Slack (Highest Priority) + Maximize Participation (Secondary) + Preference Bonus (Tertiary) --- 
    # Priority 1: Fill all positions (avoid big slack penalties ~1,000-10,000).
    # Priority 2: Maximize number of assigned shifts (distribute work, +1 each).
    # Priority 3: Position Preference Bonus (0-10 per assignment, much smaller than slack penalty).
    #
    # Since max preference = 10 and min slack penalty = 1,000:
    # Preferences can NEVER override coverage. They only break ties.

    all_assigned_vars = list(assignments.values())
    
    obj_terms = []
    if all_assigned_vars:
        obj_terms.extend(all_assigned_vars)  # +1 bonus per ANY shift assigned

    # --- Preference Bonus Terms ---
    # Note: emp 'id' is the DataFrame index, so pref_weights keys match directly
    if pref_weights:
        for (eid, p_idx, d, s), var in assignments.items():
            pos_name = positions[p_idx]['name']
            emp_prefs = pref_weights.get(eid, {})
            score = emp_prefs.get(pos_name, 0)
            if score > 0:
                obj_terms.append(score * var)

    if slacks:
        penalty_terms = [w * s_var for (_, s_var, w) in slacks]
        model.Maximize(sum(obj_terms) - sum(penalty_terms))
    elif obj_terms:
         # No slacks defined (unlikely), just maximize assignments
        model.Maximize(sum(obj_terms))

    # --- Solvers ---
    solver = cp_model.CpSolver()
    status = solver.Solve(model)
    
    result = {'status': solver.StatusName(status), 'roster': None, 'diagnostics': []}
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        data = []
        for (eid, pid, d, s), var in assignments.items():
            if solver.Value(var):
                emp_name = next(e['name'] for e in emp_list if e['id'] == eid)
                pos_name = positions[pid]['name']
                
                # Convert Shift Code to Display Name
                s_display = s
                if s == 'M': s_display = 'בוקר (07-15)'
                elif s == 'A': s_display = 'צהריים (15-23)'
                elif s == 'N': s_display = 'לילה (23-07)'
                elif s == 'DM': s_display = 'כפולה בוקר (07-19)'
                elif s == 'DN': s_display = 'כפולה לילה (19-07)'
                
                data.append({
                    "יום": d,
                    "עמדה": pos_name,
                    "משמרת": s_display,
                    "raw_shift": s,
                    "עובד": emp_name
                })
        
        # NOTE: roster DataFrame is created AFTER slacks loop below,
        # so that shortage rows injected during slack analysis are included.
            
        # Check Slacks for Warnings & Stats
        shortage_summary = {}
        # gap_recommendations will be grouped by shortage
        gap_recs_by_shortage = {}
        injected_shortages = set()  # Track (day, pos, shift_group) to avoid duplicate shortage rows
        
        # 1. Map assigned shifts per employee per day for fast lookup
        # Format: emp_assignments[emp_id][day] = list of assigned shift types ('M', 'N', 'DM'...)
        emp_assignments_map = {e['id']: {} for e in emp_list}
        for (eid, pid, d, s), var in assignments.items():
            if solver.Value(var):
                if d not in emp_assignments_map[eid]:
                    emp_assignments_map[eid][d] = []
                emp_assignments_map[eid][d].append(s)

        # 2. Analyze Slacks
        for slack_tuple in slacks:
            # Tuple structure: (label, var, weight)
            # Label format: "Day|PosName|ShiftType" e.g. "ראשון|שער ראשי|בוקר"
            label, s_var, _ = slack_tuple
            val = solver.Value(s_var)
            
            if val > 0:
                shortage_summary[label] = val
                result['diagnostics'].append(f"חסר/ים {val} עובדים ב: {label}")
                
                # --- Gap Filling Logic & Roster Injection ---
                # Parse label to understand what is missing
                parts = label.split('|')
                if len(parts) >= 3:
                   miss_day = parts[0]
                   miss_pos = parts[1]
                   miss_type_desc = parts[2]
                   
                   # Determine shift GROUP for deduplication
                   shift_group = 'other'
                   if 'בוקר' in miss_type_desc: shift_group = 'בוקר'
                   elif 'צהריים' in miss_type_desc: shift_group = 'צהריים'
                   elif 'לילה' in miss_type_desc: shift_group = 'לילה'

                   # Clean display shift name
                   display_shift = shift_group
                   if shift_group == 'בוקר': display_shift = 'בוקר (07:00-15:00)'
                   elif shift_group == 'צהריים': display_shift = 'צהריים (15:00-23:00)'
                   elif shift_group == 'לילה': display_shift = 'לילה (23:00-07:00)'
                   
                   dedup_key = (miss_day, miss_pos, shift_group)
                   
                   # Only inject ONE shortage row per raw shortage group
                   if dedup_key not in injected_shortages:
                       injected_shortages.add(dedup_key)
                       
                       data.append({
                            "יום": miss_day,
                            "עמדה": miss_pos,
                            "משמרת": display_shift,
                            "raw_shift": "SHORTAGE",
                            "עובד": f"⚠️ חוסר ({val})" 
                       })
                       
                   # --- Recommendation Logic PER UNIQUE SHORTAGE ---
                   if calc_potentials:
                       # Only generate recs if this is the first time we see this shortage group
                       # OR if we want to allow partial suggestions (let's stick to unique group to assume coverage)
                       # Actually, simple slacks loop is fine, but result key should be readable.
                       
                       shortage_key = f"{miss_pos} | {miss_day} | {display_shift}"
                       if shortage_key not in gap_recs_by_shortage:
                           gap_recs_by_shortage[shortage_key] = {'available': [], 'potential': []}

                       # Determine required shift code based on description
                       req_code = '?' 
                       if 'בוקר' in miss_type_desc: req_code = 'M'
                       if 'צהריים' in miss_type_desc: req_code = 'A'
                       if 'לילה' in miss_type_desc: req_code = 'N'
                       
                       # Scan all employees for potential match for THIS shortage
                       for e in emp_list:
                           eid = e['id']
                           ename = e['name']
                           
                           # A. Check Explicit Availability
                           day_avail = e['avail'].get(miss_day, [])
                           is_explicitly_available = req_code in day_avail
                           
                           # B. Check Hard Constraints
                           occupied_today = emp_assignments_map[eid].get(miss_day, [])
                           if occupied_today: continue
                               
                           # Rest Rules
                           violates_rest = False
                           if req_code == 'M':
                               try:
                                   d_idx = shifts.index(miss_day)
                                   if d_idx > 0:
                                       prev_day = shifts[d_idx - 1]
                                       prev_shifts = emp_assignments_map[eid].get(prev_day, [])
                                       if 'N' in prev_shifts or 'DN' in prev_shifts: violates_rest = True
                               except ValueError: pass
                           if req_code == 'N':
                               try:
                                   d_idx = shifts.index(miss_day)
                                   if d_idx < len(shifts) - 1:
                                       next_day = shifts[d_idx + 1]
                                       next_shifts = emp_assignments_map[eid].get(next_day, [])
                                       if 'M' in next_shifts or 'DM' in next_shifts: violates_rest = True
                               except ValueError: pass
                
                           if violates_rest: continue
                
                           # C. Categorize Recommendation
                           # Simplify message since it's grouped under the shortage header now
                           rec_msg = f"**{ename}**" 
                           
                           if is_explicitly_available:
                               gap_recs_by_shortage[shortage_key]['available'].append(rec_msg)
                           else:
                               gap_recs_by_shortage[shortage_key]['potential'].append(rec_msg)

        result['shortage_summary'] = shortage_summary
        result['gap_recommendations'] = gap_recs_by_shortage

        # NOW build the roster DataFrame (after shortage rows were appended to data)
        result['roster'] = pd.DataFrame(data)
        if not result['roster'].empty:
            result['roster'] = result['roster'].sort_values(by=["יום", "עמדה", "משמרת"])
        # --- Surplus Report: Available but not assigned ---
        # Employees who marked availability but weren't scheduled (all positions full)
        surplus_report = {}  # { day: [ {'name': str, 'shifts': str} ] }
        
        SHIFT_DISPLAY = {'M': 'בוקר', 'A': 'צהריים', 'N': 'לילה'}
        
        for e in emp_list:
            eid = e['id']
            ename = e['name']
            for d_i, day in enumerate(shifts):
                day_avail = e['avail'].get(day, [])
                # Only count real shift availability (M, A, N), not Can_DM/Can_DN
                real_avail = [s for s in day_avail if s in ('M', 'A', 'N')]
                
                if not real_avail:
                    continue  # Not available this day
                    
                # Check if assigned anywhere this day
                assigned_today = emp_assignments_map[eid].get(day, [])
                
                if not assigned_today:
                    # Before marking as surplus, filter out shifts blocked by rest constraints
                    # Rule: Night/DN yesterday → Morning today is blocked
                    truly_available = list(real_avail)
                    
                    if constraints.get('no_back_to_back', False) and d_i > 0:
                        prev_day = shifts[d_i - 1]
                        prev_assigned = emp_assignments_map[eid].get(prev_day, [])
                        worked_night_yesterday = any(s in ('N', 'DN') for s in prev_assigned)
                        
                        if worked_night_yesterday:
                            # Remove morning shifts — blocked by rest rule, not surplus
                            truly_available = [s for s in truly_available if s not in ('M',)]
                    
                    if not truly_available:
                        continue  # All availability was blocked by rest rules → not surplus
                    
                    # Available but NOT assigned → Surplus!
                    avail_display = ", ".join([SHIFT_DISPLAY.get(s, s) for s in truly_available])
                    if day not in surplus_report:
                        surplus_report[day] = []
                    surplus_report[day].append({
                        'name': ename,
                        'shifts': avail_display
                    })
        
        result['surplus_report'] = surplus_report

        return result
