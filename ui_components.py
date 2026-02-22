import pandas as pd
import streamlit as st

def get_shift_color(shift_type):
    """Returns the CSS class or color variables for a specific shift type."""
    # Based on the user's reference image colors:
    # Morning: Yellow/Orange
    # Afternoon: Blue/Light Blue
    # Night: Dark Blue/Purple
    
    if shift_type in ['M', 'DM']:
        return "shift-morning"
    elif shift_type in ['A']:
        return "shift-afternoon"
    elif shift_type in ['N', 'DN']:
        return "shift-night"
    return "shift-default"

def get_shift_hebrew_name(shift_type):
    if shift_type in ['M', 'DM']: return "בוקר"
    if shift_type == 'A': return "צהריים"
    if shift_type in ['N', 'DN']: return "לילה"
    return shift_type

def render_schedule_html(roster_df):
    """
    Generates a high-fidelity HTML/CSS representation of the schedule,
    mimicking the reference image provided by the user.
    """
    
    # --- 1. Data Preparation ---
    unique_positions = roster_df['עמדה'].unique()
    sorted_days = sorted(roster_df['יום'].unique())
    
    # Define the structure of rows we want to show
    # Each tuple is (Internal Code List, Display Name, Icon)
    shift_structure = [
        (['M', 'DM'], "בוקר", "🔆"),
        (['A'], "צהריים", "🌤️"),
        (['N', 'DN'], "לילה", "🌙")
    ]

    # --- 2. CSS Styles ---
    # Using a professional, grid-based CSS system with RTL support
    css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;700&display=swap');
        
        :root {
            --bg-color: #f8f9fa;
            --header-bg: #2c3e50;
            --header-text: #ffffff;
            --border-color: #e9ecef;
            
            /* Shift Colors */
            --color-morning-bg: #fef3c7;
            --color-morning-border: #f59e0b;
            --color-morning-text: #92400e;
            
            --color-afternoon-bg: #e0f2fe;
            --color-afternoon-border: #0ea5e9;
            --color-afternoon-text: #075985;
            
            --color-night-bg: #ede9fe;
            --color-night-border: #8b5cf6;
            --color-night-text: #5b21b6;
        }
        
        .schedule-container {
            font-family: 'Heebo', sans-serif;
            direction: rtl;
            width: 100%;
            margin-bottom: 2rem;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            overflow: hidden;
            border: 1px solid var(--border-color);
        }

        .pos-header {
            background-color: #1e3a8a; /* Strong Blue like reference */
            color: white;
            padding: 12px 20px;
            font-size: 1.2rem;
            font-weight: 700;
            text-align: center;
            border-bottom: 3px solid #1e40af;
        }

        .schedule-table {
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }
        
        .schedule-table th {
            background-color: #f1f5f9;
            color: #475569;
            padding: 10px;
            font-size: 0.9rem;
            text-align: center;
            border-bottom: 2px solid var(--border-color);
            border-left: 1px solid var(--border-color);
        }
        
        .column-shift-label {
            width: 80px;
            background-color: #f8fafc;
            font-weight: bold;
            color: #64748b;
        }

        .schedule-table td {
            padding: 8px;
            border-bottom: 1px solid var(--border-color);
            border-left: 1px solid var(--border-color);
            vertical-align: top;
            height: 80px; /* Min height to look uniform */
            background-color: white;
        }

        /* Shift Card Styling */
        .shift-card {
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 6px 4px;
            border-radius: 6px;
            font-size: 0.85rem;
            font-weight: 500;
            white-space: nowrap;
            margin-bottom: 4px; /* Space if multiple people */
            transition: transform 0.2s;
            border-right: 4px solid transparent; /* Colored accent */
        }
        
        .shift-card:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .shift-morning {
            background-color: var(--color-morning-bg);
            color: var(--color-morning-text);
            border-right-color: var(--color-morning-border);
        }
        
        .shift-afternoon {
            background-color: var(--color-afternoon-bg);
            color: var(--color-afternoon-text);
            border-right-color: var(--color-afternoon-border);
        }
        
        .shift-night {
            background-color: var(--color-night-bg);
            color: var(--color-night-text);
            border-right-color: var(--color-night-border);
        }

        .shift-time {
            font-size: 0.75rem;
            opacity: 0.85;
            margin-top: 2px;
        }
        
        /* Specific tweaks */
        .row-header {
            font-weight: bold;
            text-align: center;
            vertical-align: middle !important;
            background: #f8fafc;
        }
        
        .badge-shift {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 12px;
            color: white;
            font-size: 0.7rem;
            margin-bottom: 4px;
        }
        .bg-morning { background-color: #f59e0b; }
        .bg-afternoon { background-color: #0ea5e9; }
        .bg-night { background-color: #8b5cf6; }

    </style>
    """

    # --- 3. Build HTML Structure ---
    html_content = [css]

    for pos in unique_positions:
        # Header
        html_content.append(f"""
        <div class="schedule-container">
            <div class="pos-header">{pos}</div>
            <table class="schedule-table">
                <thead>
                    <tr>
                        <th class="column-shift-label">משמרת</th>
        """)
        
        # Date Headers
        for d in sorted_days:
            # You might want to format the date nicely here
            html_content.append(f"<th>{d}</th>")
        
        html_content.append("</tr></thead><tbody>")
        
        # Rows
        for raw_codes, display_name, icon in shift_structure:
            row_id = f"{pos}_{display_name}"
            
            # Identify the badge color class for the Row Header
            badge_class = "bg-morning"
            if "צהריים" in display_name: badge_class = "bg-afternoon"
            elif "לילה" in display_name: badge_class = "bg-night"
            
            html_content.append(f"""
            <tr>
                <td class="row-header">
                    <span class="badge-shift {badge_class}">{display_name}</span>
                </td>
            """)
            
            for d in sorted_days:
                # Find workers for this cell
                pos_df = roster_df[roster_df['עמדה'] == pos]
                cell_workers = pos_df[
                    (pos_df['יום'] == d) & 
                    (pos_df['raw_shift'].isin(raw_codes))
                ]
                
                html_content.append("<td>")
                
                if not cell_workers.empty:
                    for _, w in cell_workers.iterrows():
                        # Determine Card Style
                        s_type = w['raw_shift']
                        css_class = get_shift_color(s_type)
                        
                        # Format Time
                        time_range = ""
                        if s_type == 'M': time_range = "07:00 - 15:00"
                        elif s_type == 'A': time_range = "15:00 - 23:00"
                        elif s_type == 'N': time_range = "23:00 - 07:00"
                        elif s_type == 'DM': time_range = "07:00 - 19:00 (כפולה)"
                        elif s_type == 'DN': time_range = "19:00 - 07:00 (כפולה)"
                        
                        html_content.append(f"""
                        <div class="shift-card {css_class}">
                            <span>{w['עובד']}</span>
                            <span class="shift-time">{time_range}</span>
                        </div>
                        """)
                
                html_content.append("</td>")
            
            html_content.append("</tr>")
        
        html_content.append("</tbody></table></div>")

    return "\n".join(html_content)
