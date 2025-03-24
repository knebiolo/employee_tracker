# -*- coding: utf-8 -*-
"""
Created on Sun Mar 23 13:29:42 2025

@author: Kevin.Nebiolo
"""

import streamlit as st
import altair as alt
import pandas as pd
import sqlite3
import os
import sys

# Add src/, data/, and scripts/ directories to sys.path for module import
dashboard_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.abspath(os.path.join(dashboard_dir, "..", "src"))
script_path = os.path.abspath(os.path.join(dashboard_dir, "..", "scripts"))
data_path = os.path.abspath(os.path.join(dashboard_dir, "..", "data"))
sys.path.append(src_path)
sys.path.append(data_path)
sys.path.append(script_path)

# Set page config for wide layout (the theme is loaded from .streamlit/config.toml)
st.set_page_config(page_title="Employee Dashboard", layout="wide")

#data ingestion script needs to run on app startup:
from spreadsheet_manage import main as run_spreadsheet_management
run_spreadsheet_management()

# Connect to your SQLite database (adjust path if needed)
conn = sqlite3.connect(os.path.join(data_path, "employee_tracker.db"))
employee_df = pd.read_sql("SELECT * FROM employee", conn)
employee_week = pd.read_sql("SELECT * FROM employee_week", conn)
employee_mtd = pd.read_sql("SELECT * FROM employee_mtd", conn)
employee_ytd = pd.read_sql("SELECT * FROM employee_ytd", conn)
section_week = pd.read_sql("SELECT * FROM section_week", conn)
firm_week = pd.read_sql("SELECT * FROM firm_week", conn)
# PATCH 1: Convert target_pct to numeric to avoid Arrow conversion errors
for df in [employee_week, employee_mtd, employee_ytd]:
    if "target_pct" in df.columns:
         df["target_pct"] = pd.to_numeric(df["target_pct"], errors="coerce")

# --------------------------------
# Top-level employee select box
# --------------------------------
# Create two columns for the dropdowns
col_emp, col_time = st.columns(2)

with col_emp:
    selected_employee = st.selectbox("Select Employee", employee_df["name"].tolist())

with col_time:
    selected_time_range = st.selectbox("Select Time Range", ["4 Weeks", "13 Weeks", "52 Weeks"])


# Get the corresponding employee_number
emp_id = employee_df[employee_df["name"] == selected_employee].employee_number.iloc[0]

# --- PATCH 2: Compute running averages for Employee Weekly Data ---

# Filter for the selected employee and compute running averages
emp_data = employee_week[employee_week["employee_id"] == emp_id].copy()
emp_data["period_ending"] = pd.to_datetime(emp_data["period_ending"])
emp_data.sort_values("period_ending", inplace=True)
emp_data["avg_4wk"] = emp_data["actual_pct"].rolling(window=4, min_periods=1).mean()
emp_data["avg_13wk"] = emp_data["actual_pct"].rolling(window=13, min_periods=1).mean()
emp_data["avg_52wk"] = emp_data["actual_pct"].rolling(window=52, min_periods=1).mean()

# Reshape running averages for plotting
emp_avg = emp_data[["period_ending", "avg_4wk", "avg_13wk", "avg_52wk"]].copy()
emp_avg = emp_avg.melt(id_vars="period_ending", var_name="window", value_name="running_avg")
emp_avg["window"] = emp_avg["window"].map({
    "avg_4wk": "4 Weeks",
    "avg_13wk": "13 Weeks",
    "avg_52wk": "52 Weeks"
})
emp_avg["entity"] = "Employee"
emp_data_to_concat = emp_data[["period_ending", "actual_pct"]].copy()
emp_data_to_concat['entity'] = 'Employee'
emp_data_to_concat['window'] = 'Weekly'

# For Section and Firm, use the weekly data as-is
section_data = section_week[section_week["section"] == "Mekong - Fisheries/Aquatic"].copy()
section_data["period_ending"] = pd.to_datetime(section_data["period_ending"])
section_data["entity"] = "Section"
section_data["window"] = "Weekly"

firm_week["period_ending"] = pd.to_datetime(firm_week["period_ending"])
firm_week["entity"] = "Firm"
firm_week["window"] = "Weekly"

# Combine into one DataFrame for plotting
# Combine into one DataFrame for plotting
df_all = pd.concat([
    emp_data_to_concat.rename(columns={"actual_pct": "running_avg"}),
    emp_avg,
    section_data[["period_ending", "actual_pct", "entity", "window"]].rename(columns={"actual_pct": "running_avg"}),
    firm_week[["period_ending", "actual_pct", "entity", "window"]].rename(columns={"actual_pct": "running_avg"})
], ignore_index=True)

# PATCH: Add dropdown to select the time range from the present
max_date = df_all["period_ending"].max()
if selected_time_range == "4 Weeks":
    cutoff = max_date - pd.Timedelta(weeks=4)
elif selected_time_range == "13 Weeks":
    cutoff = max_date - pd.Timedelta(weeks=13)
elif selected_time_range == "52 Weeks":
    cutoff = max_date - pd.Timedelta(weeks=52)
df_all = df_all[df_all["period_ending"] >= cutoff]

# --------------------------------
# Layout columns for chart & table
# --------------------------------

col_toggles, col_chart, col_table = st.columns([1,4,4])  # 1 : 6 : 1 ratio

with col_toggles:
    st.subheader("Toggle Lines")
    show_emp_curr = st.checkbox("Employee Current", value=True)
    show_emp_4wk = st.checkbox("Employee 4wk", value=True)
    show_emp_13wk = st.checkbox("Employee 13wk", value=True)
    show_emp_52wk = st.checkbox("Employee 52wk", value=True)
    show_section = st.checkbox("Section Weekly", value=True)
    show_firm = st.checkbox("Firm Weekly", value=True)

# Filter df_all based on toggle settings
def keep_row(row):
    if row["entity"] == "Employee":
        if row["window"] == "4 Weeks" and show_emp_4wk:
            return True
        if row["window"] == "13 Weeks" and show_emp_13wk:
            return True
        if row["window"] == "52 Weeks" and show_emp_52wk:
            return True
        if row["window"] == "Weekly" and show_emp_curr:
            return True        
        return False
    elif row["entity"] == "Section" and row["window"] == "Weekly":
        return show_section
    elif row["entity"] == "Firm" and row["window"] == "Weekly":
        return show_firm
    else:
        return False

mask = df_all.apply(keep_row, axis=1)
df_plot = df_all[mask].copy()

with col_chart:
    st.subheader("Utilization Over Time")

    chart = (
        alt.Chart(df_plot)
        .mark_line(point=True)
        .encode(
            x=alt.X("period_ending:T", title="Period Ending"),
            y=alt.Y("running_avg:Q", title="Utilization"),
            color=alt.Color(
                "entity:N",
                # Keep your custom color scale if desired
                scale=alt.Scale(
                    domain=["Employee", "Section", "Firm"],
                    range=["#F08080", "#90EE90", "#87CEFA"]
                ),
                legend=alt.Legend(title="Entity", orient="bottom", direction="horizontal")
            ),
            strokeDash=alt.StrokeDash(
                "window:N",
                scale=alt.Scale(
                    domain=["4 Weeks", "13 Weeks", "52 Weeks", "Weekly"],
                    range=[[4,2], [6,3], [10,5], [1,0]]
                ),
                legend=alt.Legend(title="Window", orient="bottom", direction="horizontal")
            ),
            tooltip=["entity", "window", "period_ending", "running_avg"]
        )
        .properties(width="container", height=500)
        .interactive()
    )

    st.altair_chart(chart, use_container_width=True)


with col_table:
    st.subheader("Employee Metrics")

    # Columns to display
    metrics_cols = [
        "target_pct", "actual_pct", "target_hrs", "direct", "indirect",
        "holiday", "paid_personal_leave", "admin", "marketing",
        "business_development", "proposal", "tech_development",
        "period_ending"
    ]

    # Mapping for nicer column names
    rename_mapping = {
        "target_pct": "Target %",
        "target_hrs": "Target",
        "direct": "Direct",
        "actual_pct": "Actual %",
        "indirect": "Indirect",
        "holiday": "Holiday",
        "paid_personal_leave": "PPL",
        "admin": "Admin",
        "marketing": "MKTG",
        "business_development": "BD",
        "proposal": "Proposal",
        "tech_development": "TD",
    }

    # ---------- Last 4 Weeks (Weekly Data) ----------
    last_date = emp_data["period_ending"].max()
    cutoff = last_date - pd.Timedelta(weeks=4)
    last4weeks = emp_data.loc[emp_data["period_ending"] >= cutoff, metrics_cols].copy()
    last4weeks.sort_values("period_ending", inplace=True)
    # Convert period_ending to MM-DD and rename column to 'period'
    last4weeks["period"] = last4weeks["period_ending"].dt.strftime("%m-%d")
    last4weeks.drop(columns=["period_ending"], inplace=True)
    last4weeks.rename(columns=rename_mapping, inplace=True)
    last4weeks.reset_index(drop=True, inplace=True)
    last4weeks.set_index("period", inplace=True)
    last4weeks = last4weeks.round(2)
    last4weeks = last4weeks.style.format("{:.2f}")
    st.markdown("**Last 4 Weeks**")
    st.table(last4weeks)

    # ---------- Month-to-Date (Single Latest Row) ----------
    emp_mtd_current = employee_mtd[employee_mtd["employee_id"] == emp_id].copy()
    # Limit to the desired columns
    emp_mtd_current = emp_mtd_current[metrics_cols].copy()
    
    emp_mtd_current["period_ending"] = pd.to_datetime(emp_mtd_current["period_ending"])
    emp_mtd_current.sort_values("period_ending", inplace=True)
    emp_mtd_current = emp_mtd_current.iloc[[-1]].copy()  # Only the latest row
    emp_mtd_current["period"] = emp_mtd_current["period_ending"].dt.strftime("%m-%d")
    emp_mtd_current.drop(columns=["period_ending"], inplace=True)
    emp_mtd_current.rename(columns=rename_mapping, inplace=True)
    emp_mtd_current.reset_index(drop=True, inplace=True)
    emp_mtd_current.set_index("period", inplace=True)
    emp_mtd_current = emp_mtd_current.round(2)
    
    # Format only numeric columns to avoid errors on strings
    numeric_cols = emp_mtd_current.select_dtypes(include=["number"]).columns
    emp_mtd_styled = emp_mtd_current.style.set_table_styles(
        [{"selector": "thead", "props": [("display", "none")]}]
    ).format({col: "{:.2f}" for col in numeric_cols})
    
    st.markdown("**Month-to-Date**")
    st.table(emp_mtd_styled)

    # ---------- Year-to-Date (Single Latest Row) ----------
    emp_ytd_current = employee_ytd[employee_ytd["employee_id"] == emp_id].copy()
    # Limit to the desired columns
    emp_ytd_current = emp_ytd_current[metrics_cols].copy()
    
    emp_ytd_current["period_ending"] = pd.to_datetime(emp_ytd_current["period_ending"])
    emp_ytd_current.sort_values("period_ending", inplace=True)
    emp_ytd_current = emp_ytd_current.iloc[[-1]].copy()  # Only the latest row
    emp_ytd_current["period"] = emp_ytd_current["period_ending"].dt.strftime("%m-%d")
    emp_ytd_current.drop(columns=["period_ending"], inplace=True)
    emp_ytd_current.rename(columns=rename_mapping, inplace=True)
    emp_ytd_current.reset_index(drop=True, inplace=True)
    emp_ytd_current.set_index("period", inplace=True)
    emp_ytd_current = emp_ytd_current.round(2)
    
    # Format only numeric columns to avoid errors on strings
    numeric_cols = emp_ytd_current.select_dtypes(include=["number"]).columns
    emp_ytd_styled = emp_ytd_current.style.set_table_styles(
        [{"selector": "thead", "props": [("display", "none")]}]
    ).format({col: "{:.2f}" for col in numeric_cols})
    
    st.markdown("**Year-to-Date**")
    st.table(emp_ytd_styled)
    
    
        
st.markdown("---")  # Separator between sections

with st.container():
    # This container will span the full width of the page
    col_section_ts, col_pie = st.columns([3, 1])  # 75% and 25% split

    with col_section_ts:
        st.subheader("Section Actual Utilization Time Series")
        # Allow selection of sections to include
        available_sections = section_week["section"].unique().tolist()
        selected_sections = st.multiselect("Select Sections", available_sections, default=available_sections)
        
        # Filter section data for the selected sections
        section_ts = section_week[section_week["section"].isin(selected_sections)].copy()
        section_ts["period_ending"] = pd.to_datetime(section_ts["period_ending"])
        # Use the same time range chosen above: extract weeks from selected_time_range
        weeks = int(selected_time_range.split()[0])
        max_date_section = section_ts["period_ending"].max()
        cutoff_section = max_date_section - pd.Timedelta(weeks=weeks)
        section_ts = section_ts[section_ts["period_ending"] >= cutoff_section]
        
        # Create a time series chart for section actual utilization percent.
        section_chart = alt.Chart(section_ts).mark_line(point=True).encode(
            x=alt.X("period_ending:T", title="Period Ending"),
            y=alt.Y("actual_pct:Q", title="Actual Utilization %"),
            color=alt.Color("section:N", title="Section"),
            tooltip=["section", "period_ending:T", "actual_pct:Q"]
        ).properties(width="container", height=400).interactive()
        
        st.altair_chart(section_chart, use_container_width=True)

    with col_pie:
        st.subheader("Employee Metrics Pie Chart")
        # Dropdown to choose which data source to use for the pie chart
        pie_option = st.selectbox("Select Data for Pie Chart", ["Weekly", "Month-to-Date", "Year-to-Date"])
        
        # Determine the latest row for the chosen option
        if pie_option == "Weekly":
            latest_emp = emp_data.sort_values("period_ending").iloc[-1]
        elif pie_option == "Month-to-Date":
            latest_emp = employee_mtd[employee_mtd["employee_id"] == emp_id].copy()
            latest_emp["period_ending"] = pd.to_datetime(latest_emp["period_ending"])
            latest_emp.sort_values("period_ending", inplace=True)
            latest_emp = latest_emp.iloc[-1]
        else:  # "Year-to-Date"
            latest_emp = employee_ytd[employee_ytd["employee_id"] == emp_id].copy()
            latest_emp["period_ending"] = pd.to_datetime(latest_emp["period_ending"])
            latest_emp.sort_values("period_ending", inplace=True)
            latest_emp = latest_emp.iloc[-1]
        
        # Create a DataFrame for the pie chart using the desired columns.
        # We use target_hrs for Target.
        pie_data = pd.DataFrame({
            "Metric": ["Direct", "Holiday", "PPL", "Admin", "MKTG", "BD", "Proposal", "TD"],
            "Value": [
                latest_emp["direct"],
                latest_emp["holiday"],
                latest_emp["paid_personal_leave"],
                latest_emp["admin"],
                latest_emp["marketing"],
                latest_emp["business_development"],
                latest_emp["proposal"],
                latest_emp["tech_development"]
            ]
        })
        
        pie_chart = alt.Chart(pie_data).mark_arc(innerRadius=20).encode(
            theta=alt.Theta(field="Value", type="quantitative"),
            color=alt.Color(field="Metric", type="nominal", legend=alt.Legend(title="Metric")),
            tooltip=["Metric", "Value:Q"]
        ).properties(width=300, height=300)
        
        st.altair_chart(pie_chart, use_container_width=True)




