
# -*- coding: utf-8 -*-
"""
Module: page_reader
Contains helper functions to extract, process, and store employee utilization data.
Author: Kevin Nebiolo
"""

import pandas as pd
import os
import numpy as np
from collections import defaultdict
from sqlalchemy import create_engine
import shutil

def get_columns(page):
    """
    Extracts non-null column names and their index positions from the first row of the page.
    Standardizes column names for employee data to PEP 8 compliant names.
    """
    column_names = page.iloc[0].tolist()
    col_name_idx = {}
    # Mapping for employee sheet columns to PEP 8 compliant names
    mapping = {
        "Target": "target_pct",           # assume target is expressed as a percentage
        "Actual": "actual_pct",           # actual percentage
        "Direct": "direct",
        "Trgt\nHrs": "target_hrs",
        "Indirect": "indirect",
        "Hol": "holiday",
        "PPL": "paid_personal_leave",
        "Admin": "admin",
        "Mktg": "marketing",
        "BD": "business_development",
        "Prop": "proposal",
        "TD": "tech_development",
        "UPL/PTL": "upl_ptl",
        "Hours": "hours",
        "Stand": "standard"
    }
    for i, val in enumerate(column_names):
        if pd.notna(val):
            # Use the mapping if available, otherwise convert to a lower-case string with underscores replacing spaces
            standardized = mapping.get(val, val.strip().lower().replace(" ", "_"))
            col_name_idx[standardized] = i
    return col_name_idx

def get_firm_columns(page):
    """
    Extracts column name indices from a firm-level summary page.
    Standardizes column names for firm data to PEP 8 compliant names.
    """
    column_names = page.iloc[7].tolist()
    col_name_idx = {}
    # Mapping for firm sheet columns to PEP 8 compliant names
    mapping = {
        "Target\n%": "target_pct",
        "Actual\n%": "actual_pct",
        "Direct": "direct",
        "Target": "target_hrs",
        "Indirect": "indirect",
        "Hol": "holiday",
        "PPL": "paid_personal_leave",
        "Admin": "admin",
        "Mktg": "marketing",
        "BD": "business_development",
        "Prop": "proposal",
        "TD": "tech_development",
        "UPL/PTL": "upl_ptl",
        "Total": "total",
        "Std": "standard"
    }
    for i, val in enumerate(column_names):
        if pd.notna(val):
            standardized = mapping.get(val, val.strip().lower().replace(" ", "_"))
            col_name_idx[standardized] = i
    return col_name_idx


def get_firm_pay_period(page):
    """
    Extracts the ending date of the pay period from a firm-level summary page.
    
    Parameters:
        page (pd.DataFrame): The Excel sheet containing firm-level data.
    
    Returns:
        str: Period end date string.
    """
    for idx, row in page.iterrows():
        for i in row:
            if isinstance(i, str) and 'For the period' in i:
                period_idx = i.index('-')
                period_end = i[period_idx + 2:]
                break

def get_pay_period(page):
    """
    Extracts the ending date of the pay period from the column headers.

    Parameters:
        page (pd.DataFrame): The raw Excel sheet content.

    Returns:
        str: Period end date string.
    """
    for col in page.columns:
        if 'For the period' in col:
            period_idx = col.index('-')
            return col[period_idx + 2:]
    return None


def get_employees(page):
    """
    Extracts employee identifiers and names from the page content.

    Parameters:
        page (pd.DataFrame): The raw Excel sheet content.

    Returns:
        tuple: (employee_rows dict, employee_df DataFrame)
    """
    employee_rows = {}
    employee_data = {}

    for idx, row in page.iterrows():
        if 'Employee Number:' in str(row.iloc[0]):
            colon = row.iloc[0].index(':')
            employee_name = row.iloc[0][colon + 7:]
            number = row.iloc[0][colon + 2: colon + 6]
            last, first = employee_name.split(',')
            employee_rows[number] = idx
            employee_data[number] = [employee_name.strip(), last.strip(), first.strip(),number]

    employee_df = pd.DataFrame.from_dict(employee_data, orient='index', columns=['name', 'last', 'first','employee_number'])
    employee_df.index.name = "employee_id"
    return employee_rows, employee_df

def get_sections(page):
    """
    Extracts section identifiers and names from the page content.
    
    Parameters:
        page (pd.DataFrame): The raw Excel sheet content.
    
    Returns:
        tuple: 
            - section_rows (dict): Mapping of section names to row indices.
            - section_df (pd.DataFrame): Section metadata as a DataFrame.
    """
    section_rows = {}
    section_data = {}
    
    for idx, row in page.iterrows():
        if 'Section:' in str(row.iloc[0]):
            colon = row.iloc[0].index(':')
            section_name = row.iloc[0][colon + 2:]
            section_rows[section_name] = idx
            section_data[section_name] = [section_name]

    section_df = pd.DataFrame.from_dict(section_data, orient='index', columns=['section_name',])
    section_df.index.name = "section"
    
    return section_rows, section_df
    

def get_firm_totals(page):
    """
    Finds the row index that contains the firm's final totals.
    
    Parameters:
        page (pd.DataFrame): The Excel sheet containing firm-level data.
    
    Returns:
        int: Index of the row labeled 'Final Totals'.
    """
    #get company totals        
    for idx, row in page.iterrows():
        if 'Final Totals' in str(row.iloc[0]):
            company_row = idx
            break
    
    return company_row

def get_employee_data(page, period_end, employee_rows, col_name_idx):
    """
    Extracts week, MTD, and YTD payroll data per employee.

    Parameters:
        page (pd.DataFrame): The raw Excel sheet content.
        period_end (str): Pay period end date.
        employee_rows (dict): Mapping of employee IDs to row indices.
        col_name_idx (dict): Mapping of column names to their indices.

    Returns:
        tuple: week_df, mtd_df, ytd_df DataFrames.
    """
    week_dict = defaultdict(list)
    mtd_dict = defaultdict(list)
    ytd_dict = defaultdict(list)

    for employee in employee_rows:
        idx = employee_rows[employee]
        week_row = page.iloc[idx + 1]
        mtd_row = page.iloc[idx + 2]
        ytd_row = page.iloc[idx + 3]

        for col in col_name_idx:
            cidx = col_name_idx[col]
            week_dict[employee].append(week_row.iloc[cidx])
            mtd_dict[employee].append(mtd_row.iloc[cidx])
            ytd_dict[employee].append(ytd_row.iloc[cidx])

    week_df = pd.DataFrame.from_dict(week_dict, orient='index', columns=col_name_idx.keys())
    mtd_df = pd.DataFrame.from_dict(mtd_dict, orient='index', columns=col_name_idx.keys())
    ytd_df = pd.DataFrame.from_dict(ytd_dict, orient='index', columns=col_name_idx.keys())
    week_df.loc[:, 'actual_pct'] = week_df['actual_pct'] / 100.
    for df in [week_df, mtd_df, ytd_df]:
        df['period_ending'] = pd.to_datetime(period_end)
        df.index.name = "employee_id"
        print (df.head(10))
        
    return week_df, mtd_df, ytd_df

def get_section_data(page, period_end, section_rows, col_name_idx):
    """
    Extracts week, MTD, and YTD payroll data per section.
    
    Parameters:
        page (pd.DataFrame): The raw Excel sheet content.
        period_end (str): Pay period end date.
        section_rows (dict): Mapping of section names to row indices.
        col_name_idx (dict): Mapping of column names to their indices.
    
    Returns:
        tuple: section_week_df, section_mtd_df, section_ytd_df DataFrames.
    """

    section_week_dict = defaultdict(list)
    section_mtd_dict = defaultdict(list)
    section_ytd_dict = defaultdict(list)
    
    for section in section_rows:
        idx = section_rows[section]
        week_row = page.iloc[idx + 1]
        mtd_row = page.iloc[idx + 2]
        ytd_row = page.iloc[idx + 3]
    
        for col in col_name_idx:
            cidx = col_name_idx[col]
            section_week_dict[section].append(week_row.iloc[cidx])
            section_mtd_dict[section].append(mtd_row.iloc[cidx])
            section_ytd_dict[section].append(ytd_row.iloc[cidx])
    
    section_week_df = pd.DataFrame.from_dict(section_week_dict, orient='index', columns=col_name_idx.keys())
    section_mtd_df = pd.DataFrame.from_dict(section_mtd_dict, orient='index', columns=col_name_idx.keys())
    section_ytd_df = pd.DataFrame.from_dict(section_ytd_dict, orient='index', columns=col_name_idx.keys())
    
    for df in [section_week_df, section_mtd_df, section_ytd_df]:
        df['period_ending'] = pd.to_datetime(period_end)
        df.index.name = "section"
    
    return section_week_df, section_mtd_df, section_ytd_df

def get_firm_data(page, period_end, company_row, col_name_idx):
    """
    Extracts firm-level weekly, MTD, and YTD payroll totals.
    
    Parameters:
        page (pd.DataFrame): The Excel sheet containing firm-level data.
        period_end (str): Pay period end date.
        company_row (int): Row index corresponding to firm totals.
        col_name_idx (dict): Mapping of column names to their indices.
    
    Returns:
        tuple: firm_week_df, firm_mtd_df, firm_ytd_df DataFrames.
    """

    firm_week_dict = defaultdict(list)
    firm_mtd_dict = defaultdict(list)
    firm_ytd_dict = defaultdict(list)

    firm_week_row = page.iloc[company_row + 1]
    firm_mtd_row = page.iloc[company_row + 2]
    firm_ytd_row = page.iloc[company_row + 3]

    for col in col_name_idx:
        cidx = col_name_idx[col]
        firm_week_dict[period_end].append(firm_week_row.iloc[cidx])
        firm_mtd_dict[period_end].append(firm_mtd_row.iloc[cidx])
        firm_ytd_dict[period_end].append(firm_ytd_row.iloc[cidx])

    firm_week_df = pd.DataFrame.from_dict(firm_week_dict, orient='index', columns=col_name_idx.keys())
    firm_mtd_df = pd.DataFrame.from_dict(firm_mtd_dict, orient='index', columns=col_name_idx.keys())
    firm_ytd_df = pd.DataFrame.from_dict(firm_ytd_dict, orient='index', columns=col_name_idx.keys())

    for df in [firm_week_df, firm_mtd_df, firm_ytd_df]:
        df['period_ending'] = pd.to_datetime(period_end)
        
    return firm_week_df, firm_mtd_df, firm_ytd_df

def write_data(proj_dir, 
               period_end, 
               employee_df, 
               emp_week_df, 
               emp_mtd_df, 
               emp_ytd_df,
               sec_week_df,
               sec_mtd_df,
               sec_ytd_df,
               firm_week_df,
               firm_mtd_df,
               firm_ytd_df):
    
    """
    Writes employee, section, and firm payroll data to SQLite database.
    
    Parameters:
        proj_dir (str): Project root directory.
        period_end (str): Pay period end date.
        employee_df (pd.DataFrame): Employee master data.
        emp_week_df (pd.DataFrame): Employee weekly data.
        emp_mtd_df (pd.DataFrame): Employee month-to-date data.
        emp_ytd_df (pd.DataFrame): Employee year-to-date data.
        sec_week_df (pd.DataFrame): Section weekly data.
        sec_mtd_df (pd.DataFrame): Section month-to-date data.
        sec_ytd_df (pd.DataFrame): Section year-to-date data.
        firm_week_df (pd.DataFrame): Firm weekly data.
        firm_mtd_df (pd.DataFrame): Firm month-to-date data.
        firm_ytd_df (pd.DataFrame): Firm year-to-date data.
    """

    engine = create_engine(f"sqlite:///{os.path.join(proj_dir, 'data', 'employee_tracker.db')}")

    try:
        existing_ids = set(pd.read_sql("SELECT employee_id FROM employee", engine)['employee_id'])
    except Exception:
        existing_ids = set()

    new_employees = employee_df[~employee_df.index.isin(existing_ids)]
    if not new_employees.empty:
        new_employees.to_sql("employee", engine, if_exists="append", index=True)
        
    # PATCH: Rename columns in employee data DataFrames to match the new standardized schema.
    rename_mapping = {
        "direct\nhours": "direct",
        "indirect\nhours": "indirect",
        "hours\nhours": "hours"
    }
    emp_week_df.rename(columns=rename_mapping, inplace=True)
    emp_mtd_df.rename(columns=rename_mapping, inplace=True)
    emp_ytd_df.rename(columns=rename_mapping, inplace=True)

    # PATCH for section tables: Rename columns to match the database schema
    rename_section = {
        "direct\namount": "direct",
        "indirect\namount": "indirect",
        "total\namount": "total"
    }

    sec_week_df.rename(columns=rename_section, inplace=True)
    sec_mtd_df.rename(columns=rename_section, inplace=True)
    sec_ytd_df.rename(columns=rename_section, inplace=True)
    firm_week_df.rename(columns=rename_section, inplace=True)
    firm_mtd_df.rename(columns=rename_section, inplace=True)
    firm_ytd_df.rename(columns=rename_section, inplace=True)

    emp_week_df.to_sql("employee_week", engine, if_exists="append", index=True)
    emp_mtd_df.to_sql("employee_mtd", engine, if_exists="append", index=True)
    emp_ytd_df.to_sql("employee_ytd", engine, if_exists="append", index=True)
    sec_week_df.to_sql("section_week", engine, if_exists="append", index=True)
    sec_mtd_df.to_sql("section_mtd", engine, if_exists="append", index=True)
    sec_ytd_df.to_sql("section_ytd", engine, if_exists="append", index=True)
    firm_week_df.to_sql("firm_week", engine, if_exists="append", index=True)
    firm_mtd_df.to_sql("firm_mtd", engine, if_exists="append", index=True)
    firm_ytd_df.to_sql("firm_ytd", engine, if_exists="append", index=True)

    print(f"Data management for period ending {period_end} complete.")

def clean_up(proj_dir, worksheet):
    """
    Moves a processed spreadsheet file from the 'raw' directory to 'processed'.
    
    Parameters:
        proj_dir (str): Project root directory.
        worksheet (str): Filename of the processed spreadsheet.
    """

    src = os.path.join(proj_dir, 'data', 'raw', worksheet)
    dst = os.path.join(proj_dir, 'data', 'processed', worksheet)
    shutil.move(src, dst)
