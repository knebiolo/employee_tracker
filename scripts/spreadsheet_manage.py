
# -*- coding: utf-8 -*-
"""
Script: page_read.py
Executes the employee metric data extraction and loading pipeline.
Author: Kevin Nebiolo
"""

import os
import sys
import pandas as pd

# Add src/ directory to sys.path for module import
script_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.abspath(os.path.join(script_dir, "..", "src"))
base_path = os.path.abspath(os.path.join(script_dir,".."))
sys.path.append(src_path)
sys.path.append(base_path)


import page_reader as reader

def main():
    """
    Main function to run the data extraction and loading process.
    """
    # Set up project directory and file
    # Explicitly add the parent directory of Stryke to sys.path
    #proj_dir = r"C:\Users\knebiolo\OneDrive - Kleinschmidt Associates, Inc\Software\employee_tracker"
    
    section_page = 'Mekong - Fisheries Aquatic'
    firm_page = '$ Utilization'
    
    sheets = os.listdir(os.path.join(src_path,'data','raw'))
    
    for sheet in sheets:
        sheet_path = os.path.join(src_path, 'data', 'raw', sheet)
    
        # Load Excel page
        employee_page = pd.read_excel(sheet_path, sheet_name = section_page)
        util_page = pd.read_excel(sheet_path, sheet_name = firm_page)
    
        # Get column indices from header row
        emp_col_name_idx = reader.get_columns(employee_page)
        util_col_name_idx = reader.get_firm_columns(util_page)
    
        # Extract pay period end date
        period_end = reader.get_pay_period(employee_page)
        firm_period_end = reader.get_firm_pay_period(util_page)
        firm_row = reader.get_firm_totals(util_page)
    
        # Extract metadata
        employee_rows, employee_df = reader.get_employees(employee_page)
        section_rows, section_df = reader.get_sections(util_page)
    
        # Extract utilization data
        emp_week_df, emp_mtd_df, emp_ytd_df = reader.get_employee_data(employee_page, period_end, employee_rows, emp_col_name_idx)
        sec_week_df, sec_mtd_df, sec_ytd_df = reader.get_section_data(util_page, period_end, section_rows, util_col_name_idx)
        firm_week_df, firm_mtd_df, firm_ytd_df = reader.get_firm_data(util_page, period_end, firm_row, util_col_name_idx)
    
        # Write all data to database
        reader.write_data(src_path, 
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
                        firm_ytd_df)
    
        # Move processed file
        reader.clean_up(src_path, sheet)

if __name__ == "__main__":
    main()
