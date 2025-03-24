# -*- coding: utf-8 -*-

import os
import sys
import pandas as pd
from collections import defaultdict
from sqlalchemy import create_engine



# Add src/ directory to sys.path for module import
script_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.abspath(os.path.join(script_dir, "..", "src"))
sys.path.append(src_path)

import page_reader as reader

# Set up project directory and file
proj_dir = r"C:\Users\knebiolo\OneDrive - Kleinschmidt Associates, Inc\Software\employee_tracker"
sheet_name = 'page2.xlsx'
sheet_path = os.path.join(proj_dir, 'data', 'raw', sheet_name)

# Load Excel page
page = pd.read_excel(sheet_path)

print (page.head(10))

#%% get corporate period
for idx, row in page.iterrows():
    for i in row:
        if isinstance(i, str) and 'For the period' in i:
            period_idx = i.index('-')
            period_end = i[period_idx + 2:]
            break

#%% get sections
section_rows = {}
section_data = {}

for idx, row in page.iterrows():
    if 'Section:' in str(row.iloc[0]):
        colon = row.iloc[0].index(':')
        section_name = row.iloc[0][colon + 2:]
        section_rows[section_name] = idx
        section_data[section_name] = [section_name]

#%% get company totals        
for idx, row in page.iterrows():
    if 'Final Totals' in str(row.iloc[0]):
        company_row = idx
        break

#%% get column names and their index
column_names = page.iloc[7].tolist()
col_name_idx = {}
for i, val in enumerate(column_names):
    if pd.notna(val):
        col_name_idx[val] = i

#%% get section data
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

#%% get firm data
firm_week_dict = defaultdict(list)
firm_mtd_dict = defaultdict(list)
firm_ytd_dict = defaultdict(list)

firm_week_row = page.iloc[company_row + 1]
firm_mtd_row = page.iloc[company_row + 2]
firm_ytd_row = page.iloc[company_row + 3]

for col in col_name_idx:
    cidx = col_name_idx[col]
    firm_week_dict[period_end].append(week_row.iloc[cidx])
    firm_mtd_dict[period_end].append(mtd_row.iloc[cidx])
    firm_ytd_dict[period_end].append(ytd_row.iloc[cidx])

firm_week_df = pd.DataFrame.from_dict(firm_week_dict, orient='index', columns=col_name_idx.keys())
firm_mtd_df = pd.DataFrame.from_dict(firm_mtd_dict, orient='index', columns=col_name_idx.keys())
firm_ytd_df = pd.DataFrame.from_dict(firm_ytd_dict, orient='index', columns=col_name_idx.keys())

for df in [section_week_df, section_mtd_df, section_ytd_df]:
    df['period_ending'] = pd.to_datetime(period_end)

#%% write section data
engine = create_engine(f"sqlite:///{os.path.join(proj_dir, 'data', 'employee_tracker.db')}")

section_week_df.to_sql("section_week", engine, if_exists="append", index=True)
section_mtd_df.to_sql("section_mtd", engine, if_exists="append", index=True)
section_ytd_df.to_sql("section_ytd", engine, if_exists="append", index=True)
firm_week_df.to_sql("firm_week", engine, if_exists="append", index=True)
firm_mtd_df.to_sql("firm_mtd", engine, if_exists="append", index=True)
firm_ytd_df.to_sql("firm_ytd", engine, if_exists="append", index=True)

