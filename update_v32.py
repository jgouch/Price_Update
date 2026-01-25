import pandas as pd
import re
import os
import sys
import warnings
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

# --- 1. SETUP ---
warnings.simplefilter(action='ignore', category=FutureWarning)

# --- 2. HEADER SCANNER ---
def find_inventory_header(file_path):
    print("🔎 Scanning Inventory file for headers...")
    try:
        df_temp = pd.read_excel(file_path, header=None, nrows=20)
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return 0

    best_row = 0
    max_matches = 0
    keywords = ['SECTION', 'SPACE', 'STATUS', 'COST', 'RIGHTS', 'GARDEN']
    
    for i, row in df_temp.iterrows():
        row_str = " ".join([str(x).upper() for x in row.values if pd.notna(x)])
        matches = sum(1 for k in keywords if k in row_str)
        if matches > max_matches:
            max_matches = matches
            best_row = i
            
    if max_matches >= 2:
        print(f"✅ Found Inventory Headers at Row {best_row + 1}")
        return best_row
    else:
        print("⚠️  Could not confidently find headers. Defaulting to Row 3.")
        return 2

# --- 3. CLEANING FUNCTIONS ---
def clean_sheet_name_specific(sheet_name):
    name = re.sub(r'^\d+_', '', sheet_name)
    for word in ['Mausoleum', 'Bldg', 'Building', 'Garden']:
        name = name.replace(word, '')
    return name.strip()

def clean_sheet_name_generic(sheet_name):
    name = clean_sheet_name_specific(sheet_name)
    for word in ['Columbarium', 'Niches']:
        name = name.replace(word, '')
    return name.strip()

def clean_row_name(row_str):
    s = str(row_str).strip().upper()
    if 'ALL LEVEL' in s: return 'ALL'
    s = re.sub(r'\(.*?\)', '', s)
    s = s.replace('UNCOVERED', '').replace('COVERED', '').replace('ELEVATION', '').replace('LEVEL', '')
    s = s.replace('—', '-').replace('–', '-')
    if '-' in s: s = s.split('-')[0]
    return s.strip()

def super_clean_name(name):
    if not isinstance(name, str): return ""
    name = re.sub(r'\(.*?\)', '', name)
    name = re.sub(r'\d+', '', name)
    for word in ['GARDEN', 'SECTION', 'LOC', 'LOCATION', 'BLOCK', 'OF', 'THE']:
        name = name.upper().replace(word, '')
    name = re.sub(r'[^\w\s]', '', name)
    cleaned = name.strip().upper()
    if len(cleaned) < 2: return "" 
    return cleaned

# --- 4. ROBUST MAPPING ---
def identify_columns(df):
    cols_map = {str(c).strip().upper(): c for c in df.columns}
    mapping = {'Garden': None, 'Row': None, 'Status': None}
    
    if 'SECTION' in cols_map:
        mapping['Garden'] = cols_map['SECTION']
    else:
        for c_up, c_orig in cols_map.items():
            if 'GARDEN' in c_up or 'LOCATION' in c_up:
                mapping['Garden'] = c_orig; break

    if 'SPACE' in cols_map:
        mapping['Row'] = cols_map['SPACE']
    elif 'LOT' in cols_map:
        mapping['Row'] = cols_map['LOT']
    else:
        for c_up, c_orig in cols_map.items():
            if 'ROW' in c_up or 'TIER' in c_up:
                mapping['Row'] = c_orig; break

    for c_up, c_orig in cols_map.items():
        if 'STATUS' in c_up or 'STATE' in c_up:
            mapping['Status'] = c_orig; break
            
    return mapping

def garden_exists_in_inventory(df_inventory, garden_name, col_map):
    col_garden = col_map['Garden']
    if not col_garden: return False
    target = super_clean_name(garden_name)
    if not target: return False
    inv_clean = df_inventory[col_garden].astype(str).apply(super_clean_name)
    return inv_clean.str.contains(target, case=False, na=False).any()

# --- 5. CALCULATIONS ---
def is_grace_sidewalk(space_str):
    if not isinstance(space_str, str): return False
    match = re.search(r'Lot/Section\s+(\d+)', space_str, re.IGNORECASE)
    if not match: return False
    
    try:
        section_num = int(match.group(1))
        sidewalk_sections = [30, 40, 50, 80, 90, 100, 110, 120]
        sidewalk_sections.extend(range(60, 65))
        sidewalk_sections.extend(range(70, 75))
        return section_num in sidewalk_sections
    except:
        return False

def calculate_percent_sold(df_inventory, garden_name_full, col_map):
    col_garden, col_section, col_status = col_map['Garden'], col_map['Row'], col_map['Status']
    if not col_garden or not col_status: return None

    status_avail = ['Available', 'Serviceable', 'For Sale', 'Vacant']
    parts = re.split(r'[-–]', garden_name_full)
    target_garden = super_clean_name(parts
