import pandas as pd
import re
import os
import sys
import warnings

# --- 1. SETUP ---
warnings.simplefilter(action='ignore', category=FutureWarning)

# --- 2. CLEANING FUNCTIONS ---

def clean_sheet_name_specific(sheet_name):
    """
    Minimal cleaning. Keeps 'Columbarium' and 'Niches' to distinguish entities.
    Ex: '07_Grace Columbarium' -> 'Grace Columbarium'
    """
    # Remove numbering prefixes like "01_", "14_"
    name = re.sub(r'^\d+_', '', sheet_name)

    # Remove ONLY generic building words, but KEEP the distinction (case-insensitive)
    remove_words = ['Mausoleum', 'Bldg', 'Building']
    for word in remove_words:
        name = re.sub(re.escape(word), '', name, flags=re.IGNORECASE)

    return name.strip()

def clean_sheet_name_generic(sheet_name):
    """
    Aggressive cleaning. Removes everything to find the base garden.
    Ex: '09_Bell Tower Niches' -> 'Bell Tower'
    """
    name = clean_sheet_name_specific(sheet_name)
    remove_words = ['Columbarium', 'Niches', 'Garden']
    for word in remove_words:
        name = re.sub(re.escape(word), '', name, flags=re.IGNORECASE)
    return name.strip()

def clean_row_name(row_str):
    """
    Standardizes row names for matching.
    """
    s = str(row_str).strip().upper()
    
    if 'ALL LEVEL' in s: return 'ALL'

    # Remove Parentheses and Prefixes
    s = re.sub(r'\(.*?\)', '', s)
    s = s.replace('UNCOVERED', '').replace('COVERED', '').replace('ELEVATION', '')
    
    # Normalize dashes
    s = s.replace('—', '-').replace('–', '-')
    if ' - ' in s:
        s = s.split(' - ')[0]
        
    return s.strip()

# --- 3. MAPPING & CHECKING ---

def identify_columns(df):
    cols = [str(c) for c in df.columns]
    mapping = {'Garden': None, 'Row': None, 'Status': None}

    # GARDEN
    for c in cols:
        if 'GARDEN' in c.upper() or 'GROUP' in c.upper() or 'LOCATION' in c.upper():
            mapping['Garden'] = c
            break

    # ROW / SECTION
    candidates = [c for c in cols if any(x in c.upper() for x in ['SECTION', 'ROW', 'BLOCK', 'LOT', 'TIER'])]
    if candidates:
        best = next((x for x in candidates if 'SECTION' in x.upper()), None)
        if not best: best = next((x for x in candidates if 'ROW' in x.upper()), None)
        mapping['Row'] = best if best else candidates[0]

    # STATUS
    for c in cols:
        if 'STATUS' in c.upper() or 'STATE' in c.upper():
            mapping['Status'] = c
            break
            
    return mapping

def garden_exists_in_inventory(df_inventory, garden_name, col_map):
    """Checks if a garden name exists in the inventory."""
    col_garden = col_map['Garden']
    # Use simple string match
    mask = df_inventory[col_garden].astype(str).str.contains(garden_name, case=False, na=False, regex=False)
    return mask.any()

def is_available_status(status_value):
    """Determines if a status value is considered available."""
    if pd.isna(status_value):
        return False
    normalized = str(status_value).strip().upper()
    return normalized in {'AVAILABLE', 'SERVICEABLE', 'FOR SALE', 'VACANT'}

# --- 4. CALCULATIONS ---

def calculate_percent_sold(df_inventory, garden_name_full, col_map):
    col_garden = col_map['Garden']
    col_section = col_map['Row']
    col_status = col_map['Status']
    
    # Split Logic
    if '–' in garden_name_full: parts = garden_name_full.split('–')
    elif ' - ' in garden_name_full: parts = garden_name_full.split(' - ')
    elif '-' in garden_name_full: parts = garden_name_full.split('-')
    else: parts = [garden_name_full]
    
    main_garden = parts[0].strip()
    sub_section = parts[1].strip() if len(parts) > 1 else None

    # Filter Main Garden
    garden_mask = df_inventory[col_garden].astype(str).str.contains(main_garden, case=False, na=False, regex=False)
    garden_data = df_inventory[garden_mask]
    
    # Filter Sub-Section
    if sub_section and not garden_data.empty:
        section_mask = garden_data[col_section].astype(str).str.contains(sub_section, case=False, na=False, regex=False)
        if section_mask.any():
            garden_data = garden_data[section_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    avail_mask = garden_data[col_status].apply(is_available_status)
    avail_count = len(garden_data[avail_mask])
    
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name, col_map):
    col_garden = col_map['Garden']
    col_row = col_map['Row']
    col_status = col_map['Status']

    garden_mask = df_inventory[col_garden].astype(str).str.contains(garden_name, case=False, na=False, regex=False)
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return None

    target = clean_row_name(row_name)
    
    if target == 'ALL':
        row_data = garden_data
    else:
        row_mask = garden_data[col_row].astype(str).apply(clean_row_name) == target
        row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
    
    avail_mask = row_data[col_status].apply(is_available_status)
    return len(row_data[avail_mask])

# --- 5. FILE DETECT ---
def detect_file_types(path1, path2):
    f1 = os.path.basename(path1).lower()
    f2 = os.path.basename(path2).lower()
    if 'inventory' in f1 and 'inventory' not in f2: return path1, path2
    if 'inventory' in f2 and 'inventory' not in f1: return path2, path1
    return None, None

def infer_inventory_path(path1, path2):
    """Attempt to infer the inventory file by checking expected columns."""
    candidates = []
    for path in (path1, path2):
        try:
            df = pd.read_excel(path, header=2)
        except Exception:
            continue
        col_map = identify_columns(df)
        if None not in col_map.values():
            candidates.append(path)
    if len(candidates) == 1:
        inv_path = candidates[0]
        master_path = path2 if inv_path == path1 else path1
        return inv_path, master_path
    return None, None

# --- MAIN ---
def main():
    if len(sys.argv) < 3:
        print("Usage: python3 update_inventory_v10.py [File1] [File2]")
        return

    path_a = sys.argv[1].strip().replace("'", "").replace('"', "")
    path_b = sys.argv[2].strip().replace("'", "").replace('"', "")

    inv_path, master_path = detect_file_types(path_a, path_b)
    if not inv_path:
        inv_path, master_path = infer_inventory_path(path_a, path_b)
    if not inv_path:
        print("❌ Error: Could not verify files. One filename must include 'Inventory' or contain expected columns.")
        return

    print(f"\n📂 Inventory: {os.path.basename(inv_path)}")
    print(f"📘 Master Book: {os.path.basename(master_path)}")

    # READ INVENTORY (Row 3 Header)
    try:
        df_inv = pd.read_excel(inv_path, header=2)
        col_map = identify_columns(df_inv)
        print(f"✅ Mapped Columns: {col_map}")
        
        if None in col_map.values():
            print("❌ Error: Missing columns in inventory.")
            return
    except Exception as e:
        print(f"❌ Error reading inventory: {e}")
        return

    # READ MASTER
    try:
        sheet_dict = pd.read_excel(master_path, sheet_name=None)
    except Exception as e:
        print(f"❌ Error reading master book: {e}")
        return

    print("\nUpdating availability...")
    
    folder = os.path.dirname(master_path)
    output_path = os.path.join(folder, 'Harpeth_Hills_Master_Price_Book_UPDATED_FINAL.xlsx')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in sheet_dict.items():
            
            # --- INTELLIGENT NAME MATCHING ---
            # 1. Try Specific (e.g. "Grace Columbarium")
            search_name_specific = clean_sheet_name_specific(sheet_name)
            
            # 2. Try Generic (e.g. "Bell Tower")
            search_name_generic = clean_sheet_name_generic(sheet_name)
            
            # Decide which to use
            if garden_exists_in_inventory(df_inv, search_name_specific, col_map):
                final_search_name = search_name_specific
            else:
                final_search_name = search_name_generic

            # print(f"Sheet '{sheet_name}' -> Searching Inventory for: '{final_search_name}'")

            cols = [str(c).upper() for c in df.columns]
            
            # A. GROUND BURIAL
            percent_candidates = [
                c for c in df.columns
                if any(x in str(c).upper() for x in ['%', 'PCT', 'PERCENT'])
            ]
            if percent_candidates and 'GARDEN' in cols:
                garden_col = next(c for c in df.columns if str(c).upper() == 'GARDEN')
                sold_col = percent_candidates[0]
                
                for idx, row in df.iterrows():
                    garden_name = str(row[garden_col])
                    new_pct = calculate_percent_sold(df_inv, garden_name, col_map)
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # B. MAUSOLEUMS / NICHES
            row_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION', 'PRODUCT'])]
            qty_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['AVAIL', 'QTY'])]
            
            if row_candidates and qty_candidates:
                row_col = row_candidates[0]
                qty_col = qty_candidates[0]
                
                for idx, row in df.iterrows():
                    row_val = row[row_col]
                    if pd.isna(row_val) or str(row_val).strip() == '': continue
                    
                    count = count_row_availability(df_inv, final_search_name, str(row_val), col_map)
                    
                    if count is not None:
                        if count == 0:
                            df.at[idx, qty_col] = "Sold Out"
                        else:
                            df.at[idx, qty_col] = count

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("-" * 50)
    print(f"SUCCESS! 🚀\nFile saved: {output_path}")
    print("-" * 50)

if __name__ == "__main__":
    main()
