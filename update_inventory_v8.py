import pandas as pd
import re
import os
import sys
import warnings

# --- 1. SETUP ---
warnings.simplefilter(action='ignore', category=FutureWarning)

# --- HELPER FUNCTIONS ---
def clean_row_name(row_str):
    s = str(row_str).strip().upper()
    if ' - ' in s or ' – ' in s:
        return re.split(r'[-–]', s)[0].strip()
    if 'ELEVATION' in s:
        return s.replace('ELEVATION', '').strip()
    return s

def identify_columns(df):
    """Maps the actual column names from Row 3."""
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

def calculate_percent_sold(df_inventory, garden_name_full, col_map):
    col_garden = col_map['Garden']
    col_section = col_map['Row']
    col_status = col_map['Status']
    status_avail = ['Available', 'Serviceable', 'For Sale', 'Vacant']
    
    # Check for Split Name (e.g., "Grace - Sidewalk")
    if '–' in garden_name_full:
        parts = garden_name_full.split('–')
        main_garden = parts[0].strip()
        sub_section = parts[1].strip()
    elif '-' in garden_name_full:
        parts = garden_name_full.split('-')
        main_garden = parts[0].strip()
        sub_section = parts[1].strip()
    else:
        main_garden = garden_name_full
        sub_section = None

    # Filter Main Garden
    garden_mask = df_inventory[col_garden].astype(str).str.contains(main_garden, case=False, na=False)
    garden_data = df_inventory[garden_mask]
    
    # Filter Sub-Section if needed
    if sub_section and not garden_data.empty:
        section_mask = garden_data[col_section].astype(str).str.contains(sub_section, case=False, na=False)
        if section_mask.any():
            garden_data = garden_data[section_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    avail_mask = garden_data[col_status].astype(str).str.contains('|'.join(status_avail), case=False, na=False)
    avail_count = len(garden_data[avail_mask])
    
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name, col_map):
    col_garden = col_map['Garden']
    col_row = col_map['Row']
    col_status = col_map['Status']
    status_avail = ['Available', 'Serviceable', 'For Sale', 'Vacant']

    garden_mask = df_inventory[col_garden].astype(str).str.contains(garden_name, case=False, na=False)
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return "N/A"

    target = clean_row_name(row_name)
    row_mask = garden_data[col_row].astype(str).apply(clean_row_name) == target
    row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
    
    avail_mask = row_data[col_status].astype(str).str.contains('|'.join(status_avail), case=False, na=False)
    return len(row_data[avail_mask])

# --- AUTO-DETECT FILES ---
def detect_file_types(path1, path2):
    """
    Figures out which file is Inventory and which is Master Book based on contents.
    """
    f1_name = os.path.basename(path1).lower()
    f2_name = os.path.basename(path2).lower()
    
    inv_path = None
    master_path = None
    
    # 1. Try to guess by Filename first
    if 'inventory' in f1_name and 'inventory' not in f2_name:
        return path1, path2
    elif 'inventory' in f2_name and 'inventory' not in f1_name:
        return path2, path1
        
    # 2. If filenames are ambiguous, peek at sheet names
    print("🔍 Inspecting file contents to identify types...")
    try:
        xl1 = pd.ExcelFile(path1)
        xl2 = pd.ExcelFile(path2)
        
        sheets1 = [s.lower() for s in xl1.sheet_names]
        sheets2 = [s.lower() for s in xl2.sheet_names]
        
        # Criteria: Master Book usually has "01_ground burial" or similar
        is_master1 = any('ground burial' in s for s in sheets1)
        is_master2 = any('ground burial' in s for s in sheets2)
        
        if is_master1 and not is_master2:
            return path2, path1 # path2 is inv, path1 is master
        elif is_master2 and not is_master1:
            return path1, path2 # path1 is inv, path2 is master
            
    except Exception as e:
        print(f"Error inspecting files: {e}")
        
    return None, None

# --- MAIN LOGIC ---

def main():
    if len(sys.argv) < 3:
        print("Usage: python3 update_inventory_v8.py [File1] [File2]")
        return

    path_a = sys.argv[1].strip().replace("'", "").replace('"', "")
    path_b = sys.argv[2].strip().replace("'", "").replace('"', "")

    # SMART DETECT
    inv_path, master_path = detect_file_types(path_a, path_b)
    
    if not inv_path or not master_path:
        print("❌ Could not automatically identify which file is which.")
        print("Please make sure your Master Price Book has 'Ground Burial' in the sheet names.")
        return

    folder = os.path.dirname(master_path)
    output_path = os.path.join(folder, 'Harpeth_Hills_Master_Price_Book_UPDATED.xlsx')

    print(f"\n📂 Inventory File Detected: {os.path.basename(inv_path)}")
    print(f"📘 Master Price Book Detected: {os.path.basename(master_path)}")
    print("-" * 50)

    try:
        # HARDCODED: Read Inventory from Row 3
        df_inv = pd.read_excel(inv_path, header=2)
        
        col_map = identify_columns(df_inv)
        print(f"✅ Inventory Columns Mapped: {col_map}")
        
        if None in col_map.values():
            print("\n❌ ERROR: Still missing columns in Inventory.")
            return

    except Exception as e:
        print(f"❌ Error reading Inventory: {e}")
        return

    print("\nReading Master Price Book...")
    try:
        sheet_dict = pd.read_excel(master_path, sheet_name=None)
    except Exception as e:
        print(f"❌ Error reading Master Book: {e}")
        return

    print("Updating availability...")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in sheet_dict.items():
            
            cols = [str(c).upper() for c in df.columns]
            
            # 1. UPDATE % SOLD (With Smart Split Logic)
            if any('%' in c for c in cols) and 'GARDEN' in cols:
                garden_col = next(c for c in df.columns if str(c).upper() == 'GARDEN')
                sold_col = next(c for c in df.columns if '%' in str(c))
                
                for idx, row in df.iterrows():
                    garden_name = str(row[garden_col])
                    new_pct = calculate_percent_sold(df_inv, garden_name, col_map)
                    
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # 2. UPDATE EXACT COUNTS
            row_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION'])]
            qty_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['AVAIL', 'QTY', 'STATUS'])]
            
            if row_col_candidates and qty_col_candidates:
                row_col = row_col_candidates[0]
                qty_col = qty_col_candidates[0]
                
                clean_sheet = sheet_name
                if '_' in clean_sheet: clean_sheet = clean_sheet.split('_', 1)[1]
                clean_sheet = clean_sheet.replace('Mausoleum', '').replace('Niches', '').replace('Columbarium', '').strip()
                
                for idx, row in df.iterrows():
                    row_val = row[row_col]
                    if pd.isna(row_val) or str(row_val).strip() == '': continue
                    
                    count = count_row_availability(df_inv, clean_sheet, str(row_val), col_map)
                    
                    if count is not None and count != "N/A":
                        if count == 0:
                            df.at[idx, qty_col] = "Sold Out"
                        else:
                            df.at[idx, qty_col] = count

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("-" * 50)
    print(f"SUCCESS! 🚀\nNew file created:\n{output_path}")
    print("-" * 50)

if __name__ == "__main__":
    main()
