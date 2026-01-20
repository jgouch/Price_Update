import pandas as pd
import re
import os
import sys
import warnings

# --- 1. SETUP ---
warnings.simplefilter(action='ignore', category=FutureWarning)

# --- SMART COLUMN DETECTIVE ---
def identify_columns(df):
    """
    Attempts to find the critical columns by looking for keywords.
    Returns a dictionary of found column names.
    """
    cols = [str(c) for c in df.columns]
    mapping = {'Garden': None, 'Row': None, 'Status': None}

    # 1. FIND GARDEN COLUMN
    # Look for exact match first, then fuzzy
    if 'Garden' in cols: mapping['Garden'] = 'Garden'
    else:
        # Look for variations
        for c in cols:
            if 'GARDEN' in c.upper() or 'LOCATION' in c.upper() or 'PROPERTY GROUP' in c.upper():
                mapping['Garden'] = c
                break

    # 2. FIND ROW/SECTION COLUMN
    if 'Section' in cols: mapping['Row'] = 'Section'
    elif 'Row' in cols: mapping['Row'] = 'Row'
    else:
        for c in cols:
            if 'SECTION' in c.upper() or 'ROW' in c.upper() or 'BLOCK' in c.upper() or 'LOT' in c.upper() or 'TIER' in c.upper():
                mapping['Row'] = c
                break

    # 3. FIND STATUS COLUMN
    if 'Status' in cols: mapping['Status'] = 'Status'
    else:
        for c in cols:
            if 'STATUS' in c.upper() or 'STATE' in c.upper():
                mapping['Status'] = c
                break
                
    return mapping

# --- HELPER FUNCTIONS ---

def clean_row_name(row_str):
    """Cleans row names like 'E - Heavenly' -> 'E'."""
    s = str(row_str).strip().upper()
    if ' - ' in s:
        return s.split(' - ', 1)[0].strip()
    if ' – ' in s:
        return s.split(' – ', 1)[0].strip()
    if 'ELEVATION' in s:
        return s.replace('ELEVATION', '').strip()
    return s

def calculate_percent_sold(df_inventory, garden_name, col_map):
    """Calculates % Sold using the mapped columns."""
    col_garden = col_map['Garden']
    col_status = col_map['Status']
    status_avail = ['Available', 'Serviceable', 'For Sale'] # Add variations as needed
    
    # Filter for garden
    garden_query = str(garden_name).strip()
    if not garden_query or garden_query.lower() == 'nan':
        return None
    garden_mask = df_inventory[col_garden].astype(str).str.contains(
        re.escape(garden_query),
        case=False,
        na=False
    )
    garden_data = df_inventory[garden_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    # Check for "Available" status (fuzzy match)
    avail_mask = garden_data[col_status].astype(str).str.contains('|'.join(status_avail), case=False, na=False)
    avail_count = len(garden_data[avail_mask])
    
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name, col_map):
    """Counts available inventory using mapped columns."""
    col_garden = col_map['Garden']
    col_row = col_map['Row']
    col_status = col_map['Status']
    status_avail = ['Available', 'Serviceable', 'For Sale']

    # Filter Garden
    garden_query = str(garden_name).strip()
    if not garden_query or garden_query.lower() == 'nan':
        return "N/A"
    garden_mask = df_inventory[col_garden].astype(str).str.contains(
        re.escape(garden_query),
        case=False,
        na=False
    )
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return "N/A"

    # Filter Row
    target = clean_row_name(row_name)
    row_mask = garden_data[col_row].astype(str).apply(clean_row_name) == target
    row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
    
    # Count Available
    avail_mask = row_data[col_status].astype(str).str.contains('|'.join(status_avail), case=False, na=False)
    return len(row_data[avail_mask])

# --- MAIN LOGIC ---

def main():
    if len(sys.argv) < 3:
        print("Usage: python3 update_inventory_v3.py [Inventory_File] [Master_Price_Book]")
        return

    inv_path = sys.argv[1].strip().replace("'", "").replace('"', "")
    master_path = sys.argv[2].strip().replace("'", "").replace('"', "")

    # Define Output Path
    folder = os.path.dirname(master_path)
    output_path = os.path.join(folder, 'Harpeth_Hills_Master_Price_Book_UPDATED.xlsx')

    print(f"\nReading Inventory: {os.path.basename(inv_path)}...")

    try:
        # Read the Excel file
        df_inv = pd.read_excel(inv_path)
        
        # --- DIAGNOSTIC PRINT ---
        print("\n🔎 DIAGNOSTIC REPORT: INVENTORY COLUMNS FOUND:")
        print(list(df_inv.columns))
        print("-" * 40)
        
        # Identify Columns
        col_map = identify_columns(df_inv)
        print(f"🤖 Auto-Mapped Columns: {col_map}")
        
        if None in col_map.values():
            print("\n❌ CRITICAL ERROR: Could not identify one or more required columns.")
            print("Please copy the 'DIAGNOSTIC REPORT' list above and paste it to the chat.")
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
            
            # 1. UPDATE % SOLD
            if any('%' in c for c in cols) and 'GARDEN' in cols:
                garden_col = next(c for c in df.columns if str(c).upper() == 'GARDEN')
                sold_col_candidates = [
                    c for c in df.columns
                    if '%' in str(c)
                    and 'SOLD' in str(c).upper()
                ]
                sold_col = sold_col_candidates[0] if sold_col_candidates else next(
                    c for c in df.columns if '%' in str(c)
                )
                
                for idx, row in df.iterrows():
                    garden_name = str(row[garden_col])
                    new_pct = calculate_percent_sold(df_inv, garden_name, col_map)
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # 2. UPDATE EXACT COUNTS
            row_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION'])]
            qty_col_candidates = [
                c for c in df.columns
                if any(x in str(c).upper() for x in ['AVAIL', 'QTY', 'COUNT'])
                and 'STATUS' not in str(c).upper()
            ]
            
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
