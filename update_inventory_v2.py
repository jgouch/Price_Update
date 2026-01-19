import pandas as pd
import re
import os
import sys
import warnings

# --- 1. SETUP ---
warnings.simplefilter(action='ignore', category=FutureWarning)

# --- HELPER FUNCTIONS ---

def clean_row_name(row_str):
    """Cleans row names like 'E - Heavenly' -> 'E'."""
    s = str(row_str).strip().upper()
    if ' - ' in s or ' – ' in s:
        return re.split(r'[-–]', s)[0].strip()
    if 'ELEVATION' in s:
        return s.replace('ELEVATION', '').strip()
    return s

def calculate_percent_sold(df_inventory, garden_name, garden_col, status_col):
    """Calculates % Sold (Total - Avail / Total)."""
    status_avail = ['Available', 'Serviceable']

    # Filter for garden (flexible match)
    if garden_name is None or str(garden_name).strip() == '':
        return None

    garden_mask = df_inventory[garden_col].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    avail_count = len(garden_data[garden_data[status_col].isin(status_avail)])
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name, garden_col, row_col, status_col):
    """Counts available inventory for a specific row."""
    status_avail = ['Available', 'Serviceable']

    # Filter Garden
    if garden_name is None or str(garden_name).strip() == '':
        return None

    garden_mask = df_inventory[garden_col].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return "N/A"

    # Filter Row
    target = clean_row_name(row_name)
    # Apply cleaning to inventory row column too
    row_mask = garden_data[row_col].astype(str).apply(clean_row_name) == target
    row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
        
    return len(row_data[row_data[status_col].isin(status_avail)])

def get_inventory_column(df_inventory, candidates):
    normalized = {str(col).lower(): col for col in df_inventory.columns}
    for candidate in candidates:
        match = normalized.get(str(candidate).lower())
        if match is not None:
            return match
    return None

# --- MAIN LOGIC ---

def main():
    # Check if files were dragged in
    if len(sys.argv) < 3:
        print("\n⚠️  READY FOR DRAG & DROP ⚠️")
        print("Usage: python3 [script.py] [Inventory_File] [Master_Price_Book]")
        print("-" * 50)
        print("1. Type 'python3 ' (don't forget the space)")
        print("2. Drag this script into the terminal")
        print("3. Drag your Inventory File (.xlsm) into the terminal")
        print("4. Drag your Master Price Book (.xlsx) into the terminal")
        print("5. Press Enter")
        print("-" * 50)
        return

    # 1. Grab File Paths from Terminal Arguments
    # sys.argv[0] is the script name itself
    inv_path = sys.argv[1].strip()
    master_path = sys.argv[2].strip()
    
    # Handle Mac terminal escaping (sometimes adds backslashes before spaces)
    # Usually python handles this, but strip quotes if present
    inv_path = inv_path.replace("'", "").replace('"', "")
    master_path = master_path.replace("'", "").replace('"', "")

    if not os.path.exists(inv_path):
        print(f"❌ Error: Cannot find inventory file at: {inv_path}")
        return
    if not os.path.exists(master_path):
        print(f"❌ Error: Cannot find master file at: {master_path}")
        return

    # Define Output Path (Same folder as Master Book)
    folder = os.path.dirname(master_path)
    output_path = os.path.join(folder, 'Harpeth_Hills_Master_Price_Book_UPDATED.xlsx')

    print(f"\nProcessing:\n 1. {os.path.basename(inv_path)}\n 2. {os.path.basename(master_path)}")
    print("...")

    # 2. Load Data
    try:
        df_inv = pd.read_excel(inv_path)
        print(f"✅ Inventory Loaded ({len(df_inv)} rows)")
    except Exception as e:
        print(f"❌ Error reading Inventory: {e}")
        return

    try:
        sheet_dict = pd.read_excel(master_path, sheet_name=None)
        print(f"✅ Master Book Loaded ({len(sheet_dict)} sheets)")
    except Exception as e:
        print(f"❌ Error reading Master Book: {e}")
        return

    garden_col_inv = get_inventory_column(df_inv, ['Garden', 'GARDEN'])
    row_col_inv = get_inventory_column(df_inv, ['Section', 'SECTION', 'Row', 'ROW', 'Lot', 'LOT'])
    status_col_inv = get_inventory_column(df_inv, ['Status', 'STATUS'])

    missing_cols = []
    if garden_col_inv is None:
        missing_cols.append('Garden')
    if row_col_inv is None:
        missing_cols.append('Section/Row/Lot')
    if status_col_inv is None:
        missing_cols.append('Status')
    if missing_cols:
        print(f"❌ Error: Inventory file is missing required columns: {', '.join(missing_cols)}")
        return

    # 3. Process Updates
    print("Updating availability stats...")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in sheet_dict.items():
            
            # Identify columns
            cols = [str(c).upper() for c in df.columns]
            
            # --- UPDATE % SOLD (Ground Burial) ---
            if any('%' in c for c in cols) and 'GARDEN' in cols:
                garden_col = next(c for c in df.columns if str(c).upper() == 'GARDEN')
                sold_col = next(c for c in df.columns if '%' in str(c))

                percent_sold_cache = {}
                for garden_name in df[garden_col].dropna().unique():
                    percent_sold_cache[garden_name] = calculate_percent_sold(
                        df_inv,
                        garden_name,
                        garden_col_inv,
                        status_col_inv,
                    )

                for idx, row in df.iterrows():
                    garden_name = row[garden_col]
                    new_pct = percent_sold_cache.get(garden_name)
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # --- UPDATE EXACT COUNTS (Mausoleums/Niches) ---
            row_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION'])]
            qty_col_candidates = [
                c for c in df.columns if any(x in str(c).upper() for x in ['AVAIL', 'QTY'])
            ]
            
            if row_col_candidates and qty_col_candidates:
                row_col = row_col_candidates[0]
                qty_col = qty_col_candidates[0]
                
                # Guess Inventory Garden Name from Sheet Name
                clean_sheet = sheet_name
                if '_' in clean_sheet: clean_sheet = clean_sheet.split('_', 1)[1]
                clean_sheet = clean_sheet.replace('Mausoleum', '').replace('Niches', '').replace('Columbarium', '').strip()
                
                row_count_cache = {}
                for row_val in df[row_col].dropna().unique():
                    if str(row_val).strip() == '':
                        continue
                    row_count_cache[row_val] = count_row_availability(
                        df_inv,
                        clean_sheet,
                        str(row_val),
                        garden_col_inv,
                        row_col_inv,
                        status_col_inv,
                    )

                for idx, row in df.iterrows():
                    row_val = row[row_col]
                    # Skip empty rows or headers
                    if pd.isna(row_val) or str(row_val).strip() == '':
                        continue

                    count = row_count_cache.get(row_val)

                    if count is not None and count != "N/A":
                        if count == 0:
                            df.at[idx, qty_col] = "Sold Out"
                        else:
                            df.at[idx, qty_col] = count

            # Save the sheet to the new file
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("-" * 50)
    print(f"SUCCESS! 🚀\nNew file created:\n{output_path}")
    print("-" * 50)

if __name__ == "__main__":
    main()
