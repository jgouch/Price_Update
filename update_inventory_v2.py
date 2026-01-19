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

def calculate_percent_sold(df_inventory, garden_name):
    """Calculates % Sold (Total - Avail / Total)."""
    col_garden = 'Garden'
    col_status = 'Status'
    status_avail = ['Available', 'Serviceable']
    
    # Filter for garden (flexible match)
    garden_mask = df_inventory[col_garden].astype(str).str.contains(garden_name, case=False, na=False)
    garden_data = df_inventory[garden_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    avail_count = len(garden_data[garden_data[col_status].isin(status_avail)])
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name):
    """Counts available inventory for a specific row."""
    col_garden = 'Garden'
    col_row = 'Section' # Adjust if your report uses 'Row' or 'Lot'
    col_status = 'Status'
    status_avail = ['Available', 'Serviceable']

    # Filter Garden
    garden_mask = df_inventory[col_garden].astype(str).str.contains(garden_name, case=False, na=False)
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return "N/A"

    # Filter Row
    target = clean_row_name(row_name)
    # Apply cleaning to inventory row column too
    row_mask = garden_data[col_row].astype(str).apply(clean_row_name) == target
    row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
        
    return len(row_data[row_data[col_status].isin(status_avail)])

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
                
                for idx, row in df.iterrows():
                    garden_name = str(row[garden_col])
                    new_pct = calculate_percent_sold(df_inv, garden_name)
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # --- UPDATE EXACT COUNTS (Mausoleums/Niches) ---
            row_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION'])]
            qty_col_candidates = [c for c in df.columns if any(x in str(c).upper() for x in ['AVAIL', 'QTY', 'STATUS'])]
            
            if row_col_candidates and qty_col_candidates:
                row_col = row_col_candidates[0]
                qty_col = qty_col_candidates[0]
                
                # Guess Inventory Garden Name from Sheet Name
                clean_sheet = sheet_name
                if '_' in clean_sheet: clean_sheet = clean_sheet.split('_', 1)[1]
                clean_sheet = clean_sheet.replace('Mausoleum', '').replace('Niches', '').replace('Columbarium', '').strip()
                
                for idx, row in df.iterrows():
                    row_val = row[row_col]
                    # Skip empty rows or headers
                    if pd.isna(row_val) or str(row_val).strip() == '': continue
                    
                    count = count_row_availability(df_inv, clean_sheet, str(row_val))
                    
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
