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
    if ' - ' in s:
        return s.split(' - ', 1)[0].strip()
    if ' – ' in s:
        return s.split(' – ', 1)[0].strip()
    if 'ELEVATION' in s:
        return s.replace('ELEVATION', '').strip()
    return s

def identify_columns(df):
    """Maps the actual column names from Row 3 to our needed logical names."""
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

def normalize_status_series(series):
    return series.astype(str).str.strip().str.upper()

def calculate_percent_sold(df_inventory, garden_name, col_map):
    col_garden = col_map['Garden']
    col_status = col_map['Status']
    status_avail = {'AVAILABLE', 'SERVICEABLE', 'FOR SALE', 'VACANT'}
    
    # Filter Garden
    if pd.isna(garden_name) or str(garden_name).strip() == '':
        return None
    garden_mask = df_inventory[col_garden].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    total = len(garden_data)
    if total == 0: return None
        
    avail_mask = normalize_status_series(garden_data[col_status]).isin(status_avail)
    avail_count = avail_mask.sum()
    
    return (total - avail_count) / total

def count_row_availability(df_inventory, garden_name, row_name, col_map):
    col_garden = col_map['Garden']
    col_row = col_map['Row']
    col_status = col_map['Status']
    status_avail = {'AVAILABLE', 'SERVICEABLE', 'FOR SALE', 'VACANT'}

    if pd.isna(garden_name) or str(garden_name).strip() == '':
        return "N/A"
    garden_mask = df_inventory[col_garden].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty: return "N/A"

    target = clean_row_name(row_name)
    row_mask = garden_data[col_row].astype(str).apply(clean_row_name) == target
    row_data = garden_data[row_mask]
    
    if len(row_data) == 0: return None
    
    avail_mask = normalize_status_series(row_data[col_status]).isin(status_avail)
    return avail_mask.sum()

def select_preferred_column(df, keywords, preferred_keywords=()):
    candidates = [c for c in df.columns if any(x in str(c).upper() for x in keywords)]
    for keyword in preferred_keywords:
        for c in candidates:
            if keyword in str(c).upper():
                return c
    return candidates[0] if candidates else None

# --- MAIN LOGIC ---

def main():
    if len(sys.argv) < 3:
        print("Usage: python3 update_inventory_v6.py [Inventory_File] [Master_Price_Book]")
        return

    inv_path = sys.argv[1].strip().replace("'", "").replace('"', "")
    master_path = sys.argv[2].strip().replace("'", "").replace('"', "")

    folder = os.path.dirname(master_path)
    output_path = os.path.join(folder, 'Harpeth_Hills_Master_Price_Book_UPDATED.xlsx')

    print(f"\nProcessing Inventory: {os.path.basename(inv_path)}...")

    try:
        # HARDCODED: Read from Row 3 (header=2 means skip rows 0 and 1)
        df_inv = pd.read_excel(inv_path, header=2)
        
        # Verify Columns
        col_map = identify_columns(df_inv)
        print(f"✅ Headers found on Row 3. Columns Mapped: {col_map}")
        
        if None in col_map.values():
            print("\n❌ ERROR: Still missing columns. Here are the headers found on Row 3:")
            print(list(df_inv.columns))
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
                sold_col = select_preferred_column(
                    df,
                    ['%'],
                    preferred_keywords=('SOLD', 'PERCENT SOLD'),
                )
                if not sold_col:
                    print(f"⚠️  Skipping % sold update for '{sheet_name}' (no % column found).")
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    continue
                
                for idx, row in df.iterrows():
                    garden_name = str(row[garden_col])
                    new_pct = calculate_percent_sold(df_inv, garden_name, col_map)
                    if new_pct is not None:
                        df.at[idx, sold_col] = new_pct

            # 2. UPDATE EXACT COUNTS
            row_col = select_preferred_column(
                df,
                ['ROW', 'LEVEL', 'SECTION', 'STATION'],
                preferred_keywords=('ROW', 'SECTION'),
            )
            qty_col = select_preferred_column(
                df,
                ['AVAIL', 'QTY', 'AVAILABLE', 'QUANTITY'],
                preferred_keywords=('AVAIL', 'AVAILABLE', 'QTY', 'QUANTITY'),
            )
            
            if row_col and qty_col:
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
            elif any(x in cols for x in ['ROW', 'LEVEL', 'SECTION', 'STATION']):
                print(f"⚠️  Skipping availability counts for '{sheet_name}' (missing row or quantity column).")

            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print("-" * 50)
    print(f"SUCCESS! 🚀\nNew file created:\n{output_path}")
    print("-" * 50)

if __name__ == "__main__":
    main()
