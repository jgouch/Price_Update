import pandas as pd
import re
import os
import warnings

# --- CONFIGURATION ---
# Update these filenames to match exactly what you have in your folder
INVENTORY_FILE = 'Property - Property Inventory Listing - Single Location.xlsm'
MASTER_FILE = 'Harpeth_Hills_Master_Price_Book_2025.xlsx'
OUTPUT_FILE = 'Harpeth_Hills_Master_Price_Book_UPDATED.xlsx'

# Column Mapping (Adjust these if your inventory file uses different names)
COL_GARDEN = 'Garden'      # The high-level garden name
COL_ROW = 'Section'        # Often 'Section', 'Row', or 'Lot' in reports
COL_STATUS = 'Status'      # 'Available', 'Sold', 'Occupied', etc.
STATUS_AVAILABLE = ['Available', 'Serviceable'] # What counts as "For Sale"?
STATUS_AVAILABLE_NORMALIZED = {status.strip().upper() for status in STATUS_AVAILABLE}

# --- 1. SETUP ---
os.chdir(os.path.dirname(os.path.abspath(__file__)))
warnings.simplefilter(action='ignore', category=FutureWarning)

def clean_row_name(row_str):
    """
    Turns "E - Heavenly" into just "E" for matching.
    Turns "Elevation 101" into "101".
    """
    s = str(row_str).strip().upper()
    # If it looks like "E - Something", grab "E"
    if ' - ' in s or ' – ' in s:
        # Split by dash and take the first part
        return re.split(r'[-–]', s)[0].strip()
    # If it looks like "Elevation 101", grab "101"
    if 'ELEVATION' in s:
        return s.replace('ELEVATION', '').strip()
    return s

def _normalize_status_series(series):
    return series.astype(str).str.strip().str.upper()


def calculate_percent_sold(df_inventory, garden_name):
    """Calculates % Sold for a specific garden."""
    if not garden_name or pd.isna(garden_name) or str(garden_name).strip() == '':
        return None
    # Filter inventory for this garden
    garden_mask = df_inventory[COL_GARDEN].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    total_spaces = len(garden_data)
    if total_spaces == 0:
        return None
        
    # Count available
    avail_mask = _normalize_status_series(garden_data[COL_STATUS]).isin(STATUS_AVAILABLE_NORMALIZED)
    avail_spaces = len(garden_data[avail_mask])
    
    percent_sold = (total_spaces - avail_spaces) / total_spaces
    return percent_sold

def count_row_availability(df_inventory, garden_name, row_name):
    """Counts available spaces in a specific row/section of a garden."""
    if not garden_name or pd.isna(garden_name) or str(garden_name).strip() == '':
        return "N/A"
    # 1. Filter by Garden
    garden_mask = df_inventory[COL_GARDEN].astype(str).str.contains(
        str(garden_name),
        case=False,
        na=False,
        regex=False,
    )
    garden_data = df_inventory[garden_mask]
    
    if garden_data.empty:
        return "N/A"

    # 2. Filter by Row (Cleaned)
    target_row = clean_row_name(row_name)
    
    # We apply cleaning to the inventory column too for comparison
    # (Assuming COL_ROW contains the row info like "E", "101", etc.)
    row_mask = garden_data[COL_ROW].astype(str).apply(clean_row_name) == target_row
    row_data = garden_data[row_mask]
    
    # 3. Count Status
    avail_count = len(
        row_data[_normalize_status_series(row_data[COL_STATUS]).isin(STATUS_AVAILABLE_NORMALIZED)]
    )
    
    # If the row doesn't exist in inventory, return empty so we don't overwrite manual data
    if len(row_data) == 0:
        return None
        
    return avail_count

# --- MAIN EXECUTION ---

print("Loading files...")

# Load Inventory
try:
    df_inv = pd.read_excel(INVENTORY_FILE)
    print(f"Inventory loaded: {len(df_inv)} rows found.")
except FileNotFoundError:
    print(f"ERROR: Could not find {INVENTORY_FILE}")
    exit()

# Load Master Price Book
try:
    sheet_dict = pd.read_excel(MASTER_FILE, sheet_name=None)
    print(f"Master Price Book loaded: {len(sheet_dict)} sheets found.")
except FileNotFoundError:
    print(f"ERROR: Could not find {MASTER_FILE}")
    exit()

print("Updating availability...")

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    for sheet_name, df in sheet_dict.items():
        print(f"Processing {sheet_name}...")
        
        # Identify Columns to Update
        cols = [str(c).upper() for c in df.columns]
        
        # --- SCENARIO A: GROUND BURIAL (% SOLD) ---
        if any('%' in c for c in cols) and 'GARDEN' in cols:
            # Find the specific column names
            garden_col = next(c for c in df.columns if str(c).upper() == 'GARDEN')
            sold_col_candidates = [
                c for c in df.columns if '%' in str(c) and 'SOLD' in str(c).upper()
            ]
            if sold_col_candidates:
                sold_col = sold_col_candidates[0]
            else:
                sold_col = next(c for c in df.columns if '%' in str(c))
            
            for idx, row in df.iterrows():
                garden_name = str(row[garden_col])
                
                # Special handling for "Lawn Crypts" or subsections
                # If the garden name in Price Book is "Chapel Hill", we look for "Chapel Hill" in Inventory
                new_pct = calculate_percent_sold(df_inv, garden_name)
                
                if new_pct is not None:
                    # Update column (Format as 0.95, Excel handles the %)
                    df.at[idx, sold_col] = new_pct

        # --- SCENARIO B: MAUSOLEUMS / NICHES (EXACT COUNTS) ---
        # Look for a "Row", "Level", or "Section" column AND an "Available" or "Qty" column
        row_col_candidates = [
            c for c in df.columns
            if any(x in str(c).upper() for x in ['ROW', 'LEVEL', 'SECTION', 'STATION'])
        ]
        qty_col_candidates = [
            c for c in df.columns
            if any(x in str(c).upper() for x in ['AVAIL', 'QTY', 'QUANTITY', 'AVAILABLE'])
        ]
        
        if row_col_candidates and qty_col_candidates:
            row_col = row_col_candidates[0]
            qty_col = qty_col_candidates[0]
            
            # Clean Sheet Name to guess Garden Name (e.g. "03_Bell Tower" -> "Bell Tower")
            # This is a guess; might need manual tweaking if names don't match
            clean_sheet = sheet_name
            if '_' in clean_sheet: clean_sheet = clean_sheet.split('_', 1)[1]
            clean_sheet = clean_sheet.replace('Mausoleum', '').replace('Niches', '').replace('Columbarium', '').strip()
            
            for idx, row in df.iterrows():
                row_val = row[row_col]
                
                # Skip header rows embedded in data
                if pd.isna(row_val) or str(row_val).strip() == '': continue
                
                # Calculate Count
                # Note: We search the Inventory for the Clean Sheet Name (e.g. "Bell Tower")
                # Warning: If inventory uses "Bell Tower Mausoleum" and we search "Bell Tower", it usually works (contains).
                count = count_row_availability(df_inv, clean_sheet, str(row_val))
                
                if count is not None and count != "N/A":
                    # Update the cell
                    if count == 0:
                        df.at[idx, qty_col] = "Sold Out"
                    else:
                        df.at[idx, qty_col] = count

        # Save Sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Success! Updated Master Price Book saved as: {OUTPUT_FILE}")
