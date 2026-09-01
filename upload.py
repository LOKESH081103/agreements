import pandas as pd
from sqlalchemy import create_engine

# 1. Load Excel file (ALL SHEETS)
excel_file = "31-aug-del-dump.xlsb"  # Replace with your actual filename
all_sheets = pd.read_excel(excel_file, sheet_name=None)

# 2. Clean column names FOR EACH SHEET before combining
for sheet_name, df_sheet in all_sheets.items():
    formatted_cols = []
    
    # Loop through headers to catch dates BEFORE they get mashed together
    for col in df_sheet.columns:
        
        # 1. Catch actual Datetime/Timestamp objects 
        if hasattr(col, 'strftime'):
            # Formats to "March 23", "April 23", etc.
            formatted_cols.append(col.strftime('%B %y'))
            
        # 2. Catch strings that look like dates (e.g., "01/03/2023" or "2023-03-01 00:00:00")
        elif isinstance(col, str) and ('/' in col or '-' in col or '00:00' in col):
            try:
                # dayfirst=True ensures 01/03/2023 is March 1st, not Jan 3rd
                dt = pd.to_datetime(col, dayfirst=True)
                formatted_cols.append(dt.strftime('%B %y'))
            except Exception:
                formatted_cols.append(col)
                
        # 3. Standard text columns (like "Name", "ID")
        else:
            formatted_cols.append(str(col))
            
    # Assign the nicely formatted 'March 23' strings back to headers
    df_sheet.columns = formatted_cols

    # NOW run your string cleaning. "March 23" will safely become "march_23"
    df_sheet.columns = (
        df_sheet.columns
        .str.strip()
        .str.lower()
        .str.replace(' ', '_')
        .str.replace('[^a-zA-Z0-9_]', '', regex=True)
    )

# 3. NOW combine them. Because headers are perfectly clean, "hi" matches "hi"
df = pd.concat(all_sheets.values(), ignore_index=True)

# 3.5 Deduplicate column names 
seen = {}
new_cols = []
for col in df.columns:
    if col in seen:
        seen[col] += 1
        new_cols.append(f"{col}_{seen[col]}")
    else:
        seen[col] = 0
        new_cols.append(col)

df.columns = new_cols

# 4. Convert all date/month/year/time columns into true datetime objects
date_keywords = ['date', 'month', 'year', 'time', 'dt', 'period']

for col in df.columns:
    if any(kw in str(col).lower() for kw in date_keywords): 
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], unit='D', origin='1899-12-30', errors='coerce')
        else:
            df[col] = pd.to_datetime(df[col], format='%d/%m/%Y', dayfirst=True, errors='coerce')

# 5. Connect and re-upload to PostgreSQL
db_url = "postgresql://postgres:Virumandi007$@localhost:5432/postgres"
engine = create_engine(db_url)

df.to_sql('august31', engine, if_exists='replace', index=False, chunksize=1000)

print("Re-upload complete! Date headers are now formatted as march_23, april_23, etc.")