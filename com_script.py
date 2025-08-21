import pandas as pd
import os

# File paths (adjust folder path as needed)
folder = r'./excel_files'  # generic folder for Excel files
file1 = os.path.join(folder, 'file1.xlsx')
file2 = os.path.join(folder, 'file2.xlsx')

name1 = os.path.splitext(os.path.basename(file1))[0]  # file1 base name
name2 = os.path.splitext(os.path.basename(file2))[0]  # file2 base name

# Helper functions for cleaning
def clean_code(x: object) -> str | None:
    """Standardize company code: convert to str, trim, remove '.0', spaces, keep only digits."""
    if pd.isna(x):
        return None
    s = str(x).strip().replace('\u00A0', ' ')
    if s.endswith('.0'):
        s = s[:-2]
    s = s.replace(' ', '')
    digits = ''.join(ch for ch in s if ch.isdigit())
    return digits or None

def clean_name(x: object) -> str | None:
    """Trim names and remove empty strings."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None

def normalize_for_compare(name: str) -> str:
    """Normalize names for comparison: lowercase, single spaces."""
    if name is None:
        return ''
    s = ' '.join(name.lower().split())
    return s

def aggregate_unique(series: pd.Series) -> str | None:
    """Aggregate unique names from a series into a single string separated by ' | '."""
    vals = {clean_name(v) for v in series if pd.notna(v)}
    vals = {v for v in vals if v}
    if not vals:
        return None
    return ' | '.join(sorted(vals))

# Read specific columns
df1_raw = pd.read_excel(file1, usecols=['Registrikood', 'Nimi'], dtype={'Registrikood': str})
df2_raw = pd.read_excel(file2, usecols=['Registrikood', 'Ettevõtja nimetus'], dtype={'Registrikood': str})

# Clean and rename columns
df1 = df1_raw.copy()
df1['Code'] = df1['Registrikood'].map(clean_code)
df1['Company_Name_1'] = df1['Nimi'].map(clean_name)
df1 = df1.drop(columns=['Registrikood', 'Nimi']).dropna(subset=['Code'])

df2 = df2_raw.copy()
df2['Code'] = df2['Registrikood'].map(clean_code)
df2['Company_Name_2'] = df2['Ettevõtja nimetus'].map(clean_name)
df2 = df2.drop(columns=['Registrikood', 'Ettevõtja nimetus']).dropna(subset=['Code'])

# Aggregate duplicates by code
df1_agg = df1.groupby('Code', as_index=False)['Company_Name_1'].apply(aggregate_unique)
df2_agg = df2.groupby('Code', as_index=False)['Company_Name_2'].apply(aggregate_unique)

# Merge data on code
merged = pd.merge(df1_agg, df2_agg, on='Code', how='outer')

# Determine presence status
def get_status(row):
    in_1 = pd.notna(row['Company_Name_1']) and row['Company_Name_1'] != ''
    in_2 = pd.notna(row['Company_Name_2']) and row['Company_Name_2'] != ''
    if in_1 and in_2:
        return 'In both files'
    elif in_1:
        return f'Only in {name1}'
    else:
        return f'Only in {name2}'

merged['Status'] = merged.apply(get_status, axis=1)

# Flag different names for the same code
def names_different(n1, n2) -> str:
    if not n1 or not n2:
        return ''
    set1 = {normalize_for_compare(x) for x in str(n1).split(' | ') if x}
    set2 = {normalize_for_compare(x) for x in str(n2).split(' | ') if x}
    return 'Yes' if set1 != set2 else 'No'

merged['Different Names'] = merged.apply(lambda r: names_different(r['Company_Name_1'], r['Company_Name_2']), axis=1)

# Reorder and sort columns
merged = merged[['Code', 'Company_Name_1', 'Company_Name_2', 'Different Names', 'Status']]
merged = merged.sort_values(by='Code', kind='stable').reset_index(drop=True)

# Save result
result_file = os.path.join(folder, 'Comparison_Result.xlsx')
merged.to_excel(result_file, index=False)

print(f"✅ Done. Result saved: {result_file}")
