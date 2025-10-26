import pandas as pd
import os

# folder with Excel files
folder = './excel_files'

# input files
file1 = os.path.join(folder, 'file1.xlsx')
file2 = os.path.join(folder, 'file2.xlsx')

# file names without extension
name1 = os.path.splitext(os.path.basename(file1))[0]
name2 = os.path.splitext(os.path.basename(file2))[0]

# clean company code
def clean_code(x):
    if pd.isna(x):
        return None
    s = str(x).strip().replace('\u00A0', ' ')
    if s.endswith('.0'):
        s = s[:-2]
    s = s.replace(' ', '')
    digits = ''.join(ch for ch in s if ch.isdigit())
    return digits or None

# clean company name
def clean_name(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None

# normalize name for compare
def normalize(name):
    if not name:
        return ''
    s = ' '.join(name.lower().split())
    return s

# merge unique names
def merge_unique(series):
    vals = {clean_name(v) for v in series if pd.notna(v)}
    vals = {v for v in vals if v}
    if not vals:
        return None
    return ' | '.join(sorted(vals))

# read Excel files
df1 = pd.read_excel(file1, usecols=['Registrikood', 'Nimi'])
df2 = pd.read_excel(file2, usecols=['Registrikood', 'Ettevõtja nimetus'])

# clean and rename columns
df1['Code'] = df1['Registrikood'].map(clean_code)
df1['Name_1'] = df1['Nimi'].map(clean_name)
df1 = df1[['Code', 'Name_1']].dropna(subset=['Code'])

df2['Code'] = df2['Registrikood'].map(clean_code)
df2['Name_2'] = df2['Ettevõtja nimetus'].map(clean_name)
df2 = df2[['Code', 'Name_2']].dropna(subset=['Code'])

# combine duplicates
df1 = df1.groupby('Code', as_index=False)['Name_1'].apply(merge_unique)
df2 = df2.groupby('Code', as_index=False)['Name_2'].apply(merge_unique)

# merge by code
merged = pd.merge(df1, df2, on='Code', how='outer')

# check where company exists
def status(row):
    in1 = pd.notna(row['Name_1']) and row['Name_1'] != ''
    in2 = pd.notna(row['Name_2']) and row['Name_2'] != ''
    if in1 and in2:
        return 'In both files'
    if in1:
        return f'Only in {name1}'
    return f'Only in {name2}'

merged['Status'] = merged.apply(status, axis=1)

# check if names different
def diff(n1, n2):
    if not n1 or not n2:
        return ''
    set1 = {normalize(x) for x in str(n1).split(' | ')}
    set2 = {normalize(x) for x in str(n2).split(' | ')}
    return 'Yes' if set1 != set2 else 'No'

merged['Different Names'] = merged.apply(lambda r: diff(r['Name_1'], r['Name_2']), axis=1)

# reorder and sort
merged = merged[['Code', 'Name_1', 'Name_2', 'Different Names', 'Status']]
merged = merged.sort_values(by='Code').reset_index(drop=True)

# save result
out_file = os.path.join(folder, 'Comparison_Result.xlsx')
merged.to_excel(out_file, index=False)

print()
print("Result saved successfully")
print(f"File: {out_file}")
print()
print("Rows compared:", len(merged))
print("Comparison complete")
