import pandas as pd
import xlwings as xw

# PHASE 1 — Read raw data
df = pd.read_excel('../data/raw_data.xlsx')

# PHASE 2 — Clean data
df.columns = df.columns.str.strip()
df['City'] = df['City'].str.strip()
df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce')
df = df.dropna(subset=['Sales'])

# PHASE 3 — Create summary
summary = df.groupby('City')['Sales'].sum().reset_index()

# PHASE 4 — Save cleaned data
summary.to_excel('../dashboard/cleaned_data.xlsx', index=False)

# PHASE 5 — Update dashboard
wb = xw.Book('../dashboard/dashboard_template.xlsx')
sheet = wb.sheets['Data']   # first sheet
sheet['A1'].value = summary

wb.save('../dashboard/final_dashboard.xlsx')
wb.close()

print("Dashboard updated successfully!")
