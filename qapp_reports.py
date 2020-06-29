# QApp Reports
# 29.06.2020

# Import needed modules
import pandas as pd
from pandas import ExcelWriter

# Import the reports file
excel_file = 'reports.xlsx'
cols = ['report ID', 'staff name', 'category', 'question text', 'answer 1', 'date']
df = pd.read_excel(excel_file, usecols = cols)

# Use only the ones we need in Spain
df = df.loc[df['staff name'] == 'David Arroyo']

# Change the date format
df['date'] = pd.to_datetime(df['date'], format = '%d.%m.%Y')

# Fill the NaN with zeros
df = df.fillna(0)

# Create the "NOT_OK" dataframe
df_not_ok = df.loc[df['answer 1'] == 'NOT OK']
df_not_ok['Month'] = df_not_ok['date'].dt.to_period('M')

# Output to Excel
name = 'NOT_OK Reports.xlsx'
writer = pd.ExcelWriter(name)
df_not_ok.to_excel(writer, index = False)
writer.save()
print('Done bro!')