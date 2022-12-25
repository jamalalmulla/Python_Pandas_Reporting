import pandas as pd
from openpyxl import load_workbook
df = pd.read_excel('Filtered_Excel/filename.xlsx')
filtered_df = df[(df['Country']=='UK') & (df['Status']=='Yes')]
book = load_workbook('Filtered_Excel/filename.xlsx')
writer = pd.ExcelWriter('Filtered_Excel/filename.xlsx', engine = 'openpyxl')
writer.book = book
filtered_df.to_excel(writer, sheet_name = 'sheet2')
writer.save()
writer.close()
