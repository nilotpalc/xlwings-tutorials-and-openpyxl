#%%
"""
Write a python code for the below instructions:
    1. Use the openpyxl library for the below instructions
2. Load the workbook icarus.xlsm
3. Find the formatted table named 'iris' in worksheet 'Sheet3'
4. Do not use data_validations command
4. Copy the data in the named reference 'iris' as a pandas dataframe
5. The first row is the column header of the DataFrame
6. Close the workbook
"""

import openpyxl
import pandas as pd

tablename = 'iris'

# Load the workbook icarus.xlsm
wb = openpyxl.load_workbook(filename='icarus.xlsm')
#%%
# Find the formatted table named 'iris' in worksheet 'Sheet3'
sheet = wb['Sheet3']

# tables.items returns a dictionary of tables names and their cell references
for (i,v) in ws.tables.items():
    if i == tablename:
        table_ref = v # extract the references for the tablename
    else:
        continue

#%%
# Copy the data in the named reference 'iris' as a pandas dataframe
data = sheet[table_ref]#.destinations[0]]
rows = []
for row in data:
    values = [cell.value for cell in row]
    rows.append(values)
df = pd.DataFrame(rows[1:], columns=rows[0])

# Close the workbook
wb.close()
