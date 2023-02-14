"""
Write a python code for the below instructions:
1. Use the openpyxl library for the below instructions
2. Load the workbook iris.xlsm
3. Find the named reference named 'iris' in worksheet 'Sheet3'
4. Copy the data in the named reference 'iris' as a pandas dataframe
5. The first row is the column header of the DataFrame
6. Close the workbook
"""

#%%
import openpyxl
import pandas as pd

# Load the workbook
workbook = openpyxl.load_workbook('iris.xlsm')

# Find the named reference named 'iris'
named_range = workbook.defined_names['iris']

# Get the sheet name from the named range
sheet_name = named_range.attr_text.split('!')[0].replace("'", "")
worksheet = workbook[sheet_name]

#%%
# Copy the data in the named reference 'iris' as a pandas dataframe
range_reference = named_range.attr_text.split('!$')[1].replace('$', '')

#%%
data = worksheet[range_reference]
rows = []
for row in data:
    values = [cell.value for cell in row]
    rows.append(values)
df = pd.DataFrame(rows[1:], columns=rows[0])

# Close the workbook
workbook.close()
