# =========================================
# PYTHON SAMPLE PROJECT 2 AUTOMATION SCRIPT
# ADDDED MORE DATA INTO FROM AUTOMATION
# Author: [Rohan Sakariya] 
# =========================================

import pandas as pd 
import numpy as np

# -----------------------------------------
# STEP 1 : CREATE RAW DATA FOR 10 EMPLOYEES
# -----------------------------------------
print("GENERATING DATA FOR 10 EMPLOYEES...")

# Data for a mid-sized company department
raw_data = { 
    'Employee ID' : ['E001', 'E002', 'E003', 'E004' ,'E005', 'E006', 'E007', 'E008', 'E009', 'E0010'],
    'Name' : ['Aarav Sharma', 'Priya Singh', 'Vikram Patel', 'Ananya Reddy', 'Karan Malhotra', 'Sneha Desai', 'Rohan Metha', 'Divya Agarwal', 'Alok Kumar', 'Neha Choudary'],
    'Department' : ['IT', 'HR', 'Sales', 'Finance', 'IT', 'Marketing', 'Sales', 'Finance', 'IT', 'HR'],
    'Basic Salary' : [45000, 38000, 35000, 52000, 48000, 42000, 37000, 50000, 46000, 39000],
    'Bonus' : [5000, 3000, 7000, 6000, 5500, 4000, 6500, 5500, 5000, 3200],
    'Total Working Days' : [22, 26, 24, 22, 26, 24, 24, 22, 26, 26], # Varing month lengths
    'Days Present' : [21, 25, 22, 20, 26, 24, 22, 21, 25, 25] # Realistic absences
}

df_raw = pd.DataFrame(raw_data)

# -----------------------------------------
# STEP 2 : DATA VALIDATION & ERROR HANDLING 
# -----------------------------------------
print("PERFORMING DATA VALIDATION CHECKS...")

# Check 1 : Ensure no one was present more than total days
if (df_raw['Days Present'] > df_raw['Total Working Days']).any():
    print("WARNING: Data error found. 'Days Present' > 'Total Working Days'. Correcting...")
    # Fix it: Cap 'Days Present' at 'Total Working Days'
    df_raw['Days Present'] = np.minimum(df_raw['Days present'], df_raw['Total Working Days'])
    
# Check 2 : Ensure no division by zero error 
if (df_raw['Total Working Days'] == 0).any():
    print("WARNING: 'Total Working days' cannot be zero. Replacing with 1...")
    df_raw['Total Working Days'] = df_raw['Total Working Days'].replace(0,1)
    
# -----------------------------------------
# STEP 3 : PERFORM SALARY CALCULATIONS
# -----------------------------------------
print("CALCULATING SALARY COMPONENTS...")

df_calc = df_raw.copy()

# Perform all calculations
df_calc['Per Day Salary'] = df_calc['Basic Salary'] / df_calc['Total Working Days']
df_calc['Salary Earned'] = df_calc['Per Day Salary'] * df_calc['Days Present']
df_calc['Net Salary'] = df_calc['Salary Earned'] + df_calc['Bonus']

# Perform for cleanliness

df_calc['Per Day Salary'] = df_calc['Per Day Salary'].round(2)
df_calc['Salary Earned'] = df_calc['Salary Earned'].round(2)
df_calc['Net Salary'] = df_calc['Net Salary'].round(2)

# Reorder columns in Neat Way
column_order = [
    'Employee ID', 'Name', 'Department', 'Basic Salary', 'Bonus', 'Total Working Days', 'Days Present', 'Per Day Salary', 'Salary Earned', 'Net Salary'
]
df_calc = df_calc[column_order]

# ----------------------------------------------------
# STEP 4 : CREATE DEPARTMENT-WISE SUMMARY PIVOT TABLE 
# ----------------------------------------------------
print("BUILDING DASHBOARD WITH PIVOT TABLES...")

# It shows budget analysis.
pivot_table = pd.pivot_table(
    df_calc,
    index='Department',
    values=['Basic Salary', 'Bonus', 'Net Salary'],
    aggfunc={'Basic Salary' : 'sum', 'Bonus' : 'sum', 'Net Salary' : 'sum'},
    margins=True,
    margins_name='Grand Total'

)
# Calculate department-wise averages
avg_table = pd.pivot_table(
    df_calc,
    index='Department',
    values='Net Salary',
    aggfunc='mean'
).round(2)
avg_table.columns = ['Average Net Salary']

# ----------------------------------------------------
# STEP 5 : WRITE TO EXCEL WITH NEAT FORMATTING
# ----------------------------------------------------

output_filename = "autmate_sample_2_report.xlsx"

# Use Xlsx for advanced formatting 
with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
   # Write raw data
   df_raw.to_excel(writer, sheet_name='Raw Data', index=False)
   df_calc.to_excel(writer, sheet_name='Salary Calculations', index=False) 
   pivot_table.to_excel(writer, sheet_name='Department Summary')
   avg_table.to_excel(writer, sheet_name='Salary Insights')
   
   # Get workbook and worksheet for formatting
   workbook = writer.book
   worksheet_calc = writer.sheets['Salary Calculations']
   
   #Define money format
   money_format = workbook.add_format({'num_format': '₹#,##0.00'})
   
   # Apply formatting to salary columns (Columns D, E, H, I, J)
   for col in ['D', 'E', 'H', 'I', 'J']:
       worksheet_calc.set_column(f'{col}:{col}', 15, money_format)
    
    # Auto-adjust column widhts for all sheets
for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        # Set all columns from A to Z to width 15
        worksheet.set_column('A:Z', 15)
                
# ----------------------------------------------------
# STEP 6 : CONFIRMATION
# ----------------------------------------------------
print(f"SUCCESS! Neat report generated: '{output_filename}'")
print("\n SHEETS CREATED:")
print("1. Raw Data - input from system")
print("2. Salary Calculations - Detailed breakdown with formulas")
print("3. Department Summary - Pivot table for budget analysis")
print("4. Salary Insights - Average salaries per department")
print("5.\nThis easily system can now easily process 100+ eemployees without any changes.")