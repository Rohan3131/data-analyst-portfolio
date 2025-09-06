# =============================================================================
# Automated Payroll Management System
# Developed By: Rohan Sakariya
# Project Description:
# This system automates the complete payroll process for companies. It calculates
# employee salaries, bonuses, and generates detailed reports. I built this to
# demonstrate my ability to solve real-world business problems using programming.
# =============================================================================

import pandas as pd 
import numpy as np
from openpyxl import load_workbook

print("Starting Payroll Management System...")
print("Generating sample data for 10 employees...")

# =============================================================================
# SAMPLE DATA - Simulating real company employee data
# This represents how data would come from HR department database
# =============================================================================
raw_data = { 
    'Employee ID': ['E001', 'E002', 'E003', 'E004', 'E005', 'E006', 'E007', 'E008', 'E009', 'E010'],
    'Name': ['Aarav Sharma', 'Priya Singh', 'Vikram Patel', 'Ananya Reddy', 'Karan Malhotra', 'Sneha Desai', 'Rohan Metha', 'Divya Agarwal', 'Alok Kumar', 'Neha Choudary'],
    'Department': ['IT', 'HR', 'Sales', 'Finance', 'IT', 'Marketing', 'Sales', 'Finance', 'IT', 'HR'],
    'Basic Salary': [45000, 38000, 35000, 52000, 48000, 42000, 37000, 50000, 46000, 39000],
    'Bonus': [5000, 3000, 7000, 6000, 5500, 4000, 6500, 5500, 5000, 3200],
    'Total Working Days': [22, 26, 24, 22, 26, 24, 24, 22, 26, 26],
    'Days Present': [21, 25, 22, 20, 26, 24, 22, 21, 25, 25]
}

df_raw = pd.DataFrame(raw_data)

# =============================================================================
# DATA VALIDATION - Ensuring data quality before processing
# Attention to detail and error-handling skills
# =============================================================================
print("Checking data for possible errors...")

if (df_raw['Days Present'] > df_raw['Total Working Days']).any():
    print("Found inconsistency: Some employees have more attendance than working days")
    print("Automatically correcting the data...")
    df_raw['Days Present'] = np.minimum(df_raw['Days Present'], df_raw['Total Working Days'])
    
if (df_raw['Total Working Days'] == 0).any():
    print("Critical error: Zero working days found")
    print("Fixing to prevent calculation errors...")
    df_raw['Total Working Days'] = df_raw['Total Working Days'].replace(0, 1)

# =============================================================================
# SALARY CALCULATION - Core logic of the system
# Mathematical and logical programming skills
# =============================================================================
print("Calculating salary components...")

df_calc = df_raw.copy()

# Salary calculation formulas 
df_calc['Per Day Salary'] = df_calc['Basic Salary'] / df_calc['Total Working Days']
df_calc['Salary Earned'] = df_calc['Per Day Salary'] * df_calc['Days Present']
df_calc['Net Salary'] = df_calc['Salary Earned'] + df_calc['Bonus']

# Formatting for financial figures
df_calc['Per Day Salary'] = df_calc['Per Day Salary'].round(2)
df_calc['Salary Earned'] = df_calc['Salary Earned'].round(2)
df_calc['Net Salary'] = df_calc['Net Salary'].round(2)

# Organizing data for better readability
column_order = ['Employee ID', 'Name', 'Department', 'Basic Salary', 'Bonus', 'Total Working Days', 'Days Present', 'Per Day Salary', 'Salary Earned', 'Net Salary']
df_calc = df_calc[column_order]

# =============================================================================
# MANAGEMENT REPORTING - Creating insights for decision makers
# reate business intelligence reports
# =============================================================================
print("Generating management reports...")

# Department-wise financial summary for budget planning
pivot_table = pd.pivot_table(
    df_calc,
    index='Department',
    values=['Basic Salary', 'Bonus', 'Net Salary'],
    aggfunc='sum',
    margins=True,
    margins_name='Grand Total'
)

# Average salary analysis for HR compensation planning
avg_table = pd.pivot_table(
    df_calc,
    index='Department',
    values='Net Salary',
    aggfunc='mean'
).round(2)
avg_table.columns = ['Department Average Salary']

# =============================================================================
# EXCEL REPORT GENERATION -  output delivery
# creating user-ready business reports
# =============================================================================
print("Creating  Excel report...")

output_filename = "payroll_report.xlsx"

with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
    # Four different reports in one file
    df_raw.to_excel(writer, sheet_name='Raw Data', index=False)
    df_calc.to_excel(writer, sheet_name='Calculated Salaries', index=False) 
    pivot_table.to_excel(writer, sheet_name='Department Summary')
    avg_table.to_excel(writer, sheet_name='Salary Insights')
   
    workbook = writer.book
    rupee_format = workbook.add_format({'num_format': '₹#,##0.00'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    
    # Formatting for salary sheet
    ws_calc = writer.sheets['Calculated Salaries']
    ws_calc.set_column('A:A', 12)    # Employee ID
    ws_calc.set_column('B:B', 16)    # Employee Name
    ws_calc.set_column('C:C', 11)    # Department
    ws_calc.set_column('D:D', 16, rupee_format) # Basic Salary
    ws_calc.set_column('E:E', 16, rupee_format) # Bonus Amount
    ws_calc.set_column('F:G', 14, number_format) # Attendance Days
    ws_calc.set_column('H:J', 18, rupee_format) # Calculated Salaries
    
    # Formatting for management reports
    ws_summary = writer.sheets['Department Summary']
    ws_summary.set_column('A:A', 12)  # Department
    ws_summary.set_column('B:B', 20, rupee_format) # Basic Salary
    ws_summary.set_column('C:C', 20, rupee_format) # Bonus
    ws_summary.set_column('D:D', 20, rupee_format) # Net Salary
    
    ws_insights = writer.sheets['Salary Insights']
    ws_insights.set_column('A:A', 12)  # Department
    ws_insights.set_column('B:B', 20, rupee_format) # Average Salary

# =============================================================================
# FINAL TOUCH-UP - Ensuring perfect presentation
# Extra step 
# =============================================================================
print("Adding final  touches...")

wb = load_workbook(output_filename)

# Ensuring perfect column widths for management presentation
ws_summary = wb['Department Summary']
ws_summary.column_dimensions['B'].width = 22 # Basic Salary
ws_summary.column_dimensions['C'].width = 22 # Bonus
ws_summary.column_dimensions['D'].width = 22 # Net Salary

wb.save(output_filename)

# =============================================================================
# PROJECT COMPLETION
# =============================================================================
print("Payroll processing completed successfully!")
print("Report Generated: 'payroll_report.xlsx'")
print("\n" + "="*60)
print("REPORT CONTENTS:")
print("Raw Data - Original input data")
print("Calculated Salaries - Detailed individual calculations")
print("Department Summary - Budget analysis for management")
print("Salary Insights - Compensation trends by department")
print("="*60)

print("""
   WHY THIS PROJECT MATTERS:
- Demonstrates end-to-end automation capability
- Shows strong data validation and error handling
- Provides actionable business intelligence
- Ready for enterprise-level implementation
- Can handle 1000+ employees with same code
""")

print("I am excited about the opportunity to apply these skills")
print("In a professional environment and contribute to your team!")