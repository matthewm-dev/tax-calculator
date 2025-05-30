"""
Test Microsoft 365 Integration
Save this in OneNote as your first successful integration!
"""
import openpyxl
import os
from datetime import datetime

# Test 1: File system access
print("Current directory:", os.getcwd())
print("OneDrive accessible:", os.path.exists(os.path.expanduser("~/OneDrive")))

# Test 2: Excel integration
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'] = "Integration Test"
sheet['A2'] = f"Created on {datetime.now()}"
workbook.save("integration_test.xlsx")
print("Excel file created successfully!")

# Test 3: Access to sample data
print("\nReady to build your tax calculator!")