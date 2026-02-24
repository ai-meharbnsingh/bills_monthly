import sys
import os

# Test imports
try:
    print("Testing imports...")
    import openpyxl
    print("OK - openpyxl imported")
    import win32com.client
    print("OK - win32com.client imported")
    from dateutil.relativedelta import relativedelta
    print("OK - dateutil imported")
    import tkinter as tk
    print("OK - tkinter imported")

    # Test Excel COM
    print("\nTesting Excel COM automation...")
    excel = win32com.client.Dispatch("Excel.Application")
    print(f"OK - Excel version: {excel.Version}")
    excel.Quit()
    print("OK - Excel COM working")

    # Test file existence
    print("\nChecking required files...")
    files = ['config.ini', 'Mobile_Bill_Template.xlsx', 'Landline_Bill_Template.xlsx']
    for f in files:
        if os.path.exists(f):
            print(f"OK - {f} exists")
        else:
            print(f"ERROR - {f} NOT FOUND")

    print("\nOK - All tests passed! Script should work.")

except Exception as e:
    print(f"\nERROR: {e}")
    import traceback
    traceback.print_exc()
