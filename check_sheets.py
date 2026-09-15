import pandas as pd

EXCEL_PATH = r"C:\Users\Admin\Downloads\pharmacy_app\Ibukunolu new pharmacy database.xlsx"

xls = pd.ExcelFile(EXCEL_PATH)
print("Sheets found:")
print(xls.sheet_names)
