import pandas as pd
from config import EXCEL_PATH

def read_sheet(sheet_name):
    return pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)