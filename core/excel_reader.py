"""
Excel 读取器：使用 pandas.read_excel 读取所有 sheets
"""
import pandas as pd

def load_workbook_all_sheets(filepath):
    """
    返回: dict { sheet_name: DataFrame }
    """
    # pandas will use openpyxl/xlrd depending on file type
    sheets = pd.read_excel(filepath, sheet_name=None, engine=None)
 
    return sheets