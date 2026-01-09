"""
Excel 读取器：使用 pandas.read_excel 读取所有 sheets
"""
import pandas as pd
import logging

logger = logging.getLogger(__name__)

def load_workbook_all_sheets(filepath):
    """
    返回: dict { sheet_name: DataFrame }
    """
    logger.info(f"开始加载工作簿: {filepath}")
    try:
        # pandas will use openpyxl/xlrd depending on file type
        sheets = pd.read_excel(filepath, sheet_name=None, engine=None)
        logger.info(f"工作簿加载成功，包含 {len(sheets)} 个工作表: {list(sheets.keys())}")
        return sheets
    except Exception as e:
        logger.error(f"加载工作簿失败: {str(e)}")
        raise