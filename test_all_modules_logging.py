#!/usr/bin/env python3
"""
测试所有核心模块的日志功能
"""
import logging
import pandas as pd
from core.comparator import ExcelComparator
from core.string_comparator import StringComparator
from core.excel_reader import load_workbook_all_sheets
from core.rule_engine import RuleEngine

# 配置根日志记录器
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

def test_logging():
    logger.info("开始测试所有模块的日志功能")
    
    # 测试 string_comparator 模块
    logger.info("\n1. 测试 string_comparator 模块")
    string_comp = StringComparator()
    string_comp.exact_match("Hello", "hello", ignore_case=True)
    string_comp.fuzzy_match("Hello", "Hallo", method='levenshtein', threshold=0.8)
    
    # 测试 rule_engine 模块
    logger.info("\n2. 测试 rule_engine 模块")
    rule_engine = RuleEngine()
    df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6], 'C': [7, 8, 9]})
    df2 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6], 'C': [7, 8, 9]})
    rule_engine.validate_rule("A1 = B1", df1)
    rule_engine.validate_rule("A1 = FILE2:A1", df1, df2)
    
    # 测试 excel_reader 模块
    logger.info("\n3. 测试 excel_reader 模块")
    # 创建一个简单的测试Excel文件
    test_df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
    test_df.to_excel("test_logging.xlsx", index=False)
    try:
        sheets = load_workbook_all_sheets("test_logging.xlsx")
        logger.info(f"加载的工作表: {list(sheets.keys())}")
    except Exception as e:
        logger.error(f"加载Excel文件失败: {e}")
    
    # 测试 comparator 模块
    logger.info("\n4. 测试 comparator 模块")
    comparator = ExcelComparator()
    try:
        comparator.load_workbook("test_logging.xlsx", alias="TEST_FILE")
        logger.info(f"工作簿加载成功，别名: TEST_FILE")
    except Exception as e:
        logger.error(f"加载工作簿失败: {e}")
    
    logger.info("所有模块的日志功能测试完成")

if __name__ == "__main__":
    test_logging()