import pandas as pd
import numpy as np
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.rule_engine import RuleEngine
from core.comparator import ExcelComparator
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("bug_test.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

def test_bug_reproduction():
    """重现DataFrame布尔值错误的测试"""
    logger.info("开始测试bug重现")
    
    # 创建简单的测试数据帧
    data1 = {
        'A': [1, 2, 3],
        'B': [4, 5, 6]
    }
    df1 = pd.DataFrame(data1)
    
    data2 = {
        'A': [1, 2, 3],
        'B': [4, 5, 6]
    }
    df2 = pd.DataFrame(data2)
    
    logger.info(f"df1形状: {df1.shape}, 内容:\n{df1}")
    logger.info(f"df2形状: {df2.shape}, 内容:\n{df2}")
    
    # 创建规则引擎
    rule_engine = RuleEngine()
    comparator = ExcelComparator()
    
    # 添加规则
    rule = "FILE1:A1 = FILE1:A3"
    rule_engine.add_rule(rule)
    comparator.add_rule(rule)
    
    logger.info(f"添加规则: {rule}")
    
    try:
        # 测试rule_engine.validate_with_dataframes
        logger.info("测试rule_engine.validate_with_dataframes")
        passed, failed = rule_engine.validate_with_dataframes(df1, df2)
        logger.info(f"rule_engine测试结果: passed={passed}, failed={failed}")
    except Exception as e:
        logger.error(f"rule_engine测试失败: {e}", exc_info=True)
    
    try:
        # 测试comparator.validate_with_dataframes
        logger.info("测试comparator.validate_with_dataframes")
        passed, failed = comparator.validate_with_dataframes(df1, df2)
        logger.info(f"comparator测试结果: passed={passed}, failed={failed}")
    except Exception as e:
        logger.error(f"comparator测试失败: {e}", exc_info=True)
    
    # 测试单条规则验证
    try:
        logger.info("测试单条规则验证")
        result = rule_engine.validate_rule(rule, df1, df2)
        logger.info(f"单条规则验证结果: {result}")
    except Exception as e:
        logger.error(f"单条规则验证失败: {e}", exc_info=True)
    
    logger.info("测试结束")

if __name__ == "__main__":
    test_bug_reproduction()
