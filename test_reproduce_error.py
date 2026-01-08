import pandas as pd
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.rule_engine import RuleEngine
from core.comparator import ExcelComparator
import logging

# 配置详细日志记录
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("reproduce_error.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

def test_reproduce_error():
    """复现用户日志中的错误场景"""
    logger.info("=== 开始复现错误场景 ===")
    
    # 创建测试数据，模拟日志中的场景
    # 根据日志，用户使用了FILE1:A1= FILE1:A3这样的规则
    data1 = {'A': [10, 20, 30], 'B': [40, 50, 60]}
    df1 = pd.DataFrame(data1)
    
    # 测试规则：FILE1:A1= FILE1:A3
    rule = "FILE1:A1= FILE1:A3"
    
    try:
        logger.info(f"测试规则: {rule}")
        
        # 创建规则引擎
        rule_engine = RuleEngine()
        
        # 验证规则，使用相同的DataFrame作为df1和df2
        result = rule_engine.validate_rule(rule, df1, df1)
        
        logger.info(f"规则验证结果: {result} (类型: {type(result)})")
        logger.info("=== 错误复现测试完成 ===")
        
    except Exception as e:
        logger.error(f"测试中发生错误: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    test_reproduce_error()