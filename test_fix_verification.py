import pandas as pd
import numpy as np
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
                        logging.FileHandler("fix_test.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

def test_basic_cell_comparison():
    """测试基本的单元格比较功能"""
    logger.info("=== 测试基本单元格比较 ===")
    
    # 创建测试数据
    data1 = {'A': [10, 20, 30], 'B': [40, 50, 60]}
    df1 = pd.DataFrame(data1)
    
    data2 = {'A': [10, 25, 30], 'B': [40, 55, 60]}
    df2 = pd.DataFrame(data2)
    
    rule_engine = RuleEngine()
    
    # 应该通过的规则
    pass_rules = [
        "A1 = A1",
        "FILE1:A1 = FILE2:A1",
        "A2 > 10",
        "B3 <= 60"
    ]
    
    # 应该失败的规则
    fail_rules = [
        "A1 = A2",
        "FILE1:A2 = FILE2:A2",
        "A3 < 20",
        "B1 >= 50"
    ]
    
    for rule in pass_rules:
        try:
            result = rule_engine.validate_rule(rule, df1, df2)
            logger.info(f"规则 '{rule}' 验证结果: {result} (预期: True)")
            assert result == True, f"规则 '{rule}' 应该通过但失败了"
        except Exception as e:
            logger.error(f"规则 '{rule}' 验证出错: {e}", exc_info=True)
            raise
    
    for rule in fail_rules:
        try:
            result = rule_engine.validate_rule(rule, df1, df2)
            logger.info(f"规则 '{rule}' 验证结果: {result} (预期: False)")
            assert result == False, f"规则 '{rule}' 应该失败但通过了"
        except Exception as e:
            logger.error(f"规则 '{rule}' 验证出错: {e}", exc_info=True)
            raise
    
    logger.info("基本单元格比较测试通过")

def test_comparison_with_dataframe_result():
    """测试可能返回DataFrame的情况"""
    logger.info("=== 测试可能返回DataFrame的情况 ===")
    
    # 创建测试数据
    data1 = {'A': [10, 20, 30], 'B': [40, 50, 60]}
    df1 = pd.DataFrame(data1)
    
    # 创建一个可能导致DataFrame返回的场景
    # 例如，使用相同的列名但不同的数据结构
    rule_engine = RuleEngine()
    
    # 测试直接比较两个可能导致DataFrame的表达式
    rule = "FILE1:A1 = FILE1:A3"
    try:
        result = rule_engine.validate_rule(rule, df1, df1)
        logger.info(f"规则 '{rule}' 验证结果: {result}")
        assert isinstance(result, bool), f"结果应该是布尔值，实际是 {type(result)}"
    except Exception as e:
        logger.error(f"规则 '{rule}' 验证出错: {e}", exc_info=True)
        raise
    
    logger.info("DataFrame结果处理测试通过")

def test_edge_cases():
    """测试边界情况"""
    logger.info("=== 测试边界情况 ===")
    
    # 创建包含各种边界情况的数据
    data1 = {
        'A': [1, 2, None],  # 包含空值
        'B': [0, 1000000, -1000000],  # 包含大数值和负数
        'C': [0.1, 0.2, 0.3]  # 包含小数值
    }
    df1 = pd.DataFrame(data1)
    
    rule_engine = RuleEngine()
    
    # 测试包含空值的比较
    try:
        result = rule_engine.validate_rule("A1 = A1", df1)
        logger.info(f"空值规则验证结果: {result}")
        assert isinstance(result, bool), f"结果应该是布尔值，实际是 {type(result)}"
    except Exception as e:
        logger.error("空值规则验证出错: {e}", exc_info=True)
        raise
    
    # 测试大数值比较
    try:
        result = rule_engine.validate_rule("B2 > B1", df1)
        logger.info(f"大数值规则验证结果: {result}")
        assert result == True, "大数值比较应该通过"
    except Exception as e:
        logger.error("大数值规则验证出错: {e}", exc_info=True)
        raise
    
    # 测试小数值比较
    try:
        result = rule_engine.validate_rule("C1 + C2 = C3", df1)
        logger.info(f"小数值规则验证结果: {result}")
        # 由于浮点精度问题，0.1+0.2=0.30000000000000004，所以应该失败
        assert result == False, "小数值比较应该失败"
    except Exception as e:
        logger.error("小数值规则验证出错: {e}", exc_info=True)
        raise
    
    logger.info("边界情况测试通过")

def test_rule_engine_validation():
    """测试完整的规则引擎验证流程"""
    logger.info("=== 测试完整的规则引擎验证流程 ===")
    
    # 创建测试数据
    data1 = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df1 = pd.DataFrame(data1)
    
    data2 = {'A': [1, 2, 4], 'B': [4, 5, 7]}
    df2 = pd.DataFrame(data2)
    
    rule_engine = RuleEngine()
    comparator = ExcelComparator()
    
    # 添加多条规则
    rules = [
        "FILE1:A1 = FILE2:A1",
        "FILE1:A2 = FILE2:A2",
        "FILE1:A3 = FILE2:A3",
        "FILE1:B3 < FILE2:B3"
    ]
    
    for rule in rules:
        rule_engine.add_rule(rule)
        comparator.add_rule(rule)
    
    # 测试规则引擎的批量验证
    try:
        passed, failed = rule_engine.validate_with_dataframes(df1, df2)
        logger.info(f"批量验证结果: 通过 {len(passed)} 条，失败 {len(failed)} 条")
        logger.info(f"通过的规则: {passed}")
        logger.info(f"失败的规则: {failed}")
        assert len(passed) == 2, f"应该通过2条规则，实际通过了{len(passed)}条"
        assert len(failed) == 2, f"应该失败2条规则，实际失败了{len(failed)}条"
    except Exception as e:
        logger.error("批量验证出错: {e}", exc_info=True)
        raise
    
    # 测试比较器的批量验证
    try:
        passed, failed = comparator.validate_with_dataframes(df1, df2)
        logger.info(f"比较器批量验证结果: 通过 {len(passed)} 条，失败 {len(failed)} 条")
        assert len(passed) == 2, f"比较器应该通过2条规则，实际通过了{len(passed)}条"
        assert len(failed) == 2, f"比较器应该失败2条规则，实际失败了{len(failed)}条"
    except Exception as e:
        logger.error("比较器批量验证出错: {e}", exc_info=True)
        raise
    
    logger.info("完整规则引擎验证流程测试通过")

def test_gui_scenario():
    """测试GUI场景中的常见操作"""
    logger.info("=== 测试GUI场景 ===")
    
    # 模拟GUI中可能的数据结构
    data1 = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df1 = pd.DataFrame(data1)
    
    # 模拟GUI中添加的规则格式
    gui_rules = [
        "FILE1:A1 = FILE1:A1",
        "FILE1:A2 = FILE1:A3",
        "FILE1:B1 > FILE1:A1",
        "FILE1:B3 <= FILE1:A3"
    ]
    
    rule_engine = RuleEngine()
    
    for rule in gui_rules:
        try:
            # 清理规则中的空格
            rule = rule.strip()
            # 模拟GUI中的验证流程
            result = rule_engine.validate_rule(rule, df1, df1)
            logger.info(f"GUI规则 '{rule}' 验证结果: {result}")
            assert isinstance(result, bool), f"结果应该是布尔值，实际是 {type(result)}"
        except Exception as e:
            logger.error(f"GUI规则 '{rule}' 验证出错: {e}", exc_info=True)
            raise
    
    logger.info("GUI场景测试通过")

if __name__ == "__main__":
    logger.info("开始运行修复验证测试")
    
    try:
        test_basic_cell_comparison()
        test_comparison_with_dataframe_result()
        test_edge_cases()
        test_rule_engine_validation()
        test_gui_scenario()
        
        logger.info("\n=== 所有测试通过！修复验证成功 ===")
    except Exception as e:
        logger.error(f"\n=== 测试失败: {e} ===", exc_info=True)
        sys.exit(1)
    finally:
        logger.info("测试运行结束")
