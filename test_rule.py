#!/usr/bin/env python3
"""
测试用户报告的规则验证问题
"""
import pandas as pd
from core.rule_engine import RuleEngine
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 创建测试数据框
data = {
    'A': ['桃', 1, 2, 3],
    'B': ['梨', 4, 5, 6],
    'c': ['总计', 5, 7, 9]
}
df = pd.DataFrame(data)

# 初始化规则引擎
rule_engine = RuleEngine()

# 测试规则列表 - 使用正确的格式
test_rules = [
    '(4*FILE1:A+FILE1:B)/(FILE1:A*FILE1:B)=2',  # 修正后的规则1（列引用）
    '(4*FILE1:A2+FILE1:B2)/(FILE1:A2*FILE1:B2)=2',  # 修正后的规则2（单元格引用，第2行）
    '(4*FILE1:A3+FILE1:B3)/(FILE1:A3*FILE1:B3)=2',  # 修正后的规则3（单元格引用，第3行）
    '(4*FILE1:A4+FILE1:B4)/(FILE1:A4*FILE1:B4)=2',  # 修正后的规则4（单元格引用，第4行）
    '4*FILE1:A2+FILE1:B2=FILE1:c2'  # 简单规则作为对比（第2行）
]

# 测试每条规则
for rule in test_rules:
    print(f"\n{'='*50}")
    print(f"测试规则: {rule}")
    print(f"{'='*50}")
    try:
        is_valid, failed_cells, passed_cells = rule_engine.validate_rule(rule, df)
        print(f"验证结果: {is_valid}")
        print(f"失败单元格: {failed_cells}")
        print(f"通过单元格: {passed_cells}")
    except Exception as e:
        print(f"验证失败: {e}")
        logger.exception(f"规则验证异常: {rule}")
