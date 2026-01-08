#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单调试脚本：测试特定规则的解析和验证
"""
import pandas as pd
from core.rule_engine import RuleEngine

# 创建测试数据帧
df1 = pd.DataFrame({
    'A': [1, 2, 3, 4, 5],
    'B': [10, 20, 30, 40, 50],
    'C': [11, 22, 33, 44, 55]
})

# 初始化规则引擎
engine = RuleEngine()

# 测试具体规则
rule_to_test = "A2 * B2 + C2 = 66"
print(f"=== 测试规则: {rule_to_test} ===")

try:
    # 解析规则
    left_expr, op, right_expr = engine.parse_rule(rule_to_test)
    print(f"左侧表达式: {left_expr}")
    print(f"操作符: {op}")
    print(f"右侧表达式: {right_expr}")
    
    # 解析为RPN
    left_rpn = engine.parse_expression(left_expr)
    right_rpn = engine.parse_expression(right_expr)
    print(f"左侧RPN: {left_rpn}")
    print(f"右侧RPN: {right_rpn}")
    
    # 评估表达式
    left_value = engine.evaluate_expression(left_expr, df1)
    right_value = engine.evaluate_expression(right_expr, df1)
    print(f"左侧值: {left_value}")
    print(f"右侧值: {right_value}")
    
    # 验证规则
    result = engine.validate_rule(rule_to_test, df1)
    print(f"规则验证结果: {'通过' if result else '失败'}")
    
except Exception as e:
    print(f"错误: {e}")
    import traceback
    traceback.print_exc()
