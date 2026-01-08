#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试规则引擎的括号处理
"""
from core.rule_engine import RuleEngine

# 初始化规则引擎
engine = RuleEngine()

# 测试表达式解析
print("=== 调试表达式解析 ===")

# 测试基本表达式
test_exprs = [
    "A1 + B1",
    "A1 * (B1 + C1)",
    "A1 * B1 + C1",
    "(A1 + B1) * C1",
    "A1 * (B1 + C1) * D1"
]

for expr in test_exprs:
    try:
        rpn = engine.parse_expression(expr)
        print(f"表达式: '{expr}'")
        print(f"RPN: {rpn}")
    except Exception as e:
        print(f"表达式: '{expr}'")
        print(f"错误: {e}")
    print()

# 测试完整规则
test_rules = [
    "A1 * (B1 + C1) = 121",
    "(A1 + B1) * C1 = 121"
]

print("\n=== 调试规则解析 ===")
for rule in test_rules:
    try:
        left, op, right = engine.parse_rule(rule)
        print(f"规则: '{rule}'")
        print(f"左侧表达式: '{left}'")
        print(f"操作符: '{op}'")
        print(f"右侧表达式: '{right}'")
        
        # 测试左侧表达式的RPN转换
        left_rpn = engine.parse_expression(left)
        print(f"左侧RPN: {left_rpn}")
        
        # 测试右侧表达式的RPN转换
        right_rpn = engine.parse_expression(right)
        print(f"右侧RPN: {right_rpn}")
        
    except Exception as e:
        print(f"规则: '{rule}'")
        print(f"错误: {e}")
    print()
