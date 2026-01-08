#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细调试带有负号的表达式
"""
import pandas as pd
from core.rule_engine import RuleEngine

# 创建测试数据帧
df1 = pd.DataFrame({
    'A': [1, 2, 3, 4, 5],
    'B': [10, 20, 30, 40, 50],
    'C': [11, 22, 33, 44, 55]
})

# 创建RuleEngine的子类，添加调试输出
class DebugRuleEngine(RuleEngine):
    def parse_expression(self, expr):
        """
        重写parse_expression方法，添加详细调试信息
        """
        print(f"\n调试parse_expression:")
        print(f"  表达式: {expr}")
        
        result = super().parse_expression(expr)
        print(f"  结果RPN: {result}")
        return result
    
    def evaluate_expression(self, expr, df1, df2=None):
        """
        重写evaluate_expression方法，添加详细调试信息
        """
        print(f"\n调试evaluate_expression:")
        print(f"  表达式: {expr}")
        
        try:
            result = super().evaluate_expression(expr, df1, df2)
            print(f"  结果: {result}")
            return result
        except Exception as e:
            print(f"  错误: {e}")
            import traceback
            traceback.print_exc()
            raise

# 初始化规则引擎
engine = DebugRuleEngine()

# 测试规则
rule = "A3 * (B3 - C3) = -99"

print(f"=== 测试规则: {rule} ===")

# 解析规则
left_expr, op, right_expr = engine.parse_rule(rule)
print(f"左侧表达式: {left_expr}")
print(f"操作符: {op}")
print(f"右侧表达式: {right_expr}")

# 单独测试右侧表达式（带负号的数字）
print("\n=== 单独测试右侧表达式 '-99' ===")
try:
    value = engine.evaluate_expression(right_expr, df1)
    print(f"值: {value}")
except Exception as e:
    print(f"错误: {e}")

# 单独测试左侧表达式（带括号的减法）
print("\n=== 单独测试左侧表达式 'A3 * (B3 - C3)' ===")
try:
    value = engine.evaluate_expression(left_expr, df1)
    print(f"值: {value}")
except Exception as e:
    print(f"错误: {e}")

# 验证完整规则
print("\n=== 验证完整规则 ===")
try:
    result = engine.validate_rule(rule, df1)
    print(f"规则验证结果: {'通过' if result else '失败'}")
except Exception as e:
    print(f"验证错误: {e}")
