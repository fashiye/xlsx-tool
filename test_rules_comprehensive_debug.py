#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
带调试信息的全面测试脚本
"""
import pandas as pd
from core.rule_engine import RuleEngine

# 创建测试数据帧
df1 = pd.DataFrame({
    'A': [1, 2, 3, 4, 5],
    'B': [10, 20, 30, 40, 50],
    'C': [11, 22, 33, 44, 55]
})

df2 = pd.DataFrame({
    'A': [1, 2, 3, 4, 6],  # 最后一行与df1不同
    'B': [10, 20, 30, 40, 50],
    'C': [11, 22, 33, 44, 56]  # 最后一行与df1不同
})

# 创建RuleEngine的子类，添加调试输出
class DebugRuleEngine(RuleEngine):
    def validate_rule(self, rule, df1, df2=None):
        """
        重写validate_rule方法，添加详细调试信息
        """
        print(f"\n--- 调试规则验证: {rule} ---")
        try:
            result = super().validate_rule(rule, df1, df2)
            print(f"结果: {'通过' if result else '失败'}")
            return result
        except Exception as e:
            print(f"错误: {e}")
            import traceback
            traceback.print_exc()
            return False

# 初始化规则引擎
engine = DebugRuleEngine()

print("=== 全面测试规则引擎功能 ===")

# 测试5: 测试复杂表达式
print("\n5. 测试复杂算术表达式")
complex_rules = [
    "A1 * (B1 + C1) = 121",  # 1 * (10 + 11) = 21（失败）
    "A2 * B2 + C2 = 66",     # 2 * 20 + 22 = 62（失败）
    "A3 * (B3 - C3) = -99"   # 3 * (30 - 33) = -9（失败）
]

for rule in complex_rules:
    result = engine.validate_rule(rule, df1)
    status = "✓ 通过" if result else "✗ 失败"
    print(f"规则 '{rule}': {status}")

print("\n=== 全面测试完成 ===")
