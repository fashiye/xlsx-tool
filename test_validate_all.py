#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试validate_all_rules方法在同时传递两个数据帧时的行为
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
    'A': [1, 2, 3, 4, 6],
    'B': [10, 20, 30, 40, 50],
    'C': [11, 22, 33, 44, 56]
})

# 初始化规则引擎
engine = RuleEngine()

# 测试规则列表
test_rules = [
    "A1 + B1 = C1",  # 通过
    "A2 * B2 + C2 = 66",  # 应该失败：2 * 20 + 22 = 62
    "A3 * (B3 - C3) = -99"  # 应该失败：3 * (30 - 33) = -9
]

print("=== 测试validate_all_rules方法 ===")

# 逐个验证规则
print("\n1. 逐个验证规则:")
for rule in test_rules:
    try:
        result = engine.validate_rule(rule, df1, df2)
        print(f"规则 '{rule}': {'通过' if result else '失败'}")
    except Exception as e:
        print(f"规则 '{rule}': 错误 - {e}")

# 添加规则并批量验证
print("\n2. 批量验证规则:")
engine.clear_rules()
for rule in test_rules:
    engine.add_rule(rule)

try:
    passed, failed = engine.validate_all_rules(df1, df2)
    print(f"总规则数: {len(test_rules)}")
    print(f"通过的规则数: {len(passed)}")
    print(f"失败的规则数: {len(failed)}")
    print(f"通过的规则: {passed}")
    print(f"失败的规则: {failed}")
except Exception as e:
    print(f"批量验证错误: {e}")
    import traceback
    traceback.print_exc()
