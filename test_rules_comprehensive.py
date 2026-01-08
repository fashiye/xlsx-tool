#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全面测试自定义规则引擎功能
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

# 初始化规则引擎
engine = RuleEngine()

print("=== 全面测试规则引擎功能 ===")

# 测试1: 应该通过的规则
print("\n1. 测试应该通过的规则")
pass_rules = [
    "A1 + B1 = C1",  # 1 + 10 = 11
    "A2 + B2 = C2",  # 2 + 20 = 22
    "A3 < B3",       # 3 < 30
    "A4 * B4 = C4 * 3.6363636363636365"  # 4 * 40 = 44 * 3.6363636363636365
]

for rule in pass_rules:
    result = engine.validate_rule(rule, df1)
    status = "✓ 通过" if result else "✗ 失败（预期通过）"
    print(f"规则 '{rule}': {status}")
    if not result:
        engine.add_rule(rule)  # 只添加失败的规则用于后续测试

# 测试2: 应该失败的规则
print("\n2. 测试应该失败的规则")
fail_rules = [
    "A1 + B1 = C2",  # 1 + 10 = 22（错误）
    "A5 > B5",       # 5 > 50（错误）
    "A1 * B1 = C1 + 1"  # 1 * 10 = 11 + 1（错误）
]

for rule in fail_rules:
    result = engine.validate_rule(rule, df1)
    status = "✗ 失败" if not result else "✓ 通过（预期失败）"
    print(f"规则 '{rule}': {status}")
    engine.add_rule(rule)  # 添加到规则列表用于后续测试

# 测试3: 跨文件测试
print("\n3. 测试跨文件规则")
cross_file_rules = [
    "FILE1:A1 = FILE2:A1",  # 1 = 1（通过）
    "FILE1:A5 = FILE2:A5",  # 5 = 6（失败）
    "FILE1:A1 + FILE1:B1 = FILE2:C1"  # 1 + 10 = 11（通过）
]

for rule in cross_file_rules:
    result = engine.validate_rule(rule, df1, df2)
    status = "✓ 通过" if result else "✗ 失败"
    print(f"规则 '{rule}': {status}")
    engine.add_rule(rule)  # 添加到规则列表用于后续测试

# 测试4: 验证所有规则
print("\n4. 测试批量验证所有规则")
engine.clear_rules()  # 先清空之前的规则

# 添加混合规则
test_rules = [
    "A1 + B1 = C1",  # 通过
    "A2 + B2 = C2",  # 通过
    "A5 + B5 = C5 + 1",  # 失败：5 + 50 = 55 + 1
    "FILE1:A1 = FILE2:A1",  # 通过
    "FILE1:A5 = FILE2:A5"  # 失败：5 = 6
]

for rule in test_rules:
    engine.add_rule(rule)

passed, failed = engine.validate_all_rules(df1, df2)
print(f"总规则数: {len(test_rules)}")
print(f"通过的规则数: {len(passed)}")
print(f"失败的规则数: {len(failed)}")
print(f"\n通过的规则: {passed}")
print(f"\n失败的规则: {failed}")

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
