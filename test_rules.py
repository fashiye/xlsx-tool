#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试自定义规则引擎功能
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

print("=== 测试规则引擎功能 ===")

# 测试1: 基本算术规则
print("\n1. 测试基本算术规则")
rule1 = "A1 + B1 = C1"
engine.add_rule(rule1)
result = engine.validate_rule(rule1, df1)
print(f"规则 '{rule1}': {'通过' if result else '失败'}")

# 测试2: 比较运算符
print("\n2. 测试比较运算符")
rule2 = "A1 < B1"
engine.add_rule(rule2)
result = engine.validate_rule(rule2, df1)
print(f"规则 '{rule2}': {'通过' if result else '失败'}")

# 测试3: 复杂算术表达式
print("\n3. 测试复杂算术表达式")
rule3 = "A1 * B1 + 1 = C1"
engine.add_rule(rule3)
result = engine.validate_rule(rule3, df1)
print(f"规则 '{rule3}': {'通过' if result else '失败'}")

# 测试4: 跨文件比较
print("\n4. 测试跨文件比较")
rule4 = "FILE1:A1 = FILE2:A1"
engine.add_rule(rule4)
result = engine.validate_rule(rule4, df1, df2)
print(f"规则 '{rule4}': {'通过' if result else '失败'}")

# 测试5: 跨文件算术比较
print("\n5. 测试跨文件算术比较")
rule5 = "FILE1:A1 + FILE1:B1 = FILE2:A1 + FILE2:B1"
engine.add_rule(rule5)
result = engine.validate_rule(rule5, df1, df2)
print(f"规则 '{rule5}': {'通过' if result else '失败'}")

# 测试6: 验证所有规则
print("\n6. 测试验证所有规则")
engine.clear_rules()
engine.add_rule("A1 + B1 = C1")
engine.add_rule("A2 + B2 = C2")
engine.add_rule("A5 + B5 = C5")  # 这个应该失败
passed, failed = engine.validate_all_rules(df1)
print(f"通过的规则: {passed}")
print(f"失败的规则: {failed}")

print("\n=== 测试完成 ===")
