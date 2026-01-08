#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细调试evaluate_expression方法的执行过程
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

# 创建RuleEngine的子类，添加调试输出
class DebugRuleEngine(RuleEngine):
    def evaluate_expression(self, expr, df1, df2=None):
        """
        重写evaluate_expression方法，添加详细调试信息
        """
        print(f"\n调试evaluate_expression:")
        print(f"  表达式: {expr}")
        print(f"  df1: {'存在' if df1 is not None else '不存在'}")
        print(f"  df2: {'存在' if df2 is not None else '不存在'}")
        
        try:
            rpn = self.parse_expression(expr)
            print(f"  RPN: {rpn}")
            
            stack = []
            
            for i, token in enumerate(rpn):
                print(f"  步骤 {i+1}: 处理标记 '{token}'")
                print(f"    当前栈: {stack}")
                
                if isinstance(token, (int, float)):
                    print(f"    类型: 数字")
                    stack.append(token)
                    print(f"    栈更新: {stack}")
                elif token in self.operators:
                    print(f"    类型: 操作符")
                    if len(stack) < 2:
                        raise ValueError("无效的表达式")
                    b = stack.pop()
                    a = stack.pop()
                    print(f"    弹出操作数: {a}, {b}")
                    result = self.operators[token][1](a, b)
                    print(f"    计算结果: {a} {token} {b} = {result}")
                    stack.append(result)
                    print(f"    栈更新: {stack}")
                elif isinstance(token, str):
                    print(f"    类型: 单元格引用")
                    if token.startswith('FILE1:'):
                        print(f"    前缀: FILE1: - 使用df1")
                        cell_ref = token[6:]
                        cell_value = self.get_cell_value(cell_ref, df1)
                    elif token.startswith('FILE2:'):
                        print(f"    前缀: FILE2: - 使用df2")
                        if df2 is None:
                            raise ValueError("需要df2参数来处理FILE2:前缀的单元格引用")
                        cell_ref = token[6:]
                        cell_value = self.get_cell_value(cell_ref, df2)
                    else:
                        print(f"    前缀: 无 - 使用df1")
                        cell_value = self.get_cell_value(token, df1)
                    print(f"    获取单元格值: {token} = {cell_value}")
                    stack.append(cell_value)
                    print(f"    栈更新: {stack}")
                else:
                    raise ValueError(f"无效的标记：{token}")
            
            if len(stack) != 1:
                raise ValueError("无效的表达式")
            
            result = stack[0]
            print(f"  最终结果: {result}")
            return result
            
        except Exception as e:
            print(f"  错误: {e}")
            import traceback
            traceback.print_exc()
            raise

# 初始化调试规则引擎
engine = DebugRuleEngine()

# 测试规则
rule = "A2 * B2 + C2 = 66"

# 只传递df1测试
print("\n=== 只传递df1测试 ===")
try:
    result = engine.validate_rule(rule, df1)
    print(f"规则验证结果: {'通过' if result else '失败'}")
except Exception as e:
    print(f"验证错误: {e}")

# 同时传递df1和df2测试
print("\n=== 同时传递df1和df2测试 ===")
try:
    result = engine.validate_rule(rule, df1, df2)
    print(f"规则验证结果: {'通过' if result else '失败'}")
except Exception as e:
    print(f"验证错误: {e}")
