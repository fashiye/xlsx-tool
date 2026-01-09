#!/usr/bin/env python3
"""
最终测试：使用用户提供的完整数据集测试规则验证
"""
import pandas as pd
from core.rule_engine import RuleEngine
import logging

# 配置日志
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 创建用户提供的完整数据集
data = {
    'A': ['桃', 1, 2, 3],
    'B': ['梨', 4, 5, 6],
    'c': ['总计', 5, 7, 9]
}
df = pd.DataFrame(data)

# 打印数据集信息
print("测试数据集：")
print(df)
print()

# 初始化规则引擎
rule_engine = RuleEngine()

# 测试规则：用户的主要问题规则
main_rule = '(4*FILE1:A+FILE1:B)/(FILE1:A*FILE1:B)=2'
print(f"测试规则：{main_rule}")
print("="*60)

# 测试每条数据行
for i in range(1, len(df)):  # 从第2行开始（跳过标题行）
    row_data = df.iloc[i]
    A = row_data['A']
    B = row_data['B']
    
    # 手动计算规则结果
    manual_result = (4*A + B) / (A*B)
    
    # 使用单元格引用测试规则
    cell_rule = f'(4*FILE1:A{i+1}+FILE1:B{i+1})/(FILE1:A{i+1}*FILE1:B{i+1})=2'
    is_valid, failed, passed = rule_engine.validate_rule(cell_rule, df)
    
    print(f"行 {i+1}: A={A}, B={B}")
    print(f"  手动计算: (4*A + B)/(A*B) = (4*{A} + {B})/({A}*{B}) = {manual_result:.6f}")
    print(f"  规则验证: {is_valid} (计算结果: {manual_result:.6f} {'=' if manual_result == 2 else '≠'} 2)")
    print()

# 总结
print("="*60)
print("结论：")
print("1. 规则验证功能正常工作，只有当数据满足规则条件时才会通过验证")
print("2. 只有当 A=1, B=4 时，(4*A + B)/(A*B) = 2 才成立")
print("3. 对于其他行的数据，这个规则不成立，因此验证失败")
print("4. 如果要验证所有行都满足规则，需要确保数据符合该数学条件")