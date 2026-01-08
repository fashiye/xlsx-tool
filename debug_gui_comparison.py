import pandas as pd
import os
import traceback
from core.comparator import ExcelComparator
from core.rule_engine import RuleEngine
import core.excel_reader as excel_reader

# 创建测试数据
data = {
    'A': [1, 2, 3],
    'B': [4, 5, 6]
}

df = pd.DataFrame(data)

# 创建测试文件
if not os.path.exists('test_debug.xlsx'):
    df.to_excel('test_debug.xlsx', index=False)
    print(f"创建测试文件: test_debug.xlsx")

# 模拟完整的GUI比较流程
print("=== 调试GUI比较流程 ===")
print(f"当前工作目录: {os.getcwd()}")
print(f"测试文件存在: {os.path.exists('test_debug.xlsx')}")

# 1. 初始化比较器
comparator = ExcelComparator()

# 2. 模拟GUI中的文件加载
file_path = 'test_debug.xlsx'
alias = 'file1'

print(f"\n1. 加载文件: {file_path}")
try:
    # 加载工作簿
    comparator.load_workbook(file_path, alias)
    print(f"   工作簿加载成功")
    
    # 获取工作表列表
    sheets = comparator.list_sheets(alias)
    print(f"   工作表列表: {sheets}")
    
    # 加载工作表数据
    sheet_name = sheets[0] if sheets else None
    if sheet_name:
        df1 = comparator.get_sheet_dataframe(alias, sheet_name)
        print(f"   工作表数据加载成功，形状: {df1.shape}")
        print(f"   数据内容:\n{df1}")
    
    # 再加载一次作为file2
    comparator.load_workbook(file_path, 'file2')
    df2 = comparator.get_sheet_dataframe('file2', sheet_name)
    
except Exception as e:
    print(f"   错误: {e}")
    traceback.print_exc()

# 3. 模拟添加规则
print(f"\n2. 添加规则")
rule = "FILE1:A1= FILE1:A3"
print(f"   规则: {rule}")

try:
    comparator.add_rule(rule)
    print(f"   规则添加成功")
    print(f"   当前规则列表: {comparator.get_rules()}")
except Exception as e:
    print(f"   错误: {e}")
    traceback.print_exc()

# 4. 调试规则解析和验证
print(f"\n3. 调试规则解析和验证")
try:
    # 直接使用rule_engine进行测试
    rule_engine = comparator.rule_engine
    
    # 解析规则
    print(f"   解析规则: {rule}")
    left_expr, op, right_expr = rule_engine.parse_rule(rule)
    print(f"   解析结果: left='{left_expr}', op='{op}', right='{right_expr}'")
    
    # 求值表达式
    print(f"   求值表达式")
    left_value = rule_engine.evaluate_expression(left_expr, df1, df2)
    print(f"   左表达式 '{left_expr}' = {left_value}, 类型: {type(left_value)}")
    
    right_value = rule_engine.evaluate_expression(right_expr, df1, df2)
    print(f"   右表达式 '{right_expr}' = {right_value}, 类型: {type(right_value)}")
    
    # 验证规则
    print(f"   验证规则")
    result = rule_engine.validate_rule(rule, df1, df2)
    print(f"   验证结果: {result}")
    
except Exception as e:
    print(f"   错误: {e}")
    traceback.print_exc()

# 5. 调试validate_with_dataframes方法
print(f"\n4. 调试validate_with_dataframes方法")
try:
    print(f"   调用comparator.validate_with_dataframes(df1, df2)")
    passed_rules, failed_rules = comparator.validate_with_dataframes(df1, df2)
    print(f"   比较结果: 通过{len(passed_rules)}条，失败{len(failed_rules)}条")
    print(f"   通过规则: {passed_rules}")
    print(f"   失败规则: {failed_rules}")
except Exception as e:
    print(f"   错误: {e}")
    traceback.print_exc()
    print(f"   完整错误信息:")
    traceback.print_exc()

# 清理测试文件
if os.path.exists('test_debug.xlsx'):
    os.remove('test_debug.xlsx')
    print(f"\n清理测试文件: test_debug.xlsx")

print(f"\n=== 调试完成 ===")
