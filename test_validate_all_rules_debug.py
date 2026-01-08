import pandas as pd
from core.rule_engine import RuleEngine

# 创建测试数据框
df1 = pd.DataFrame({
    'A': [1, 2, 3, 4],
    'B': [5, 6, 7, 8],
    'C': [9, 10, 11, 12]
})

df2 = pd.DataFrame({
    'A': [1, 2, 5, 4],
    'B': [5, 7, 7, 8],
    'C': [9, 10, 12, 12]
})

# 调试测试
def debug_validate_all_rules():
    engine = RuleEngine()
    
    # 添加规则
    rules = [
        "A1 + B1 = C1",
        "A2 + B2 = C2", 
        "A3 + B3 = C3",
        "FILE1:A1 = FILE2:A1",
        "FILE1:A3 = FILE2:A3"
    ]
    
    for rule in rules:
        engine.add_rule(rule)
        # 逐个测试每个规则
        result = engine.validate_rule(rule, df1, df2)
        print(f"Rule: {rule} -> Result: {result}")
    
    # 验证所有规则
    passed, failed = engine.validate_all_rules(df1, df2)
    print(f"\nPassed rules: {passed}")
    print(f"Failed rules: {failed}")
    print(f"Number of passed rules: {len(passed)}")
    print(f"Number of failed rules: {len(failed)}")

if __name__ == "__main__":
    debug_validate_all_rules()