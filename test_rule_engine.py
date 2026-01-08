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

def test_parse_rule():
    """测试规则解析功能"""
    engine = RuleEngine()
    
    # 测试基本规则解析
    left, op, right = engine.parse_rule("A1 + B1 = C1")
    assert left == "A1 + B1"
    assert op == "="
    assert right == "C1"
    
    # 测试不同操作符
    left, op, right = engine.parse_rule("A1 > B1")
    assert op == ">"
    
    left, op, right = engine.parse_rule("A1 < B1")
    assert op == "<"
    
    left, op, right = engine.parse_rule("A1 != B1")
    assert op == "!="
    
    left, op, right = engine.parse_rule("A1 >= B1")
    assert op == ">="
    
    left, op, right = engine.parse_rule("A1 <= B1")
    assert op == "<="

def test_evaluate_expression():
    """测试表达式求值功能"""
    engine = RuleEngine()
    
    # 测试单个单元格引用
    assert engine.evaluate_expression("A1", df1) == 1.0
    assert engine.evaluate_expression("B2", df1) == 6.0
    
    # 测试基本算术运算
    assert engine.evaluate_expression("A1 + B1", df1) == 6.0
    assert engine.evaluate_expression("B1 - A1", df1) == 4.0
    assert engine.evaluate_expression("A1 * B1", df1) == 5.0
    assert engine.evaluate_expression("B1 / A1", df1) == 5.0
    
    # 测试复杂表达式
    assert engine.evaluate_expression("A1 + B1 * C1", df1) == 46.0
    assert engine.evaluate_expression("(A1 + B1) * C1", df1) == 54.0
    
    # 测试跨文件引用
    assert engine.evaluate_expression("FILE1:A1", df1) == 1.0
    assert engine.evaluate_expression("FILE2:A3", df1, df2) == 5.0
    
    # 测试混合引用
    assert engine.evaluate_expression("FILE1:A1 + FILE2:B2", df1, df2) == 1.0 + 7.0

def test_validate_rule():
    """测试规则验证功能"""
    engine = RuleEngine()
    
    # 测试通过的规则
    assert engine.validate_rule("A1 + B1 = C1", df1) is False
    assert engine.validate_rule("A1 = FILE2:A1", df1, df2) is True
    
    # 测试失败的规则
    assert engine.validate_rule("A1 + B1 = C2", df1) is False
    assert engine.validate_rule("A1 > B1", df1) is False
    
    # 测试跨文件比较
    assert engine.validate_rule("FILE1:A3 = FILE2:A3", df1, df2) is False
    assert engine.validate_rule("FILE1:A1 = FILE2:A1", df1, df2) is True
    
    # 测试不同操作符
    assert engine.validate_rule("A1 < B1", df1) is True
    assert engine.validate_rule("B1 > A1", df1) is True
    assert engine.validate_rule("A1 <= B1", df1) is True
    assert engine.validate_rule("B1 >= A1", df1) is True
    assert engine.validate_rule("A1 != B1", df1) is True

def test_validate_all_rules():
    """测试验证所有规则功能"""
    engine = RuleEngine()
    
    # 添加规则
    engine.add_rule("A1 + B1 = C1")
    engine.add_rule("A2 + B2 = C2")
    engine.add_rule("A3 + B3 = C3")
    engine.add_rule("FILE1:A1 = FILE2:A1")
    engine.add_rule("FILE1:A3 = FILE2:A3")
    
    # 验证规则
    passed, failed = engine.validate_all_rules(df1, df2)
    
    # 检查结果
    assert len(passed) == 1  # 只有FILE1:A1 = FILE2:A1通过
    assert len(failed) == 4  # 其他四条规则应该失败

def test_range_references():
    """测试范围引用处理（应该抛出错误）"""
    engine = RuleEngine()
    
    # 测试范围引用解析（应该抛出错误）
    try:
        engine.parse_expression("A1:B2")
        assert False, "应该抛出ValueError"
    except ValueError:
        assert True
    
    try:
        engine.parse_expression("FILE1:A1:B2")
        assert False, "应该抛出ValueError"
    except ValueError:
        assert True
    
    try:
        engine.get_cell_value("A1:B2", df1)
        assert False, "应该抛出ValueError"
    except ValueError:
        assert True

def test_invalid_rules():
    """测试无效规则处理"""
    engine = RuleEngine()
    
    # 测试无效单元格引用
    assert engine.validate_rule("Z1 = A1", df1) is False  # 列超出范围
    assert engine.validate_rule("A10 = A1", df1) is False  # 行超出范围
    
    # 测试无效表达式
    assert engine.validate_rule("A1 + = C1", df1) is False  # 缺少操作数
    assert engine.validate_rule("A1 + B1 * = C1", df1) is False  # 无效表达式

def test_float_comparison():
    """测试浮点比较"""
    df_float = pd.DataFrame({
        'A': [1.0, 2.0, 3.0000001],
        'B': [1.0000001, 2.0, 3.0]
    })
    
    engine = RuleEngine()
    
    # 测试浮点容差
    assert engine.validate_rule("A1 = B1", df_float) is True  # 在容差范围内
    assert engine.validate_rule("A3 = B3", df_float) is True  # 在容差范围内

if __name__ == "__main__":
    # 运行所有测试
    test_parse_rule()
    test_evaluate_expression()
    test_validate_rule()
    test_validate_all_rules()
    test_range_references()
    test_invalid_rules()
    test_float_comparison()
    print("所有测试通过！")
