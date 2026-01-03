"""
数学运算验证器：使用 utils.formula_parser.evaluate_formula
"""
from utils.formula_parser import evaluate_formula

def validate_formula(cells_dict, formula, expected_value, tolerance=0.0):
    """
    cells_dict: { 'A1': 1.2, 'B1': 3 }
    formula: 'A1 + B1 * 2'
    expected_value: numeric expected
    """
    result = evaluate_formula(formula, cells_dict)
    if isinstance(result, (int, float)) and isinstance(expected_value, (int, float)):
        diff = abs(result - expected_value)
        return diff <= tolerance, result, expected_value
    else:
        return result == expected_value, result, expected_value