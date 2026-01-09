"""
数学运算验证器：使用 utils.formula_parser.evaluate_formula
"""
from utils.formula_parser import evaluate_formula
import logging

logger = logging.getLogger(__name__)

def validate_formula(cells_dict, formula, expected_value, tolerance=0.0):
    """
    cells_dict: { 'A1': 1.2, 'B1': 3 }
    formula: 'A1 + B1 * 2'
    expected_value: numeric expected
    """
    logger.info(f"验证公式: {formula}, 单元格值: {cells_dict}, 期望值: {expected_value}, 容差: {tolerance}")
    try:
        result = evaluate_formula(formula, cells_dict)
        logger.info(f"公式计算结果: {result}")
        
        if isinstance(result, (int, float)) and isinstance(expected_value, (int, float)):
            diff = abs(result - expected_value)
            is_valid = diff <= tolerance
            logger.info(f"数值验证结果: 差值 = {diff}, 在容差范围内 = {is_valid}")
            return is_valid, result, expected_value
        else:
            is_valid = result == expected_value
            logger.info(f"非数值验证结果: {is_valid}")
            return is_valid, result, expected_value
    except Exception as e:
        logger.error(f"公式验证失败: {str(e)}")
        return False, None, expected_value