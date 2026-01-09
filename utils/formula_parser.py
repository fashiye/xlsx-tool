"""
安全公式解析与计算（仅支持数值运算、变量、+ - * / ** and parentheses）
基于 ast.NodeVisitor，避免使用 eval。
"""
import ast
import operator as op
import logging

logger = logging.getLogger(__name__)

# 支持的运算符映射
_allowed_operators = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Pow: op.pow,
}

class _FormulaEvaluator(ast.NodeVisitor):
    def __init__(self, variables):
        self.variables = variables
        self.result = None

    def visit_Expr(self, node):
        self.visit(node.value)

    def visit_BinOp(self, node):
        left = self.visit(node.left)
        right = self.visit(node.right)
        operator = _allowed_operators[type(node.op)]
        self.result = operator(left, right)

    def visit_Num(self, node):
        self.result = node.n

    def visit_Name(self, node):
        if node.id not in self.variables:
            raise ValueError(f"Unknown variable: {node.id}")
        self.result = self.variables[node.id]

def evaluate_formula(formula, variables):
    """
    安全地评估数学公式，支持变量引用。
    
    参数:
        formula: 数学公式字符串，如 "A1 + B1 * 2"
        variables: 变量字典，如 {"A1": 10, "B1": 20}
    
    返回:
        float 或 int: 公式计算结果
    
    异常:
        ValueError: 如果公式包含不支持的操作或变量
    """
    try:
        # 解析公式为AST
        tree = ast.parse(formula, mode='eval')
        
        # 创建评估器并访问AST
        evaluator = _FormulaEvaluator(variables)
        evaluator.visit(tree)
        
        return evaluator.result
    except SyntaxError as e:
        logger.error(f"Formula syntax error: {str(e)}")
        raise ValueError(f"Invalid formula syntax: {str(e)}") from e
    except KeyError as e:
        logger.error(f"Unknown variable in formula: {str(e)}")
        raise ValueError(f"Unknown variable: {str(e)}") from e
    except Exception as e:
        logger.error(f"Error evaluating formula: {str(e)}")
        raise ValueError(f"Error evaluating formula: {str(e)}") from e