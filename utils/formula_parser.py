"""
安全公式解析与计算（仅支持数值运算、变量、+ - * / ** and parentheses）
基于 ast.NodeVisitor，避免使用 eval。
"""
import ast
import operator as op

# 支持的运算符映射
_allowed_operators = {
    ast.Add: op.add,
    ast.Sub: op.sub,
    ast.Mult: op.mul,
    ast.Div: op.truediv,
    ast.Pow: op.pow,
    ast.USub: op.neg,
    ast.UAdd: op.pos,
    ast.Mod: op.mod,
}

class FormulaEvaluator(ast.NodeVisitor):
    def __init__(self, variables):
        self.vars = variables or {}

    def visit(self, node):
        if isinstance(node, ast.Expression):
            return self.visit(node.body)
        return super().visit(node)

    def visit_BinOp(self, node):
        left = self.visit(node.left)
        right = self.visit(node.right)
        op_type = type(node.op)
        if op_type in _allowed_operators:
            return _allowed_operators[op_type](left, right)
        raise ValueError(f"不支持的二元运算符: {op_type}")

    def visit_UnaryOp(self, node):
        operand = self.visit(node.operand)
        op_type = type(node.op)
        if op_type in _allowed_operators:
            return _allowed_operators[op_type](operand)
        raise ValueError(f"不支持的一元运算符: {op_type}")

    def visit_Num(self, node):
        return node.n

    def visit_Constant(self, node):
        # Python3.8+: ast.Constant
        if isinstance(node.value, (int, float)):
            return node.value
        raise ValueError("只支持数值常量")

    def visit_Name(self, node):
        if node.id in self.vars:
            val = self.vars[node.id]
            if val is None:
                raise ValueError(f"变量 {node.id} 值为空")
            try:
                return float(val)
            except:
                raise ValueError(f"变量 {node.id} 不能转换为数值: {val}")
        raise ValueError(f"未提供变量: {node.id}")

    def generic_visit(self, node):
        raise ValueError(f"不支持表达式类型: {type(node)}")

def evaluate_formula(formula, variables):
    """
    例: formula="A + B * C", variables={'A':1, 'B':2, 'C':3}
    返回计算结果（float）
    """
    try:
        tree = ast.parse(formula, mode='eval')
        evaluator = FormulaEvaluator(variables)
        return evaluator.visit(tree)
    except Exception as e:
        raise