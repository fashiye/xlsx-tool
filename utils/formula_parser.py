"""
安全公式解析与计算（仅支持数值运算、变量、+ - * / ** and parentheses）
基于 ast.NodeVisitor，避免使用 eval。
"""
import ast
import operator as op
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

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
        logger.info(f"初始化公式计算器，变量: {self.vars}")

    def visit(self, node):
        if isinstance(node, ast.Expression):
            return self.visit(node.body)
        return super().visit(node)

    def visit_BinOp(self, node):
        left = self.visit(node.left)
        right = self.visit(node.right)
        op_type = type(node.op)
        if op_type in _allowed_operators:
            op_name = op_type.__name__
            result = _allowed_operators[op_type](left, right)
            logger.info(f"执行二元运算: {left} {op_name} {right} = {result}")
            return result
        raise ValueError(f"不支持的二元运算符: {op_type}")

    def visit_UnaryOp(self, node):
        operand = self.visit(node.operand)
        op_type = type(node.op)
        if op_type in _allowed_operators:
            op_name = op_type.__name__
            result = _allowed_operators[op_type](operand)
            logger.info(f"执行一元运算: {op_name} {operand} = {result}")
            return result
        raise ValueError(f"不支持的一元运算符: {op_type}")

    def visit_Num(self, node):
        logger.info(f"处理数值常量: {node.n}")
        return node.n

    def visit_Constant(self, node):
        # Python3.8+: ast.Constant
        if isinstance(node.value, (int, float)):
            logger.info(f"处理常量: {node.value}")
            return node.value
        raise ValueError("只支持数值常量")

    def visit_Name(self, node):
        logger.info(f"处理变量: {node.id}")
        if node.id in self.vars:
            val = self.vars[node.id]
            if val is None:
                raise ValueError(f"变量 {node.id} 值为空")
            try:
                num_val = float(val)
                logger.info(f"变量 {node.id} 值: {val} -> 转换为 {num_val}")
                return num_val
            except Exception as e:
                logger.error(f"变量 {node.id} 转换失败: {val}")
                raise ValueError(f"变量 {node.id} 不能转换为数值: {val}") from e
        raise ValueError(f"未提供变量: {node.id}")

    def generic_visit(self, node):
        logger.error(f"不支持的表达式类型: {type(node)}")
        raise ValueError(f"不支持表达式类型: {type(node)}")


def evaluate_formula(formula, variables):
    """
    例: formula="A + B * C", variables={'A':1, 'B':2, 'C':3}
    返回计算结果（float）
    """
    logger.info(f"开始评估公式: {formula}, 变量: {variables}")
    try:
        tree = ast.parse(formula, mode='eval')
        evaluator = FormulaEvaluator(variables)
        result = evaluator.visit(tree)
        logger.info(f"公式评估完成，结果: {result}")
        return result
    except Exception as e:
        logger.error(f"公式评估失败: {str(e)}")
        raise