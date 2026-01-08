"""
自定义规则引擎：解析和执行数据校验规则
"""
import re
import pandas as pd
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

class RuleEngine:
    """
    规则引擎，用于解析和执行自定义数据校验规则
    """
    def __init__(self):
        self.rules = []
        # 支持的运算符优先级（从低到高）
        self.operators = {
            '=': (1, lambda a, b: a == b),
            '!=': (1, lambda a, b: a != b),
            '<': (1, lambda a, b: a < b),
            '<=': (1, lambda a, b: a <= b),
            '>': (1, lambda a, b: a > b),
            '>=': (1, lambda a, b: a >= b),
            '+': (2, lambda a, b: a + b),
            '-': (2, lambda a, b: a - b),
            '*': (3, lambda a, b: a * b),
            '/': (3, lambda a, b: a / b if b != 0 else float('inf')),
            '%': (3, lambda a, b: a % b if b != 0 else 0)
        }
        
    def add_rule(self, rule):
        """
        添加规则到规则列表
        
        参数:
            rule: 规则字符串，如 "A1 + B1 = C1"
        """
        self.rules.append(rule)
    
    def clear_rules(self):
        """清空规则列表"""
        self.rules.clear()
    
    def parse_rule(self, rule):
        """
        解析规则字符串为左侧表达式、操作符和右侧表达式

        参数:
            rule: 规则字符串，如 "A1 + B1 = C1"

        返回:
            tuple: (left_expr, operator, right_expr)
        """
        logger.info(f"解析规则: {rule}")
        # 查找比较操作符
        for op in ['>=', '<=', '!=', '=', '>', '<']:
            if op in rule:
                parts = rule.split(op, 1)
                result = (parts[0].strip(), op, parts[1].strip())
                logger.info(f"规则解析结果: {result}")
                return result
        logger.error(f"无效的规则格式：缺少比较操作符: {rule}")
        raise ValueError("无效的规则格式：缺少比较操作符")
    
    def parse_expression(self, expr):
        """
        使用逆波兰表达式（RPN）解析算术表达式
        
        参数:
            expr: 算术表达式字符串，如 "A1 + B1 * C1"
            
        返回:
            list: 逆波兰表达式
        """
        tokens = []
        i = 0
        n = len(expr)
        
        while i < n:
            # 跳过空格
            if expr[i].isspace():
                i += 1
                continue
            
            # 解析负号（一元运算符）
            if expr[i] == '-' and (i == 0 or expr[i-1].isspace() or expr[i-1] in '+-*/%()'):
                # 负号后面应该跟数字或括号
                i += 1
                # 检查是否是负数
                if i < n and (expr[i].isdigit() or expr[i] == '.'):
                    # 是负数，解析数字
                    j = i
                    while j < n and (expr[j].isdigit() or expr[j] == '.'):
                        j += 1
                    num_str = '-' + expr[i:j]
                    tokens.append(float(num_str) if '.' in num_str else int(num_str))
                    i = j
                else:
                    # 不是负数，可能是二元减号
                    tokens.append('-')
            
            # 解析FILE1:或FILE2:前缀的单元格引用
            elif i + 6 <= n and expr[i:i+6] == 'FILE1:':
                j = i + 6
                # 提取单元格引用
                while j < n and (expr[j].isalpha() or expr[j].isdigit() or expr[j] == ':'):
                    j += 1
                token = expr[i:j]
                # 检查是否是范围引用
                if token.count(':') > 1:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格引用）")
                # 进一步检查是否是范围引用（例如FILE1:A1:B2）
                if ':' in token[6:]:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格引用）")
                tokens.append(token)  # 包括FILE1:前缀
                i = j
            elif i + 6 <= n and expr[i:i+6] == 'FILE2:':
                j = i + 6
                # 提取单元格引用
                while j < n and (expr[j].isalpha() or expr[j].isdigit() or expr[j] == ':'):
                    j += 1
                token = expr[i:j]
                # 检查是否是范围引用
                if token.count(':') > 1:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格引用）")
                # 进一步检查是否是范围引用（例如FILE2:A1:B2）
                if ':' in token[6:]:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格引用）")
                tokens.append(token)  # 包括FILE2:前缀
                i = j
            # 解析普通单元格引用（如A1）
            elif expr[i].isalpha():
                j = i
                while j < n and (expr[j].isalpha() or expr[j].isdigit()):
                    j += 1
                token = expr[i:j]
                # 检查是否是范围引用
                if ':' in token:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格引用）")
                tokens.append(token)
                i = j
            # 解析数字
            elif expr[i].isdigit() or expr[i] == '.':
                j = i
                while j < n and (expr[j].isdigit() or expr[j] == '.'):
                    j += 1
                tokens.append(float(expr[i:j]) if '.' in expr[i:j] else int(expr[i:j]))
                i = j
            # 解析操作符
            elif expr[i] in self.operators:
                # 处理双字符操作符
                if i + 1 < n and expr[i:i+2] in self.operators:
                    tokens.append(expr[i:i+2])
                    i += 2
                else:
                    tokens.append(expr[i])
                    i += 1
            # 解析括号
            elif expr[i] == '(':
                tokens.append('(')
                i += 1
            elif expr[i] == ')':
                tokens.append(')')
                i += 1
            else:
                raise ValueError(f"无效的字符：{expr[i]}")
        
        # 转换为逆波兰表达式
        output = []
        stack = []
        
        for token in tokens:
            if isinstance(token, (int, float)) or (isinstance(token, str) and (token.isalnum() or token.startswith('FILE1:') or token.startswith('FILE2:'))):
                # 数字或单元格引用直接输出
                output.append(token)
            elif token == '(':
                # 左括号入栈
                stack.append(token)
            elif token == ')':
                # 右括号：弹出栈顶元素直到遇到左括号
                while stack and stack[-1] != '(':
                    output.append(stack.pop())
                if not stack:
                    raise ValueError("括号不匹配")
                stack.pop()  # 弹出左括号
            elif token in self.operators:
                # 操作符：弹出栈顶优先级更高或相等的操作符
                while stack and stack[-1] != '(' and \
                      self.operators[stack[-1]][0] >= self.operators[token][0]:
                    output.append(stack.pop())
                stack.append(token)
        
        # 弹出剩余操作符
        while stack:
            if stack[-1] == '(':
                raise ValueError("括号不匹配")
            output.append(stack.pop())
        
        return output
    
    def evaluate_expression(self, expr, df1, df2=None):
        """
        评估表达式的值，支持FILE1:和FILE2:前缀的单元格引用
        
        参数:
            expr: 算术表达式字符串，如 "A1 + B1" 或 "FILE1:A1 + FILE2:B1"
            df1: 第一个数据帧（默认数据帧）
            df2: 第二个数据帧（可选，用于跨文件比较）
            
        返回:
            float: 表达式的值
        """
        rpn = self.parse_expression(expr)
        stack = []
        
        for token in rpn:
            if isinstance(token, (int, float)):
                # 数字直接入栈
                stack.append(token)
            elif token in self.operators:
                # 操作符：弹出两个操作数，计算结果后入栈
                if len(stack) < 2:
                    raise ValueError("无效的表达式")
                b = stack.pop()
                a = stack.pop()
                
                # 确保操作数是标量值
                if hasattr(a, 'shape'):
                    if hasattr(a, 'iloc'):
                        a = a.iloc[0, 0] if a.shape[0] > 0 and a.shape[1] > 0 else 0.0
                    elif hasattr(a, 'item'):
                        a = a.item()
                    else:
                        a = float(a[0]) if len(a) > 0 else 0.0
                
                if hasattr(b, 'shape'):
                    if hasattr(b, 'iloc'):
                        b = b.iloc[0, 0] if b.shape[0] > 0 and b.shape[1] > 0 else 0.0
                    elif hasattr(b, 'item'):
                        b = b.item()
                    else:
                        b = float(b[0]) if len(b) > 0 else 0.0
                
                result = self.operators[token][1](a, b)
                
                # 确保结果是标量值
                if hasattr(result, 'shape'):
                    if hasattr(result, 'iloc'):
                        result = result.iloc[0, 0] if result.shape[0] > 0 and result.shape[1] > 0 else 0.0
                    elif hasattr(result, 'item'):
                        result = result.item()
                    else:
                        result = float(result[0]) if len(result) > 0 else 0.0
                
                stack.append(result)
            elif isinstance(token, str):
                # 单元格引用：获取值
                if token.startswith('FILE1:'):
                    # FILE1前缀，使用df1
                    cell_ref = token[6:]
                    cell_value = self.get_cell_value(cell_ref, df1)
                elif token.startswith('FILE2:'):
                    # FILE2前缀，使用df2
                    if df2 is None:
                        raise ValueError("需要df2参数来处理FILE2:前缀的单元格引用")
                    cell_ref = token[6:]
                    cell_value = self.get_cell_value(cell_ref, df2)
                else:
                    # 默认使用df1
                    cell_value = self.get_cell_value(token, df1)
                stack.append(cell_value)
            else:
                raise ValueError(f"无效的标记：{token}")
        
        if len(stack) != 1:
            raise ValueError("无效的表达式")
        
        # 确保最终结果是标量值
        final_result = stack[0]
        if hasattr(final_result, 'shape'):
            if hasattr(final_result, 'iloc'):
                final_result = final_result.iloc[0, 0] if final_result.shape[0] > 0 and final_result.shape[1] > 0 else 0.0
            elif hasattr(final_result, 'item'):
                final_result = final_result.item()
            else:
                final_result = float(final_result[0]) if len(final_result) > 0 else 0.0
        
        return final_result
    
    def get_cell_value(self, cell_ref, df):
        """
        从数据帧中获取单元格值
        
        参数:
            cell_ref: 单元格引用，如 "A1"（不支持范围）
            df: 数据帧
            
        返回:
            float: 单元格的值
        """
        # 解析列字母和行号
        match = re.match(r'^([A-Za-z]+)(\d+)$', cell_ref)
        if not match:
            raise ValueError(f"无效的单元格引用：{cell_ref}（不支持范围引用）")
        
        col_letters = match.group(1)
        row_str = match.group(2)
        
        # 转换列字母为索引（A=0, B=1, ..., AA=26）
        col_idx = 0
        for ch in col_letters.upper():
            col_idx = col_idx * 26 + (ord(ch) - ord('A') + 1)
        col_idx -= 1
        
        # 转换行号为索引（1-based -> 0-based）
        row_idx = int(row_str) - 1
        
        # 检查范围
        if col_idx < 0 or col_idx >= df.shape[1]:
            raise ValueError(f"列索引超出范围：{col_letters}")
        if row_idx < 0 or row_idx >= df.shape[0]:
            raise ValueError(f"行索引超出范围：{row_str}")
        
        # 获取值
        value = df.iloc[row_idx, col_idx]
        
        # 转换为数值
        if pd.isna(value):
            return 0.0
        elif isinstance(value, (int, float)):
            return float(value)
        else:
            try:
                return float(value)
            except ValueError:
                return 0.0
    
    def validate_rule(self, rule, df1, df2=None):
        """
        验证单条规则，支持单DataFrame或双DataFrame比较

        参数:
            rule: 规则字符串，如 "A1 + B1 = C1" 或 "FILE1:A1 = FILE2:A1"
            df1: 第一个数据帧（默认数据帧）
            df2: 第二个数据帧（可选，用于跨文件比较）

        返回:
            bool: 规则是否通过验证
        """
        logger.info(f"验证规则: {rule}")
        try:
            left_expr, op, right_expr = self.parse_rule(rule)
            left_value = self.evaluate_expression(left_expr, df1, df2)
            right_value = self.evaluate_expression(right_expr, df1, df2)
            
            logger.info(f"表达式求值结果 - 左: {left_expr} = {left_value}, 右: {right_expr} = {right_value}")
            
            # 确保操作数是标量值，避免DataFrame布尔上下文问题
            if hasattr(left_value, 'shape'):
                # 如果是DataFrame或Series，尝试转换为标量
                if hasattr(left_value, 'iloc'):
                    left_value = left_value.iloc[0, 0] if left_value.shape[0] > 0 and left_value.shape[1] > 0 else 0.0
                elif hasattr(left_value, 'item'):
                    left_value = left_value.item()
                else:
                    left_value = float(left_value[0]) if len(left_value) > 0 else 0.0
            
            if hasattr(right_value, 'shape'):
                # 如果是DataFrame或Series，尝试转换为标量
                if hasattr(right_value, 'iloc'):
                    right_value = right_value.iloc[0, 0] if right_value.shape[0] > 0 and right_value.shape[1] > 0 else 0.0
                elif hasattr(right_value, 'item'):
                    right_value = right_value.item()
                else:
                    right_value = float(right_value[0]) if len(right_value) > 0 else 0.0
            
            # 执行比较操作
            try:
                if op == '=':
                    result = abs(left_value - right_value) < 1e-6  # 考虑浮点误差
                elif op == '!=':
                    result = abs(left_value - right_value) >= 1e-6
                elif op == '<':
                    result = left_value < right_value
                elif op == '<=':
                    result = left_value <= right_value
                elif op == '>':
                    result = left_value > right_value
                elif op == '>=':
                    result = left_value >= right_value
                else:
                    return False
                
                # 最严格的结果处理：确保结果是标量
                if isinstance(result, (bool, int, float)):
                    final_result = bool(result)
                elif hasattr(result, 'any'):
                    # 如果是DataFrame或Series，使用any()获取单个布尔值
                    final_result = bool(result.any())
                elif hasattr(result, 'all'):
                    final_result = bool(result.all())
                elif hasattr(result, 'iloc'):
                    # 如果仍然是DataFrame，尝试获取第一个值
                    final_result = bool(result.iloc[0, 0] if result.shape[0] > 0 and result.shape[1] > 0 else False)
                elif hasattr(result, 'item'):
                    final_result = bool(result.item())
                elif hasattr(result, '__len__') and len(result) > 0:
                    final_result = bool(result[0])
                else:
                    final_result = bool(result)
                
                logger.info(f"规则验证结果: {rule} -> {final_result}")
                return final_result
            except Exception as e:
                # 如果比较操作出错，返回False
                logger.error(f"比较操作错误：{e}")
                return False
        except Exception as e:
            # 解析或计算错误时返回False
            logger.error(f"规则验证错误：{e}")
            return False
    
    def validate_all_rules(self, df1, df2=None):
        """
        验证所有规则，支持单DataFrame或双DataFrame比较
        
        参数:
            df1: 第一个数据帧（默认数据帧）
            df2: 第二个数据帧（可选，用于跨文件比较）
            
        返回:
            tuple: (passed_rules, failed_rules)
                - passed_rules: 通过的规则列表
                - failed_rules: 失败的规则列表
        """
        passed = []
        failed = []
        
        for rule in self.rules:
            if self.validate_rule(rule, df1, df2):
                passed.append(rule)
            else:
                failed.append(rule)
        
        return passed, failed
    
    def validate_with_dataframes(self, df1, df2):
        """
        使用两个数据帧验证规则（支持跨文件比较）
        
        参数:
            df1: 文件1的数据帧
            df2: 文件2的数据帧
            
        返回:
            tuple: (passed_rules, failed_rules)
        """
        return self.validate_all_rules(df1, df2)
