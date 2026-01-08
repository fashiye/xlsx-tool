"""
自定义规则引擎：解析和执行数据校验规则
"""
import re
import pandas as pd
import logging

# 配置日志记录
logging.basicConfig(level=logging.DEBUG, 
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
            
            # 解析FILE1:或FILE2:前缀的单元格或列引用
            elif i + 6 <= n and expr[i:i+6] == 'FILE1:':
                j = i + 6
                # 提取单元格或列引用
                while j < n and (expr[j].isalpha() or expr[j].isdigit() or expr[j] == ':'):
                    j += 1
                token = expr[i:j]
                # 检查是否是复杂范围引用（超过一个冒号）
                if token.count(':') > 1:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格或列引用）")
                tokens.append(token)  # 包括FILE1:前缀
                i = j
            elif i + 6 <= n and expr[i:i+6] == 'FILE2:':
                j = i + 6
                # 提取单元格或列引用
                while j < n and (expr[j].isalpha() or expr[j].isdigit() or expr[j] == ':'):
                    j += 1
                token = expr[i:j]
                # 检查是否是复杂范围引用（超过一个冒号）
                if token.count(':') > 1:
                    raise ValueError(f"不支持的范围引用：{token}（请使用单个单元格或列引用）")
                tokens.append(token)  # 包括FILE2:前缀
                i = j
            # 解析普通单元格或列引用（如A1或A）
            elif expr[i].isalpha():
                j = i
                while j < n and (expr[j].isalpha() or expr[j].isdigit()):
                    j += 1
                token = expr[i:j]
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
        logger.debug(f"evaluate_expression - 输入表达式: {expr}")
        logger.debug(f"evaluate_expression - df1类型: {type(df1)}, df1形状: {df1.shape}")
        logger.debug(f"evaluate_expression - df2类型: {type(df2)}, df2形状: {df2.shape if df2 is not None else 'None'}")
        
        rpn = self.parse_expression(expr)
        logger.debug(f"evaluate_expression - 解析后的RPN表达式: {rpn}")
        
        stack = []
        
        for token in rpn:
            logger.debug(f"evaluate_expression - 处理标记: {token}, 类型: {type(token)}")
            
            if isinstance(token, (int, float)):
                # 数字直接入栈
                logger.debug(f"evaluate_expression - 数字标记，直接入栈: {token}")
                stack.append(token)
            elif token in self.operators:
                # 操作符：弹出两个操作数，计算结果后入栈
                if len(stack) < 2:
                    logger.error(f"evaluate_expression - 操作符{token}需要两个操作数，但栈中只有{len(stack)}个元素")
                    raise ValueError("无效的表达式")
                
                b = stack.pop()
                a = stack.pop()
                logger.debug(f"evaluate_expression - 弹出操作数: a={a} (类型: {type(a)}), b={b} (类型: {type(b)})")
                
                # 确保操作数是标量值
                if hasattr(a, 'shape'):
                    logger.debug(f"evaluate_expression - 操作数a是DataFrame/Series类型，需要转换为标量")
                    if hasattr(a, 'iloc'):
                        a = a.iloc[0, 0] if a.shape[0] > 0 and a.shape[1] > 0 else 0.0
                    elif hasattr(a, 'item'):
                        a = a.item()
                    elif hasattr(a, '__len__'):
                        a = float(a[0]) if len(a) > 0 else 0.0
                    else:
                        a = 0.0
                    logger.debug(f"evaluate_expression - 转换后a={a} (类型: {type(a)})")
                
                if hasattr(b, 'shape'):
                    logger.debug(f"evaluate_expression - 操作数b是DataFrame/Series类型，需要转换为标量")
                    if hasattr(b, 'iloc'):
                        b = b.iloc[0, 0] if b.shape[0] > 0 and b.shape[1] > 0 else 0.0
                    elif hasattr(b, 'item'):
                        b = b.item()
                    elif hasattr(b, '__len__'):
                        b = float(b[0]) if len(b) > 0 else 0.0
                    else:
                        b = 0.0
                    logger.debug(f"evaluate_expression - 转换后b={b} (类型: {type(b)})")
                
                # 执行运算
                op_func = self.operators[token][1]
                logger.debug(f"evaluate_expression - 执行运算: {a} {token} {b}")
                result = op_func(a, b)
                logger.debug(f"evaluate_expression - 运算结果: {result} (类型: {type(result)})")
                
                # 确保结果是标量值
                if hasattr(result, 'shape'):
                    logger.debug(f"evaluate_expression - 运算结果是DataFrame/Series类型，需要转换为标量")
                    if hasattr(result, 'iloc'):
                        result = result.iloc[0, 0] if result.shape[0] > 0 and result.shape[1] > 0 else 0.0
                    elif hasattr(result, 'item'):
                        result = result.item()
                    elif hasattr(result, '__len__'):
                        result = float(result[0]) if len(result) > 0 else 0.0
                    else:
                        result = 0.0
                    logger.debug(f"evaluate_expression - 转换后结果: {result} (类型: {type(result)})")
                
                stack.append(result)
            elif isinstance(token, str):
                # 单元格引用：获取值
                logger.debug(f"evaluate_expression - 单元格引用标记: {token}")
                
                if token.startswith('FILE1:'):
                    # FILE1前缀，使用df1
                    cell_ref = token[6:]
                    cell_value = self.get_cell_value(cell_ref, df1)
                    logger.debug(f"evaluate_expression - FILE1单元格引用: {cell_ref} = {cell_value} (类型: {type(cell_value)})")
                elif token.startswith('FILE2:'):
                    # FILE2前缀，使用df2
                    if df2 is None:
                        logger.error(f"evaluate_expression - 需要df2参数来处理FILE2:前缀的单元格引用")
                        raise ValueError("需要df2参数来处理FILE2:前缀的单元格引用")
                    cell_ref = token[6:]
                    cell_value = self.get_cell_value(cell_ref, df2)
                    logger.debug(f"evaluate_expression - FILE2单元格引用: {cell_ref} = {cell_value} (类型: {type(cell_value)})")
                else:
                    # 默认使用df1
                    cell_value = self.get_cell_value(token, df1)
                    logger.debug(f"evaluate_expression - 默认单元格引用: {token} = {cell_value} (类型: {type(cell_value)})")
                
                stack.append(cell_value)
            else:
                logger.error(f"evaluate_expression - 无效的标记: {token} (类型: {type(token)})")
                raise ValueError(f"无效的标记：{token}")
            
            logger.debug(f"evaluate_expression - 当前栈状态: {stack}")
        
        if len(stack) != 1:
            logger.error(f"evaluate_expression - 表达式求值完成后栈中应有1个元素，但有{len(stack)}个: {stack}")
            raise ValueError("无效的表达式")
        
        # 确保最终结果是标量值
        final_result = stack[0]
        logger.debug(f"evaluate_expression - 求值结果: {final_result} (类型: {type(final_result)})")
        
        if hasattr(final_result, 'shape'):
            logger.debug(f"evaluate_expression - 最终结果是DataFrame/Series类型，需要转换为标量")
            if hasattr(final_result, 'iloc'):
                final_result = final_result.iloc[0, 0] if final_result.shape[0] > 0 and final_result.shape[1] > 0 else 0.0
            elif hasattr(final_result, 'item'):
                final_result = final_result.item()
            elif hasattr(final_result, '__len__'):
                final_result = float(final_result[0]) if len(final_result) > 0 else 0.0
            else:
                final_result = 0.0
            logger.debug(f"evaluate_expression - 转换后最终结果: {final_result} (类型: {type(final_result)})")
        
        return final_result
    
    def get_cell_value(self, cell_ref, df):
        """
        从数据帧中获取单元格值或列数据
        
        参数:
            cell_ref: 单元格引用（如 "A1"）或列引用（如 "A"）
            df: 数据帧
            
        返回:
            float 或 pd.Series: 单元格的值（标量）或列数据（Series）
        """
        # 解析单元格引用（如A1）
        cell_match = re.match(r'^([A-Za-z]+)(\d+)$', cell_ref)
        if cell_match:
            # 单个单元格引用
            col_letters = cell_match.group(1)
            row_str = cell_match.group(2)
            
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
            
            logger.debug(f"获取单元格值 - 引用: {cell_ref}, 行: {row_idx}, 列: {col_idx}, 原始值: {value}, 类型: {type(value)}")
            
            # 确保值是标量
            if hasattr(value, 'shape'):
                # 如果是DataFrame或Series，提取第一个元素
                logger.warning(f"单元格值不是标量: {value}, 类型: {type(value)}, 形状: {getattr(value, 'shape', '未知')}")
                if hasattr(value, 'iloc'):
                    # DataFrame
                    if value.shape[0] > 0 and value.shape[1] > 0:
                        value = value.iloc[0, 0]
                    else:
                        value = 0.0
                elif hasattr(value, 'item'):
                    # Series或numpy数组
                    try:
                        value = value.item()
                    except (ValueError, TypeError):
                        value = 0.0
                elif hasattr(value, '__len__') and len(value) > 0:
                    # 其他可迭代对象
                    value = value[0]
                else:
                    value = 0.0
            
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
        
        # 解析列引用（如A）
        col_match = re.match(r'^([A-Za-z]+)$', cell_ref)
        if col_match:
            # 整列引用
            col_letters = col_match.group(1)
            
            # 转换列字母为索引（A=0, B=1, ..., AA=26）
            col_idx = 0
            for ch in col_letters.upper():
                col_idx = col_idx * 26 + (ord(ch) - ord('A') + 1)
            col_idx -= 1
            
            # 检查范围
            if col_idx < 0 or col_idx >= df.shape[1]:
                raise ValueError(f"列索引超出范围：{col_letters}")
            
            # 获取整列数据
            col_data = df.iloc[:, col_idx]
            
            logger.debug(f"获取列数据 - 引用: {cell_ref}, 列: {col_idx}, 数据类型: {type(col_data)}")
            
            # 转换为数值类型，无法转换的设为NaN
            col_data = pd.to_numeric(col_data, errors='coerce')
            
            # 填充NaN为0
            col_data = col_data.fillna(0)
            
            return col_data
        
        # 既不是单元格引用也不是列引用
        raise ValueError(f"无效的引用格式：{cell_ref}")
    
    def validate_rule(self, rule, df1, df2=None):
        """
        验证单条规则，支持单DataFrame或双DataFrame比较
        支持单元格引用（如A1）和列引用（如A）

        参数:
            rule: 规则字符串，如 "A1 + B1 = C1" 或 "FILE1:A = FILE2:A"
            df1: 第一个数据帧（默认数据帧）
            df2: 第二个数据帧（可选，用于跨文件比较）

        返回:
            tuple: (is_valid, failed_cells, passed_cells)
                - is_valid: 布尔值，表示规则是否通过验证
                - failed_cells: 失败的单元格列表（如[(row1, col1), (row2, col2)]）
                - passed_cells: 通过的单元格列表
        """
        logger.info(f"验证规则: {rule}")
        try:
            # 确保df1和df2是DataFrame类型
            if not hasattr(df1, 'iloc'):
                logger.error(f"参数df1不是DataFrame类型: {type(df1)}")
                return False, [], []
            if df2 is not None and not hasattr(df2, 'iloc'):
                logger.error(f"参数df2不是DataFrame类型: {type(df2)}")
                return False, [], []
                
            left_expr, op, right_expr = self.parse_rule(rule)
            logger.info(f"解析后的规则组件: 左表达式={left_expr}, 操作符={op}, 右表达式={right_expr}")
            
            left_value = self.evaluate_expression(left_expr, df1, df2)
            right_value = self.evaluate_expression(right_expr, df1, df2)
            
            logger.info(f"表达式求值结果类型 - 左值类型: {type(left_value)}, 右值类型: {type(right_value)}")
            logger.info(f"表达式求值结果 - 左: {left_expr} = {left_value}, 右: {right_expr} = {right_value}")
            
            # 检查是否是列规则（左或右值是Series类型）
            is_column_rule = isinstance(left_value, pd.Series) or isinstance(right_value, pd.Series)
            
            failed_cells = []
            passed_cells = []
            
            if is_column_rule:
                # 处理列规则
                logger.info(f"处理列规则: {rule}")
                
                # 确保两边都是Series类型
                if not isinstance(left_value, pd.Series):
                    # 将标量转换为与右值相同长度的Series
                    left_value = pd.Series([left_value] * len(right_value), index=right_value.index)
                elif not isinstance(right_value, pd.Series):
                    # 将标量转换为与左值相同长度的Series
                    right_value = pd.Series([right_value] * len(left_value), index=left_value.index)
                
                # 确保两个Series长度相同
                if len(left_value) != len(right_value):
                    logger.error(f"列长度不匹配: 左={len(left_value)}, 右={len(right_value)}")
                    return False, [], []
                
                # 执行逐行比较
                for i in range(len(left_value)):
                    lv = left_value.iloc[i]
                    rv = right_value.iloc[i]
                    
                    try:
                        # 执行比较
                        if op == '=':
                            row_result = abs(float(lv) - float(rv)) < 1e-6  # 考虑浮点误差
                        elif op == '!=':
                            row_result = abs(float(lv) - float(rv)) >= 1e-6
                        elif op == '<':
                            row_result = float(lv) < float(rv)
                        elif op == '<=':
                            row_result = float(lv) <= float(rv)
                        elif op == '>':
                            row_result = float(lv) > float(rv)
                        elif op == '>=':
                            row_result = float(lv) >= float(rv)
                        else:
                            logger.error(f"未知的比较操作符: {op}")
                            row_result = False
                        
                        if row_result:
                            passed_cells.append((i, None))  # None表示列引用，没有具体列索引
                        else:
                            failed_cells.append((i, None))  # None表示列引用，没有具体列索引
                    except Exception as e:
                        logger.error(f"行比较错误 (行{i+1}): {e}")
                        failed_cells.append((i, None))
                
                # 检查是否所有行都通过
                all_passed = len(failed_cells) == 0
                logger.info(f"列规则验证结果: {rule} -> {all_passed}, 失败行数: {len(failed_cells)}, 通过行数: {len(passed_cells)}")
                return all_passed, failed_cells, passed_cells
            else:
                # 处理单元格规则（保持原有逻辑）
                logger.info(f"处理单元格规则: {rule}")
                
                # 确保操作数是标量值
                def ensure_scalar(value):
                    """确保值是标量"""
                    logger.debug(f"ensure_scalar输入: {value}, 类型: {type(value)}")
                    
                    # 已经是标量值
                    if isinstance(value, (bool, int, float)):
                        logger.debug(f"已经是标量值: {value}")
                        return value
                    
                    # 检查是否是DataFrame或Series
                    elif hasattr(value, 'shape'):
                        logger.warning(f"发现非标量值: {value}, 类型: {type(value)}, 形状: {getattr(value, 'shape', '未知')}")
                        
                        # DataFrame类型
                        if hasattr(value, 'iloc'):
                            logger.debug("处理DataFrame类型")
                            if value.shape[0] > 0 and value.shape[1] > 0:
                                scalar_value = value.iloc[0, 0]  # 提取首个单元格值
                                logger.debug(f"从DataFrame提取值: {scalar_value}, 类型: {type(scalar_value)}")
                                # 递归调用确保最终返回标量
                                return ensure_scalar(scalar_value)
                            else:
                                logger.debug("空DataFrame，返回0.0")
                                return 0.0
                        
                        # Series或numpy数组类型
                        elif hasattr(value, 'item'):
                            logger.debug("处理Series或numpy数组类型")
                            try:
                                scalar_value = value.item()
                                logger.debug(f"从Series/数组提取值: {scalar_value}, 类型: {type(scalar_value)}")
                                return ensure_scalar(scalar_value)
                            except (ValueError, TypeError):
                                logger.warning("无法使用item()提取值")
                                if hasattr(value, '__len__') and len(value) > 0:
                                    scalar_value = float(value[0])
                                    logger.debug(f"从Series/数组提取第一个元素: {scalar_value}")
                                    return ensure_scalar(scalar_value)
                                else:
                                    logger.debug("空Series/数组，返回0.0")
                                    return 0.0
                        
                        # 其他具有len()的类型
                        elif hasattr(value, '__len__') and len(value) > 0:
                            logger.debug("处理具有len()的类型")
                            try:
                                scalar_value = float(value[0])
                                logger.debug(f"提取第一个元素: {scalar_value}")
                                return ensure_scalar(scalar_value)
                            except (ValueError, TypeError):
                                logger.warning("无法转换为浮点数，返回0.0")
                                return 0.0
                        
                        # 其他情况
                        else:
                            logger.debug("其他情况，返回0.0")
                            return 0.0
                    
                    # 尝试转换为数值
                    else:
                        logger.debug("尝试转换为数值")
                        try:
                            scalar_value = float(value)
                            logger.debug(f"转换为浮点数: {scalar_value}")
                            return scalar_value
                        except (ValueError, TypeError):
                            logger.warning(f"无法转换为浮点数: {value}, 返回0.0")
                            return 0.0
                
                # 转换值为标量
                left_value = ensure_scalar(left_value)
                right_value = ensure_scalar(right_value)
                
                logger.info(f"转换后的值 - 左: {left_value} (类型: {type(left_value)}), 右: {right_value} (类型: {type(right_value)})")
                
                # 执行比较操作
                try:
                    # 再次确保值是标量，双重保险
                    left_scalar = ensure_scalar(left_value)
                    right_scalar = ensure_scalar(right_value)
                    
                    logger.info(f"比较前的最终标量值 - 左: {left_scalar} (类型: {type(left_scalar)}), 右: {right_scalar} (类型: {type(right_scalar)})")
                    
                    # 确保值是可比较的类型
                    if not isinstance(left_scalar, (int, float)) or not isinstance(right_scalar, (int, float)):
                        logger.error(f"比较值不是数值类型: 左={left_scalar} (类型: {type(left_scalar)}), 右={right_scalar} (类型: {type(right_scalar)})")
                        return False, [], []
                    
                    # 执行比较
                    if op == '=':
                        result = abs(float(left_scalar) - float(right_scalar)) < 1e-6  # 考虑浮点误差
                    elif op == '!=':
                        result = abs(float(left_scalar) - float(right_scalar)) >= 1e-6
                    elif op == '<':
                        result = float(left_scalar) < float(right_scalar)
                    elif op == '<=':
                        result = float(left_scalar) <= float(right_scalar)
                    elif op == '>':
                        result = float(left_scalar) > float(right_scalar)
                    elif op == '>=':
                        result = float(left_scalar) >= float(right_scalar)
                    else:
                        logger.error(f"未知的比较操作符: {op}")
                        return False, [], []
                    
                    logger.info(f"比较结果: {result}, 类型: {type(result)}")
                    
                    # 确保最终结果是标量布尔值
                    final_result = bool(result)
                    
                    # 解析单元格引用以获取行列信息
                    cell_match = re.search(r'([A-Za-z]+)(\d+)', rule)
                    if cell_match:
                        col_letters = cell_match.group(1)
                        row_str = cell_match.group(2)
                        
                        # 转换为索引
                        col_idx = 0
                        for ch in col_letters.upper():
                            col_idx = col_idx * 26 + (ord(ch) - ord('A') + 1)
                        col_idx -= 1
                        row_idx = int(row_str) - 1
                        
                        if final_result:
                            passed_cells.append((row_idx, col_idx))
                        else:
                            failed_cells.append((row_idx, col_idx))
                    
                    logger.info(f"规则验证结果: {rule} -> {final_result}")
                    return final_result, failed_cells, passed_cells
                except Exception as e:
                    # 如果比较操作出错，返回False并记录详细错误
                    logger.error(f"比较操作错误：{e}", exc_info=True)
                    logger.error(f"比较失败的详细信息 - 左值: {left_value} (类型: {type(left_value)}), 右值: {right_value} (类型: {type(right_value)}), 操作符: {op}")
                    return False, [], []
        except Exception as e:
            # 解析或计算错误时返回False
            logger.error(f"规则验证错误：{e}", exc_info=True)
            return False, [], []
    
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
