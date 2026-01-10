#!/usr/bin/env python3
"""
Excel数据对比服务 - 业务逻辑层

负责处理Excel数据对比的核心业务逻辑，包括：
1. 工作簿加载与管理
2. 规则管理与验证
3. 数据比较（直接比较和规则比较）
4. 结果处理与导出

该服务类设计为可以独立于GUI运行，实现前后端分离
"""
import pandas as pd
import logging
from core.comparator import ExcelComparator

logger = logging.getLogger(__name__)

class ComparisonService:
    """Excel数据对比服务类，封装所有核心业务逻辑"""
    
    def __init__(self):
        """
        初始化比较服务
        
        创建Excel比较器实例，初始化数据存储
        """
        self.comparator = ExcelComparator()
        self.rules = []  # 存储用户定义的比较规则，每个元素是一个字典 {'rule': '规则文本', 'comment': '备注'}
        
        # 数据存储
        self.file1_df = {}  # 文件1的所有工作表，格式：{sheet_name: DataFrame}
        self.file2_df = {}  # 文件2的所有工作表，格式：{sheet_name: DataFrame}
        self.result_df = None  # 比较结果的数据框
        self.result_map = None  # 结果状态映射 {(行,列): 'equal'/'diff'}
    
    def load_workbook(self, file_path, alias="file1"):
        """
        加载Excel工作簿
        
        参数:
            file_path: Excel文件路径
            alias: 工作簿别名，默认为"file1"
            
        返回:
            list: 工作簿中的工作表名称列表
        """
        logger.info(f"加载工作簿: {file_path}，别名为: {alias}")
        self.comparator.load_workbook(file_path, alias)
        return self.comparator.list_sheets(alias)
    
    def get_workbook_sheets(self, alias="file1"):
        """
        获取指定工作簿的所有工作表名称
        
        参数:
            alias: 工作簿别名
            
        返回:
            list: 工作表名称列表
        """
        return self.comparator.list_sheets(alias)
    
    def load_sheet_data(self, alias, sheet_name):
        """
        加载指定工作表的数据
        
        参数:
            alias: 工作簿别名
            sheet_name: 工作表名称
            
        返回:
            pandas.DataFrame: 工作表的数据框
        """
        logger.info(f"加载工作表数据: 工作簿={alias}，工作表={sheet_name}")
        df = self.comparator.get_sheet_dataframe(alias, sheet_name)
        
        # 保存到对应的文件数据框字典
        if alias == "file1":
            self.file1_df[sheet_name] = df
        elif alias == "file2":
            self.file2_df[sheet_name] = df
        
        return df
    
    def add_rule(self, rule_text, comment=""):
        """
        添加比较规则
        
        参数:
            rule_text: 规则文本，例如："A1 + B1 = C1" 或 "FILE1:A1 = FILE2:A1"
            comment: 规则备注
            
        返回:
            bool: 规则添加是否成功
        """
        logger.info(f"尝试添加规则: {rule_text}，备注: {comment}")
        try:
            # 验证规则格式
            self.comparator.rule_engine.parse_rule(rule_text)
            
            # 添加到规则列表
            self.rules.append({'rule': rule_text, 'comment': comment})
            logger.info(f"规则添加成功: {rule_text}")
            return True
        except Exception as e:
            logger.error(f"规则添加失败: {str(e)}")
            raise Exception(f"规则格式无效: {str(e)}") from e
    
    def clear_rules(self):
        """清除所有已添加的规则"""
        self.rules.clear()
    
    def get_rules(self):
        """
        获取所有已添加的规则
        
        返回:
            list: 规则文本列表
        """
        return self.rules
    
    def import_rules(self, file_path):
        """
        从文件导入规则
        支持格式：规则文本 # 备注
        
        参数:
            file_path: 规则文件路径
            
        返回:
            bool: 导入是否成功
        """
        logger.info(f"从文件导入规则: {file_path}")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip() and not line.startswith('#')]
            
            # 解析并添加规则
            imported_count = 0
            for line in lines:
                # 分割规则和备注
                if '#' in line:
                    parts = line.split('#', 1)
                    rule_text = parts[0].strip()
                    comment = parts[1].strip()
                else:
                    rule_text = line.strip()
                    comment = ""
                
                if rule_text:
                    self.add_rule(rule_text, comment)
                    imported_count += 1
            
            logger.info(f"成功导入{imported_count}条规则")
            return True
        except Exception as e:
            logger.error(f"导入规则失败: {str(e)}")
            raise Exception(f"导入规则失败: {str(e)}") from e
    
    def run_comparison(self, use_rules=True, options=None):
        """
        执行比较操作
        
        参数:
            use_rules: 是否使用用户定义的规则进行比较
            options: 比较选项字典（仅在直接比较时使用）
            
        返回:
            tuple: (result_text, result_df, result_map)
                - result_text: 比较结果的文本描述
                - result_df: 比较结果的数据框
                - result_map: 结果状态映射
        """
        logger.info("开始执行比较操作")
        logger.debug(f"run_comparison参数 - use_rules: {use_rules}, options: {options}")
        
        # 检查数据框是否可用
        if use_rules and self.rules:
            # 规则比较可以使用单表或双表
            if not self.file1_df:
                logger.warning("比较失败：未选择文件")
                raise Exception("请先选择至少一个文件进行比较")
            # 检查是否有任何工作表数据
            has_data = any(not df.empty for df in self.file1_df.values())
            if not has_data:
                logger.warning("比较失败：文件1没有可用数据")
                raise Exception("请先选择至少一个文件进行比较")
        else:
            # 直接比较需要两个文件
            if not self.file1_df or not self.file2_df:
                logger.warning("比较失败：未选择两个文件")
                raise Exception("请先选择两个文件进行比较")
            # 检查是否有任何工作表数据
            has_file1_data = any(not df.empty for df in self.file1_df.values())
            has_file2_data = any(not df.empty for df in self.file2_df.values())
            if not has_file1_data or not has_file2_data:
                logger.warning("比较失败：其中一个文件没有可用数据")
                raise Exception("请先选择两个文件进行比较")
        
        result_text = ""
        
        # 根据是否使用规则执行不同的比较
        logger.debug(f"use_rules类型: {type(use_rules)}, use_rules值: {use_rules}")
        logger.debug(f"self.rules类型: {type(self.rules)}, self.rules值: {self.rules}")
        logger.debug(f"self.rules长度: {len(self.rules) if hasattr(self.rules, '__len__') else '不可测'}")
        
        if use_rules and self.rules:
            logger.info(f"使用{len(self.rules)}条用户定义规则进行比较")
            # 清除比较器中已有的规则
            self.comparator.clear_rules()
            
            # 添加用户定义的规则
            for rule_dict in self.rules:
                self.comparator.add_rule(rule_dict['rule'])
            
            # 验证所有规则
            # 如果file2_df不存在、为空或没有可用数据，则使用单表比较（将None作为df2参数）
            if not self.file2_df:
                logger.info("使用单表比较模式进行规则验证")
                passed_rules, failed_rules, all_failed_cells, all_passed_cells = self.comparator.validate_with_dataframes(self.file1_df, None)
            else:
                # 检查file2_df是否有可用数据
                has_file2_data = any(not df.empty for df in self.file2_df.values())
                if not has_file2_data:
                    logger.info("文件2没有可用数据，使用单表比较模式进行规则验证")
                    passed_rules, failed_rules, all_failed_cells, all_passed_cells = self.comparator.validate_with_dataframes(self.file1_df, None)
                else:
                    logger.info("使用双表比较模式进行规则验证")
                    passed_rules, failed_rules, all_failed_cells, all_passed_cells = self.comparator.validate_with_dataframes(self.file1_df, self.file2_df)
            
            # 生成规则比较结果
            result_text = f"规则比较结果：\n"
            result_text += f"总规则数: {len(self.rules)}\n"
            result_text += f"通过规则数: {len(passed_rules)}\n"
            result_text += f"失败规则数: {len(failed_rules)}\n\n"
            
            if passed_rules:
                result_text += "通过的规则：\n"
                for rule in passed_rules:
                    result_text += f"  ✓ {rule}\n"
                result_text += "\n"
            
            if failed_rules:
                result_text += "失败的规则：\n"
                for rule in failed_rules:
                    result_text += f"  ✗ {rule}\n"
                result_text += "\n"
            
            # 为规则比较结果创建DataFrame，包含详细的行通过数据
            rules_data = []
            
            # 处理通过的规则
            for rule in passed_rules:
                # 获取该规则对应的通过单元格
                rule_passed_cells = [cell for cell in all_passed_cells if cell[0] == rule]
                
                # 收集通过的行号
                passed_rows = set()
                for cell in rule_passed_cells:
                    if len(cell) == 3:  # 格式：(rule, row_idx, col_idx)
                        # 转换为Excel行号：索引+2（+1用于从0开始到从1开始的转换，+1用于跳过header行）
                        passed_rows.add(cell[1] + 2)
                    elif len(cell) == 2:  # 格式：(row_idx, col_idx)
                        passed_rows.add(cell[0] + 2)
                
                # 生成通过行号的字符串表示
                if passed_rows:
                    passed_rows_str = ', '.join(map(str, sorted(passed_rows)))
                else:
                    passed_rows_str = '所有行'
                
                rules_data.append({
                    '规则': rule,
                    '状态': '通过',
                    '详细信息': f'通过行: {passed_rows_str}'
                })
            
            # 处理失败的规则
            for rule in failed_rules:
                # 获取该规则对应的失败单元格
                rule_failed_cells = [cell for cell in all_failed_cells if cell[0] == rule]
                
                # 收集失败的行号
                failed_rows = set()
                for cell in rule_failed_cells:
                    if len(cell) == 3:  # 格式：(rule, row_idx, col_idx)
                        # 转换为Excel行号：索引+2（+1用于从0开始到从1开始的转换，+1用于跳过header行）
                        failed_rows.add(cell[1] + 2)
                    elif len(cell) == 2:  # 格式：(row_idx, col_idx)
                        failed_rows.add(cell[0] + 2)
                
                # 生成失败行号的字符串表示
                if failed_rows:
                    failed_rows_str = ', '.join(map(str, sorted(failed_rows)))
                else:
                    failed_rows_str = '无'
                
                rules_data.append({
                    '规则': rule,
                    '状态': '失败',
                    '详细信息': f'失败行: {failed_rows_str}'
                })
            
            result_df = pd.DataFrame(rules_data)
            
            # 创建结果映射，用于GUI高亮显示
            result_map = {'failed_cells': all_failed_cells, 'passed_cells': all_passed_cells}
            return result_text, result_df, result_map
        else:
            logger.info("执行直接比较")
            # 默认比较选项
            options = options or {
                'tolerance': '0',
                'ignore_case': False
            }
            
            # 执行直接比较
            logger.debug(f"直接比较参数 - file1_df: {self.file1_df.shape}, file2_df: {self.file2_df.shape}")
            logger.debug(f"file1_df类型: {type(self.file1_df)}")
            logger.debug(f"file2_df类型: {type(self.file2_df)}")
            logger.debug(f"file1_df内容: {self.file1_df}")
            logger.debug(f"file2_df内容: {self.file2_df}")
            
            # 调用compare_direct
            try:
                compare_result = self.comparator.compare_direct(self.file1_df, self.file2_df, options)
                logger.debug(f"compare_direct返回类型: {type(compare_result)}")
                logger.debug(f"compare_direct返回值: {compare_result}")
                
                # 解包返回值
                if compare_result and len(compare_result) == 2:
                    self.result_df, self.result_map = compare_result
                    logger.debug(f"解包后 - result_df: {self.result_df is not None}, result_map: {self.result_map is not None}")
                    if self.result_df is not None:
                        logger.debug(f"result_df形状: {self.result_df.shape}, result_df类型: {type(self.result_df)}")
                    if self.result_map is not None:
                        logger.debug(f"result_map类型: {type(self.result_map)}, result_map长度: {len(self.result_map) if hasattr(self.result_map, '__len__') else '不可测'}")
                else:
                    logger.error(f"compare_direct返回值格式错误: {compare_result}")
                    self.result_df = None
                    self.result_map = None
            except Exception as e:
                logger.error(f"compare_direct调用失败: {str(e)}", exc_info=True)
                logger.error(f"file1_df类型: {type(self.file1_df)}, file1_df形状: {self.file1_df.shape if hasattr(self.file1_df, 'shape') else '未知'}")
                logger.error(f"file2_df类型: {type(self.file2_df)}, file2_df形状: {self.file2_df.shape if hasattr(self.file2_df, 'shape') else '未知'}")
                raise Exception(f"比较失败: {str(e)}") from e
            
            # 格式化直接比较结果
            result_text = self._format_direct_comparison_result()
            
            return result_text, self.result_df, self.result_map
    
    def _format_direct_comparison_result(self):
        """
        格式化直接比较结果

        返回:
            str: 格式化的比较结果文本
        """
        logger.debug(f"格式化直接比较结果 - result_df: {self.result_df is not None}, result_map: {self.result_map is not None}")
        if self.result_df is None or self.result_df.empty or not self.result_map:
            return "无比较结果"
        
        # 计算差异统计
        try:
            if not self.result_map:
                total_cells = 0
                equal_cells = 0
                diff_cells = 0
                diff_rate = 0.0
            else:
                total_cells = len(self.result_map)
                equal_cells = sum(1 for status in self.result_map.values() if status == 'equal')
                diff_cells = total_cells - equal_cells
                diff_rate = diff_cells / total_cells if total_cells > 0 else 0
        except Exception as e:
            logger.error(f"计算差异统计时出错: {str(e)}")
            raise
        
        # 收集差异位置
        diff_positions = []
        for (row, col), status in self.result_map.items():
            if status == 'diff':
                col_letter = self._col_index_to_letter(col)
                cell_pos = f"{col_letter}{row+1}"
                diff_positions.append(cell_pos)
        
        # 格式化结果
        result_text = f"比较结果统计:\n"
        result_text += f"总单元格数: {total_cells}\n"
        result_text += f"相同单元格数: {equal_cells}\n"
        result_text += f"差异单元格数: {diff_cells}\n"
        result_text += f"差异率: {diff_rate:.2%}\n"
        
        if diff_positions:
            result_text += f"\n差异位置:\n"
            result_text += ", ".join(diff_positions[:10])  # 只显示前10个差异位置
            if len(diff_positions) > 10:
                result_text += f"... 等{len(diff_positions)}处差异"
        else:
            result_text += f"\n所有单元格完全相同"
        
        return result_text
    
    def _col_index_to_letter(self, col_index):
        """
        将0-based列索引转换为Excel列字母，例如0 -> A, 25 -> Z, 26 -> AA
        
        参数:
            col_index: 列索引（0-based）
            
        返回:
            str: Excel列字母
        """
        if col_index < 0:
            return ""
        letters = ""
        while col_index >= 0:
            col_index, rem = divmod(col_index, 26)
            letters = chr(rem + ord('A')) + letters
            col_index = col_index - 1
        return letters
    
    def save_results(self, result_df, file_path):
        """
        保存比较结果到文件
        
        参数:
            result_df: 比较结果的数据框
            file_path: 保存路径
            
        返回:
            bool: 保存是否成功
        """
        if result_df is None or result_df.empty:
            logger.warning("没有可保存的结果")
            return False
        
        try:
            if file_path.endswith('.xlsx'):
                self.comparator.export_results(result_df, file_path, format='excel')
            elif file_path.endswith('.csv'):
                self.comparator.export_results(result_df, file_path, format='csv')
            logger.info(f"结果已保存到: {file_path}")
            return True
        except Exception as e:
            logger.error(f"保存结果失败: {str(e)}")
            raise Exception(f"保存结果失败: {str(e)}") from e
    
    def save_original_with_highlights(self, df, file_path, failed_cells=None, passed_cells=None):
        """
        将原始表格另存为并添加颜色标记
        
        参数:
            df: 原始数据框
            file_path: 保存路径
            failed_cells: 失败的单元格列表，格式为[(row1, col1), (row2, col2)]
            passed_cells: 通过的单元格列表，格式为[(row1, col1), (row2, col2)]
            
        返回:
            bool: 保存是否成功
        """
        if df is None or df.empty:
            logger.warning("没有可保存的原始表格")
            return False
        
        try:
            if not file_path.endswith('.xlsx'):
                logger.error("只能保存为Excel文件(.xlsx)格式")
                return False
            
            logger.info(f"将原始表格另存为并添加颜色标记: {file_path}")
            self.comparator.export_with_highlights(df, file_path, failed_cells, passed_cells)
            logger.info(f"带颜色标记的原始表格已保存到: {file_path}")
            return True
        except Exception as e:
            logger.error(f"保存带颜色标记的原始表格失败: {str(e)}")
            raise Exception(f"保存带颜色标记的原始表格失败: {str(e)}") from e
