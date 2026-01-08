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

# 配置日志记录
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

class ComparisonService:
    """Excel数据对比服务类，封装所有核心业务逻辑"""
    
    def __init__(self):
        """
        初始化比较服务
        
        创建Excel比较器实例，初始化数据存储
        """
        self.comparator = ExcelComparator()
        self.rules = []  # 存储用户定义的比较规则
        
        # 数据存储
        self.file1_df = None  # 文件1的数据框
        self.file2_df = None  # 文件2的数据框
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
        
        # 保存到对应的文件数据框
        if alias == "file1":
            self.file1_df = df
        elif alias == "file2":
            self.file2_df = df
        
        return df
    
    def add_rule(self, rule_text):
        """
        添加比较规则
        
        参数:
            rule_text: 规则文本，例如："A1 + B1 = C1" 或 "FILE1:A1 = FILE2:A1"
            
        返回:
            bool: 规则添加是否成功
        """
        logger.info(f"尝试添加规则: {rule_text}")
        try:
            # 验证规则格式
            self.comparator.rule_engine.parse_rule(rule_text)
            
            # 添加到规则列表
            self.rules.append(rule_text)
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
        
        if not self.file1_df or not self.file2_df:
            logger.warning("比较失败：未选择两个文件")
            raise Exception("请先选择两个文件进行比较")
        
        result_text = ""
        
        # 根据是否使用规则执行不同的比较
        if use_rules and self.rules:
            logger.info(f"使用{len(self.rules)}条用户定义规则进行比较")
            # 清除比较器中已有的规则
            self.comparator.clear_rules()
            
            # 添加用户定义的规则
            for rule in self.rules:
                self.comparator.add_rule(rule)
            
            # 验证所有规则
            passed_rules, failed_rules = self.comparator.validate_with_dataframes(self.file1_df, self.file2_df)
            
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
            
            # 规则比较没有结果数据框和映射
            return result_text, None, None
        else:
            logger.info("执行直接比较")
            # 默认比较选项
            options = options or {
                'tolerance': '0',
                'ignore_case': False
            }
            
            # 执行直接比较
            self.result_df, self.result_map = self.comparator.compare_direct(self.file1_df, self.file2_df, options)
            
            # 格式化直接比较结果
            result_text = self._format_direct_comparison_result()
            
            return result_text, self.result_df, self.result_map
    
    def _format_direct_comparison_result(self):
        """
        格式化直接比较结果
        
        返回:
            str: 格式化的比较结果文本
        """
        if not self.result_df or not self.result_map:
            return "无比较结果"
        
        # 计算差异统计
        total_cells = len(self.result_map)
        equal_cells = sum(1 for status in self.result_map.values() if status == 'equal')
        diff_cells = total_cells - equal_cells
        diff_rate = diff_cells / total_cells if total_cells > 0 else 0
        
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
        if not result_df:
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
