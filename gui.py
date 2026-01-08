"""
Excel数据对比工具 - GUI界面
根据设计要求重新设计的中文界面
"""
import os
import html
import difflib
import logging
from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QTableView, QCheckBox, QMessageBox, QSplitter, QTextEdit, QRadioButton, QButtonGroup, QGroupBox, QComboBox
import pandas as pd

from core.comparison_service import ComparisonService
from core.diff_highlighter import DiffHighlighter
from core.string_comparator import StringComparator

# 配置日志记录
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

# ------------------ 辅助函数 ------------------
def 列索引转字母(col_index):
    """将0-based列索引转换为Excel列字母，例如0 -> A, 25 -> Z, 26 -> AA"""
    if col_index < 0:
        return ""
    letters = ""
    while col_index >= 0:
        col_index, rem = divmod(col_index, 26)
        letters = chr(rem + ord('A')) + letters
        col_index = col_index - 1
    return letters

def 选择索引转Excel范围(indexes):
    """
    将QModelIndex列表转换为Excel范围字符串
    例如: 多个连续索引 -> 'A1:C5'
    单个索引 -> 'B2'
    """
    if not indexes:
        return ""
    rows = [idx.row() for idx in indexes]
    cols = [idx.column() for idx in indexes]
    start_row = min(rows)
    end_row = max(rows)
    start_col = min(cols)
    end_col = max(cols)
    
    # Excel行号从1开始
    start_cell = f"{列索引转字母(start_col)}{start_row+1}"
    end_cell = f"{列索引转字母(end_col)}{end_row+1}"
    
    if start_cell == end_cell:
        return start_cell
    return f"{start_cell}:{end_cell}"

# ------------------ Qt数据模型 ------------------
class PandasDataModel(QAbstractTableModel):
    """将Pandas DataFrame转换为Qt TableView可用的数据模型"""
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df.fillna('')

    def update_data(self, new_df):
        """更新数据模型中的数据"""
        self.beginResetModel()
        self._df = new_df.fillna('')
        self.endResetModel()

    def rowCount(self, parent=None):
        """返回行数"""
        return len(self._df.index)

    def columnCount(self, parent=None):
        """返回列数"""
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        """返回指定索引和角色的数据"""
        if not index.isValid():
            return QVariant()
        value = str(self._df.iat[index.row(), index.column()])
        if role == Qt.DisplayRole:
            return value
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """返回表头数据"""
        if role != Qt.DisplayRole:
            return QVariant()
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])

class ResultDataModel(QAbstractTableModel):
    """比较结果数据模型，支持差异高亮"""
    def __init__(self, df=pd.DataFrame(), result_map=None, parent=None):
        super().__init__(parent)
        self._df = df.fillna('')
        self.result_map = result_map or {}
        # 定义高亮颜色
        self.color_map = {
            'equal': QBrush(QColor("#CCFFCC")),  # 相等 - 绿色
            'diff': QBrush(QColor("#FFCCCC")),   # 差异 - 红色
            'error': QBrush(QColor("#FFDAB9")),  # 错误 - 浅黄色
            'empty': QBrush(QColor("#FFFFFF"))   # 空值 - 白色
        }

    def update_data(self, new_df, new_result_map):
        """更新结果数据和映射"""
        self.beginResetModel()
        self._df = new_df.fillna('')
        self.result_map = new_result_map or {}
        self.endResetModel()

    def rowCount(self, parent=None):
        """返回行数"""
        return len(self._df.index)

    def columnCount(self, parent=None):
        """返回列数"""
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        """返回指定索引和角色的数据，支持背景色高亮"""
        if not index.isValid():
            return QVariant()
        value = str(self._df.iat[index.row(), index.column()])
        if role == Qt.DisplayRole:
            return value
        if role == Qt.BackgroundRole:
            status = self.result_map.get((index.row(), index.column()), 'empty')
            return self.color_map.get(status, self.color_map['empty'])
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """返回表头数据"""
        if role != Qt.DisplayRole:
            return QVariant()
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])

# ------------------ 差异显示面板 ------------------
class DiffDisplayPanel(QWidget):
    """差异显示面板，简化版只显示比较结果文本"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighter = DiffHighlighter()
        self.setup_ui()

    def setup_ui(self):
        """初始化差异显示面板界面"""
        layout = QVBoxLayout()
        
        # 差异显示区域 - 仅保留一个文本编辑区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Courier", 10))
        layout.addWidget(self.result_text)

        # 差异统计标签
        self.stats_label = QLabel()
        layout.addWidget(self.stats_label)
        self.setLayout(layout)
    
    def set_diff_content(self, content):
        """
        设置差异显示面板的内容
        参数:
            content: 要显示的文本内容
        """
        self.result_text.setPlainText(content)
        diff_count = content.count('差异') if '差异' in content else 0
        self.stats_label.setText(f"比较完成，{diff_count} 处差异")

# ------------------ 主应用程序 ------------------
class ComparisonTool(QMainWindow):
    """Excel数据对比工具主窗口"""
    def __init__(self):
        """
        初始化比较工具主窗口
        
        创建比较服务实例，初始化界面布局
        """
        super().__init__()
        self.service = ComparisonService()
        self.setup_ui()
        # 数据存储
        self.file1_df = None  # 文件1的数据框
        self.file2_df = None  # 文件2的数据框
        self.result_df = None  # 比较结果的数据框
        self.result_map = None  # 结果状态映射 {(行,列): 'equal'/'diff'}

    def setup_ui(self):
        """
        初始化主窗口界面
        
        设置窗口标题、尺寸，创建文件选择区域和结果显示区域
        以及状态栏
        """
        self.setWindowTitle("Excel数据对比工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 主控件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
          
        # 主内容区域 - 文件选择区域
        content_layout = QHBoxLayout()
        main_layout.addLayout(content_layout)
        
        # 文件1区域
        file1_panel = QGroupBox("文件1区域")
        file1_layout = QVBoxLayout(file1_panel)
        
        # 文件1控件
        file1_controls_layout = QHBoxLayout()
        self.file1_path_label = QLabel("未选择文件")
        file1_controls_layout.addWidget(self.file1_path_label)
        file1_browse_btn = QPushButton("浏览...")
        file1_browse_btn.clicked.connect(lambda: self.open_workbook(alias="file1"))
        file1_controls_layout.addWidget(file1_browse_btn)
        file1_layout.addLayout(file1_controls_layout)
        
        # 文件1工作表选择
        file1_sheet_layout = QHBoxLayout()
        file1_sheet_layout.addWidget(QLabel("工作表:"))
        self.file1_sheet_input = QComboBox()
        # 连接工作表选择变化信号
        self.file1_sheet_input.currentIndexChanged.connect(lambda: self.load_sheet_data("file1", self.file1_sheet_input.currentText()))
        file1_sheet_layout.addWidget(self.file1_sheet_input)
        file1_layout.addLayout(file1_sheet_layout)
        
        # 文件1表格视图
        self.file1_table = QTableView()
        self.file1_table.setSelectionMode(self.file1_table.ExtendedSelection)
        file1_layout.addWidget(self.file1_table)
        
        content_layout.addWidget(file1_panel, 1)
        
        # 文件2区域
        file2_panel = QGroupBox("文件2区域")
        file2_layout = QVBoxLayout(file2_panel)
        
        # 文件2控件
        file2_controls_layout = QHBoxLayout()
        self.file2_path_label = QLabel("未选择文件")
        file2_controls_layout.addWidget(self.file2_path_label)
        file2_browse_btn = QPushButton("浏览...")
        file2_browse_btn.clicked.connect(lambda: self.open_workbook(alias="file2"))
        file2_controls_layout.addWidget(file2_browse_btn)
        file2_layout.addLayout(file2_controls_layout)
        
        # 文件2工作表选择
        file2_sheet_layout = QHBoxLayout()
        file2_sheet_layout.addWidget(QLabel("工作表:"))
        self.file2_sheet_input = QComboBox()
        # 连接工作表选择变化信号
        self.file2_sheet_input.currentIndexChanged.connect(lambda: self.load_sheet_data("file2", self.file2_sheet_input.currentText()))
        file2_sheet_layout.addWidget(self.file2_sheet_input)
        file2_layout.addLayout(file2_sheet_layout)
        
        # 文件2表格视图
        self.file2_table = QTableView()
        self.file2_table.setSelectionMode(self.file2_table.ExtendedSelection)
        file2_layout.addWidget(self.file2_table)
        
        content_layout.addWidget(file2_panel, 1)
        
        # 比较控制区域
        compare_controls_layout = QHBoxLayout()
        self.start_compare_btn = QPushButton("开始比较")
        self.start_compare_btn.clicked.connect(self.run_comparison)
        compare_controls_layout.addWidget(self.start_compare_btn)
        
        self.export_btn = QPushButton("导出结果")
        self.export_btn.clicked.connect(self.export_results)
        compare_controls_layout.addWidget(self.export_btn)
        
        main_layout.addLayout(compare_controls_layout)
        
        # 下方区域分为左右两栏
        bottom_layout = QHBoxLayout()
        main_layout.addLayout(bottom_layout, 1)
        
        # 左下方：规则输入区域
        rules_group_box = QGroupBox("规则输入")
        rules_layout = QVBoxLayout(rules_group_box)
        bottom_layout.addWidget(rules_group_box, 1)
        
        # 规则输入控件
        self.rule_input = QLineEdit()
        self.rule_input.setPlaceholderText("输入比较规则，例如：A1 + B1 = C1 或 FILE1:A1 = FILE2:A1")
        rules_layout.addWidget(self.rule_input)
        
        # 添加规则按钮
        self.add_rule_btn = QPushButton("添加规则")
        self.add_rule_btn.clicked.connect(self.add_rule)
        rules_layout.addWidget(self.add_rule_btn)
        
        # 规则列表
        self.rules_list = QTextEdit()
        self.rules_list.setReadOnly(True)
        self.rules_list.setPlaceholderText("已添加的规则将显示在这里")
        rules_layout.addWidget(self.rules_list)
        
        # 右下方：结果显示区域
        results_group_box = QGroupBox("比较结果")
        results_layout = QVBoxLayout(results_group_box)
        bottom_layout.addWidget(results_group_box, 1)
        
        # 差异显示面板
        self.diff_panel = DiffDisplayPanel()
        results_layout.addWidget(self.diff_panel)
        
        # 状态栏
        self.statusBar()
        self.statusBar().showMessage("就绪")
        
    def open_workbook(self, alias="file1"):
        """
        打开Excel工作簿并通过比较服务加载
        参数:
            alias: 工作簿别名，默认为"file1"
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "打开Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        
        logger.info(f"打开工作簿: {file_path}，别名为: {alias}")
        try:
            # 使用服务加载工作簿
            sheets = self.service.load_workbook(file_path, alias)
            if alias == "file1":
                self.file1_path_label.setText(file_path)
                # 更新工作表下拉列表
                self.file1_sheet_input.clear()
                for sheet in sheets:
                    self.file1_sheet_input.addItem(sheet)
                # 默认加载第一个工作表
                if sheets:
                    self.file1_sheet_input.setCurrentText(sheets[0])
                    self.load_sheet_data(alias, sheets[0])
            else:  # file2
                self.file2_path_label.setText(file_path)
                # 更新工作表下拉列表
                self.file2_sheet_input.clear()
                for sheet in sheets:
                    self.file2_sheet_input.addItem(sheet)
                # 默认加载第一个工作表
                if sheets:
                    self.file2_sheet_input.setCurrentText(sheets[0])
                    self.load_sheet_data(alias, sheets[0])
            
            logger.info(f"已加载文件: {file_path}")
            self.statusBar().showMessage(f"已加载文件: {file_path}")
        except Exception as e:
            logger.error(f"加载文件失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"加载文件失败: {str(e)}")
            self.statusBar().showMessage("加载文件失败")
    
    def load_sheet_data(self, alias, sheet_name):
        """
        加载工作表数据到对应的表格视图
        参数:
            alias: 工作簿别名
            sheet_name: 工作表名称
        """
        logger.info(f"加载工作表数据: 工作簿={alias}，工作表={sheet_name}")
        try:
            df = self.service.load_sheet_data(alias, sheet_name)
            model = PandasDataModel(df)
            if alias == "file1":
                self.file1_table.setModel(model)
                # 连接选择信号
                self.file1_table.selectionModel().selectionChanged.connect(self.update_file1_selection)
                # 保存数据到GUI实例变量
                self.file1_df = df
            else:  # file2
                self.file2_table.setModel(model)
                # 连接选择信号
                self.file2_table.selectionModel().selectionChanged.connect(self.update_file2_selection)
                # 保存数据到GUI实例变量
                self.file2_df = df
            logger.info(f"工作表数据加载成功: 工作簿={alias}，工作表={sheet_name}，形状={df.shape}")
        except Exception as e:
            logger.error(f"加载工作表失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"加载工作表失败: {str(e)}")
    
    def update_file1_selection(self, selected, deselected):
        """
        更新文件1的单元格选择
        将选择的单个单元格自动插入到规则输入框中
        """
        selected_indexes = selected.indexes()
        if not selected_indexes:
            return
        
        # 只处理单个单元格选择
        if len(selected_indexes) > 1:
            return
            
        range_str = 选择索引转Excel范围(selected_indexes)
        if range_str:
            current_text = self.rule_input.text()
            new_text = current_text + f" FILE1:{range_str}" if current_text else f"FILE1:{range_str}"
            self.rule_input.setText(new_text)
    
    def update_file2_selection(self, selected, deselected):
        """
        更新文件2的单元格选择
        将选择的单个单元格自动插入到规则输入框中
        """
        selected_indexes = selected.indexes()
        if not selected_indexes:
            return
        
        # 只处理单个单元格选择
        if len(selected_indexes) > 1:
            return
            
        range_str = 选择索引转Excel范围(selected_indexes)
        if range_str:
            current_text = self.rule_input.text()
            new_text = current_text + f" FILE2:{range_str}" if current_text else f"FILE2:{range_str}"
            self.rule_input.setText(new_text)
    
    def add_rule(self):
        """
        添加比较规则到规则列表
        """
        rule_text = self.rule_input.text().strip()
        if not rule_text:
            QMessageBox.warning(self, "警告", "请输入有效的比较规则")
            return
            
        logger.info(f"尝试添加规则: {rule_text}")
        try:
            # 使用服务添加规则
            self.service.add_rule(rule_text)
            
            # 更新规则列表显示
            self.update_rules_list()
            
            # 清空输入框
            self.rule_input.clear()
            
            logger.info(f"规则添加成功: {rule_text}")
            self.statusBar().showMessage(f"规则添加成功: {rule_text}")
        except Exception as e:
            logger.error(f"规则添加失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"规则格式无效: {str(e)}")
            self.statusBar().showMessage("规则添加失败")
    
    def update_rules_list(self):
        """
        更新规则列表显示
        """
        rules = self.service.get_rules()
        if not rules:
            self.rules_list.setPlainText("无规则")
            return
            
        rules_text = "已添加规则:\n"
        for i, rule in enumerate(rules, 1):
            rules_text += f"{i}. {rule}\n"
            
        self.rules_list.setPlainText(rules_text)
    

    
    def run_comparison(self):
        """
        执行比较操作
        
        调用比较服务执行比较，将结果显示在diff_panel中，并支持在原表格中标红差异
        """
        logger.info("开始执行比较操作")
        try:
            self.statusBar().showMessage("比较中...")
            
            # 检查是否使用规则
            use_rules = len(self.service.get_rules()) > 0
            
            # 执行比较
            result_text, result_df, result_map = self.service.run_comparison(use_rules=use_rules)
            
            # 保存结果到实例变量
            self.result_df = result_df
            self.result_map = result_map
            
            # 在原表格中标红差异
            if result_map:
                self.highlight_differences_in_original_tables()
            
            # 显示结果
            self.diff_panel.set_diff_content(result_text)
            logger.info("比较结果显示完成")
            
            # 更新状态栏
            if "规则" in result_text:
                passed_count = result_text.count("✓")
                failed_count = result_text.count("✗")
                self.statusBar().showMessage(f"规则比较完成 - 通过{passed_count}条，失败{failed_count}条")
            elif "差异单元格数" in result_text:
                import re
                match = re.search(r"差异单元格数: (\d+)", result_text)
                diff_cells = int(match.group(1)) if match else 0
                self.statusBar().showMessage(f"比较完成 - 发现{diff_cells}处差异")
            
        except Exception as e:
            logger.error(f"比较失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"比较失败: {str(e)}")
            self.statusBar().showMessage("比较失败")
    
    def highlight_differences_in_original_tables(self):
        """
        在原表格中标红差异
        
        使用ResultDataModel替换原有的PandasDataModel，实现差异高亮显示
        """
        if not self.result_map:
            return
            
        # 为文件1表格创建结果模型，仅显示左侧数据但保持右侧表格的差异标记
        if self.file1_df is not None:
            # 确保file1_df和result_df的形状兼容
            rows = min(self.file1_df.shape[0], self.result_df.shape[0])
            cols = min(self.file1_df.shape[1], self.result_df.shape[1])
            
            # 创建与file1_df形状相同的差异映射
            file1_result_map = {}
            for r in range(rows):
                for c in range(cols):
                    file1_result_map[(r, c)] = self.result_map.get((r, c), 'empty')
            
            # 创建结果模型并应用到表格
            result_model = ResultDataModel(self.file1_df, file1_result_map)
            self.file1_table.setModel(result_model)
        
        # 为文件2表格创建结果模型，仅显示右侧数据但保持左侧表格的差异标记
        if self.file2_df is not None:
            # 确保file2_df和result_df的形状兼容
            rows = min(self.file2_df.shape[0], self.result_df.shape[0])
            cols = min(self.file2_df.shape[1], self.result_df.shape[1])
            
            # 创建与file2_df形状相同的差异映射
            file2_result_map = {}
            for r in range(rows):
                for c in range(cols):
                    file2_result_map[(r, c)] = self.result_map.get((r, c), 'empty')
            
            # 创建结果模型并应用到表格
            result_model = ResultDataModel(self.file2_df, file2_result_map)
            self.file2_table.setModel(result_model)
    
    def format_comparison_result(self):
        """
        将比较结果格式化为字符串
        """
        if not self.result_df or not self.result_map:
            return "无比较结果"
        
        # 计算统计信息
        total_cells = len(self.result_map)
        equal_cells = sum(1 for status in self.result_map.values() if status == 'equal')
        diff_cells = total_cells - equal_cells
        diff_rate = diff_cells / total_cells if total_cells > 0 else 0
        
        # 收集差异位置
        diff_positions = []
        for (row, col), status in self.result_map.items():
            if status == 'diff':
                col_letter = 列索引转字母(col)
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
    
    def save_results(self):
        """
        保存比较结果到文件
        """
        if not self.result_df:
            QMessageBox.warning(self, "警告", "没有可保存的结果")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(self, "保存结果", "", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if not file_path:
            return
        
        try:
            self.service.save_results(self.result_df, file_path)
            self.statusBar().showMessage(f"结果已保存到: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存结果失败: {str(e)}")
            self.statusBar().showMessage("保存结果失败")
    
    def export_results(self):
        """
        导出比较结果，与保存结果功能相同
        """
        self.save_results()
    

