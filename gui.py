"""
Excel数据对比工具 - GUI界面
根据设计要求重新设计的中文界面
"""

import logging
from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QTableView, QCheckBox, QMessageBox, QSplitter, QTextEdit, QRadioButton, QButtonGroup, QGroupBox, QComboBox, QScrollArea
import pandas as pd

from core.comparison_service import ComparisonService
from core.diff_highlighter import DiffHighlighter

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
    例如: 整列选择 -> 'A'
    多个连续索引 -> 'A1:C5'
    单个索引 -> 'B2'
    """
    if not indexes:
        return ""
    
    # 获取选择的行列信息
    rows = [idx.row() for idx in indexes]
    cols = [idx.column() for idx in indexes]
    
    # 如果所有索引都来自同一个单元格
    if len(set(zip(rows, cols))) == 1:
        return f"{列索引转字母(cols[0])}{rows[0]+1}"
    
    # 检查是否是整列选择（所有行都被选中）
    # 注意：这里的整列选择需要满足所有索引来自同一列
    if len(set(cols)) == 1:
        # 获取选择的行范围
        min_row = min(rows)
        max_row = max(rows)
        selected_rows_count = max_row - min_row + 1
        
        # 如果选择的所有索引都来自同一列，并且选择了至少3行连续数据，就认为是整列选择
        # 这适用于大多数用户点击列标题进行整列选择的情况
        if selected_rows_count >= 3:
            return 列索引转字母(cols[0])
    
    # 处理常规的单元格范围选择
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
        self.file1_table.setSelectionBehavior(self.file1_table.SelectItems)
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
        self.file2_table.setSelectionBehavior(self.file2_table.SelectItems)
        file2_layout.addWidget(self.file2_table)
        
        content_layout.addWidget(file2_panel, 1)
        
        # 比较控制按钮将移到结果显示区域
        
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
        
        # 备注输入控件
        self.comment_input = QLineEdit()
        self.comment_input.setPlaceholderText("输入规则备注（可选）")
        rules_layout.addWidget(self.comment_input)
        
        # 添加规则按钮
        self.add_rule_btn = QPushButton("添加规则")
        self.add_rule_btn.clicked.connect(self.add_rule)
        rules_layout.addWidget(self.add_rule_btn)
        
        # 导入导出规则按钮
        import_export_layout = QHBoxLayout()
        
        self.import_rule_btn = QPushButton("导入规则")
        self.import_rule_btn.clicked.connect(self.import_rule)
        import_export_layout.addWidget(self.import_rule_btn)
        
        self.export_rule_btn = QPushButton("导出规则")
        self.export_rule_btn.clicked.connect(self.export_rule)
        import_export_layout.addWidget(self.export_rule_btn)
        
        rules_layout.addLayout(import_export_layout)
        
        # 规则列表容器 - 使用QScrollArea和QVBoxLayout实现动态规则项和滚动条
        self.rules_scroll_area = QScrollArea()
        self.rules_scroll_area.setWidgetResizable(True)  # 允许内容大小自适应
        
        self.rules_container = QWidget()
        self.rules_layout = QVBoxLayout(self.rules_container)
        self.rules_layout.setSpacing(10)  # 设置规则项之间的间距
        self.rules_layout.setAlignment(Qt.AlignTop)  # 内容顶部对齐
        
        self.rules_scroll_area.setWidget(self.rules_container)  # 将容器放入滚动区域
        rules_layout.addWidget(self.rules_scroll_area)  # 添加滚动区域到主布局
        
        # 右下方：结果显示区域
        results_group_box = QGroupBox("比较结果")
        results_layout = QVBoxLayout(results_group_box)
        bottom_layout.addWidget(results_group_box, 1)
        
        # 比较控制按钮区域
        compare_controls_layout = QHBoxLayout()
        self.start_compare_btn = QPushButton("开始比较")
        self.start_compare_btn.clicked.connect(self.run_comparison)
        compare_controls_layout.addWidget(self.start_compare_btn)
        
        self.export_btn = QPushButton("导出结果")
        self.export_btn.clicked.connect(self.export_results)
        compare_controls_layout.addWidget(self.export_btn)
        
        self.save_original_btn = QPushButton("保存原始表格（带颜色标记）")
        self.save_original_btn.clicked.connect(self.save_original_with_highlights)
        compare_controls_layout.addWidget(self.save_original_btn)
        
        results_layout.addLayout(compare_controls_layout)
        
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
        将选择的单元格或列自动插入到规则输入框中
        """
        selected_indexes = selected.indexes()
        if not selected_indexes:
            return
            
        range_str = 选择索引转Excel范围(selected_indexes)
        if range_str:
            # 获取当前选择的工作表名称
            sheet_name = self.file1_sheet_input.currentText()
            current_text = self.rule_input.text()
            # 格式化为 FILE1:SheetName:range_str
            new_text = current_text + f" FILE1:{sheet_name}:{range_str}" if current_text else f"FILE1:{sheet_name}:{range_str}"
            self.rule_input.setText(new_text)
    
    def update_file2_selection(self, selected, deselected):
        """
        更新文件2的单元格选择
        将选择的单元格或列自动插入到规则输入框中
        """
        selected_indexes = selected.indexes()
        if not selected_indexes:
            return
            
        range_str = 选择索引转Excel范围(selected_indexes)
        if range_str:
            # 获取当前选择的工作表名称
            sheet_name = self.file2_sheet_input.currentText()
            current_text = self.rule_input.text()
            # 格式化为 FILE2:SheetName:range_str
            new_text = current_text + f" FILE2:{sheet_name}:{range_str}" if current_text else f"FILE2:{sheet_name}:{range_str}"
            self.rule_input.setText(new_text)
    
    def add_rule(self):
        """
        添加比较规则到规则列表
        """
        rule_text = self.rule_input.text().strip()
        comment = self.comment_input.text().strip()
        if not rule_text:
            QMessageBox.warning(self, "警告", "请输入有效的比较规则")
            return
            
        logger.info(f"尝试添加规则: {rule_text}，备注: {comment}")
        try:
            # 使用服务添加规则
            self.service.add_rule(rule_text, comment)
            
            # 更新规则列表显示
            self.update_rules_list()
            
            # 清空输入框
            self.rule_input.clear()
            self.comment_input.clear()
            
            logger.info(f"规则添加成功: {rule_text}")
            self.statusBar().showMessage(f"规则添加成功: {rule_text}")
        except Exception as e:
            logger.error(f"规则添加失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"规则格式无效: {str(e)}")
            self.statusBar().showMessage("规则添加失败")
    
    def update_rules_list(self):
        """
        更新规则列表显示，动态生成每个规则项的控件
        """
        rules = self.service.get_rules()
        
        # 清除现有规则项控件
        while self.rules_layout.count() > 0:
            item = self.rules_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        # 如果没有规则，添加提示标签
        if not rules:
            no_rules_label = QLabel("无规则")
            no_rules_label.setAlignment(Qt.AlignCenter)
            no_rules_label.setStyleSheet("color: gray;")
            self.rules_layout.addWidget(no_rules_label)
            return
            
        # 动态生成每个规则项
        for i, rule_dict in enumerate(rules, 1):
            rule_text = rule_dict['rule']
            comment = rule_dict['comment']
            
            # 创建规则项容器
            rule_item = QWidget()
            rule_item_layout = QHBoxLayout(rule_item)
            rule_item_layout.setContentsMargins(5, 5, 5, 5)  # 设置内边距
            rule_item.setStyleSheet("border: 1px solid #ddd; border-radius: 5px;")
            
            # 规则序号和文本
            rule_label = QLabel(f"{i}. {rule_text}")
            rule_label.setStyleSheet("font-weight: bold;")
            rule_item_layout.addWidget(rule_label, 3)  # 设置伸缩因子
            
            # 备注输入框
            comment_input = QLineEdit()
            comment_input.setPlaceholderText("规则备注")
            if comment:
                comment_input.setText(comment)
            # 连接文本变化信号，用于更新备注
            comment_input.textChanged.connect(lambda text, idx=i-1: self.update_rule_comment(idx, text))
            rule_item_layout.addWidget(comment_input, 2)
            
            # 删除按钮
            delete_btn = QPushButton("删除")
            delete_btn.setStyleSheet("background-color: #ff6b6b; color: white;")
            delete_btn.setFixedWidth(60)
            # 连接删除按钮点击信号，使用lambda传递索引
            delete_btn.clicked.connect(lambda checked, idx=i-1: self.remove_rule(idx))
            rule_item_layout.addWidget(delete_btn)
            
            # 添加规则项到布局
            self.rules_layout.addWidget(rule_item)
    
    def remove_rule(self, index):
        """
        删除指定索引的规则
        参数:
            index: 规则索引
        """
        rules = self.service.get_rules()
        if 0 <= index < len(rules):
            rule_text = rules[index]['rule']
            self.service.rules.pop(index)  # 从服务中删除规则
            self.update_rules_list()  # 更新UI显示
            logger.info(f"删除规则成功: {rule_text}")
            self.statusBar().showMessage(f"已删除规则: {rule_text}")
    
    def update_rule_comment(self, index, comment):
        """
        更新指定索引规则的备注
        参数:
            index: 规则索引
            comment: 新的备注文本
        """
        rules = self.service.get_rules()
        if 0 <= index < len(rules):
            rules[index]['comment'] = comment  # 更新备注
            logger.debug(f"更新规则备注: 索引={index}, 新备注={comment}")
    
    def import_rule(self):
        """
        从文件导入规则
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "导入规则文件", "", "文本文件 (*.txt);;所有文件 (*.*)")
        if not file_path:
            return
            
        try:
            self.service.import_rules(file_path)
            self.update_rules_list()
            self.statusBar().showMessage(f"规则导入成功: {file_path}")
        except Exception as e:
            logger.error(f"规则导入失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"规则导入失败: {str(e)}")
    
    def export_rule(self):
        """
        将当前规则列表导出到文本文件
        """
        rules = self.service.get_rules()
        if not rules:
            QMessageBox.warning(self, "警告", "没有可导出的规则")
            return
            
        # 打开文件保存对话框
        file_path, _ = QFileDialog.getSaveFileName(self, "导出规则文件", "rules.txt", "文本文件 (*.txt);;所有文件 (*.*)")
        if not file_path:
            return
            
        try:
            # 写入规则到文件
            with open(file_path, 'w', encoding='utf-8') as f:
                # 写入文件头注释
                f.write("# Excel数据对比工具 - 规则文件\n")
                f.write("# 格式：规则文本 # 备注（可选）\n\n")
                
                # 写入每条规则
                for rule_dict in rules:
                    rule_text = rule_dict['rule']
                    comment = rule_dict['comment']
                    if comment:
                        f.write(f"{rule_text}  # {comment}\n")
                    else:
                        f.write(f"{rule_text}\n")
            
            logger.info(f"规则导出成功: {file_path}")
            self.statusBar().showMessage(f"规则导出成功: {file_path}")
            QMessageBox.information(self, "成功", f"规则已导出到: {file_path}")
        except Exception as e:
            logger.error(f"规则导出失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"规则导出失败: {str(e)}")
    

    
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
        if self.result_df is None or self.result_df.empty:
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
    
    def save_original_with_highlights(self):
        """
        保存带有颜色标记的原始表格
        
        允许用户选择保存文件1或文件2的原始表格，并添加颜色标记
        支持单表操作
        """
        # 检查是否有至少一个文件和比较结果
        has_file1 = self.file1_df is not None and not self.file1_df.empty
        has_file2 = self.file2_df is not None and not self.file2_df.empty
        
        if not has_file1 and not has_file2:
            QMessageBox.warning(self, "警告", "请先加载至少一个文件进行比较")
            return
        
        if self.result_map is None:
            QMessageBox.warning(self, "警告", "请先执行比较操作")
            return
        
        # 询问用户要保存哪个文件的原始表格
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Question)
        msg_box.setWindowTitle("选择保存的表格")
        
        # 根据加载的文件显示不同的选项
        if has_file1 and has_file2:
            msg_box.setText("请选择要保存的原始表格：")
            msg_box.addButton("文件1", QMessageBox.ActionRole)
            msg_box.addButton("文件2", QMessageBox.ActionRole)
            msg_box.addButton(QMessageBox.Cancel)
        elif has_file1:
            msg_box.setText("将保存文件1的原始表格")
            msg_box.addButton("确定", QMessageBox.ActionRole)
            msg_box.addButton(QMessageBox.Cancel)
        else:  # has_file2
            msg_box.setText("将保存文件2的原始表格")
            msg_box.addButton("确定", QMessageBox.ActionRole)
            msg_box.addButton(QMessageBox.Cancel)
        
        choice = msg_box.exec_()
        if choice == 1:  # Cancel
            return
        
        # 根据用户选择获取对应的DataFrame
        if not has_file2 or (has_file1 and (choice == 0 or not has_file2)):
            df_to_save = self.file1_df
            file_name = "file1"
        else:
            df_to_save = self.file2_df
            file_name = "file2"
        
        # 显示文件保存对话框
        file_path, _ = QFileDialog.getSaveFileName(self, "保存原始表格（带颜色标记）", 
                                                   f"{file_name}_with_highlights.xlsx", 
                                                   "Excel Files (*.xlsx)")
        if not file_path:
            return
        
        try:
            # 处理result_map以获取failed_cells和passed_cells
            failed_cells = []
            passed_cells = []
            
            if isinstance(self.result_map, dict):
                if 'failed_cells' in self.result_map:
                    # 规则比较的情况
                    rule_failed = self.result_map.get('failed_cells', [])
                    rule_passed = self.result_map.get('passed_cells', [])
                    
                    # 转换格式：[(rule, row_idx, col_idx), ...] → [(row_idx, col_idx), ...]
                    for item in rule_failed:
                        if len(item) == 3:
                            failed_cells.append((item[1], item[2]))  # 只取row_idx和col_idx
                        elif len(item) == 2:
                            failed_cells.append(item)
                    
                    for item in rule_passed:
                        if len(item) == 3:
                            passed_cells.append((item[1], item[2]))  # 只取row_idx和col_idx
                        elif len(item) == 2:
                            passed_cells.append(item)
                else:
                    # 直接比较的情况
                    for (row, col), status in self.result_map.items():
                        if status == 'diff':
                            failed_cells.append((row, col))
                        elif status == 'equal':
                            passed_cells.append((row, col))
            
            # 调用服务保存带颜色标记的原始表格
            self.service.save_original_with_highlights(df_to_save, file_path, failed_cells, passed_cells)
            
            self.statusBar().showMessage(f"带颜色标记的原始表格已保存到: {file_path}")
            QMessageBox.information(self, "成功", f"带颜色标记的原始表格已保存到: {file_path}")
        except Exception as e:
            logger.error(f"保存带颜色标记的原始表格失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"保存失败: {str(e)}")
            self.statusBar().showMessage("保存失败")
    

