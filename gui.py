"""
增强版 GUI（基于之前的 MVP）：
- 可视化矩形选择（从 QTableView 选区生成 Excel 范围）
- 并排/内联/统一差异面板：在结果表格点击单元格时展示
- 使用 core/diff_highlighter.py 提供高亮效果
"""
import os
import html
from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QFileDialog, QLabel, QLineEdit, QTableView, QSpinBox, QCheckBox, QMessageBox,
    QSplitter, QTextEdit, QRadioButton, QButtonGroup, QGroupBox
)
import pandas as pd

from core.comparator import ExcelComparator
from core.diff_highlighter import DiffHighlighter
from core.string_comparator import StringComparator

# ------------------ Helpers ------------------
def col_index_to_letters(n):
    """0 -> A, 25 -> Z, 26 -> AA"""
    if n < 0:
        return ""
    letters = ""
    while n >= 0:
        n, rem = divmod(n, 26)
        letters = chr(rem + ord('A')) + letters
        n = n - 1
    return letters

def selection_indexes_to_excel_range(indexes):
    """
    indexes: list of QModelIndex
    返回 Excel 范围字符串，例如 'A1:C5'
    如果只有一个单元格：'B2'
    """
    if not indexes:
        return ""
    rows = [idx.row() for idx in indexes]
    cols = [idx.column() for idx in indexes]
    r1, r2 = min(rows), max(rows)
    c1, c2 = min(cols), max(cols)
    # Excel rows are 1-based
    start = f"{col_index_to_letters(c1)}{r1+1}"
    end = f"{col_index_to_letters(c2)}{r2+1}"
    if start == end:
        return start
    return f"{start}:{end}"

# ------------------ Qt Models ------------------
class PandasModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df.fillna('')

    def update(self, df):
        self.beginResetModel()
        self._df = df.fillna('')
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return QVariant()
        value = str(self._df.iat[index.row(), index.column()])
        if role == Qt.DisplayRole:
            return value
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return QVariant()
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])

class ResultModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), result_map=None, parent=None):
        super().__init__(parent)
        self._df = df.fillna('')
        self.result_map = result_map or {}
        self.colors = {
            'equal': QBrush(QColor("#CCFFCC")),
            'diff': QBrush(QColor("#FFCCCC")),
            'error': QBrush(QColor("#FFDAB9")),
            'empty': QBrush(QColor("#FFFFFF"))
        }

    def update(self, df, result_map):
        self.beginResetModel()
        self._df = df.fillna('')
        self.result_map = result_map or {}
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self._df.index)

    def columnCount(self, parent=None):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return QVariant()
        value = str(self._df.iat[index.row(), index.column()])
        if role == Qt.DisplayRole:
            return value
        if role == Qt.BackgroundRole:
            status = self.result_map.get((index.row(), index.column()), 'empty')
            return self.colors.get(status, self.colors['empty'])
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return QVariant()
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        else:
            return str(self._df.index[section])

# ------------------ Diff Display Panel ------------------
class DiffDisplayPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighter = DiffHighlighter()
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        # mode toggles
        mode_box = QGroupBox("显示模式")
        mode_layout = QHBoxLayout()
        self.side_by_side_radio = QRadioButton("并排对比")
        self.inline_radio = QRadioButton("内联高亮")
        self.unified_radio = QRadioButton("统一差异")
        self.side_by_side_radio.setChecked(True)
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.side_by_side_radio)
        self.mode_group.addButton(self.inline_radio)
        self.mode_group.addButton(self.unified_radio)
        mode_layout.addWidget(self.side_by_side_radio)
        mode_layout.addWidget(self.inline_radio)
        mode_layout.addWidget(self.unified_radio)
        mode_box.setLayout(mode_layout)
        layout.addWidget(mode_box)

        # splitter with left/right/unified
        splitter = QSplitter(Qt.Horizontal)
        self.left_text = QTextEdit()
        self.left_text.setReadOnly(True)
        self.left_text.setFont(QFont("Courier", 10))
        self.right_text = QTextEdit()
        self.right_text.setReadOnly(True)
        self.right_text.setFont(QFont("Courier", 10))
        self.unified_text = QTextEdit()
        self.unified_text.setReadOnly(True)
        self.unified_text.setFont(QFont("Courier", 10))

        splitter.addWidget(self.left_text)
        splitter.addWidget(self.right_text)
        layout.addWidget(splitter)
        layout.addWidget(self.unified_text)
        self.unified_text.hide()

        self.stats_label = QLabel()
        layout.addWidget(self.stats_label)
        self.setLayout(layout)

        # signals
        self.side_by_side_radio.toggled.connect(self.update_display_mode)
        self.inline_radio.toggled.connect(self.update_display_mode)
        self.unified_radio.toggled.connect(self.update_display_mode)

    def update_display_mode(self):
        if self.side_by_side_radio.isChecked():
            self.left_text.show(); self.right_text.show(); self.unified_text.hide()
        elif self.unified_radio.isChecked():
            self.left_text.hide(); self.right_text.hide(); self.unified_text.show()
        else:
            self.left_text.show(); self.right_text.show(); self.unified_text.hide()

    def display_diff(self, left, right, diff_type='line'):
        """
        left/right: raw strings
        diff_type: 'line' or 'char'
        """
        s_left = "" if left is None else str(left)
        s_right = "" if right is None else str(right)

        if diff_type == 'line':
            lines1 = s_left.splitlines()
            lines2 = s_right.splitlines()
            # stats
            diff = list(difflib.unified_diff(lines1, lines2, lineterm=''))
            added = sum(1 for line in diff if line.startswith('+') and not line.startswith('+++'))
            removed = sum(1 for line in diff if line.startswith('-') and not line.startswith('---'))
            self.stats_label.setText(f"差异统计: +{added}行, -{removed}行")

            if self.side_by_side_radio.isChecked():
                html_sb = self.highlighter.side_by_side_html(lines1, lines2)
                # split into left/right by finding our columns (we return two columns in one html block)
                # For simplicity, put whole html into both panes but they contain both columns visually;
                # instead put left lines in left_text, right lines in right_text in plain format
                left_html = '<div style="font-family:monospace;">' + ''.join(f'<div>{html.escape(l)}</div>' for l in lines1) + '</div>'
                right_html = '<div style="font-family:monospace;">' + ''.join(f'<div>{html.escape(l)}</div>' for l in lines2) + '</div>'
                self.left_text.setHtml(left_html)
                self.right_text.setHtml(right_html)
            elif self.unified_radio.isChecked():
                uni = self.highlighter.unified_diff_html(lines1, lines2)
                self.unified_text.setHtml(uni)
            else:
                # inline: highlight differences using character-level highlighter per line
                left_html, right_html = self.highlighter.highlight_text_diff(s_left, s_right)
                self.left_text.setHtml(f'<div style="font-family:monospace;">{left_html}</div>')
                self.right_text.setHtml(f'<div style="font-family:monospace;">{right_html}</div>')
        else:
            # character mode - show inline highlighting
            left_html, right_html = self.highlighter.highlight_text_diff(s_left, s_right)
            self.left_text.setHtml(f'<div style="font-family:monospace;">{left_html}</div>')
            self.right_text.setHtml(f'<div style="font-family:monospace;">{right_html}</div>')
            self.stats_label.setText("字符级差异显示")