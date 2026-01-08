#!/usr/bin/env python3
"""
入口：启动 GUI 应用
"""
import sys
from PyQt5.QtWidgets import QApplication
from gui import ComparisonTool
debug = False
def main():
    """
    主函数，应用程序的入口点
    
    创建QApplication实例，初始化比较工具窗口并显示
    启动Qt事件循环，直到应用程序退出
    """
    app = QApplication(sys.argv)
    window = ComparisonTool()
    window.show()
    if debug:    
            window.setStyleSheet("""
            QWidget {
                border: 1px solid red;
            }
            QPushButton {
                border: 1px solid blue;
            }
        """)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
