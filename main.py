#!/usr/bin/env python3
"""
入口：启动 GUI 应用
"""
import sys
from PyQt5.QtWidgets import QApplication
from gui import ComparisonTool

def main():
    app = QApplication(sys.argv)
    window = ComparisonTool()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()