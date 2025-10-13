#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QAutoDoc - Word文档内容自动填充工具
专为Word文档设计的内容自动填充工具，支持格式保持和批量处理
"""

import sys
import os
from PyQt5.QtWidgets import QApplication
from gui.main_window import MainWindow

def main():
    """主程序入口"""
    app = QApplication(sys.argv)
    app.setApplicationName("QAutoDoc")
    app.setApplicationVersion("1.0.0")
    
    # 创建主窗口
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()