#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主窗口界面
"""

import os
from PyQt5.QtWidgets import (QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, 
                             QPushButton, QLabel, QLineEdit, QTextEdit, 
                             QFileDialog, QMessageBox, QProgressBar, QTabWidget,
                             QGroupBox, QFormLayout, QScrollArea, QSplitter,
                             QRadioButton, QTableWidget, QTableWidgetItem, QHeaderView,
                             QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from core.document_processor import DocumentProcessor
from core.field_extractor import FieldExtractor


class ProcessingThread(QThread):
    """处理线程，用于后台处理文档"""
    progress_updated = pyqtSignal(int)
    processing_finished = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, template_path, output_dir, field_values):
        super().__init__()
        self.template_path = template_path
        self.output_dir = output_dir
        self.field_values = field_values
        
    def run(self):
        try:
            processor = DocumentProcessor()
            result = processor.process_document(
                self.template_path, 
                self.output_dir, 
                self.field_values
            )
            self.processing_finished.emit(result)
        except Exception as e:
            self.error_occurred.emit(str(e))


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self.template_path = ""
        self.field_values = {}
        self.init_ui()
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("QAutoDoc - Word文档批量填充工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 应用现代化样式
        self.apply_modern_style()
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 创建标签页
        tab_widget = QTabWidget()
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                border: 1px solid #e0e0e0;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px 16px;
                margin-right: 2px;
                min-width: 80px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 1px solid white;
                margin-bottom: -1px;
            }
            QTabBar::tab:hover {
                background-color: #e8f4fd;
            }
        """)
        main_layout.addWidget(tab_widget)
        
        # 批量处理标签页（主界面）
        batch_tab = self.create_batch_processing_tab()
        tab_widget.addTab(batch_tab, "批量处理")
        
        # 关于标签页
        about_tab = self.create_about_tab()
        tab_widget.addTab(about_tab, "关于")
    
    def apply_modern_style(self):
        """应用现代化样式"""
        style = """
        QMainWindow {
            background-color: #f8f9fa;
            font-family: "Segoe UI", "Microsoft YaHei", sans-serif;
        }
        
        QWidget {
            font-size: 11pt;
        }
        
        QGroupBox {
            font-weight: bold;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 15px;
            background-color: white;
        }
        
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top center;
            padding: 0 8px;
            background-color: white;
        }
        
        QLineEdit, QComboBox {
            border: 1px solid #d0d0d0;
            border-radius: 4px;
            padding: 8px 12px;
            font-size: 10pt;
            background-color: white;
        }
        
        QLineEdit:focus, QComboBox:focus {
            border-color: #2196F3;
            outline: none;
        }
        
        QPushButton {
            background-color: #2196F3;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 8px 16px;
            font-weight: bold;
            font-size: 10pt;
            min-height: 32px;
        }
        
        QPushButton:hover {
            background-color: #1976D2;
        }
        
        QPushButton:pressed {
            background-color: #0D47A1;
        }
        
        QPushButton:disabled {
            background-color: #b0bec5;
            color: #78909c;
        }
        
        QProgressBar {
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            text-align: center;
            background-color: #f5f5f5;
        }
        
        QProgressBar::chunk {
            background-color: #4CAF50;
            border-radius: 2px;
        }
        
        QTextEdit {
            border: 1px solid #d0d0d0;
            border-radius: 4px;
            padding: 8px;
            background-color: white;
            font-family: "Consolas", "Courier New", monospace;
            font-size: 9pt;
        }
        
        QLabel {
            color: #424242;
        }
        
        QRadioButton {
            spacing: 8px;
        }
        
        QRadioButton::indicator {
            width: 16px;
            height: 16px;
            border-radius: 8px;
            border: 2px solid #b0b0b0;
        }
        
        QRadioButton::indicator:checked {
            background-color: #2196F3;
            border-color: #2196F3;
        }
        
        QTabWidget::pane {
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            background-color: white;
        }
        
        QTabBar::tab {
            background-color: #f5f5f5;
            border: 1px solid #e0e0e0;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 8px 16px;
            margin-right: 2px;
            min-width: 80px;
        }
        
        QTabBar::tab:selected {
            background-color: white;
            border-bottom: 1px solid white;
            margin-bottom: -1px;
        }
        
        QTabBar::tab:hover {
            background-color: #e8f4fd;
        }
        """
        self.setStyleSheet(style)
        

        
    def create_batch_processing_tab(self):
        """创建批量处理标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)
        
        # 模板选择区域
        template_group = QGroupBox("📄 模板文档")
        template_layout = QHBoxLayout(template_group)
        template_layout.setContentsMargins(12, 20, 12, 12)
        
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setPlaceholderText("请选择Word模板文档...")
        self.template_path_edit.setMinimumHeight(36)
        template_layout.addWidget(self.template_path_edit)
        
        browse_btn = QPushButton("📁 浏览")
        browse_btn.setMinimumHeight(36)
        browse_btn.clicked.connect(self.browse_template)
        template_layout.addWidget(browse_btn)
        
        layout.addWidget(template_group)
        
        # 数据输入区域
        data_group = QGroupBox("📊 批量数据")
        data_layout = QVBoxLayout(data_group)
        data_layout.setContentsMargins(12, 20, 12, 12)
        
        # 数据输入方式选择
        input_method_layout = QHBoxLayout()
        input_method_layout.addWidget(QLabel("数据输入方式："))
        
        self.csv_radio = QRadioButton("📁 CSV文件导入")
        self.manual_radio = QRadioButton("⌨️ 手动输入（每个字段一个标签页）")
        self.manual_radio.setChecked(True)  # 默认选择手动输入
        
        input_method_layout.addWidget(self.csv_radio)
        input_method_layout.addWidget(self.manual_radio)
        input_method_layout.addStretch()
        
        data_layout.addLayout(input_method_layout)
        
        # CSV文件导入区域
        self.csv_group = QGroupBox("CSV文件导入")
        csv_layout = QHBoxLayout(self.csv_group)
        csv_layout.setContentsMargins(12, 15, 12, 12)
        
        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("请选择CSV数据文件...")
        self.csv_path_edit.setMinimumHeight(36)
        csv_layout.addWidget(self.csv_path_edit)
        
        csv_browse_btn = QPushButton("📁 浏览")
        csv_browse_btn.setMinimumHeight(36)
        csv_browse_btn.clicked.connect(self.browse_csv)
        csv_layout.addWidget(csv_browse_btn)
        
        data_layout.addWidget(self.csv_group)
        self.csv_group.setVisible(False)  # 默认隐藏CSV导入
        
        # 手动输入区域 - 每个字段一个标签页（无边框）
        manual_layout = QVBoxLayout()
        manual_layout.setContentsMargins(0, 0, 0, 0)
        
        # 字段标签页控件
        self.field_tabs = QTabWidget()
        self.field_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #e0e0e0;
                border-radius: 6px;
            }
            QTabBar::tab {
                background-color: #f8f9fa;
                border: 1px solid #e0e0e0;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 6px 12px;
                margin-right: 1px;
                min-width: 60px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 1px solid white;
            }
        """)
        manual_layout.addWidget(self.field_tabs)
        
        data_layout.addLayout(manual_layout)
        
        # 连接单选按钮信号
        self.csv_radio.toggled.connect(self.toggle_input_method)
        self.manual_radio.toggled.connect(self.toggle_input_method)
        
        layout.addWidget(data_group)
        
        # 初始化手动输入表格
        self.init_manual_table()
        
        # 处理按钮区域
        buttons_layout = QHBoxLayout()
        
        process_btn = QPushButton("🚀 批量生成文档")
        process_btn.setMinimumHeight(42)
        process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 12pt;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        process_btn.clicked.connect(self.process_batch_documents)
        
        buttons_layout.addWidget(process_btn)
        
        # 文档命名字段下拉框（放在按钮右边）
        buttons_layout.addSpacing(10)
        buttons_layout.addWidget(QLabel("文档命名字段："))
        
        self.naming_field_combo = QComboBox()
        self.naming_field_combo.setMinimumHeight(36)
        self.naming_field_combo.addItem("自动选择（默认）")
        buttons_layout.addWidget(self.naming_field_combo)
        
        buttons_layout.addStretch()
        
        layout.addLayout(buttons_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(8)
        layout.addWidget(self.progress_bar)
        
        # 结果区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setVisible(False)
        self.result_text.setMinimumHeight(120)
        layout.addWidget(self.result_text)
        
        # 版权信息
        copyright_label = QLabel("© 2025 QAutoDoc - 让文档处理更高效")
        copyright_label.setStyleSheet("""
            QLabel {
                font-size: 9pt;
                color: #757575;
                text-align: center;
                margin-top: 10px;
                margin-bottom: 5px;
            }
        """)
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(copyright_label)
        
        return tab
        
    def create_about_tab(self):
        """创建关于标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # 标题区域
        title_label = QLabel("QAutoDoc - Word文档内容自动填充工具")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 24pt;
                font-weight: bold;
                color: #2196F3;
                text-align: center;
                margin-bottom: 20px;
            }
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 版本信息卡片
        version_group = QGroupBox("📋 基本信息")
        version_layout = QVBoxLayout(version_group)
        version_layout.setContentsMargins(15, 15, 15, 15)
        
        version_label = QLabel("<b>版本：</b>1.0.0")
        version_label.setStyleSheet("font-size: 11pt; margin: 5px;")
        version_layout.addWidget(version_label)
        
        layout.addWidget(version_group)
        
        # 功能特点卡片
        features_group = QGroupBox("✨ 功能特点")
        features_layout = QVBoxLayout(features_group)
        features_layout.setContentsMargins(15, 15, 15, 15)
        
        features_text = """
        <ul style="margin: 0; padding-left: 20px;">
        <li><b>格式保持：</b>在填充内容时完全保留原始Word文档的格式</li>
        <li><b>字段识别：</b>自动识别 {{字段名}} 格式的动态字段</li>
        <li><b>内容替换：</b>精确替换指定字段内容而不影响文档其他部分</li>
        <li><b>批量处理：</b>支持根据多组数据一次性生成多个文档</li>
        </ul>
        """
        features_label = QLabel(features_text)
        features_label.setStyleSheet("font-size: 11pt; line-height: 1.6;")
        features_label.setWordWrap(True)
        features_layout.addWidget(features_label)
        
        layout.addWidget(features_group)
        
        # 应用场景卡片
        scenarios_group = QGroupBox("🎯 应用场景")
        scenarios_layout = QVBoxLayout(scenarios_group)
        scenarios_layout.setContentsMargins(15, 15, 15, 15)
        
        scenarios_text = """
        <ul style="margin: 0; padding-left: 20px;">
        <li>批量生成合同、通知书、邀请函等</li>
        <li>自动化办公文档处理</li>
        <li>减少重复性文档编辑工作</li>
        <li>批量生成学生档案、成绩单等</li>
        <li>自动化报告生成</li>
        </ul>
        """
        scenarios_label = QLabel(scenarios_text)
        scenarios_label.setStyleSheet("font-size: 11pt; line-height: 1.6;")
        scenarios_label.setWordWrap(True)
        scenarios_layout.addWidget(scenarios_label)
        
        layout.addWidget(scenarios_group)
        
        # 添加弹性空间
        layout.addStretch()
        
        # 底部信息
        footer_label = QLabel("© 2025 QAutoDoc - 让文档处理更高效")
        footer_label.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #757575;
                text-align: center;
                margin-top: 20px;
            }
        """)
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(footer_label)
        
        return tab
        
    def browse_template(self):
        """浏览并选择模板文档"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Word模板文档", "", "Word文档 (*.docx)"
        )
        
        if file_path:
            self.template_path = file_path
            self.template_path_edit.setText(file_path)
            print(f"模板路径已设置：{self.template_path}")  # 调试信息
            fields = self.extract_fields()
            print(f"提取到的字段：{fields}")  # 调试信息
            

        
    def browse_csv(self):
        """浏览并选择CSV文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择CSV数据文件", "", "CSV文件 (*.csv);;所有文件 (*.*)"
        )
        
        if file_path:
            self.csv_path_edit.setText(file_path)
            
    def toggle_input_method(self):
        """切换数据输入方式"""
        if self.csv_radio.isChecked():
            self.csv_group.setVisible(True)
            self.field_tabs.setVisible(False)
        else:
            self.csv_group.setVisible(False)
            self.field_tabs.setVisible(True)
            # 当切换到手动输入时，重新提取字段以更新表格
            if self.template_path:
                self.extract_fields()
            
    def init_manual_table(self):
        """初始化手动输入表格"""
        # 清空现有标签页
        self.field_tabs.clear()
        self.field_text_edits = {}  # 存储每个字段的多行文本框
        

                    
    def extract_fields(self):
        """从模板文档中提取字段"""
        if not self.template_path:
            return []
            
        try:
            extractor = FieldExtractor()
            fields = extractor.extract_fields(self.template_path)
            
            # 更新命名字段下拉框
            self.naming_field_combo.clear()
            self.naming_field_combo.addItem("自动选择（默认）")
            for field in fields:
                self.naming_field_combo.addItem(field)
            
            # 更新手动输入界面
            if self.manual_radio.isChecked():
                self.create_field_tabs(fields)
                print(f"字段标签页创建完成，当前字段文本框数量：{len(self.field_text_edits)}")
                
            return fields
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"提取字段时出错：{str(e)}")
            return []
            
    def create_field_tabs(self, fields):
        """为每个字段创建标签页"""
        # 清空现有标签页
        self.field_tabs.clear()
        self.field_text_edits = {}
        
        print(f"开始创建字段标签页，字段数量：{len(fields)}")
        
        for field in fields:
            # 创建标签页
            tab = QWidget()
            layout = QVBoxLayout(tab)
            layout.setContentsMargins(10, 10, 10, 10)
            layout.setSpacing(8)
            
            # 添加说明标签
            info_label = QLabel(f"📝 在下方输入 <b>{field}</b> 的值，每行对应一个文档：")
            info_label.setStyleSheet("""
                QLabel {
                    font-size: 10pt;
                    color: #424242;
                    margin-bottom: 8px;
                }
            """)
            layout.addWidget(info_label)
            
            # 创建多行文本框
            text_edit = QTextEdit()
            text_edit.setPlaceholderText(f"请输入 {field} 的值，每行一个...")
            text_edit.setStyleSheet("""
                QTextEdit {
                    border: 1px solid #d0d0d0;
                    border-radius: 6px;
                    padding: 8px;
                    font-family: "Consolas", "Courier New", monospace;
                    font-size: 10pt;
                    background-color: white;
                }
                QTextEdit:focus {
                    border-color: #2196F3;
                }
            """)
            layout.addWidget(text_edit)
            
            # 存储文本框引用
            self.field_text_edits[field] = text_edit
            print(f"创建字段 '{field}' 的文本框，内存地址：{id(text_edit)}")
            
            # 添加到标签页
            self.field_tabs.addTab(tab, field)
            
        print(f"字段标签页创建完成，存储的文本框数量：{len(self.field_text_edits)}")
            
    def get_field_values_from_table(self):
        """从手动输入标签页获取字段值"""
        if not self.field_text_edits:
            print("没有字段文本框，返回空列表")
            return []
            
        print(f"字段文本框数量：{len(self.field_text_edits)}")
        print(f"当前存储的字段文本框：{list(self.field_text_edits.keys())}")
        
        # 检查当前活动的标签页
        current_tab_index = self.field_tabs.currentIndex()
        current_tab_text = self.field_tabs.tabText(current_tab_index)
        print(f"当前活动标签页索引：{current_tab_index}，标签页文本：'{current_tab_text}'")
        
        # 获取每个字段的文本行
        field_lines = {}
        max_lines = 0
        
        for field_name, text_edit in self.field_text_edits.items():
            # 检查文本框是否有效
            print(f"检查字段 '{field_name}' 的文本框，内存地址：{id(text_edit)}")
            
            text = text_edit.toPlainText().strip()
            print(f"字段 '{field_name}' 的文本内容：'{text}'")
            
            if text:
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                field_lines[field_name] = lines
                max_lines = max(max_lines, len(lines))
                print(f"字段 '{field_name}' 解析出的行：{lines}")
            else:
                field_lines[field_name] = []
                print(f"字段 '{field_name}' 为空")
                
        print(f"最大行数：{max_lines}")
        
        # 如果没有数据行，返回空列表
        if max_lines == 0:
            print("没有数据行，返回空列表")
            return []
            
        # 构建数据行
        data_rows = []
        for i in range(max_lines):
            row_data = {}
            for field_name, lines in field_lines.items():
                if i < len(lines):
                    row_data[field_name] = lines[i]
                else:
                    row_data[field_name] = ""  # 如果该字段行数不够，填充空值
            
            print(f"第 {i+1} 行数据：{row_data}")
            
            # 只有当行中有实际数据时才添加
            if any(row_data.values()):
                data_rows.append(row_data)
                print(f"第 {i+1} 行有数据，已添加到结果")
            else:
                print(f"第 {i+1} 行没有数据，跳过")
            
        print(f"最终数据行数：{len(data_rows)}")
        return data_rows
        
    def get_field_values_from_csv(self):
        """从CSV文件中获取字段值"""
        csv_path = self.csv_path_edit.text()
        if not csv_path:
            return []
            
        try:
            import csv
            field_values_list = []
            
            with open(csv_path, 'r', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    # 过滤空值行
                    if any(row.values()):
                        field_values_list.append(row)
                        
            return field_values_list
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取CSV文件时出错：{str(e)}")
            return []
            
    def get_naming_field(self):
        """获取用户选择的命名字段"""
        current_text = self.naming_field_combo.currentText()
        if current_text == "自动选择（默认）":
            return None
        return current_text
        
    def process_batch_documents(self):
        """批量处理文档"""
        print("开始批量处理文档...")
        
        if not self.template_path:
            print("模板路径为空")
            QMessageBox.warning(self, "警告", "请先选择模板文档")
            return
            
        print(f"模板路径：{self.template_path}")
        
        # 检查是否已经有字段文本框
        if not self.field_text_edits:
            print("没有字段文本框，需要重新提取字段")
            # 提取字段
            fields = self.extract_fields()
            if not fields:
                print("未提取到字段")
                QMessageBox.warning(self, "警告", "未找到可替换的字段")
                return
            print(f"提取到的字段：{fields}")
        else:
            print("使用现有的字段文本框")
            fields = list(self.field_text_edits.keys())
            print(f"现有字段：{fields}")
        
        # 获取数据
        if self.csv_radio.isChecked():
            print("使用CSV模式")
            field_values_list = self.get_field_values_from_csv()
        else:
            print("使用手动输入模式")
            field_values_list = self.get_field_values_from_table()
            
        print(f"获取到的数据行数：{len(field_values_list)}")
        
        if not field_values_list:
            print("没有可用的数据")
            QMessageBox.warning(self, "警告", "没有可用的数据")
            return
            
        # 选择输出目录
        output_dir = QFileDialog.getExistingDirectory(
            self, "选择输出目录"
        )
        
        if not output_dir:
            return
            
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(field_values_list))
        self.progress_bar.setValue(0)
        
        # 获取命名字段
        naming_field = self.get_naming_field()
        
        # 创建处理线程
        self.batch_processing_thread = BatchProcessingThread(
            self.template_path, 
            output_dir, 
            field_values_list,
            naming_field
        )
        self.batch_processing_thread.progress_updated.connect(self.progress_bar.setValue)
        self.batch_processing_thread.processing_finished.connect(self.on_batch_processing_finished)
        self.batch_processing_thread.error_occurred.connect(self.on_batch_processing_error)
        self.batch_processing_thread.start()
        
    def on_batch_processing_finished(self, result_files):
        """批量处理完成回调"""
        self.progress_bar.setValue(self.progress_bar.maximum())
        
        # 显示结果
        result_text = f"批量处理完成！共生成 {len(result_files)} 个文档：\n\n"
        for file_path in result_files:
            result_text += f"• {os.path.basename(file_path)}\n"
            
        self.result_text.setText(result_text)
        self.result_text.setVisible(True)
        
        QMessageBox.information(self, "完成", f"批量处理完成！\n共生成 {len(result_files)} 个文档")
        
    def on_batch_processing_error(self, error_msg):
        """批量处理错误回调"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "错误", f"批量处理时出错：{error_msg}")


class BatchProcessingThread(QThread):
    """批量处理线程"""
    progress_updated = pyqtSignal(int)
    processing_finished = pyqtSignal(list)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, template_path, output_dir, field_values_list, naming_field=None):
        super().__init__()
        self.template_path = template_path
        self.output_dir = output_dir
        self.field_values_list = field_values_list
        self.naming_field = naming_field
        
    def run(self):
        try:
            processor = DocumentProcessor()
            result_files = []
            
            print(f"批量处理开始，共 {len(self.field_values_list)} 行数据")
            print(f"命名字段设置：{self.naming_field}")
            
            for i, field_values in enumerate(self.field_values_list):
                try:
                    print(f"\n=== 处理第 {i+1} 行数据 ===")
                    print(f"字段值：{field_values}")
                    
                    output_path = processor.process_document(
                        self.template_path, 
                        self.output_dir, 
                        field_values,
                        self.naming_field,
                        i  # 传递序号参数
                    )
                    
                    print(f"生成的文档：{output_path}")
                    result_files.append(output_path)
                    self.progress_updated.emit(i + 1)
                    
                except Exception as e:
                    print(f"处理第 {i+1} 组数据时出错：{str(e)}")
                    
            print(f"\n批量处理完成，共生成 {len(result_files)} 个文档")
            self.processing_finished.emit(result_files)
            
        except Exception as e:
            print(f"批量处理线程出错：{str(e)}")
            self.error_occurred.emit(str(e))