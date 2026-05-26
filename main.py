#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
QAutoDoc - 智能文档生成器
支持格式: DOCX / PPTX / XLSX
"""

import sys
import os
import re
import csv
import logging
import platform
import subprocess
from datetime import datetime
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
                             QPushButton, QLabel, QLineEdit, QTextEdit,
                             QFileDialog, QProgressBar, QToolTip,
                             QRadioButton, QButtonGroup, QComboBox, QCheckBox, QScrollArea,
                             QStackedLayout, QFrame, QGroupBox, QTabWidget, QStatusBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings, QPropertyAnimation, QEasingCurve, QTimer, QParallelAnimationGroup
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent
from PyQt5.QtWidgets import QGraphicsOpacityEffect

# ======================== 依赖安全预检 ========================
HAS_DOCX = False
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    pass

HAS_PPTX = False
try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    pass

HAS_OPENPYXL = False
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    pass


# ======================== 日志配置 (防写保护崩溃) ========================
try:
    LOG_DIR = Path("logs")
    LOG_DIR.mkdir(exist_ok=True)
    LOG_FILE = LOG_DIR / "app.log"
    # 尝试写入测试
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        pass
except (PermissionError, OSError):
    # 若无当前目录写入权限，重定向至用户主目录
    LOG_DIR = Path(os.path.expanduser("~")) / ".qautodoc"
    LOG_DIR.mkdir(exist_ok=True)
    LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)
logger = logging.getLogger(__name__)


# ======================== 字段提取器 ========================
class FieldExtractor:
    def __init__(self):
        # 高容错正则：支持半角 {{ }}、全角 ｛｛ ｝｝ 以及混合匹配
        self.field_pattern = re.compile(r'[\{｛]{2}(.*?)[\}｝]{2}')

    def extract_fields(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        
        # 显式依赖预检，防止 NameError 闪退
        if ext == '.docx' and not HAS_DOCX:
            raise ImportError("缺少依赖库: python-docx。请在终端执行: pip install python-docx")
        if ext == '.pptx' and not HAS_PPTX:
            raise ImportError("缺少依赖库: python-pptx。请在终端执行: pip install python-pptx")
        if ext == '.xlsx' and not HAS_OPENPYXL:
            raise ImportError("缺少依赖库: openpyxl。请在终端执行: pip install openpyxl")

        try:
            if ext == '.docx':
                return self._extract_docx(file_path)
            elif ext == '.pptx':
                return self._extract_pptx(file_path)
            elif ext == '.xlsx':
                return self._extract_xlsx(file_path)
            else:
                raise ValueError(f"不支持的格式: {ext}")
        except Exception as e:
            logger.exception("提取字段失败")
            raise Exception(f"解析失败: {str(e)}")

    def _extract_docx(self, path):
        doc = Document(path)
        fields = []
        seen = set()
        for p in doc.paragraphs:
            self._collect_fields(p.text, fields, seen)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self._collect_fields(p.text, fields, seen)
        return fields

    def _extract_pptx(self, path):
        prs = Presentation(path)
        fields = []
        seen = set()
        for slide in prs.slides:
            self._extract_shapes(slide.shapes, fields, seen)
        return fields

    def _extract_shapes(self, shapes, fields, seen):
        for shape in shapes:
            if hasattr(shape, "shapes"):
                self._extract_shapes(shape.shapes, fields, seen)
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    self._collect_fields(p.text, fields, seen)
            if hasattr(shape, "has_table") and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            self._collect_fields(p.text, fields, seen)

    def _extract_xlsx(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        fields = []
        seen = set()
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                for cell in row:
                    if isinstance(cell, str):
                        self._collect_fields(cell, fields, seen)
        return fields

    def _collect_fields(self, text, fields_list, seen_set):
        if not text:
            return
        for m in self.field_pattern.findall(text):
            f = m.strip()
            if f and f not in seen_set:
                fields_list.append(f)
                seen_set.add(f)


# ======================== 文档处理器 ========================
class DocumentProcessor:
    def __init__(self):
        self.field_pattern = re.compile(r'[\{｛]{2}(.*?)[\}｝]{2}')

    def process_document(self, template_path, output_dir, field_values, naming_field=None, index=None):
        if not os.path.exists(template_path):
            raise FileNotFoundError("模板文件丢失")
        if not field_values:
            raise ValueError("渲染数据为空")

        ext = os.path.splitext(template_path)[1].lower()
        filename = self._gen_filename(template_path, field_values, naming_field, index)
        out_path = os.path.join(output_dir, filename)

        try:
            if ext == '.docx':
                self._process_docx(template_path, out_path, field_values)
            elif ext == '.pptx':
                self._process_pptx(template_path, out_path, field_values)
            elif ext == '.xlsx':
                self._process_xlsx(template_path, out_path, field_values)
            else:
                raise ValueError(f"不支持的格式: {ext}")
            return out_path
        except Exception as e:
            logger.exception("文档处理失败")
            raise Exception(f"构建失败: {str(e)}")

    def _process_docx(self, tpl, out, fv):
        doc = Document(tpl)
        for p in doc.paragraphs:
            self._replace_in_para(p, fv)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        self._replace_in_para(p, fv)
        doc.save(out)

    def _process_pptx(self, tpl, out, fv):
        prs = Presentation(tpl)
        for slide in prs.slides:
            self._process_shapes(slide.shapes, fv)
        prs.save(out)

    def _process_shapes(self, shapes, fv):
        for shape in shapes:
            if hasattr(shape, "shapes"):
                self._process_shapes(shape.shapes, fv)
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    self._replace_in_para(p, fv)
            if hasattr(shape, "has_table") and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            self._replace_in_para(p, fv)

    def _process_xlsx(self, tpl, out, fv):
        wb = openpyxl.load_workbook(tpl)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        # 同时支持全角与半角花括号的正则替换
                        matches = list(self.field_pattern.finditer(cell.value))
                        if matches:
                            new_text = cell.value
                            for m in sorted(matches, key=lambda x: x.start(), reverse=True):
                                name = m.group(1).strip()
                                if name in fv:
                                    # 只重组字符串值，不触碰单元格底层Style样式
                                    new_text = new_text[:m.start()] + str(fv[name]) + new_text[m.end():]
                            cell.value = new_text
        wb.save(out)

    def _replace_in_para(self, paragraph, fv):
        if not paragraph.text:
            return
        matches = list(self.field_pattern.finditer(paragraph.text))
        if not matches: 
            return
        # 反向迭代匹配项，确保偏移量不变
        matches.sort(key=lambda x: x.start(), reverse=True)
        for m in matches:
            name = m.group(1).strip()
            if name in fv:
                self._replace_fmt_runs(paragraph, m.start(), m.end(), str(fv[name]))

    def _replace_fmt_runs(self, paragraph, sp, ep, repl):
        """
        高稳定性 Run 替换算法：
        1. 仅针对修改交叉区域的 Run.text
        2. 引入对非文本（如图片、图表、空白）Run 的 None 校验保护
        """
        runs = paragraph.runs
        if not runs:
            return

        pos = 0
        affected = []
        for i, r in enumerate(runs):
            run_text = r.text if r.text else ""
            rl = len(run_text)
            if pos < ep and pos + rl > sp:
                rs = max(0, sp - pos)
                re_ = min(rl, ep - pos)
                affected.append((i, rs, re_))
            pos += rl

        if not affected:
            return

        # 获取首个受影响的 Run 索引，用于附着新值并保持样式
        first_idx = affected[0][0]
        affected.sort(key=lambda x: x[0], reverse=True)

        for idx, rs, re_ in affected:
            run = runs[idx]
            run_text = run.text if run.text else ""
            before = run_text[:rs]
            after = run_text[re_:]
            if idx == first_idx:
                run.text = before + repl + after
            else:
                run.text = before + after

    def _gen_filename(self, tpl, fv, nf, idx):
        tn, ext = os.path.splitext(os.path.basename(tpl))
        if nf and nf in fv and fv[nf]:
            base = f"{tn}_{fv[nf]}"
        else:
            base = tn
            for f in ['姓名', '名称', '标题', 'subject', 'title', 'name']:
                if f in fv and fv[f]:
                    base = f"{tn}_{fv[f]}"
                    break
            if base == tn:
                base = f"{tn}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        if idx is not None:
            base += f"_{idx + 1:02d}"
        return self._sanitize(f"{base}{ext}")

    def _sanitize(self, fn):
        # 移除非法字符与控制字符
        for c in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
            fn = fn.replace(c, '_')
        fn = "".join(ch for ch in fn if ord(ch) >= 32)
        # 处理 Windows 文件名尾部空格及点
        name, ext = os.path.splitext(fn)
        name = name.strip().rstrip('.')
        if not name:
            name = "output"
        return f"{name}{ext}"


# ======================== UI 动画组件 ========================
class AnimatedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.opacity_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.opacity_effect)
        self.opacity_effect.setOpacity(1.0)
        
        self.animation_group = QParallelAnimationGroup()
        self.position_animation = QPropertyAnimation(self, b"geometry")
        self.position_animation.setDuration(150)
        self.position_animation.setEasingCurve(QEasingCurve.OutCubic)
        
        self.opacity_animation = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.opacity_animation.setDuration(200)
        self.opacity_animation.setEasingCurve(QEasingCurve.InOutCubic)
        
        self.animation_group.addAnimation(self.position_animation)
        self.animation_group.addAnimation(self.opacity_animation)
        self.is_hovered = False

    def enterEvent(self, event):
        if not self.is_hovered and self.isEnabled():
            self.is_hovered = True
            current_geometry = self.geometry()
            self.position_animation.setStartValue(current_geometry)
            self.position_animation.setEndValue(current_geometry.adjusted(0, -2, 0, -2))
            self.opacity_animation.setStartValue(1.0)
            self.opacity_animation.setEndValue(0.85)
            self.animation_group.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.is_hovered and self.isEnabled():
            self.is_hovered = False
            current_geometry = self.geometry()
            self.position_animation.setStartValue(current_geometry)
            self.position_animation.setEndValue(current_geometry.adjusted(0, 2, 0, 2))
            self.opacity_animation.setStartValue(0.85)
            self.opacity_animation.setEndValue(1.0)
            self.animation_group.start()
        super().leaveEvent(event)


class DropArea(QFrame):
    fileDropped = pyqtSignal(str)
    messageRequested = pyqtSignal(str, str) 

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.NoFrame)  
        self.setCursor(Qt.PointingHandCursor) 
        self.setStyleSheet("""
            QFrame { background-color: #f5f5f5; border: 2px dashed #cccccc; border-radius: 10px; }
            QFrame:hover { background-color: #e8e8e8; border: 2px dashed #999999; }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        tip_text = ("将模板文件拖拽至此 (支持 DOCX / PPTX / XLSX)\n\n"
                    "💡 提示：在文档中使用 {{变量名}} 占位，例如：{{姓名}}")
        self.label = QLabel(tip_text)
        self.label.setAlignment(Qt.AlignCenter)  
        self.label.setStyleSheet("color: #666666; font-size: 14px; font-weight: bold; border: none; line-height: 1.5;")
        layout.addWidget(self.label)
        self.setLayout(layout)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("QFrame { background-color: #e8e8e8; border: 2px dashed #666666; border-radius: 10px; }")

    def dragLeaveEvent(self, event):
        self.setStyleSheet("QFrame { background-color: #f5f5f5; border: 2px dashed #cccccc; border-radius: 10px; } QFrame:hover { background-color: #e8e8e8; border: 2px dashed #999999; }")

    def dropEvent(self, event: QDropEvent):
        self.dragLeaveEvent(event)
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if path.lower().endswith(('.docx', '.pptx', '.xlsx')):
                self.fileDropped.emit(path)
                self.show_success(os.path.basename(path))
            else:
                self.messageRequested.emit("格式拦截: 仅支持 DOCX / PPTX / XLSX", "error")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:  
            fp, _ = QFileDialog.getOpenFileName(self, "选择模板", "", "支持格式 (*.docx *.pptx *.xlsx)")
            if fp:
                self.fileDropped.emit(fp)
                self.show_success(os.path.basename(fp))

    def show_success(self, filename):
        self.label.setText(f"已加载: {filename}")
        self.label.setStyleSheet("color: #2196F3; font-size: 14px; font-weight: bold; border: none;")

    def reset(self):
        tip_text = ("将模板文件拖拽至此 (支持 DOCX / PPTX / XLSX)\n\n"
                    "💡 提示：在文档中使用 {{变量名}} 占位，例如：{{姓名}}")
        self.label.setText(tip_text)
        self.label.setStyleSheet("color: #666666; font-size: 14px; font-weight: bold; border: none; line-height: 1.5;")


# ======================== 批量处理线程 ========================
class BatchProcessingThread(QThread):
    progress_updated = pyqtSignal(int)
    processing_finished = pyqtSignal(list, list)

    def __init__(self, template_path, output_dir, field_values_list, naming_field=None):
        super().__init__()
        self.template_path = template_path
        self.output_dir = output_dir
        self.field_values_list = field_values_list
        self.naming_field = naming_field
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        processor = DocumentProcessor()
        result_files = []
        failed_items = []
        for i, fv in enumerate(self.field_values_list):
            if self._is_cancelled:
                break
            try:
                out = processor.process_document(self.template_path, self.output_dir, fv, self.naming_field, i)
                result_files.append(out)
            except Exception as e:
                failed_items.append(f"数据索引[{i+1}]: {str(e)}")
            self.progress_updated.emit(i + 1)
        self.processing_finished.emit(result_files, failed_items)


# ======================== 关于面板 ========================
class AboutPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(24)
        layout.setContentsMargins(30, 30, 30, 30)

        info_group = QGroupBox("关于")
        info_layout = QVBoxLayout()
        info_layout.setSpacing(16)
        info_layout.setContentsMargins(20, 25, 20, 20)

        title = QLabel("QAutoDoc")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #333333;")
        info_layout.addWidget(title)

        desc = QLabel("基于 {{变量名}} 的轻量文档批量渲染工具。\n"
                      "结合结构化数据，一键实现多文档自动化处理。\n\n"
                      "• 格式支持: Word (.docx) / PowerPoint (.pptx) / Excel (.xlsx)\n"
                      "• 格式安全: 独创安全算法，仅在文本节点进行替换，不破坏源文件底层样式结构\n"
                      "• 静默流交互: 精简传统弹窗阻断，全部流程通知托付于底部状态栏")
        desc.setStyleSheet("color: #666666; font-size: 13px; line-height: 1.6;")
        info_layout.addWidget(desc)
        
        info_layout.addStretch()
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.back_btn = AnimatedButton("返回")
        self.back_btn.setStyleSheet("""
            QPushButton { background: #f5f5f5; color: #666666; border: 1px solid #e0e0e0; padding: 8px 24px; border-radius: 6px; font-size: 14px; min-width: 100px; height: 36px; }
            QPushButton:hover { background: #e0e0e0; } QPushButton:pressed { background: #d0d0d0; }
        """)
        self.back_btn.clicked.connect(self.return_to_main)
        btn_layout.addWidget(self.back_btn)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def return_to_main(self):
        if isinstance(self.parent, MainWindow):
            self.parent.show_main_panel()


# ======================== 主窗口 ========================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("QAutoDoc", "QAutoDoc")
        self.template_path = ""
        self.output_dir = ""
        self.batch_thread = None
        self.field_text_edits = {}
        self.field_apply_checkboxes = {}

        self.init_style()
        self._load_settings()
        self.init_ui()

    def init_style(self):
        self.setWindowTitle("QAutoDoc_1.1.0 智能文档生成器")
        self.setStyleSheet("""
            QMainWindow { background-color: #ffffff; }
            QLabel { color: #333333; font-size: 13px; }
            QGroupBox { border: 1px solid #e0e0e0; border-radius: 8px; margin-top: 10px; padding: 10px; font-weight: bold; color: #333333; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; color: #333333; }
            QLineEdit, QComboBox, QTextEdit { padding: 6px 12px; border: 1px solid #e0e0e0; border-radius: 6px; font-size: 13px; background: white; color: #333333; }
            QLineEdit:hover, QComboBox:hover, QTextEdit:hover { border-color: #999999; }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border-color: #2196F3; }
            QComboBox::drop-down { border: none; width: 20px; }
            QComboBox::down-arrow { image: none; border: none; width: 0; height: 0; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 4px solid #999999; margin-right: 8px; }
            QComboBox QAbstractItemView { border: 1px solid #e0e0e0; background: white; selection-background-color: #f5f5f5; selection-color: #333333; }
            QRadioButton, QCheckBox { color: #333333; font-size: 13px; padding: 4px 0; }
            QRadioButton::indicator, QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #e0e0e0; border-radius: 8px; }
            QCheckBox::indicator { border-radius: 4px; }
            QRadioButton::indicator:hover, QCheckBox::indicator:hover { border-color: #999999; }
            QRadioButton::indicator:checked, QCheckBox::indicator:checked { background-color: #2196F3; border-color: #2196F3; }
            QTabWidget::pane { border: 1px solid #e0e0e0; background: white; border-radius: 6px; top: -1px; }
            QTabBar::tab { background: #f5f5f5; border: 1px solid #e0e0e0; padding: 8px 16px; margin-right: 2px; border-top-left-radius: 6px; border-top-right-radius: 6px; color: #666666; }
            QTabBar::tab:selected { background: white; border-bottom-color: white; color: #2196F3; font-weight: bold; }
            QProgressBar { border: 1px solid #e0e0e0; border-radius: 6px; background-color: #f5f5f5; text-align: center; color: #333333; }
            QProgressBar::chunk { background-color: #2196F3; border-radius: 5px; }
            QStatusBar { background-color: #f5f5f5; color: #666666; border-top: 1px solid #e0e0e0; padding: 2px; }
            QToolTip { background-color: #333333; color: white; border: none; padding: 8px 12px; border-radius: 4px; font-size: 12px; font-family: "Microsoft YaHei", sans-serif; }
        """)

    def _load_settings(self):
        if self.settings.value("geometry"):
            self.restoreGeometry(self.settings.value("geometry"))
        else:
            self.setGeometry(150, 100, 600, 750)

    def closeEvent(self, event):
        # 线程安全回收：确保主窗口退出时，后台渲染线程被同步终止并释放资源
        if self.batch_thread and self.batch_thread.isRunning():
            self.batch_thread.cancel()
            self.batch_thread.wait()
        self.settings.setValue("geometry", self.saveGeometry())
        event.accept()

    def init_ui(self):
        self.setMinimumSize(550, 700)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.stacked_layout = QStackedLayout(central_widget)

        # ----------------- 主面板 -----------------
        self.main_panel = QWidget()
        layout = QVBoxLayout(self.main_panel)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        self.drop_area = DropArea(self)
        self.drop_area.setFixedHeight(120)
        self.drop_area.fileDropped.connect(self.on_template_file_selected)
        self.drop_area.messageRequested.connect(self.show_status_message)
        layout.addWidget(self.drop_area)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setContentsMargins(0,0,0,0)
        scroll_layout.setSpacing(15)

        # 数据输入
        data_group = QGroupBox("数据源")
        data_layout = QVBoxLayout()
        data_layout.setSpacing(12)
        data_layout.setContentsMargins(15, 20, 15, 15)

        radio_layout = QHBoxLayout()
        self.manual_radio = QRadioButton("手动录入")
        self.csv_radio = QRadioButton("批量导入")
        self.manual_radio.setChecked(True)
        radio_layout.addWidget(self.manual_radio)
        radio_layout.addWidget(self.csv_radio)
        radio_layout.addStretch()
        
        # 快速指引悬浮按钮
        help_btn = QPushButton("💡 快速指引")
        help_btn.setCursor(Qt.PointingHandCursor)
        help_text = ("【快速上手指南】\n\n"
                     "1. 准备模板：在源文件中，将需要变化的地方写为 {{变量名}}。\n"
                     "   例如：尊敬的 {{姓名}}，您的余额为 {{金额}} 元。\n\n"
                     "2. 载入模板：将该文件拖入上方虚线框内，系统会自动识别变量。\n\n"
                     "3. 注入数据：\n"
                     "   - 手动录入：在下方出现的文本框内按行输入数据。\n"
                     "   - 批量导入：点击“下载数据模板”，在Excel填好后浏览导入。\n\n"
                     "4. 一键生成：点击右下角执行即可！")
        help_btn.setToolTip(help_text)
        help_btn.setStyleSheet("QPushButton { background: transparent; color: #2196F3; border: none; font-weight: bold; font-size: 13px; } QPushButton:hover { color: #1976D2; text-decoration: underline; }")
        help_btn.clicked.connect(lambda: QToolTip.showText(help_btn.mapToGlobal(help_btn.rect().bottomLeft()), help_text, help_btn))
        radio_layout.addWidget(help_btn)

        data_layout.addLayout(radio_layout)

        self.csv_widget = QWidget()
        csv_lay = QHBoxLayout(self.csv_widget)
        csv_lay.setContentsMargins(0, 0, 0, 0)
        
        self.export_csv_btn = QPushButton("① 下载数据模板")
        self.export_csv_btn.setStyleSheet("QPushButton { background: #e8f0fe; color: #1a73e8; border: 1px solid #bbdefb; padding: 6px 16px; border-radius: 4px; font-weight: bold;} QPushButton:hover { background: #d2e3fc; }")
        self.export_csv_btn.clicked.connect(self.export_csv_template)
        
        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("② 选择填写好的数据文件 (.csv)...")
        
        csv_btn = QPushButton("浏览")
        csv_btn.setStyleSheet("QPushButton { background: #f5f5f5; border: 1px solid #e0e0e0; padding: 6px 16px; border-radius: 4px; } QPushButton:hover { background: #e8e8e8; }")
        csv_btn.clicked.connect(self.browse_csv)
        
        csv_lay.addWidget(self.export_csv_btn)
        csv_lay.addWidget(self.csv_path_edit)
        csv_lay.addWidget(csv_btn)
        
        data_layout.addWidget(self.csv_widget)
        self.csv_widget.setVisible(False)

        self.field_container = QWidget()
        self.field_container_layout = QVBoxLayout(self.field_container)
        self.field_container_layout.setContentsMargins(0, 0, 0, 0)
        data_layout.addWidget(self.field_container)

        self.csv_radio.toggled.connect(self.toggle_input_method)
        self.init_manual_inputs()

        data_group.setLayout(data_layout)
        scroll_layout.addWidget(data_group)

        # 输出设置
        output_group = QGroupBox("输出设置")
        output_layout = QVBoxLayout()
        output_layout.setSpacing(12)
        output_layout.setContentsMargins(15, 20, 15, 15)

        naming_layout = QHBoxLayout()
        naming_label = QLabel("文件命名:")
        self.naming_field_combo = QComboBox()
        self.naming_field_combo.addItem("[默认推断]")
        self.naming_field_combo.setMinimumWidth(200)
        naming_layout.addWidget(naming_label)
        naming_layout.addWidget(self.naming_field_combo)
        naming_layout.addStretch()
        output_layout.addLayout(naming_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(12)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setVisible(False)
        output_layout.addWidget(self.progress_bar)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setVisible(False)
        self.result_text.setFixedHeight(80)
        output_layout.addWidget(self.result_text)

        output_group.setLayout(output_layout)
        scroll_layout.addWidget(output_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        # 底部栏
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 5, 0, 0)

        self.about_btn = AnimatedButton("关于")
        self.about_btn.setStyleSheet(self._get_secondary_btn_style())
        self.about_btn.clicked.connect(self.show_about_panel)
        button_layout.addWidget(self.about_btn)

        self.reset_btn = AnimatedButton("重置")
        self.reset_btn.setStyleSheet(self._get_secondary_btn_style())
        self.reset_btn.clicked.connect(self.reset_workspace)
        button_layout.addWidget(self.reset_btn)

        button_layout.addStretch()

        self.open_folder_btn = AnimatedButton("打开目录")
        self.open_folder_btn.setStyleSheet("QPushButton { background: #4CAF50; color: white; border: none; padding: 8px 16px; border-radius: 6px; font-weight: bold; } QPushButton:hover { background: #43A047; }")
        self.open_folder_btn.setVisible(False)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        button_layout.addWidget(self.open_folder_btn)

        self.action_btn = AnimatedButton("执行生成")
        self.action_btn.setStyleSheet(self._get_action_style("blue"))
        self.action_btn.clicked.connect(self.toggle_action)
        button_layout.addWidget(self.action_btn)

        layout.addLayout(button_layout)
        self.stacked_layout.addWidget(self.main_panel)

        self.about_panel = AboutPanel(self)
        self.stacked_layout.addWidget(self.about_panel)
        self.stacked_layout.setCurrentWidget(self.main_panel)

        # 状态栏
        self.statusBar = self.statusBar()
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #666666; font-size: 13px; padding: 2px 6px;")
        self.statusBar.addWidget(self.status_label)

    def _get_secondary_btn_style(self):
        return "QPushButton { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #f0f0f0, stop:1 #e8e8e8); color: #333333; border: 1px solid #e0e0e0; padding: 8px 16px; border-radius: 6px; } QPushButton:hover { background: #e8e8e8; } QPushButton:pressed { background: #d0d0d0; }"

    def _get_action_style(self, color_type):
        if color_type == "blue":
            return "QPushButton { background: #2196F3; color: white; border: none; padding: 8px 24px; border-radius: 6px; font-weight: bold; } QPushButton:hover { background: #1976D2; } QPushButton:disabled { background: #BDBDBD; color: #E0E0E0; }"
        return "QPushButton { background: #F44336; color: white; border: none; padding: 8px 24px; border-radius: 6px; font-weight: bold; } QPushButton:hover { background: #D32F2F; }"

    def show_status_message(self, message, msg_type="info", duration=5000):
        cmap = {"info": "#666666", "success": "#4CAF50", "warning": "#FF9800", "error": "#F44336"}
        self.status_label.setStyleSheet(f"color: {cmap.get(msg_type, '#666666')}; font-weight: bold; padding: 2px 6px;")
        self.status_label.setText(message)
        if duration > 0:
            if hasattr(self, '_status_timer'): 
                self._status_timer.stop()
            self._status_timer = QTimer(self)
            self._status_timer.setSingleShot(True)
            self._status_timer.timeout.connect(self._reset_status_message)
            self._status_timer.start(duration)

    def _reset_status_message(self):
        if self.batch_thread and self.batch_thread.isRunning(): 
            return
        self.status_label.setStyleSheet("color: #666666; font-weight: normal; padding: 2px 6px;")
        self.status_label.setText("就绪" if self.template_path else "请载入源模板")

    def show_about_panel(self):
        self._slide_panel(self.about_panel)

    def show_main_panel(self):
        self._slide_panel(self.main_panel)

    def _slide_panel(self, tgt):
        self.animation = QPropertyAnimation(self.stacked_layout.currentWidget(), b"geometry")
        self.animation.setDuration(250)
        self.animation.setEasingCurve(QEasingCurve.OutCubic)
        self.animation.setStartValue(self.stacked_layout.currentWidget().geometry())
        self.stacked_layout.setCurrentWidget(tgt)
        self.animation.setEndValue(tgt.geometry())
        self.animation.start()

    def reset_workspace(self):
        self.template_path = ""
        self.settings.setValue("template_path", "")
        self.drop_area.reset()
        self.csv_path_edit.clear()
        self.naming_field_combo.clear()
        self.naming_field_combo.addItem("[默认推断]")
        self._clear_field_tabs()
        self.progress_bar.setVisible(False)
        self.result_text.setVisible(False)
        self.open_folder_btn.setVisible(False)
        self.show_status_message("工作区已重置", "success")

    def export_csv_template(self):
        if not self.template_path: 
            return self.show_status_message("中止: 模板未就绪", "warning")
        fields = list(self.field_text_edits.keys()) if self.field_text_edits else self.extract_fields()
        if not fields: 
            return self.show_status_message("中止: 未检测到变量", "warning")
            
        save_path, _ = QFileDialog.getSaveFileName(self, "下载数据模板", "数据模板.csv", "CSV (*.csv)")
        if save_path:
            try:
                with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
                    csv.writer(f).writerow(fields) 
                self.show_status_message(f"下载完毕: {os.path.basename(save_path)}", "success", 8000)
            except Exception as e:
                self.show_status_message(f"下载受阻: {str(e)}", "error")

    def open_output_folder(self):
        if self.output_dir and os.path.exists(self.output_dir):
            try:
                if platform.system() == "Windows": 
                    os.startfile(self.output_dir)
                elif platform.system() == "Darwin": 
                    subprocess.Popen(["open", self.output_dir])
                else: 
                    subprocess.Popen(["xdg-open", self.output_dir])
            except Exception: 
                pass

    def on_template_file_selected(self, path):
        self.template_path = path
        self.settings.setValue("template_path", path)
        self.extract_fields()
        self.open_folder_btn.setVisible(False)

    def browse_csv(self):
        fp, _ = QFileDialog.getOpenFileName(self, "定位数据", "", "CSV (*.csv)")
        if fp: 
            self.csv_path_edit.setText(fp)

    def toggle_input_method(self):
        is_csv = self.csv_radio.isChecked()
        self.csv_widget.setVisible(is_csv)
        self.field_container.setVisible(not is_csv)
        if not is_csv and self.template_path: 
            self.extract_fields()

    def init_manual_inputs(self):
        self.field_tabs_widget = None

    def _clear_field_tabs(self):
        if self.field_tabs_widget is not None:
            self.field_container_layout.removeWidget(self.field_tabs_widget)
            self.field_tabs_widget.deleteLater()
            self.field_tabs_widget = None
        self.field_text_edits = {}
        self.field_apply_checkboxes = {}

    def extract_fields(self):
        if not self.template_path: 
            return []
        try:
            fields = FieldExtractor().extract_fields(self.template_path)
            self.naming_field_combo.clear()
            self.naming_field_combo.addItem("[默认推断]")
            for f in fields: 
                self.naming_field_combo.addItem(f)
            if self.manual_radio.isChecked(): 
                self.create_field_tabs(fields)
            
            if fields:
                self.show_status_message(f"解析完成: 识别到 {len(fields)} 个变量", "success")
            else:
                self.show_status_message("解析完成，但未在模板中检测到 {{变量}} 占位符", "warning")
            return fields
        except Exception as e:
            self.show_status_message(str(e), "error")
            return []

    def create_field_tabs(self, fields):
        self._clear_field_tabs()
        if not fields: 
            return
        self.field_tabs_widget = QTabWidget()
        for field in fields:
            tab = QWidget()
            tl = QVBoxLayout(tab)
            tl.setSpacing(10)
            tl.addWidget(QLabel(f"变量 <b>{field}</b> (每行对应一份文件):"))
            te = QTextEdit()
            te.setPlaceholderText("第一项\n第二项")
            tl.addWidget(te)
            cb = QCheckBox("全局锁定 (所有文件应用首行值)")
            tl.addWidget(cb)
            self.field_text_edits[field], self.field_apply_checkboxes[field] = te, cb
            self.field_tabs_widget.addTab(tab, field)
        self.field_container_layout.addWidget(self.field_tabs_widget)

    def get_field_values_from_manual(self):
        if not self.field_text_edits: 
            return []
        fl, fa, mx = {}, {}, 0
        for fn, te in self.field_text_edits.items():
            txt = te.toPlainText().strip()
            fa[fn] = self.field_apply_checkboxes.get(fn, QCheckBox()).isChecked()
            if txt:
                if fa[fn]:
                    fl[fn] = [txt]
                else:
                    lines = [l.strip() for l in txt.split('\n')]
                    # 仅移除尾部的空行，保留中间的有意义空项，防止索引错位
                    while lines and not lines[-1]:
                        lines.pop()
                    fl[fn] = lines
                mx = max(mx, len(fl[fn]))
            else: 
                fl[fn] = []
        
        for fn, lines in fl.items():
            if not fa.get(fn, False): 
                mx = max(mx, len(lines))
        if mx == 0: 
            return []
        
        rows = []
        for i in range(mx):
            rd = {}
            for fn, lines in fl.items():
                if fa.get(fn, False) and lines: 
                    rd[fn] = lines[0]
                elif i < len(lines): 
                    rd[fn] = lines[i]
                else: 
                    rd[fn] = ""
            if any(rd.values()): 
                rows.append(rd)
        return rows

    def get_field_values_from_csv(self):
        """
        高鲁棒性 CSV 读入器：
        依次尝试多类编码，有效防御中文字符集引发的 UnicodeDecodeError。
        """
        cp = self.csv_path_edit.text()
        if not cp or not os.path.exists(cp): 
            return []
        
        encodings = ['utf-8-sig', 'gb18030', 'utf-8', 'ansi']
        for enc in encodings:
            try:
                res = []
                with open(cp, 'r', encoding=enc) as f:
                    for row in csv.DictReader(f):
                        cln = {k.strip(): v.strip() if v else "" for k, v in row.items()}
                        if any(cln.values()): 
                            res.append(cln)
                return res
            except UnicodeDecodeError:
                continue
            except Exception as e:
                self.show_status_message(f"解析流异常: {str(e)}", "error")
                return []
        
        self.show_status_message("解析流异常: 无法自动识别的文件编码，请确保文件保存为 UTF-8 或 GBK 格式", "error")
        return []

    def toggle_action(self):
        if self.action_btn.text() == "执行生成": 
            self.start_processing()
        else: 
            self.cancel_processing()

    def start_processing(self):
        if not self.template_path: 
            return self.show_status_message("中止: 模板未就绪", "warning")
        fvl = self.get_field_values_from_csv() if self.csv_radio.isChecked() else self.get_field_values_from_manual()
        if not fvl: 
            return self.show_status_message("中止: 数据源无效", "error")
            
        outdir = QFileDialog.getExistingDirectory(self, "定向输出路径")
        if not outdir: 
            return
        self.output_dir = outdir
        self.settings.setValue("output_dir", outdir)

        self.action_btn.setText("终止任务")
        self.action_btn.setStyleSheet(self._get_action_style("red"))
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(fvl))
        self.progress_bar.setValue(0)
        self.result_text.setVisible(False)
        self.open_folder_btn.setVisible(False)
        self.show_status_message("执行引擎启动...", "info", 0)

        nf = None if self.naming_field_combo.currentText() == "[默认推断]" else self.naming_field_combo.currentText()
        self.batch_thread = BatchProcessingThread(self.template_path, outdir, fvl, nf)
        self.batch_thread.progress_updated.connect(self.update_progress)
        self.batch_thread.processing_finished.connect(self.on_processing_finished)
        self.batch_thread.start()

    def update_progress(self, val):
        self.progress_bar.setValue(val)
        self.show_status_message(f"渲染中 [{val}/{self.progress_bar.maximum()}]...", "info", 0)

    def cancel_processing(self):
        if self.batch_thread and self.batch_thread.isRunning():
            self.batch_thread.cancel()
            self.action_btn.setEnabled(False)
            self.show_status_message("挂起指令下达中...", "warning", 0)

    def on_processing_finished(self, success, fails):
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.action_btn.setEnabled(True)
        self.action_btn.setText("执行生成")
        self.action_btn.setStyleSheet(self._get_action_style("blue"))

        if fails:
            self.show_status_message(f"调度完结: 成功 {len(success)} / 阻断 {len(fails)}", "warning", 8000)
            txt = f"执行概览 | 投递成功: {len(success)}\n\n[异常队列] {len(fails)} 宗 (可能部分输出目标文件已被系统独占):\n" + "\n".join(fails)
            self.result_text.setText(txt)
            self.result_text.setVisible(True)
        else:
            self.show_status_message(f"渲染达标: 成功输出 {len(success)} 份构建文件", "success", 8000)
            self.result_text.setVisible(False)
        
        if success: 
            self.open_folder_btn.setVisible(True)


# ======================== 挂载点 ========================
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Microsoft YaHei", 9))
    app.setApplicationName("QAutoDoc")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()