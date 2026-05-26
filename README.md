# QAutoDoc - Word文档内容自动填充工具

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.7+-blue.svg" alt="Python Version">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="License">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg" alt="Platform">
</p>

<p align="center">
  <strong>QAutoDoc是一款专为办公文档设计的内容自动填充工具，具有格式保持、字段识别和批量处理等核心功能。</strong>
</p>

---

## 📖 项目简介

**QAutoDoc** 是一款免配置的桌面端文档自动化填充工具。它支持通过极简的数据表（CSV）或手工分行录入，一键将结构化数据批量渲染进 Word (`.docx`)、PowerPoint (`.pptx`) 和 Excel (`.xlsx`) 模板中。

该工具的核心优势在于采用**段落 Run 级格式锁定算法**，在注入文本数据的同时，最大程度保护源模板底层的字体、行距、颜色与排版结构不发生形变，非常适合用于批量制作合同、通知书、邀请函、标准化报告等场景。

---

## 🚀 核心特性

*   **多格式原生兼容**
    *   **Word (`.docx`)**：支持段落和表格单元格内文本的精准匹配与替换。
    *   **PowerPoint (`.pptx`)**：支持穿透多层复杂“组合图层 (Group Shapes)”及幻灯片表格进行精准填充。
    *   **Excel (`.xlsx`)**：提供安全的单元格数值重组，不破坏原有的边框、背景及条件格式。
*   **格式高保真锁定**
    *   **物理分块（Split Runs）对齐**：针对 Word 和 PPT 在编辑过程中占位符常被物理切碎的问题，采用基于段落字符偏移的映射算法，不删除或重构底层 XML 节点。
    *   **非文本运行块保护**：自动识别并安全略过文档内嵌的图片、图表、公式等非文本运行块，保障排版复杂文档时的稳定性。
*   **多编码容错 CSV 导入**
    *   导入 CSV 时自动依次尝试 `utf-8-sig`、`gb18030`、`utf-8`、`ansi` 编码，有效解决 Windows 系统下 Excel 默认导出中文 GBK 编码导致闪退的问题。
*   **异步多线程架构**
    *   渲染引擎运行于独立的后台线程中，支持高并发处理，提供进度反馈，并支持中途安全“终止任务”，主窗口退出时自动回收后台资源。

---

## 🛠️ 安装与部署

### 1. 环境准备
*   Python 3.7 或更高版本
*   推荐使用 Windows 10/11（亦支持 macOS 和 Linux 桌面环境）

### 2. 安装依赖项
```bash
pip install PyQt5 python-docx python-pptx openpyxl

### 许可证

本项目采用MIT许可证。

### 贡献

欢迎提交Issue和Pull Request来改进项目。

### 联系方式

如有问题或建议，请通过项目Issue页面联系。
