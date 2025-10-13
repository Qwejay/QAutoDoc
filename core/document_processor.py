#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档处理器
用于处理Word文档的字段替换和格式保持
"""

import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import parse_xml

from .field_extractor import FieldExtractor


class DocumentProcessor:
    """文档处理器类"""
    
    def __init__(self):
        self.field_extractor = FieldExtractor()
        self.field_pattern = re.compile(r'\{\{(.*?)\}\}')
        
    def process_document(self, template_path, output_dir, field_values, naming_field=None, index=None):
        """
        处理文档并进行字段替换
        
        Args:
            template_path: 模板文档路径
            output_dir: 输出目录
            field_values: 字段值字典
            naming_field: 命名字段，如果为None则自动选择
            index: 文档序号（用于批量处理时区分同名文件）
            
        Returns:
            str: 生成的文档路径
        """
        try:
            # 验证输入参数
            self._validate_inputs(template_path, output_dir, field_values)
            
            # 加载模板文档
            doc = Document(template_path)
            
            # 执行字段替换
            self._replace_fields_in_document(doc, field_values)
            
            # 生成输出文件名
            output_filename = self._generate_output_filename(template_path, field_values, naming_field, index)
            output_path = os.path.join(output_dir, output_filename)
            
            # 保存文档
            doc.save(output_path)
            
            return output_path
            
        except Exception as e:
            raise Exception(f"处理文档时出错：{str(e)}")
            
    def _validate_inputs(self, template_path, output_dir, field_values):
        """验证输入参数"""
        if not os.path.exists(template_path):
            raise Exception("模板文件不存在")
            
        if not os.path.isdir(output_dir):
            raise Exception("输出目录不存在")
            
        if not field_values:
            raise Exception("字段值为空")
            
    def _replace_fields_in_document(self, doc, field_values):
        """
        在文档中替换字段
        
        Args:
            doc: Word文档对象
            field_values: 字段值字典
        """
        # 替换段落中的字段
        for paragraph in doc.paragraphs:
            self._replace_fields_in_paragraph(paragraph, field_values)
            
        # 替换表格中的字段
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_fields_in_paragraph(paragraph, field_values)
                        
    def _replace_fields_in_paragraph(self, paragraph, field_values):
        """
        在段落中替换字段，保持格式
        
        Args:
            paragraph: 段落对象
            field_values: 字段值字典
        """
        # 获取段落文本
        original_text = paragraph.text
        
        # 查找所有字段位置
        matches = list(self.field_pattern.finditer(original_text))
        
        if not matches:
            return
            
        # 按位置从后往前替换，避免位置偏移问题
        matches.sort(key=lambda x: x.start(), reverse=True)
        
        for match in matches:
            field_name = match.group(1).strip()
            if field_name in field_values:
                replacement_text = field_values[field_name]
                
                # 获取字段在段落中的位置
                start_pos = match.start()
                end_pos = match.end()
                
                # 执行替换，保持格式
                self._replace_text_with_formatting(paragraph, start_pos, end_pos, replacement_text)
                
    def _replace_text_with_formatting(self, paragraph, start_pos, end_pos, replacement_text):
        """
        替换文本并保持格式
        
        Args:
            paragraph: 段落对象
            start_pos: 开始位置
            end_pos: 结束位置
            replacement_text: 替换文本
        """
        # 清空段落内容
        p = paragraph._p
        
        # 获取段落的所有运行(run)
        runs = paragraph.runs
        
        if not runs:
            # 如果没有运行，直接设置文本
            paragraph.text = replacement_text
            return
            
        # 计算文本位置
        current_pos = 0
        text_to_replace = ""
        runs_to_remove = []
        
        # 找到需要替换的运行
        for i, run in enumerate(runs):
            run_text = run.text
            run_length = len(run_text)
            
            # 检查当前运行是否在替换范围内
            if current_pos + run_length > start_pos and current_pos < end_pos:
                # 计算在运行中的开始和结束位置
                run_start = max(0, start_pos - current_pos)
                run_end = min(run_length, end_pos - current_pos)
                
                # 提取要替换的文本
                text_to_replace += run_text[run_start:run_end]
                
                # 标记需要处理的运行
                runs_to_remove.append((i, run_start, run_end))
                
            current_pos += run_length
            
        # 从后往前处理运行，避免索引问题
        runs_to_remove.sort(reverse=True)
        
        # 替换文本
        for run_idx, run_start, run_end in runs_to_remove:
            run = runs[run_idx]
            run_text = run.text
            
            # 分割运行文本
            before_text = run_text[:run_start]
            after_text = run_text[run_end:]
            
            # 更新运行文本
            if run_idx == runs_to_remove[0][0]:  # 第一个处理的运行
                run.text = before_text + replacement_text + after_text
            else:
                run.text = before_text + after_text
                
        # 清理空运行
        self._cleanup_empty_runs(paragraph)
        
    def _cleanup_empty_runs(self, paragraph):
        """清理空运行"""
        runs = paragraph.runs
        empty_runs = []
        
        for i, run in enumerate(runs):
            if not run.text.strip():
                empty_runs.append(i)
                
        # 从后往前删除空运行
        for run_idx in sorted(empty_runs, reverse=True):
            run = runs[run_idx]
            r = run._r
            r.getparent().remove(r)
            
    def _generate_output_filename(self, template_path, field_values, naming_field=None, index=None):
        """
        生成输出文件名
        
        Args:
            template_path: 模板路径
            field_values: 字段值字典
            naming_field: 命名字段，如果为None则自动选择
            index: 文档序号（用于批量处理时区分同名文件）
            
        Returns:
            str: 输出文件名
        """
        # 获取模板文件名（不含扩展名）
        template_name = os.path.splitext(os.path.basename(template_path))[0]
        
        # 如果指定了命名字段，使用该字段的值
        if naming_field and naming_field in field_values and field_values[naming_field]:
            base_filename = f"{template_name}_{field_values[naming_field]}"
            # 如果有序号，添加序号
            if index is not None:
                base_filename += f"_{index+1:02d}"
            filename = f"{base_filename}.docx"
            # 清理文件名中的非法字符
            filename = self._sanitize_filename(filename)
            return filename
        
        # 尝试使用有意义的字段值作为文件名的一部分
        meaningful_fields = ['姓名', '名称', '标题', 'subject', 'title', 'name']
        
        for field in meaningful_fields:
            if field in field_values and field_values[field]:
                base_filename = f"{template_name}_{field_values[field]}"
                # 如果有序号，添加序号
                if index is not None:
                    base_filename += f"_{index+1:02d}"
                filename = f"{base_filename}.docx"
                # 清理文件名中的非法字符
                filename = self._sanitize_filename(filename)
                return filename
                
        # 如果没有有意义的字段，使用时间戳和序号
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = f"{template_name}_{timestamp}"
        # 如果有序号，添加序号
        if index is not None:
            base_filename += f"_{index+1:02d}"
        return f"{base_filename}.docx"
        
    def _sanitize_filename(self, filename):
        """清理文件名中的非法字符"""
        # Windows文件名非法字符
        illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        
        for char in illegal_chars:
            filename = filename.replace(char, '_')
            
        return filename
    
    def batch_process(self, template_path, output_dir, field_values_list):
        """
        批量处理文档
        
        Args:
            template_path: 模板路径
            output_dir: 输出目录
            field_values_list: 字段值字典列表
            
        Returns:
            list: 生成的文档路径列表
        """
        generated_files = []
        
        for i, field_values in enumerate(field_values_list):
            try:
                output_path = self.process_document(template_path, output_dir, field_values)
                generated_files.append(output_path)
            except Exception as e:
                print(f"处理第 {i+1} 组数据时出错：{str(e)}")
                
        return generated_files