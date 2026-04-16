#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据"项目信息.txt"中的配置，替换"标准化模板"目录下所有docx文件中的{字段}标记，结果保存到"输出文件"目录
"""
import os
import shutil
import re
import configparser
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# 读取配置文件
config = configparser.ConfigParser()
config_path = "config.ini"
if os.path.exists(config_path):
    config.read(config_path, encoding="utf-8")

# 配置，优先使用配置文件中的值，否则使用默认值
TEMPLATE_DIR = config.get("paths", "template_dir", fallback="标准化模板")
OUTPUT_DIR = config.get("paths", "output_dir", fallback="输出文件")
PROJECT_INFO_FILE = config.get("paths", "project_info_file", fallback="项目信息.txt")

def load_project_info():
    """加载项目信息，返回字段字典"""
    info = {}
    if not os.path.exists(PROJECT_INFO_FILE):
        print(f"错误：{PROJECT_INFO_FILE} 文件不存在！")
        return None

    try:
        with open(PROJECT_INFO_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if ":" in line:
                    key, value = line.split(":", 1)
                    key = key.strip()
                    value = value.strip()
                    info[key] = value
        print(f"成功加载 {len(info)} 个项目字段")
        return info
    except Exception as e:
        print(f"读取 {PROJECT_INFO_FILE} 失败: {str(e)}")
        return None

def replace_text_in_doc(doc, info):
    """替换文档中的所有字段"""
    # 替换段落中的字段
    for para in doc.paragraphs:
        for key, value in info.items():
            placeholder = "{" + key + "}"
            if placeholder in para.text:
                para.text = para.text.replace(placeholder, value)

    # 替换表格中的字段
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in info.items():
                        placeholder = "{" + key + "}"
                        if placeholder in para.text:
                            para.text = para.text.replace(placeholder, value)

    return doc

def replace_text_in_excel(wb, info):
    """替换Excel文件中的所有字段"""
    # 遍历所有工作表
    for ws in wb.worksheets:
        # 遍历所有单元格
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for key, value in info.items():
                        placeholder = "{" + key + "}"
                        if placeholder in cell.value:
                            cell.value = cell.value.replace(placeholder, value)
    return wb

def process_docx_file(source_path, target_path, info):
    """处理单个docx文件，替换字段后保存到目标路径"""
    try:
        # 跳过临时文件和损坏文件
        if os.path.basename(source_path).startswith("~$"):
            print(f"跳过临时文件: {source_path}")
            return False
        doc = Document(source_path)
        doc = replace_text_in_doc(doc, info)
        doc.save(target_path)
        return True
    except Exception as e:
        print(f"处理文件失败 {source_path}: {str(e)}")
        return False

def process_excel_file(source_path, target_path, info):
    """处理单个Excel文件，替换字段后保存到目标路径"""
    try:
        # 跳过临时文件
        if os.path.basename(source_path).startswith("~$"):
            print(f"跳过临时文件: {source_path}")
            return False
        # 加载Excel文件，keep_vba=True保留宏（支持.xlsm）
        wb = load_workbook(source_path, keep_vba=True)
        wb = replace_text_in_excel(wb, info)
        wb.save(target_path)
        return True
    except InvalidFileException as e:
        print(f"无效Excel文件 {source_path}: {str(e)}")
        return False
    except Exception as e:
        print(f"处理Excel文件失败 {source_path}: {str(e)}")
        return False

def main():
    # 加载项目信息
    info = load_project_info()
    if not info:
        return

    # 检查模板目录
    if not os.path.exists(TEMPLATE_DIR):
        print(f"错误：{TEMPLATE_DIR} 目录不存在！请先运行 doc_to_docx.py 生成标准化模板")
        return

    # 创建输出目录
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # 统计信息
    total_files = 0
    processed_docx = 0
    processed_excel = 0
    copied_files = 0
    failed_files = 0

    # 遍历模板目录
    for root, dirs, files in os.walk(TEMPLATE_DIR):
        # 计算相对路径
        rel_path = os.path.relpath(root, TEMPLATE_DIR)
        target_root = os.path.join(OUTPUT_DIR, rel_path)

        # 创建输出子目录
        Path(target_root).mkdir(parents=True, exist_ok=True)

        for file in files:
            total_files += 1
            source_file = os.path.join(root, file)
            target_file = os.path.join(target_root, file)

            # 跳过临时文件
            if file.startswith("~$"):
                print(f"跳过临时文件: {source_file}")
                continue

            # 处理docx文件
            if file.lower().endswith(".docx"):
                print(f"处理Word: {source_file} -> {target_file}")
                if process_docx_file(source_file, target_file, info):
                    processed_docx += 1
                else:
                    failed_files += 1
                    # 处理失败的话复制原文件
                    print(f"复制原文件: {source_file} -> {target_file}")
                    shutil.copy2(source_file, target_file)
            # 处理Excel文件
            elif file.lower().endswith((".xlsx", ".xlsm")):
                print(f"处理Excel: {source_file} -> {target_file}")
                if process_excel_file(source_file, target_file, info):
                    processed_excel += 1
                else:
                    failed_files += 1
                    # 处理失败的话复制原文件
                    print(f"复制原文件: {source_file} -> {target_file}")
                    shutil.copy2(source_file, target_file)
            else:
                # 其他文件直接复制
                print(f"复制: {source_file} -> {target_file}")
                shutil.copy2(source_file, target_file)
                copied_files += 1

    # 输出统计
    print(f"\n处理完成！")
    print(f"总文件数: {total_files}")
    print(f"成功处理Word文件: {processed_docx}")
    print(f"成功处理Excel文件: {processed_excel}")
    print(f"直接复制其他文件: {copied_files}")
    print(f"处理失败文件: {failed_files}")
    print(f"结果已保存到: {OUTPUT_DIR} 目录")

if __name__ == "__main__":
    main()
