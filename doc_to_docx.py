#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将"文件模板"目录下的所有.doc文件转换为.docx格式，保持目录结构，输出到"标准化模板"目录
"""
import os
import shutil
import subprocess
import sys
import configparser
from pathlib import Path

# 读取配置文件
config = configparser.ConfigParser()
config_path = "config.ini"
if os.path.exists(config_path):
    config.read(config_path, encoding="utf-8")

# 配置，优先使用配置文件中的值，否则使用默认值
SOURCE_DIR = config.get("paths", "source_dir", fallback="文件模板")
TARGET_DIR = config.get("paths", "template_dir", fallback="标准化模板")

def convert_doc_to_docx_win(doc_path, docx_path):
    """Windows平台使用pywin32转换doc到docx"""
    try:
        import win32com.client as win32
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs(os.path.abspath(docx_path), FileFormat=16)  # 16 = docx format
        doc.Close()
        word.Quit()
        return True
    except Exception as e:
        print(f"转换失败 {doc_path}: {str(e)}")
        return False

def convert_doc_to_docx_unix(doc_path, docx_path):
    """Unix/Linux/Mac平台使用libreoffice转换doc到docx"""
    try:
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "docx",
            "--outdir", os.path.dirname(os.path.abspath(docx_path)),
            os.path.abspath(doc_path)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            # 重命名文件（libreoffice会自动生成.docx后缀）
            generated_file = os.path.splitext(os.path.basename(doc_path))[0] + ".docx"
            generated_path = os.path.join(os.path.dirname(os.path.abspath(docx_path)), generated_file)
            if os.path.exists(generated_path) and generated_path != os.path.abspath(docx_path):
                os.rename(generated_path, os.path.abspath(docx_path))
            return True
        else:
            print(f"转换失败 {doc_path}: {result.stderr}")
            return False
    except Exception as e:
        print(f"转换失败 {doc_path}: {str(e)}")
        return False

def convert_doc_to_docx(doc_path, docx_path):
    """根据平台选择转换方式"""
    if sys.platform.startswith("win"):
        return convert_doc_to_docx_win(doc_path, docx_path)
    else:
        return convert_doc_to_docx_unix(doc_path, docx_path)

def main():
    # 创建目标目录
    Path(TARGET_DIR).mkdir(parents=True, exist_ok=True)

    # 统计信息
    total_files = 0
    converted_files = 0
    copied_files = 0
    failed_files = 0

    # 遍历源目录
    for root, dirs, files in os.walk(SOURCE_DIR):
        # 计算相对路径
        rel_path = os.path.relpath(root, SOURCE_DIR)
        target_root = os.path.join(TARGET_DIR, rel_path)

        # 创建目标子目录
        Path(target_root).mkdir(parents=True, exist_ok=True)

        for file in files:
            total_files += 1
            source_file = os.path.join(root, file)
            target_file = os.path.join(target_root, file)

            # 跳过临时文件和隐藏文件
            if file.startswith("~$") or file.startswith("."):
                continue

            # 处理doc文件
            if file.lower().endswith(".doc") and not file.lower().endswith(".docx"):
                # 生成目标docx文件名
                target_docx = os.path.splitext(target_file)[0] + ".docx"

                # 检查目标文件是否已存在，存在则跳过
                if os.path.exists(target_docx):
                    print(f"跳过: {target_docx} 已存在")
                    copied_files += 1
                    continue

                print(f"转换: {source_file} -> {target_docx}")
                if convert_doc_to_docx(source_file, target_docx):
                    converted_files += 1
                else:
                    failed_files += 1
                    # 转换失败的话检查原文件是否存在，不存在才复制
                    if not os.path.exists(target_file):
                        print(f"复制原文件: {source_file} -> {target_file}")
                        shutil.copy2(source_file, target_file)
                    else:
                        print(f"跳过: {target_file} 已存在")
            else:
                # 其他文件直接复制（包括.docx、.xlsx、.xlsm等格式）
                # 检查目标文件是否已存在，存在则跳过
                if os.path.exists(target_file):
                    print(f"跳过: {target_file} 已存在")
                    copied_files += 1
                    continue

                if file.lower().endswith(".docx"):
                    print(f"复制docx文件: {source_file} -> {target_file}")
                else:
                    print(f"复制: {source_file} -> {target_file}")
                shutil.copy2(source_file, target_file)
                copied_files += 1

    # 输出统计
    print(f"\n处理完成！")
    print(f"总文件数: {total_files}")
    print(f"成功转换doc文件: {converted_files}")
    print(f"直接复制文件: {copied_files}")
    print(f"转换失败文件: {failed_files}")

if __name__ == "__main__":
    main()
