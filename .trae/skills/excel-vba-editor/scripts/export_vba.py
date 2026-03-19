#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导出 Excel 文件中的 VBA 代码到 vba_src/ 目录

工作流步骤:
    第2步: 导出这些模块的代码到 vba_src/
    （在分析修改范围后执行，Git commit 前执行）

用法:
    # 导出所有模块
    python .trae/skills/excel-vba-editor/scripts/export_vba.py "工作簿名称.xlsm"
    
    # 导出指定模块（推荐）
    python .trae/skills/excel-vba-editor/scripts/export_vba.py "工作簿名称.xlsm" "ModuleName"
    python .trae/skills/excel-vba-editor/scripts/export_vba.py "工作簿名称.xlsm" "Module1,Module2,Module3"
    
    # 交互式输入
    python .trae/skills/excel-vba-editor/scripts/export_vba.py

完整工作流:
    1. 分析本次任务会修改哪些VBA代码(模块、类模块)
    2. 导出这些模块的代码到 vba_src/  <- 本脚本
    3. Git commit
    4. AI修改代码
    5. 写回Excel
"""

import sys
import json
from pathlib import Path
from datetime import datetime

from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_component_type, sanitize_filename, get_module
)


def normalize_line_endings(code):
    """
    规范化换行符，避免产生额外空行
    VBA代码使用 \r\n，需要统一处理
    """
    if not code:
        return ""
    # 先将所有换行符统一为 \n
    code = code.replace('\r\n', '\n').replace('\r', '\n')
    # 移除末尾的空白行
    code = code.rstrip()
    # 再转换回 \r\n（Windows风格）
    return code.replace('\n', '\r\n')


def export_module(component, book_name, book_dir):
    """导出单个模块到文件"""
    comp_type = get_component_type(component.Type)
    comp_name = component.Name

    try:
        # 读取代码内容
        if component.CodeModule.CountOfLines > 0:
            code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
        else:
            code = ""

        # 规范化换行符，避免额外空行
        code = normalize_line_endings(code)

        # 确定文件扩展名
        if component.Type == 1:  # Standard Module
            ext = '.bas'
        elif component.Type == 2:  # Class Module
            ext = '.cls'
        elif component.Type == 3:  # Form
            ext = '.frm'
        else:  # Workbook/Worksheet Code
            ext = '.vba'

        # 保存到文件
        filename = sanitize_filename(comp_name) + ext
        filepath = book_dir / filename

        with open(filepath, 'w', encoding='utf-8', newline='') as f:
            f.write(f"' 模块名称: {comp_name}\r\n")
            f.write(f"' 模块类型: {comp_type}\r\n")
            f.write(f"' 来源工作簿: {book_name}\r\n")
            f.write(f"' 导出时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\r\n")
            f.write("' " + "="*50 + "\r\n")
            if code:
                f.write(code)

        return {
            'name': comp_name,
            'type': comp_type,
            'file': str(filepath)
        }

    except Exception as e:
        return {
            'name': comp_name,
            'type': comp_type,
            'file': None,
            'error': str(e)
        }


def export_all_modules(book_name, module_names=None, output_dir="vba_src"):
    """
    导出工作簿中模块的代码到指定目录
    
    参数:
        book_name: 工作簿名称
        module_names: 要导出的模块名称列表，None表示导出所有模块
        output_dir: 输出目录
    """
    app, error = get_excel_app()
    if error:
        return {"error": error}

    book, error = get_workbook(app, book_name)
    if error:
        return {"error": error}

    try:
        vb_proj = book.api.VBProject

        # 创建输出目录
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)

        # 创建工作簿专属子目录
        book_dir = output_path / sanitize_filename(book_name.replace('.xlsm', '').replace('.xls', ''))
        book_dir.mkdir(exist_ok=True)

        exported = []

        if module_names:
            # 导出指定模块
            for module_name in module_names:
                component, error = get_module(vb_proj, module_name)
                if error:
                    exported.append({
                        'name': module_name,
                        'type': None,
                        'file': None,
                        'error': error
                    })
                else:
                    result = export_module(component, book_name, book_dir)
                    exported.append(result)
        else:
            # 遍历所有组件
            for component in vb_proj.VBComponents:
                result = export_module(component, book_name, book_dir)
                exported.append(result)

        return {
            "success": True,
            "workbook": book_name,
            "output_dir": str(book_dir),
            "exported_count": len(exported),
            "modules": exported
        }

    except Exception as e:
        return {"error": f"访问 VBA 项目时出错: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) > 2:
        book_name = sys.argv[1]
        # 支持逗号分隔的模块名列表
        module_names = [name.strip() for name in sys.argv[2].split(',')]
    elif len(sys.argv) > 1:
        book_name = sys.argv[1]
        module_names = None
    else:
        book_name = input("请输入工作簿名称 (例如: 'svn跨分支合表工具.xlsm'): ")
        modules_input = input("请输入要导出的模块名称（多个用逗号分隔，留空导出所有）: ").strip()
        module_names = [name.strip() for name in modules_input.split(',')] if modules_input else None

    result = export_all_modules(book_name, module_names)
    print(json.dumps(result, ensure_ascii=False, indent=2))
