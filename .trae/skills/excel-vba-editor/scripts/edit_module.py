#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接在Excel工作簿中编辑VBA模块代码

这是推荐的工作流脚本，用于直接在Excel中修改代码，而不是编辑导出的文件。

用法:
    # 替换整个模块
    python .trae/skills/excel-vba-editor/scripts/edit_module.py "workbook.xlsm" "ModuleName" "FILE:new_code.bas"
    
    # 从字符串替换
    python .trae/skills/excel-vba-editor/scripts/edit_module.py "workbook.xlsm" "ModuleName" "Public Sub Hello(): MsgBox \"Hi\": End Sub"

工作流:
    1. 导出代码（作为Git对比基线）
    2. Git commit
    3. 使用本脚本直接在Excel中编辑
    4. 再次导出，查看修改差异
"""

import sys
import json
from pathlib import Path

from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_module, read_code_from_file, write_module_code
)


def edit_module(book_name, module_name, new_code):
    """
    直接在Excel工作簿中编辑VBA模块代码
    
    参数:
        book_name: 工作簿名称（如 "workbook.xlsm"）
        module_name: 模块名称（如 "Module1"）
        new_code: 新代码，支持 FILE:前缀从文件读取
    
    返回:
        操作结果字典
    """
    # 获取Excel应用
    app, error = get_excel_app()
    if error:
        return {"error": error}

    # 获取工作簿
    book, error = get_workbook(app, book_name)
    if error:
        return {"error": error}

    # 获取VBA项目
    vb_proj, error = get_vb_project(book)
    if error:
        return {"error": error}

    # 获取模块
    component, error = get_module(vb_proj, module_name)
    if error:
        return {"error": error}

    # 读取新代码（支持FILE:前缀）
    new_code, error = read_code_from_file(new_code)
    if error:
        return {"error": error}

    # 写入新代码
    error = write_module_code(component, new_code)
    if error:
        return {"error": error}

    return {
        "success": True,
        "message": f"成功修改模块: {module_name}",
        "workbook": book_name,
        "module": module_name,
        "code_lines": len(new_code.split('\n')) if new_code else 0
    }


if __name__ == "__main__":
    if len(sys.argv) > 3:
        book_name = sys.argv[1]
        module_name = sys.argv[2]
        new_code = sys.argv[3]
    else:
        print("用法: python edit_module.py \"workbook.xlsm\" \"ModuleName\" \"FILE:code.bas 或代码字符串\"")
        print()
        print("示例:")
        print('  python edit_module.py "svn跨分支合表工具.xlsm" "MCompareTool" "FILE:MCompareTool.bas"')
        print('  python edit_module.py "workbook.xlsm" "Module1" "Public Sub Test(): MsgBox \"Hello\": End Sub"')
        sys.exit(1)

    result = edit_module(book_name, module_name, new_code)
    print(json.dumps(result, ensure_ascii=False, indent=2))
