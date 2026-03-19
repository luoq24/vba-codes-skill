#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
列出工作簿中所有 VBA 模块及其代码

工作流步骤:
    第1步: 分析本次任务会修改哪些VBA代码(模块、类模块)
    （用于查看所有模块，确定需要修改的范围）

用法:
    python .trae/skills/excel-vba-editor/scripts/list_modules.py "工作簿名称.xlsm"
    python .trae/skills/excel-vba-editor/scripts/list_modules.py  # 交互式输入

完整工作流:
    1. 分析本次任务会修改哪些VBA代码(模块、类模块)  <- 本脚本
    2. 导出这些模块的代码到 vba_src/
    3. Git commit
    4. AI修改代码
    5. 写回Excel
"""

import sys
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_component_type, read_module_code
)


def list_modules(book_name):
    """读取工作簿中所有模块的代码"""
    app, error = get_excel_app()
    if error:
        return {"error": error}

    book, error = get_workbook(app, book_name)
    if error:
        return {"error": error}

    vb_proj, error = get_vb_project(book)
    if error:
        return {"error": error}

    try:
        modules = []
        for component in vb_proj.VBComponents:
            modules.append({
                'name': component.Name,
                'type': get_component_type(component.Type),
                'code': read_module_code(component)
            })
        return {"success": True, "workbook": book_name, "modules": modules}
    except Exception as e:
        return {"error": f"Error accessing VBA project: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) > 1:
        book_name = sys.argv[1]
    else:
        book_name = input("Enter workbook name (e.g., 'Book1.xlsm'): ")

    result = list_modules(book_name)
    print(json.dumps(result, ensure_ascii=False, indent=2))
