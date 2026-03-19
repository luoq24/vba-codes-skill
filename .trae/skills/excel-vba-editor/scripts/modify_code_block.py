import sys
import re
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project, get_module
)


def modify_code_block(book_name, module_name, pattern, replacement, use_regex=False):
    """修改模块中匹配特定模式的代码块"""
    app, error = get_excel_app()
    if error:
        return {"error": error}

    book, error = get_workbook(app, book_name)
    if error:
        return {"error": error}

    vb_proj, error = get_vb_project(book)
    if error:
        return {"error": error}

    component, error = get_module(vb_proj, module_name)
    if error:
        return {"error": error}

    try:
        code_module = component.CodeModule
        total_lines = code_module.CountOfLines
        full_code = code_module.Lines(1, total_lines)

        if use_regex:
            modified_code = re.sub(pattern, replacement, full_code, flags=re.DOTALL)
        else:
            modified_code = full_code.replace(pattern, replacement)

        code_module.DeleteLines(1, total_lines)
        code_module.AddFromString(modified_code)

        return {"success": True, "message": f"Successfully modified code block in module '{module_name}'"}
    except Exception as e:
        return {"error": f"Error modifying code block: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("用法:")
        print("  字符串替换: python modify_code_block.py <工作簿> <模块> <搜索模式> <替换内容>")
        print("  正则替换: python modify_code_block.py <工作簿> <模块> <模式> <替换内容> --regex")
        sys.exit(1)

    book_name = sys.argv[1]
    module_name = sys.argv[2]
    pattern = sys.argv[3]
    replacement = sys.argv[4]
    use_regex = "--regex" in sys.argv or "-r" in sys.argv

    result = modify_code_block(book_name, module_name, pattern, replacement, use_regex)
    print(json.dumps(result, ensure_ascii=False, indent=2))
