import sys
import re
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_module, read_code_from_file
)


def modify_method(book_name, module_name, method_name, new_code, method_type="Sub"):
    """修改模块中特定方法的代码"""
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

    new_code, error = read_code_from_file(new_code)
    if error:
        return {"error": error}

    try:
        code_module = component.CodeModule
        total_lines = code_module.CountOfLines
        full_code = code_module.Lines(1, total_lines)

        pattern = rf'(Public|Private|Friend)?\s*{method_type}\s+{method_name}\s*\([^)]*\)[^{{]*?End {method_type}'
        modified_code = re.sub(pattern, new_code, full_code, flags=re.DOTALL)

        if modified_code == full_code:
            return {"error": f"Method '{method_name}' not found in module '{module_name}'"}

        code_module.DeleteLines(1, total_lines)
        code_module.AddFromString(modified_code)

        return {"success": True, "message": f"Successfully modified method '{method_name}'"}
    except Exception as e:
        return {"error": f"Error modifying method: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("用法: python modify_method.py <工作簿> <模块> <方法名> <新代码> [--type=Function]")
        sys.exit(1)

    book_name = sys.argv[1]
    module_name = sys.argv[2]
    method_name = sys.argv[3]
    new_code = sys.argv[4]

    method_type = "Sub"
    for arg in sys.argv[5:]:
        if arg.startswith("--type="):
            method_type = arg[7:]

    result = modify_method(book_name, module_name, method_name, new_code, method_type)
    print(json.dumps(result, ensure_ascii=False, indent=2))
