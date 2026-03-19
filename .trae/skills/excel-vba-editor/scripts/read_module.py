import sys
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_module, get_component_type, read_module_code
)


def read_module(book_name, module_name):
    """读取指定模块的代码"""
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
        return {
            "success": True,
            "workbook": book_name,
            "module": module_name,
            "type": get_component_type(component.Type),
            "code": read_module_code(component)
        }
    except Exception as e:
        return {"error": f"Error reading module: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) > 2:
        book_name = sys.argv[1]
        module_name = sys.argv[2]
    else:
        book_name = input("Enter workbook name: ")
        module_name = input("Enter module name: ")

    result = read_module(book_name, module_name)
    print(json.dumps(result, ensure_ascii=False, indent=2))
