import sys
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project,
    get_module, read_code_from_file, write_module_code
)


def modify_module(book_name, module_name, new_code):
    """替换指定模块的全部代码"""
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

    error = write_module_code(component, new_code)
    if error:
        return {"error": error}

    return {"success": True, "message": f"Successfully modified module: {module_name}"}


if __name__ == "__main__":
    if len(sys.argv) > 3:
        book_name = sys.argv[1]
        module_name = sys.argv[2]
        new_code = sys.argv[3]
    else:
        book_name = input("Enter workbook name: ")
        module_name = input("Enter module name: ")
        new_code = input("Enter new code (or type 'FILE:filename'): ")

    result = modify_module(book_name, module_name, new_code)
    print(json.dumps(result, ensure_ascii=False, indent=2))
