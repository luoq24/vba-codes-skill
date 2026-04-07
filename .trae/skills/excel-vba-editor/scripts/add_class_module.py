import sys
import json
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project, read_code_from_file
)


def add_class_module(book_name, module_name, code):
    """向工作簿添加新类模块"""
    app, error = get_excel_app()
    if error:
        return {"error": error}

    book, error = get_workbook(app, book_name)
    if error:
        return {"error": error}

    code, error = read_code_from_file(code)
    if error:
        return {"error": error}

    vb_proj, error = get_vb_project(book)
    if error:
        return {"error": error}

    try:
        vb_proj.VBComponents.Add(2)  # 添加类模块 (vbext_ct_ClassModule)
        vb_proj.VBComponents(vb_proj.VBComponents.Count).Name = module_name
        vb_proj.VBComponents(module_name).CodeModule.AddFromString(code)
        return {"success": True, "message": f"Successfully added class module: {module_name}"}
    except Exception as e:
        return {"error": f"Error adding class module: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) > 3:
        book_name = sys.argv[1]
        module_name = sys.argv[2]
        code = sys.argv[3]
    else:
        book_name = input("Enter workbook name: ")
        module_name = input("Enter new class module name: ")
        code = input("Enter code (or type 'FILE:filename' to read from file): ")

    result = add_class_module(book_name, module_name, code)
    print(json.dumps(result, ensure_ascii=False, indent=2))