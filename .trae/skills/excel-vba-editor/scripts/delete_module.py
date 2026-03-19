import sys
import json
from vba_utils import get_excel_app, get_workbook, get_vb_project


def delete_module(book_name, module_name):
    """从工作簿删除指定模块"""
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
        vb_proj.VBComponents.Remove(vb_proj.VBComponents(module_name))
        return {"success": True, "message": f"Successfully deleted module: {module_name}"}
    except Exception as e:
        return {"error": f"Error deleting module: {str(e)}"}


if __name__ == "__main__":
    if len(sys.argv) > 2:
        book_name = sys.argv[1]
        module_name = sys.argv[2]
    else:
        book_name = input("Enter workbook name: ")
        module_name = input("Enter module name to delete: ")

    result = delete_module(book_name, module_name)
    print(json.dumps(result, ensure_ascii=False, indent=2))
