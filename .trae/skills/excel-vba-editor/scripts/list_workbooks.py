import json
from vba_utils import get_excel_app


def list_workbooks():
    """列出所有已打开的 Excel 工作簿"""
    app, error = get_excel_app()
    if error:
        return {"error": error}

    books = [book.name for book in app.books]
    return {"success": True, "workbooks": books}


if __name__ == "__main__":
    result = list_workbooks()
    print(json.dumps(result, ensure_ascii=False, indent=2))
