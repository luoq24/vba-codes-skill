import sys
from vba_utils import (
    get_excel_app, get_workbook, get_vb_project, get_component_type
)


def list_module_names(book_name):
    """列出工作簿中所有模块的名称"""
    app, error = get_excel_app()
    if error:
        print(f"Error: {error}")
        return

    book, error = get_workbook(app, book_name)
    if error:
        print(f"Error: {error}")
        return

    vb_proj, error = get_vb_project(book)
    if error:
        print(f"Error: {error}")
        return

    try:
        categories = {
            'Standard Module': [],
            'Class Module': [],
            'Form': [],
            'Workbook/Worksheet Code': []
        }

        for component in vb_proj.VBComponents:
            comp_type = get_component_type(component.Type)
            if comp_type in categories:
                categories[comp_type].append(component.Name)

        print(f"\n=== VBA Modules in '{book_name}' ===\n")

        type_names = {
            'Standard Module': '【标准模块】',
            'Class Module': '【类模块】',
            'Form': '【窗体】',
            'Workbook/Worksheet Code': '【工作表代码】'
        }

        total = 0
        for eng_name, chn_name in type_names.items():
            modules = categories[eng_name]
            if modules:
                print(chn_name)
                for i, name in enumerate(modules, 1):
                    print(f"  {i}. {name}")
                print()
                total += len(modules)

        print(f"总计: {total} 个模块")

    except Exception as e:
        print(f"Error accessing VBA project: {str(e)}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        book_name = sys.argv[1]
    else:
        book_name = input("Enter workbook name: ")

    list_module_names(book_name)
