#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA 编辑器公共工具模块

提供 xlwings 操作 Excel VBA 的公共函数
"""

import xlwings as xw
import os


# 组件类型映射
COMPONENT_TYPES = {
    1: 'Standard Module',
    2: 'Class Module',
    3: 'Form',
    100: 'Workbook/Worksheet Code'
}

COMPONENT_TYPES_CN = {
    1: '模块',
    2: '类模块',
    3: '窗体',
    100: '工作簿/工作表代码'
}


def get_excel_app():
    """获取活动的 Excel 应用程序实例"""
    app = xw.apps.active
    if not app:
        return None, "No active Excel application found. Please start Excel and open the .xlsm file."
    return app, None


def get_workbook(app, book_name):
    """
    获取指定名称的工作簿
    
    参数:
        app: xlwings App 实例
        book_name: 工作簿名称
    
    返回:
        (workbook, error) - 如果出错，workbook 为 None，error 为错误信息
    """
    try:
        book = app.books[book_name]
        return book, None
    except Exception as e:
        available = [b.name for b in app.books]
        return None, f"Workbook '{book_name}' not found. Available: {available}"


def get_vb_project(book):
    """
    获取工作簿的 VBProject
    
    参数:
        book: xlwings Book 实例
    
    返回:
        (vb_project, error) - 如果出错，vb_project 为 None
    """
    try:
        return book.api.VBProject, None
    except Exception as e:
        return None, f"Error accessing VBA project: {str(e)}"


def get_component_type(comp_type, use_chinese=False):
    """
    获取组件类型的名称
    
    参数:
        comp_type: 组件类型代码 (1, 2, 3, 100)
        use_chinese: 是否返回中文名称
    """
    types = COMPONENT_TYPES_CN if use_chinese else COMPONENT_TYPES
    return types.get(comp_type, 'Unknown')


def read_code_from_file(code_input):
    """
    从文件读取代码（支持 FILE: 前缀）
    
    参数:
        code_input: 代码字符串或 FILE:filepath
    
    返回:
        (code, error) - 如果出错，code 为 None
    """
    if code_input.startswith("FILE:"):
        filename = code_input[5:]
        if not os.path.exists(filename):
            return None, f"File not found: {filename}"
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                return f.read(), None
        except Exception as e:
            return None, f"Error reading file: {str(e)}"
    return code_input, None


def get_module(vb_proj, module_name):
    """
    获取指定模块
    
    参数:
        vb_proj: VBProject 对象
        module_name: 模块名称
    
    返回:
        (component, error) - 如果出错，component 为 None
    """
    try:
        return vb_proj.VBComponents(module_name), None
    except Exception as e:
        return None, f"Module '{module_name}' not found"


def read_module_code(component):
    """
    读取模块代码
    
    参数:
        component: VBComponent 对象
    
    返回:
        代码字符串
    """
    try:
        if component.CodeModule.CountOfLines > 0:
            return component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
        return ""
    except Exception as e:
        return f"[Error reading code: {str(e)}]"


def write_module_code(component, code):
    """
    写入模块代码（先清空再写入）
    
    参数:
        component: VBComponent 对象
        code: 代码字符串
    
    返回:
        error - 如果成功返回 None
    """
    try:
        code_module = component.CodeModule
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        if code.strip():
            # 确保使用 Windows 风格的换行符 (\r\n)，避免产生额外空行
            code = code.replace('\r\n', '\n').replace('\n', '\r\n')
            code_module.AddFromString(code)
        return None
    except Exception as e:
        return f"Error writing code: {str(e)}"


def sanitize_filename(name):
    """清理文件名，移除非法字符"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name
