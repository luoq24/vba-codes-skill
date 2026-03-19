@echo off
chcp 65001 >nul
title VBA Git 自动提交工具

:: VBA Git 自动提交脚本
:: 工作流: 分析修改范围 -> 导出模块 -> Git commit -> AI修改 -> 写回Excel
:: 用法: git_auto_commit.bat [工作簿名称] [提交信息]
:: 示例: git_auto_commit.bat "svn跨分支合表工具.xlsm" "导出原始代码"
::
:: 注意：此脚本执行第3步（Git commit），前提是已完成：
::   第1步: 分析任务，确定需要修改的模块
::   第2步: 使用 export_vba.py 导出模块到 vba_src/

setlocal enabledelayedexpansion

:: 设置 Python 解释器路径（使用虚拟环境）
set "PYTHON_EXE=e:\python_space\.venv_work\Scripts\python.exe"
set "SCRIPTS_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPTS_DIR%..\..\..\.."

:: 解析参数
if "%~1"=="" (
    set /p BOOK_NAME="请输入工作簿名称 (例如: svn跨分支合表工具.xlsm): "
) else (
    set "BOOK_NAME=%~1"
)

if "%~2"=="" (
    set "COMMIT_MSG=自动导出 VBA 代码 - %date% %time%"
) else (
    set "COMMIT_MSG=%~2"
)

echo ========================================
echo    VBA Git 自动提交工具
echo ========================================
echo.
echo 工作簿: %BOOK_NAME%
echo 提交信息: %COMMIT_MSG%
echo.

:: 检查 Python 解释器
if not exist "%PYTHON_EXE%" (
    echo [错误] 未找到 Python 解释器: %PYTHON_EXE%
    echo 请检查虚拟环境路径是否正确。
    pause
    exit /b 1
)

:: 检查 export_vba.py 是否存在
if not exist "%SCRIPTS_DIR%export_vba.py" (
    echo [错误] 未找到 export_vba.py 脚本
    pause
    exit /b 1
)

echo [1/3] 正在导出 VBA 代码...
echo.

:: 导出 VBA 代码
cd /d "%PROJECT_ROOT%"
"%PYTHON_EXE%" "%SCRIPTS_DIR%export_vba.py" "%BOOK_NAME%"

if %errorlevel% neq 0 (
    echo.
    echo [错误] 导出 VBA 代码失败
    pause
    exit /b 1
)

echo.
echo [2/3] 正在添加到 Git...
echo.

:: 检查是否是 Git 仓库
if not exist "%PROJECT_ROOT%\.git" (
    echo [初始化] 创建 Git 仓库...
    cd /d "%PROJECT_ROOT%"
    git init
    echo.
)

:: 添加所有 vba_src/ 目录下的更改
cd /d "%PROJECT_ROOT%"
git add vba_src/

echo.
echo [3/3] 正在提交...
echo.

:: 提交更改
git commit -m "%COMMIT_MSG%"

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    提交成功！
    echo ========================================
    echo.
    echo 已导出并提交 VBA 代码。
    echo 现在可以让 AI 修改代码了。
    echo.
    echo 查看修改:
    echo   git diff HEAD~1    (查看上次提交的差异)
    echo   git log --oneline  (查看提交历史)
    echo.
) else (
    echo.
    echo [提示] 没有新的更改需要提交
    echo.
)

pause
