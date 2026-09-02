@echo off
setlocal
title VBA Git 自动提交工具（本地 vba_src 子仓库）

:: 用途：导出 Excel 中指定工作簿的 VBA 代码到 vba_src 子仓库并提交。
::
:: vba_src 是独立的本地 Git 子仓库（无远程仓库），提交只发生在本地，
:: 不会进入主仓库，也不会推送到远程。任务结束后可运行
:: reset_vba_src.bat 一键清空重置。
::
:: 用法: git_auto_commit.bat [工作簿名称] [提交信息]
:: 示例: git_auto_commit.bat "svn跨分支合表工具.xlsm" "导出原始代码"
::
:: 前提: 已完成第2步 export_vba.py 导出模块到 vba_src/

set "PYTHON_EXE=e:\python_space\.venv_work\Scripts\python.exe"
set "SCRIPTS_DIR=%~dp0"
set "PROJECT_ROOT=%SCRIPTS_DIR%..\..\..\.."
set "VBA_SRC=%PROJECT_ROOT%\vba_src"

:: 解析参数：工作簿名称
if not "%~1"=="" goto :book_ok
set /p BOOK_NAME="请输入工作簿名称 (例如: svn跨分支合表工具.xlsm): "
goto :parse_msg

:book_ok
set "BOOK_NAME=%~1"

:parse_msg
:: 解析参数：提交信息
if not "%~2"=="" goto :msg_set
set "COMMIT_MSG=自动导出 VBA 代码 - %date% %time%"
goto :begin

:msg_set
set "COMMIT_MSG=%~2"

:begin
echo ========================================
echo   VBA Git 自动提交工具（本地 vba_src 子仓库）
echo ========================================
echo.
echo 工作簿: %BOOK_NAME%
echo 提交信息: %COMMIT_MSG%
echo.

:: 检查 Python 解释器
if exist "%PYTHON_EXE%" goto :check_script
echo [错误] 未找到 Python 解释器: %PYTHON_EXE%
echo 请检查虚拟环境路径是否正确。
pause
exit /b 1

:check_script
:: 检查 export_vba.py 是否存在
if exist "%SCRIPTS_DIR%export_vba.py" goto :do_export
echo [错误] 未找到 export_vba.py 脚本
pause
exit /b 1

:do_export
echo [1/3] 正在导出 VBA 代码...
echo.
cd /d "%PROJECT_ROOT%"
"%PYTHON_EXE%" "%SCRIPTS_DIR%export_vba.py" "%BOOK_NAME%"
if %errorlevel% equ 0 goto :prepare_repo
echo.
echo [错误] 导出 VBA 代码失败
pause
exit /b 1

:prepare_repo
echo.
echo [2/3] 检查 vba_src 子仓库...
:: vba_src 尚未初始化为 Git 仓库时自动初始化（本地仓库，无远程）
if exist "%VBA_SRC%\.git" goto :do_commit
if not exist "%VBA_SRC%" mkdir "%VBA_SRC%"
cd /d "%VBA_SRC%"
git init -b main >nul 2>&1
if %errorlevel% neq 0 goto :init_failed
if not exist "%VBA_SRC%\vba_codes_will_export_here.txt" type nul > "%VBA_SRC%\vba_codes_will_export_here.txt"
git add .
git commit -m "init: vba_src local sub repo" >nul 2>&1
echo   已自动初始化 vba_src 子仓库。

:do_commit
echo.
echo [3/3] 正在提交到 vba_src 子仓库...
echo.
cd /d "%VBA_SRC%"
git add .
git commit -m "%COMMIT_MSG%"
if %errorlevel% equ 0 goto :commit_ok
echo.
echo [提示] 没有新的更改需要提交。
goto :finish

:commit_ok
echo.
echo ========================================
echo    提交成功！
echo ========================================
echo.
echo 已导出并提交到本地 vba_src 子仓库（无远程仓库）。
echo 现在可以让 AI 修改代码了。
echo.
echo 查看修改（需先 cd 进入 vba_src 目录）:
echo   git diff HEAD~1    (查看上次提交的差异)
echo   git log --oneline  (查看提交历史)
echo.

:finish
echo 提示: 本次任务结束后，运行 reset_vba_src.bat 可一键清空重置 vba_src。
echo.
pause
exit /b 0

:init_failed
echo [错误] vba_src 子仓库初始化失败，请确认 Git 已安装并加入 PATH。
pause
exit /b 1
