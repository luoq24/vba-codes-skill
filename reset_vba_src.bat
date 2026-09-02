@echo off
setlocal
title vba_src 子仓库重置工具

set "ROOT=%~dp0"
set "VBA_SRC=%ROOT%vba_src"
set "PLACEHOLDER=%VBA_SRC%\vba_codes_will_export_here.txt"

:: 静默模式: reset_vba_src.bat auto   跳过确认提示与暂停
set "SILENT=%~1"

echo ============================================
echo     vba_src 子仓库重置工具
echo ============================================
echo.
echo  vba_src 是独立的本地 Git 子仓库（无远程），
echo  仅用于存放导出的 VBA 代码做差异对比，用完即弃。
echo  本工具会删除整个 vba_src 目录，并重新初始化
echo  为一个全新的本地 Git 子仓库。
echo.

if /i "%SILENT%"=="auto" goto :run

set /p CONFIRM="确认清空并重置 vba_src? [Y/N]: "
if /i not "%CONFIRM%"=="Y" goto :cancelled
goto :run

:cancelled
echo 已取消，未做任何更改。
pause
exit /b 0

:run
cd /d "%ROOT%"

echo [1/3] 正在删除 vba_src ...
if exist "%VBA_SRC%" rd /s /q "%VBA_SRC%"
if exist "%VBA_SRC%" goto :delete_failed

echo [2/3] 正在初始化全新 Git 子仓库 ...
mkdir "%VBA_SRC%"
if errorlevel 1 goto :mkdir_failed
cd /d "%VBA_SRC%"
git init -b main >nul 2>&1
if errorlevel 1 goto :init_failed

:: 创建空占位文件，保证仓库有初始提交基线
type nul > "%PLACEHOLDER%"

echo [3/3] 正在建立初始提交 ...
git add .
git commit -m "init: reset vba_src" >nul 2>&1
if errorlevel 1 goto :commit_failed

echo.
echo 完成：vba_src 已重置为全新本地子仓库，main 分支，1 个初始提交。
echo 后续 VBA 任务的临时提交只会发生在这里，不会进入主仓库，
echo 也不会推送到远程仓库。
echo.
if /i not "%SILENT%"=="auto" pause
exit /b 0

:delete_failed
echo [错误] 删除 vba_src 失败，可能被其他程序占用，请关闭后重试。
pause
exit /b 1

:mkdir_failed
echo [错误] 创建目录失败。
pause
exit /b 1

:init_failed
echo [错误] git init 失败，请确认 Git 已安装并加入 PATH。
pause
exit /b 1

:commit_failed
echo [警告] 初始提交失败，请先配置 Git 用户信息：
echo         git config --global user.name  "Your Name"
echo         git config --global user.email "you@example.com"
echo.
if /i not "%SILENT%"=="auto" pause
exit /b 1
