@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo.
echo ============================================================
echo   PMI 比对诊断（比对结果不对劲时先跑这个）
echo ============================================================
echo.

set "A=%~1"
set "B=%~2"

rem 允许两个参数顺序颠倒：自动把 .md 认成开发侧文件
if /i "%A:~-4%"==".md" (
    set "MD=%A%"
    set "XLSX=%B%"
) else (
    set "XLSX=%A%"
    set "MD=%B%"
)

if "%XLSX%"=="" (
    echo 用法（任选一种）：
    echo   1. 把 SFA 报告的 xlsx 文件拖到本 bat 上
    echo   2. 把 xlsx 和 markdown 两个文件一起选中，拖到本 bat 上
    echo   3. 直接在下面粘贴路径
    echo.
    set /p "XLSX=请输入 SFA 报告 xlsx 路径："
)

if "%XLSX%"=="" (
    echo.
    echo 没有拿到 xlsx 路径，已退出。
    pause
    exit /b 1
)

echo   SFA 报告   : %XLSX%
if not "%MD%"=="" echo   开发 markdown: %MD%
echo.
echo 正在诊断，请稍候...
echo.

if "%MD%"=="" (
    "%PY%" doctor.py "%XLSX%"
) else (
    "%PY%" doctor.py "%XLSX%" "%MD%"
)

echo.
echo ------------------------------------------------------------
echo 诊断完成。完整报告已写入本目录下的 doctor_report.txt
echo 把这个文件发给维护者即可，不需要再补充其他信息。
echo ------------------------------------------------------------
pause
