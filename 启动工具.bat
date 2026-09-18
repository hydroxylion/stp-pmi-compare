@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 PMI 数据一致性比对工具...
".venv\Scripts\python.exe" -m streamlit run pmi_compare_work.py
pause
