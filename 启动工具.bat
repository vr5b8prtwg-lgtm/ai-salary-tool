@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 AI 标注员工薪酬结算工具...
echo 启动完成后请在浏览器打开： http://127.0.0.1:5000
"C:\Users\24620\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" app.py
pause