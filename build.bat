@echo off
chcp 65001 >nul
echo ========================================
echo   PDF 工具箱 — 打包脚本
echo ========================================

if not exist "dist" mkdir dist

pyinstaller --onefile --windowed --name "PDF工具箱" --clean main.py

echo.
echo 打包完成: dist\PDF工具箱.exe
pause
