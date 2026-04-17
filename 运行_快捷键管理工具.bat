@echo off
chcp 65001 >nul
title AutoCAD 快捷键管理工具 v2.3
cd /d "%~dp0"
echo ==========================================
echo    AutoCAD 快捷键管理工具 v2.3
echo    开源项目 ^|^| 作者: jjyy7783
echo    支持 2004-2026 全版本 ^|^| 用户配置系统
echo ==========================================
echo.
python autocad_shortcut_manager.py
pause