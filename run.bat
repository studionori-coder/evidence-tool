@echo off
chcp 65001 >nul
rem Python 環境でそのまま起動する（コンソールなし）
cd /d %~dp0
start "" pythonw evidence_tool.py
