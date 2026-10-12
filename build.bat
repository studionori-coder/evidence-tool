@echo off
chcp 65001 >nul
rem EvidenceTool.exe を dist フォルダに作成する
cd /d %~dp0
python -m pip install -r requirements.txt pyinstaller || goto :err
python -m PyInstaller --noconfirm --onefile --windowed --name EvidenceTool --icon assets\icon.ico --hidden-import pystray._win32 evidence_tool.py || goto :err
echo.
echo 完了: dist\EvidenceTool.exe
pause
exit /b 0
:err
echo ビルドに失敗しました。
pause
exit /b 1
