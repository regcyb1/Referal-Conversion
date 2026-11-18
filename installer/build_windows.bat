@echo off
rem Build "Assessment PDF to Excel.exe" (single portable file).
rem Usage (from anywhere, on Windows):  installer\build_windows.bat
setlocal
cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv || python -m venv .venv || goto :error
)
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip || goto :error
".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt -r installer\requirements-build.txt || goto :error

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
".venv\Scripts\python.exe" -m PyInstaller installer\converter.spec --noconfirm --clean || goto :error

echo Built: dist\Assessment PDF to Excel.exe
exit /b 0

:error
echo Build failed.
exit /b 1
