@echo off
rem Build dist\md-converter-ui.exe with the same Python that runs the app.
rem (The bare "pyinstaller" command can belong to another Python install that
rem lacks the converter packages; the .exe then cannot convert documents.)
setlocal
cd /d "%~dp0"

echo Using:
python -c "import sys; print('  ' + sys.executable + '  (Python ' + sys.version.split()[0] + ')')" || goto :nopython

echo Installing / checking the required packages...
python -m pip install --quiet -r requirements.txt pyinstaller || goto :failed

python -m PyInstaller --noconfirm md-converter-ui.spec || goto :failed

echo.
echo Done: %~dp0dist\md-converter-ui.exe
exit /b 0

:nopython
echo Python was not found on PATH.
exit /b 1

:failed
echo.
echo The build failed; see the messages above.
exit /b 1
