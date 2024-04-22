pyinstaller --onefile .\version_xlsx.py
del "*.spec"
del "*.exe"
rd /s /q "build"
copy .\dist\version_xlsx.exe version_xlsx.exe
rd /s /q "dist"
pause