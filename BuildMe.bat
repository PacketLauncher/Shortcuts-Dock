@echo off

python -m PyInstaller --onefile --windowed --clean --name "Shortcuts Dock" --icon "Shortcuts_Dock.ico" --version-file "version_info.txt" "Shortcuts_Dock.py"

if exist "Shortcuts Dock.exe" del "Shortcuts Dock.exe"

move "dist\Shortcuts Dock.exe"



