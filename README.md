# Shortcuts Dock

A lightweight Windows dock mainly for program shortcuts, for cleaning up the taskbar for more important things, such as opened programs.

## Changelog
v1.0.0.0
November 10, 2026
- Initial release.

## Requirements

- Windows 10 or Windows 11

## Installation

1. Download and extract the release ZIP file to a folder, such as "C:\Program Files\Shortcuts Dock".
2. Launch the program and right click on its tray icon to learn of the varies settings.

## Usage

- Drag files and folders onto the dock to make shortcuts to them.
- Right-click shortcuts on the dock to either open, open program file location, rename or to remove them from the dock.
- Right click on the dock and select "Add Windows app..." to add a Microsoft Store app shortcut.
- Right-click on the dock and select "Add separator" to separate icons from one another.
- Right-click on a separator and select "Remove separator" to remove it.
- Drag shortcuts and separators in the dock amongst each other to change their orders.
- Drag and drop a shortcut from the dock to the desktop/another folder to remove it from the dock and to make a copy of it there.
- Use mouse wheel when dock edges are beyond the screen width/height for scrolling to offset portions of the dock.


## Author

Gal Shemesh (PacketLauncher)
https://github.com/packetlauncher


The following info is for developers who wish to make further adjustments to the program for their own personal use:

1. Install Python if it's not already installed.
2. Open CMD window in the at the py script directory and install the following libraries if you don't have them:

python -m pip install PySide6 pywin32
python -m pip install pyinstaller

3. Test the Python script either by running the "TestMe.bat" script, or from CMD by running:
python "Shortcuts_Dock.py"

4. Before creating the EXE, you may want to update its release number to showing in both the EXE details and in the program's About section. To doing so, edit the version_info.txt file and make sure to update ALL 4 lines accordingly before compiling the EXE:
	- filevers=(1, 0, 0, 0),
    - prodvers=(1, 0, 0, 0)
	- StringStruct('FileVersion', '1.0.0.0'),
	- StringStruct('ProductVersion', '1.0.0.0')

5. Create the EXE either by running the "BuildMe.bat" script, or from CMD by running:
python -m PyInstaller --onefile --windowed --clean --name "Shortcuts Dock" --icon "Shortcuts_Dock.ico" --version-file "version_info.txt" "Shortcuts_Dock.py"

4. Once you're done, you may use the Build_Leftovers_CleanUp.bat script for cleaning up any Python leftovers afterwards.