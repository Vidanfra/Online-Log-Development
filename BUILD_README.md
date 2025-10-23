# Building the Executable

This guide explains how to build a standalone executable (.exe) file for the Online Logger application.

## Requirements

- Python 3.10+ installed
- All dependencies from `Python Script/requirements.txt` installed
- PyInstaller (will be auto-installed by the build script if missing)

## Quick Build

Simply double-click `build_exe.bat` or run it from the command line:

```batch
build_exe.bat
```

The script will:
1. Check and install PyInstaller if needed
2. Clean previous build artifacts
3. Build the executable with the Online Logger icon
4. Copy all necessary files (Excel templates, guides, settings)
5. Create the folder structure for monitored folders

## Output

The executable and all required files will be in:
```
dist\Online_Logger_v2.0\
```

The main executable is:
```
dist\Online_Logger_v2.0\Online_Logger_v2.0.exe
```

## Console Window (Debugging)

By default, the console window is **enabled** to help with debugging. This means:
- You'll see a command prompt window alongside the GUI
- All print statements and errors are visible
- Useful for troubleshooting issues

### To Disable Console (Production Build)

Edit `build_exe.bat` and change:
```batch
--console ^
```
to:
```batch
--noconsole ^
```

Or use the `--windowed` flag instead.

## Icon

The application icon is set to:
```
_repositoryfiles\OnlineLoggerLogo.ico
```

The `.ico` format is recommended for Windows executables and provides the best compatibility.

## Build Options Explained

The build script uses these PyInstaller options:

- `--name="Online_Logger_v2.0"` - Name of the executable
- `--icon="_repositoryfiles\OnlineLoggerLogo.ico"` - Application icon
- `--console` - Keep console window visible (for debugging)
- `--onedir` - Create a directory with all files (not a single .exe)
- `--noconfirm` - Overwrite output directory without asking
- `--add-data "settings;settings"` - Include settings folder (embedded in _internal)
- `--add-data "_repositoryfiles;_repositoryfiles"` - Include logo/images
- `--hidden-import` - Ensure critical modules are included

**Note:** The settings folder is also copied to the root of the distribution folder for easy access to example project JSONs and configuration files.

## Troubleshooting

### PyInstaller not found
The script will automatically try to install it. If that fails:
```batch
pip install pyinstaller
```

### Missing modules in executable
Add them to the `--hidden-import` list in `build_exe.bat`:
```batch
--hidden-import=your_module_name ^
```

### Icon not showing
- Ensure the icon file exists in `_repositoryfiles\`
- Convert the logo to `.ico` format for best compatibility
- Check that the path in the script is correct

### Excel integration not working
Make sure Microsoft Excel is installed on the target machine. The executable requires Excel for xlwings to function.

## Distribution

When distributing the application:

1. Zip the entire `dist\Online_Logger_v2.0\` folder
2. Include the `README.md` and user guide
3. Ensure recipients have:
   - Windows OS
   - Microsoft Excel installed
   - Write permissions for the folders

## Advanced: Single File Executable

To create a single .exe file (slower startup but easier to distribute):

Edit `build_exe.bat` and change:
```batch
--onedir ^
```
to:
```batch
--onefile ^
```

**Note:** Single file mode extracts files to a temporary directory on each run, which can be slower.

## Clean Build

To start fresh, the script automatically removes:
- `dist/` folder
- `build/` folder  
- `.spec` files

You can also manually delete these folders before building.
