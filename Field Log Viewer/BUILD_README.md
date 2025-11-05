# Field Log Viewer v2.0 - Build Instructions

## Quick Build

### Simple Way:
```
Double-click: build_Field_Log_Viewer.bat
```

This will:
1. Check for PyInstaller (install if needed)
2. Clean previous builds
3. Create a single `.exe` file with settings included
4. Generate a deployment-ready folder

### Build Time:
- First build: ~2-3 minutes (includes dependency scanning)
- Subsequent builds: ~1-2 minutes

---

## What Gets Created

After running the build script:

```
Field Log Viewer/
├── dist/
│   └── Field_Log_Viewer_v2.0.exe  <-- Main executable
├── %APP_NAME%/    <-- Ready to deploy!
│   ├── Field_Log_Viewer_v2.0.exe
│   └── flv_settings/
│       └── (default settings structure)
└── build/  (temporary build files - can delete)
```

---

## Deployment Package

The `%APP_NAME%` folder is ready to zip and send:

### What to Send:
```
%APP_NAME%.zip
├── Field_Log_Viewer_v2.0.exe
└── flv_settings/
```

### Instructions for Colleagues:

1. **Extract the ZIP file**
   - Extract to any local folder (C:\ or D:\)
   - ⚠️ DO NOT extract to network drive

2. **Run the Executable**
   - Double-click `Field_Log_Viewer_v2.0.exe`
   - No Python installation required!

3. **First Use**
   - Browse to select SQLite database
   - Select table from dropdown
   - Choose target Excel file
   - Configure keywords/colors as needed
   - Settings auto-save in `flv_settings` folder

---

## File Sizes

Typical sizes:
- **Executable:** ~60-80 MB (includes Python runtime + all dependencies)
- **Settings folder:** <1 KB
- **Total deployment:** ~60-80 MB

---

## Requirements

### For Building (Your Machine):
- Python 3.13+
- PyInstaller (auto-installed by script if missing)
- Required packages: tkinter, sqlite3, xlwings, threading, json, os, time

### For Running (Colleague's Machine):
- Windows 10/11
- Excel (for xlwings integration)
- **No Python needed!** (included in .exe)

---

## Build Options Explained

The build script uses these PyInstaller options:

| Option | Purpose |
|--------|---------|
| `--onefile` | Creates single .exe (not a folder) |
| `--windowed` | No console window (GUI only) |
| `--name` | Output filename |
| `--add-data` | Include settings folder in exe |
| `--exclude-module` | Remove unused packages (faster build, smaller size) |

---

## Troubleshooting Build

### "PyInstaller not found"
**Solution:** Script will auto-install. If that fails:
```powershell
pip install pyinstaller
```

### Build is very slow (>5 minutes)
**Solution:** This is normal for first build. PyInstaller scans all dependencies. Subsequent builds are faster.

### Executable is too large (>150 MB)
**Solution:** The script already excludes unnecessary modules (matplotlib, PIL, plotly). Size is expected to be ~60-80 MB.

### "Failed to execute script"
**Solution:** Run the build script again - sometimes first build has temporary issues.

---

## Testing the Executable

After building:

1. **Navigate to deployment folder:**
   ```
   cd %APP_NAME%
   ```

2. **Run the executable:**
   - Double-click `Field_Log_Viewer_v2.0.exe`
   - Should open without errors

3. **Test basic functionality:**
   - Browse to a database file
   - Verify table dropdown works
   - Check settings save correctly

---

## Version Updates

To build a new version:

1. **Edit the build script:**
   ```bat
   SET VERSION=2.1  REM Change this line
   ```

2. **Run the build:**
   ```
   build_Field_Log_Viewer.bat
   ```

3. **Output will be:**
   ```
   Field_Log_Viewer_v2.1.exe
   ```

---

## Clean Build

If you need to force a completely clean build:

```powershell
# Delete build artifacts
rmdir /s /q build
rmdir /s /q dist
rmdir /s /q %APP_NAME%
del Field_Log_Viewer_v*.spec

# Run build again
build_Field_Log_Viewer.bat
```

---

## Notes

- **Portable:** The .exe can be moved anywhere (settings folder moves with it)
- **Self-contained:** All dependencies included (except Excel)
- **No installation:** Just extract and run
- **Settings persist:** Saved in `flv_settings` next to executable

---

## Support

If build fails:
1. Check Python version: `python --version` (should be 3.13+)
2. Update PyInstaller: `pip install --upgrade pyinstaller`
3. Check for error messages in build output
4. Try clean build (see above)
