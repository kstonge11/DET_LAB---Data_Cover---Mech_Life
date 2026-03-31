# Building DET Lab Cover Sheet Tool for Windows

This guide explains how to create a standalone Windows executable (.exe) from the DET Lab Cover Sheet Tool.

## Prerequisites

1. **Windows 10/11** (64-bit)
2. **Python 3.11+** installed and in PATH
3. **Microsoft Excel** installed (required for the printing feature at runtime)

## Quick Build (Recommended)

The easiest way to build is using the provided batch script:

```cmd
build.bat
```

This will:
- Create a virtual environment
- Install all dependencies
- Build the executable

The output will be in `dist\DET_Lab_CoverSheet\`.

### Clean Build

To do a fresh build (removes previous build artifacts):

```cmd
build.bat clean
```

## Manual Build

If you prefer to build manually:

### 1. Create Virtual Environment

```cmd
python -m venv venv
venv\Scripts\activate
```

### 2. Install Dependencies

```cmd
pip install -r requirements.txt
```

### 3. Build Executable

**Option A: Folder Distribution** (faster startup, larger folder)
```cmd
pyinstaller DETLAB.spec --noconfirm
```

**Option B: Single File** (slower startup, easier to share)
```cmd
pyinstaller DETLAB_onefile.spec --noconfirm
```

## Output

After building:

| Build Type | Output Location | Size (approx) |
|------------|-----------------|---------------|
| Folder | `dist\DET_Lab_CoverSheet\` | ~150-200 MB |
| One File | `dist\DET_Lab_CoverSheet.exe` | ~80-100 MB |

## Distribution

### Folder Distribution
Copy the entire `dist\DET_Lab_CoverSheet\` folder to the target machine. Users run `DET_Lab_CoverSheet.exe` from within the folder.

### Single File Distribution
Copy just `dist\DET_Lab_CoverSheet.exe`. Note: First launch will be slower as it extracts to a temp folder.

## Runtime Requirements

On the target machine:
- **Microsoft Excel** must be installed for cover sheet printing
- No Python installation required
- No additional DLLs needed (all bundled)

## Troubleshooting

### Build Errors

**"ModuleNotFoundError: No module named 'PyQt6'"**
```cmd
pip install PyQt6
```

**"ModuleNotFoundError: No module named 'win32com'"**
```cmd
pip install pywin32
```

After installing pywin32, run:
```cmd
python -m pywin32_postinstall -install
```

### Runtime Errors

**App won't start / black screen**
Build with console enabled for debugging:
1. Edit the `.spec` file
2. Change `console=False` to `console=True`
3. Rebuild

**"Excel COM print error"**
- Ensure Microsoft Excel is installed
- Try running as Administrator

**Missing qdarktheme styles**
The spec file should include qdarktheme data files. If styles are missing, verify:
```python
qdarktheme_datas = collect_data_files('qdarktheme')
```
is in the spec file and `datas=qdarktheme_datas` is in the Analysis section.

## Adding an Icon

1. Create or obtain a `.ico` file
2. Place it in the project directory as `icon.ico`
3. Uncomment the `icon='icon.ico'` line in the spec file
4. Rebuild

## Version Info (Optional)

To add Windows version info, install:
```cmd
pip install pyinstaller-versionfile
```

Then create `version.yml` and reference it in the spec file.

---

Built with PyInstaller 6.x for Python 3.11+
