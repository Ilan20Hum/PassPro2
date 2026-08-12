# PassPro2

Desktop password manager for Windows — save and search names, passwords, emails, and remarks locally on your computer.

**Passwords stay only in `assets/data2.xlsx` on your PC. That file is gitignored and is never pushed to GitHub.**

## Setup (development)

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

First run needs a local workbook (gitignored):

```bat
copy assets\data2.template.xlsx assets\data2.xlsx
```

Or let the app create it automatically on first launch.

## Run (dev)

```bat
python main.py
```

## Fix a broken EXE / rebuild (recommended)

If you see **"Error loading Python DLL / python311.dll"**:

1. You were likely opening an old/incomplete build under `build\...` (or a OneDrive-synced copy).
2. Double-click **`build.bat`** in the project folder.
3. Run only: **`dist\PassPro2\PassPro2.exe`**
4. If the project lives under OneDrive, right-click the folder → **Always keep on this device**, or copy `dist\PassPro2` to a short local path like `C:\PassPro2`.
5. If it still fails, install [VC++ Redistributable x64](https://aka.ms/vs/17/release/vc_redist.x64.exe).

`build.bat` will also copy your existing `assets\data2.xlsx` next to the new EXE **without overwriting** a file that is already there.

Manual build:

```bat
pyinstaller --noconfirm PassPro2.spec
```

Output: `dist\PassPro2\PassPro2.exe`

## Data location

| Mode | Password file |
|------|----------------|
| Dev (`python main.py`) | `assets\data2.xlsx` next to `main.py` |
| Built EXE | `assets\data2.xlsx` next to `PassPro2.exe` |

Sheets:

- **Passwords** — Name, Password, Email, Remarks
- **Shortcuts** — main email, second email, identity number
