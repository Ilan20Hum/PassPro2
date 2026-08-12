# PassPro2

Desktop password manager for Windows — save and search names, passwords, emails, and remarks locally on your computer.

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Copy the data template before first run (your local file is gitignored):

```bash
copy assets\data2.xlsx.template assets\data2.xlsx
```

Or create `assets/data2.xlsx` with two sheets: **Passwords** (columns: Name, Password, Email, Remarks) and **Shortcuts**.

## Run

```bash
python main.py
```

## Build executable

```bash
pyinstaller PassPro2.spec
```

Output: `dist/PassPro2/PassPro2.exe`
