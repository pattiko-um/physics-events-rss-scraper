Windows build instructions for `physics-events-rss-scraper`

This repository contains `main.py` (the scraper). To produce a single-file Windows executable (.exe) there are two recommended approaches:

1) Build on a Windows machine (recommended)

- Install Python (3.8+ recommended) and Git.
- Open a PowerShell or CMD prompt in the project folder.
- Create and activate a virtual environment:

```powershell
python -m venv venv
venv\Scripts\Activate.ps1  # PowerShell
```

- Install dependencies and PyInstaller:

```powershell
pip install --upgrade pip
pip install -r requirements.txt pyinstaller
```

- Build the single-file exe:

```powershell
pyinstaller --noconfirm --onefile --name physics-events-scraper main.py
```

- The produced exe will be in `dist\physics-events-scraper.exe`.

2) Use GitHub Actions to build on `windows-latest` (CI)

I included a workflow (`.github/workflows/build-windows.yml`) that builds the exe and uploads it as an artifact. Push to your repo or trigger the workflow via the Actions tab to get a Windows-built `.exe` you can download.

Notes
- PyInstaller must run on the target OS (Windows) to reliably produce a native `.exe`. Cross-compiling from macOS is not supported out-of-the-box.
- If you need the exe built automatically, use the included GitHub Actions workflow. If you prefer building locally but you don't have Windows, set up a Windows VM or use CI.
