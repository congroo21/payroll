# Payroll Paystub Generator

This project extracts payroll data from a PDF and generates individual pay-stub PDFs.

## Building the Windows Executable

1. Install [PyInstaller](https://pyinstaller.org/):
   ```sh
   pip install pyinstaller
   ```
2. Run PyInstaller on `bootstrap.py`:
   ```sh
   pyinstaller --onefile --add-data "NanumGothic.ttf;." --add-data "main.py;." bootstrap.py
   ```
   - Both the font and `main.py` need to be packaged or distributed alongside the executable.
   - The resulting `bootstrap.exe` can be distributed to users.

The bootstrap script will install required Python packages and a local JDK if they
are missing, then launch `main.py`.
