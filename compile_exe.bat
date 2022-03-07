call env\Scripts\activate.bat
pyinstaller --onefile --hiddenimport openpyxl --noconsole src\app.py
pause