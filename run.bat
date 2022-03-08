set PYTHONDONTWRITEBYTECODE=1
call env\Scripts\activate.bat
python compile_ui.py
cd src
python app.py
pause