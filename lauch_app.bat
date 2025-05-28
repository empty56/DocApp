@echo off
SETLOCAL

:: Activate venv
CALL venv\Scripts\activate.bat

:: Launch browser
START http://127.0.0.1:8000

:: Run Django dev server
python manage.py runserver