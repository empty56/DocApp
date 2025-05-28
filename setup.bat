@echo off
SETLOCAL

echo Setting up FormatChecker...

:: Create virtual environment if not exists
IF NOT EXIST "venv\" (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate virtual environment
CALL venv\Scripts\activate.bat

:: Install requirements
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt

:: Generate .env if needed
echo Creating .env file...
python generate_env.py

echo Setup complete. You can now run the app using FormatCheckerLauncher.exe.
pause