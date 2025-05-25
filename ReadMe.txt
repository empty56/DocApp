# FormatChecker

This is a web-based document formatting and grammar checker built with Django. It uses Microsoft Word for formatting checks and LanguageTool for grammar validation.

## Features

- Format checking (font, spacing, indents, etc.)
- Grammar checking via LanguageTool API (with exception words)
- File upload with styled UI
- Download results as .txt
- Requires Microsoft Word (via COM interface)

## Requirements

- Windows OS
- Microsoft Word installed (local)
- Python 3.10+
- Virtual environment recommended
- Don't try to work in Word while performing checks as it can crash the program

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/yourname/format-checker.git
cd format-checker

# 2. Create a virtual environment
python -m venv venv
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Create .env and set up your values
SECRET_KEY=your-secret-key-here # SECURITY WARNING: keep the secret key used in production secret!
DEBUG=True # SECURITY WARNING: don't run with debug turned on in production!
ALLOWED_HOSTS=your-values # for local machine: 127.0.0.1,localhost

# 5. Run the server
python manage.py runserver