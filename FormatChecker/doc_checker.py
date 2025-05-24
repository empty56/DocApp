import tempfile
import pythoncom
import win32com.client as win32

from FormatChecker.checkers import main_part_checker, extras_checker, ai_utils

def is_win32_doc_empty(doc):
    try:
        for p in doc.Paragraphs:
            text = p.Range.Text.strip()
            if text and text != '\x07' and text != '\r':
                return False
        return True
    except Exception:
        return True

def check_document_rules(file_stream, document_part, formatting_check=True, grammar_check=True, exception_words=None):
    if exception_words is None:
        exception_words = []
    pythoncom.CoInitialize()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
        temp_file.write(file_stream.getvalue())
        temp_file_path = temp_file.name

    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False
    doc = word_app.Documents.Open(temp_file_path)

    checkers = {
        "tech_assignment": extras_checker,
        "testing_methodology": extras_checker,
        "user_manual": extras_checker,
        "main_part": main_part_checker,
    }

    try:
        if document_part not in checkers:
            return {"error": "Unknown document part"}
        if is_win32_doc_empty(doc):
            return {"error": "The uploaded document appears to be empty."}
        checker = checkers[document_part]
        result = {}

        if formatting_check:
            result["formatting"] = checker.check_formatting(doc)

        if grammar_check:
            result["grammar"] = ai_utils.check_document_spelling(doc, exception_words)

        return result if result else {"error": "No checks performed"}
    finally:
        try:
            doc.Close()
        except Exception as e:
            print("Failed to close document:", e)
        try:
            word_app.Quit()
        except Exception as quit_error:
            print("Warning: Word quit failed:", quit_error)
        pythoncom.CoUninitialize()

