import tempfile
import pythoncom
import win32com.client as win32
from .utils import process_file
from FormatChecker.checkers import entry_checker, main_part_checker, extras_checker, ai_utils


def check_document_rules(file_stream, document_part, formatting_check=True, grammar_check=True, exception_words=None):
    if exception_words is None:
        exception_words = []
    pythoncom.CoInitialize()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
        temp_file.write(file_stream.getvalue())
        temp_file_path = temp_file.name

    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep Word application hidden
    doc = word_app.Documents.Open(temp_file_path)

    checkers = {
        "entry": entry_checker,
        "tech_assignment": extras_checker,
        "testing_methodology": extras_checker,
        "user_manual": extras_checker,
        "main_part": main_part_checker,
    }

    try:
        if document_part not in checkers:
            return {"error": "Unknown document part"}

        checker = checkers[document_part]
        result = {}

        if formatting_check:
            result["formatting"] = checker.check_formatting(doc)

        if grammar_check:
            result["grammar"] = ai_utils.check_document_spelling(doc, exception_words)

        return result if result else "No checks performed"

    finally:
        doc.Close()
        word_app.Quit()
        pythoncom.CoUninitialize()


# def main():
    # Get the DOCX file path
    # file_path = "D:/Diploma/Bakalavrat/Docs/TestDoc3.docx"
    # document_part = "tech_assignment"
    # process_file(file_path)
    # check_formatting_rules(file_path, document_part)
    # file_path = "D:/Diploma/Bakalavrat/Docs/Main_part_copy4.docx"
    # document_part = "main_part"
    # process_file(file_path)
    # check_document_rules(file_path, document_part)
#
# if __name__ == "__main__":
#     main()
