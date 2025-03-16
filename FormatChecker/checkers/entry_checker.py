import win32com.client as win32
from FormatChecker.checkers.doc_utils import check_page_attributes

def check_alignment(file_path):
    # Start Word application
    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep Word application hidden during processing
    # Open the document
    doc = word_app.Documents.Open(file_path)
    # Access page setup (which includes margins)
    check_page_attributes(doc)

    doc.Close()
    word_app.Quit()
    # Close the document and quit Wor