# import win32com.client as win32
import re
from docx import Document

wdHeaderFooterPrimary = 1
wdListBullet = 2  # For bullet lists
wdListNumber = 3  # For numbered lists

# Function to convert points to centimeters
def points_to_cm(points):
    return points * 0.0352778

def check_page_attributes(doc):

    # Access page setup (which includes margins)
    page_setup = doc.sections[0].PageSetup  # Get page setup
    errors = []

    # Expected values
    expected_margins = {
        "Left": 3.00,
        "Right": 2.00,
        "Top": 2.00,
        "Bottom": 2.00
    }
    expected_page_size = (21.0, 29.7)  # A4 size in cm

    # Check margins
    actual_margins = {
        "Left": points_to_cm(page_setup.LeftMargin),
        "Right": points_to_cm(page_setup.RightMargin),
        "Top": points_to_cm(page_setup.TopMargin),
        "Bottom": points_to_cm(page_setup.BottomMargin)
    }

    for side, expected_value in expected_margins.items():
        if round(actual_margins[side], 2) != expected_value:
            errors.append(f"{side} Margin is {actual_margins[side]:.2f} cm (should be {expected_value} cm)")

    # Check page size (width and height)
    actual_width = points_to_cm(page_setup.PageWidth)
    actual_height = points_to_cm(page_setup.PageHeight)

    if round(actual_width, 2) != expected_page_size[0] or round(actual_height, 2) != expected_page_size[1]:
        errors.append(f"Page size is {actual_width:.2f} cm x {actual_height:.2f} cm (should be 21 cm x 29.7 cm)")

    if errors:
        print("\n".join(errors))
    else:
        print("Margins and page size are correct.")

# Function to check if all text is Times New Roman, 14 pt
def check_font_and_size(doc, expected_font="Times New Roman", expected_size=14):
    """Check if the font and size for each paragraph in the document is Times New Roman, 14 pt."""
    issues = []

    # Iterate over paragraphs
    for paragraph in doc.Paragraphs:
        # Access the font of the entire paragraph's range
        text = paragraph.Range.Text.strip()

        # Skip if the paragraph is empty or contains only whitespace/special characters
        if not text or text in ['\x07', '\x0c']:  # Common non-visible characters like '\x07' (bell), '\x0c' (form feed)
            continue

        font = paragraph.Range.Font

        # Skip paragraphs with absurd font sizes
        if font.Size == 9999999.0:
            print(f"Skipping paragraph with abnormal font size: {text}")
            continue

        # Check the font name
        if font.Name != expected_font:
            issues.append(f"Incorrect font: {font.Name} in paragraph: {text}")

        # Check the font size
        if font.Size != expected_size:
            issues.append(f"Incorrect font size: {font.Size} pt in paragraph: {text}")

    return issues

# Function to check if text is in full caps and bold
def check_full_caps_bold(paragraph):
    # Print to debug the paragraph text
    # print("Checking paragraph:", paragraph.Range.Text.strip())

    # Check if the paragraph is empty
    if not paragraph.Range.Text.strip():
        return False

    # Get the entire text of the paragraph
    paragraph_text = paragraph.Range.Text.strip()

    # Check if the paragraph text is in uppercase
    if paragraph_text.isupper() and paragraph.Range.Font.Bold:
        return True
    return False

def check_list_formatting(doc, headers):
    """Check the formatting of both manually typed and Word-generated lists, and verify correct indents."""

    for paragraph in doc.Paragraphs:
        paragraph_format = paragraph.Format
        text = paragraph.Range.Text.strip()
        left_indent = round(paragraph_format.LeftIndent / 28.35, 2)  # Convert from points to cm
        first_line_indent = round(paragraph_format.FirstLineIndent / 28.35, 2) # Convert from points to cm
        list_type = paragraph.Range.ListFormat.ListType

        # Skip if it's not a list
        if list_type == 0:
            continue

        # Check if it's a heading (using provided headers list)
        is_heading = text in headers or (paragraph.Range.Font.Bold == True and paragraph.Range.Text.isupper())

        # For List Type 3
        if list_type == 3:
            if not is_heading:  # Exclude headings from this indent check
                if left_indent != 1.75 or first_line_indent != -0.5:
                    print(f"Incorrect indents in List Type 3 (non-heading) paragraph: {text}")
                    print(f"Left Indent: {left_indent:.2f} cm, First Line Indent: {first_line_indent:.2f} cm")
            # else:
            #     print(f"Heading found (excluded from List Type 3 check): {text}")

        # For List Type 4
        elif list_type == 4:
            if left_indent != 2.25 or first_line_indent != -0.45:
                print(f"Incorrect indents in List Type 4 paragraph: {text}")
                print(f"Left Indent: {left_indent:.2f} cm, First Line Indent: {first_line_indent:.2f} cm")

        # List marker and spacing check (same as your previous implementation)
        list_patterns = [
            r'^\d+\.\s',  # Number with dot (e.g., 1. )
            r'^\d+\)\s',  # Number with bracket (e.g., 1) )
            r'^[*•–-]\s',  # Symbols (e.g., *, •, – (long dash), - (short dash))
        ]
        for pattern in list_patterns:
            match = re.match(pattern, text)
            if match:
                # Ensure there is exactly one space after the list marker
                if not re.match(pattern + r'\S', text):
                    print(f"Incorrect spacing after manually typed list marker in paragraph: {text}")
                break  # No need to check other patterns if one matches

def check_table_format(doc_path):
    """
    Checks whether tables in the document meet the specified formatting requirements.

    Args:
        doc_path (str): Path to the Word document.

    Returns:
        list: A list of issues found in table formatting.
    """
    doc = Document(doc_path)
    issues = []
    tables = doc.tables
    paragraphs = doc.paragraphs
    table_count = 0
    previous_table_num = None

    for i, para in enumerate(paragraphs):
        text = para.text.strip()

        # Check for table number format "Таблиця X.Y"
        if text.startswith("Таблиця"):
            table_count += 1
            expected_num = f"Таблиця {table_count}."
            if not text.startswith(expected_num):
                issues.append(f"❌ Table numbering incorrect: '{text}' (expected '{expected_num}')")

            # Ensure right alignment (assuming we have a way to check)
            if para.alignment != 2:  # 2 means right-aligned in python-docx
                issues.append(f"❌ Table number '{text}' is not right-aligned.")

            previous_table_num = text  # Store the last valid table number

        # Check for table continuation
        elif text.startswith("Продовження табл."):
            if previous_table_num and text != f"Продовження {previous_table_num}":
                issues.append(
                    f"❌ Incorrect continuation format: '{text}' (expected 'Продовження {previous_table_num}')")

    # Checking table width and text formatting
    for idx, table in enumerate(tables):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # Check font and size
                    for run in paragraph.runs:
                        if run.font.name != "Times New Roman" or run.font.size is None or run.font.size.pt != 14:
                            issues.append(f"❌ Table {idx + 1} contains text not in Times New Roman 14 pt.")

        # Width check can be done if we extract page width, but it's tricky without exact measurements.

    return issues
