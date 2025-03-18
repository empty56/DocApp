# import win32com.client as win32
import re
from docx import Document

wdHeaderFooterPrimary = 1
wdListBullet = 2  # For bullet lists
wdListNumber = 3  # For numbered lists

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
def check_font_and_size(doc, expected_font="Times New Roman", expected_size=14, exclude_after=None):
    """
    Check if paragraphs in the document use the expected font and size.
    Skips the "ДОДАТКИ" section if exclude_after is set to "ДОДАТКИ".
    """
    exclude_mode = False  # Track whether we should start excluding text

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        if not text or text in ['\x07', '\x0c']:  # Skip empty/special characters
            continue

        # If we encounter the "ДОДАТКИ" section, we stop checking
        if exclude_after and exclude_after in text.upper():
            exclude_mode = True

        if exclude_mode:
            continue  # Skip everything after "ДОДАТКИ"

        font = paragraph.Range.Font

        # Skip paragraphs with absurd font sizes
        if font.Size == 9999999.0:
            print(f"Skipping paragraph with abnormal font size: {text}")
            continue

        # Check the font name
        if font.Name != expected_font:
            print(f"Incorrect font: {font.Name} in paragraph: {text}")

        # Check the font size
        if font.Size != expected_size:
            print(f"Incorrect font size: {font.Size} pt in paragraph: {text}")

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

def check_table_format(doc):
    previous_table_num = None  # Store the last detected table number
    table_index = 0  # Track actual table numbers as detected by our logic

    paragraphs = list(doc.Paragraphs)  # Convert to a list for indexing

    for i, paragraph in enumerate(paragraphs):
        text = paragraph.Range.Text.strip()

        # Identify a table number (Таблиця X.Y)
        if text.startswith("Таблиця"):
            table_index += 1  # Increment our table counter
            previous_table_num = text  # Store the last valid table number

            # Ensure right indent is 0 (proper alignment check for table number)
            if round(paragraph.Range.ParagraphFormat.RightIndent / 28.35, 2) != 0.0:
                print(f"Incorrect right indent for table number: '{text}' (should be 0.0 cm).")

            # Ensure the table name (next row) is CENTERED
            if i + 1 < len(paragraphs):  # Check the next paragraph safely
                next_paragraph = paragraphs[i + 1]
                next_text = next_paragraph.Range.Text.strip()

                if next_paragraph.Range.ParagraphFormat.Alignment != 1:  # 1 means centered
                    print(f"Incorrect alignment for table name: '{next_text}' (should be centered).")

        # Check for table continuation format
        elif text.startswith("Продовження табл."):
            match = re.match(r"Продовження табл\. (\d+(\.\d+)?)", text)  # Extract table number
            if match:
                table_number = match.group(1)  # Extracted number from continuation
                expected_continuation = f"Продовження табл. {table_number}"
                if text != expected_continuation:
                    print(f"Incorrect continuation format: '{text}' (expected '{expected_continuation}')")

    # Checking table width, text formatting, and handling merged cells
    for idx, table in enumerate(doc.Tables, start=1):  # `idx` now directly refers to table count
        checked_cells = set()  # Avoid duplicate font/size checks

        try:
            for row in table.Rows:
                for cell in row.Cells:
                    # Iterate through paragraphs in the cell to check their text
                    cell_text = ""
                    for para in cell.Range.Paragraphs:
                        para_text = para.Range.Text.strip()
                        # Clean up control characters like \r, \x07
                        cell_text = re.sub(r'[\x00-\x1F\x7F]', '', para_text)
                        # Remove extra spaces between words
                        cell_text = re.sub(r'\s+', ' ', cell_text)

                    # # Check if we got valid text from the cell
                    # if len(cell_text.strip()) == 0:
                    #     print(f"Empty or invalid text in Table {idx} at Row {row.Index}, Column {cell.ColumnIndex}")
                    #     continue

                    if cell_text and cell_text not in checked_cells:
                        checked_cells.add(cell_text)  # Mark cell as checked

                        font = cell.Range.Font
                        if font.Name != "Times New Roman" or font.Size != 14:
                            print(
                                f"Incorrect font or size in Table {idx}: {repr(cell_text)}")  # Use `idx` for actual table count

        except Exception:
            print(f"Skipping Table {idx} due to merged cell issue.")  # Use `idx` for actual table count

    print("Table formatting check completed.")


def clean_topic_name(topic, to_upper=False, to_lower=False):
    cleaned_topic = ''.join([i for i in topic if not i.isdigit()]).replace('.', '').replace('\t', '').strip()

    if to_upper:
        return cleaned_topic.upper()

    elif to_lower:
        return cleaned_topic.lower()

    return cleaned_topic