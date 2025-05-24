import re
import math

wdHeaderFooterPrimary = 1
wdListBullet = 2  # For bullet lists
wdListNumber = 3  # For numbered lists

def points_to_cm(points):
    return points * 0.0352778

def check_page_attributes(doc):
    page_setup = doc.sections[0].PageSetup
    result_text = ""
    tolerance = 0.05  # ±0.05 cm tolerance

    expected_margins = {
        "Left": 3.00,
        "Right": 2.00,
        "Top": 2.00,
        "Bottom": 2.00
    }
    expected_page_size = (21.0, 29.7)  # A4 size in cm

    actual_margins = {
        "Left": points_to_cm(page_setup.LeftMargin),
        "Right": points_to_cm(page_setup.RightMargin),
        "Top": points_to_cm(page_setup.TopMargin),
        "Bottom": points_to_cm(page_setup.BottomMargin)
    }

    for side, expected_value in expected_margins.items():
        actual = actual_margins[side]
        if not math.isclose(actual, expected_value, abs_tol=tolerance):
            result_text += f"{side} Margin is {actual:.2f} cm (should be {expected_value} cm)\n"
    actual_width = points_to_cm(page_setup.PageWidth)
    actual_height = points_to_cm(page_setup.PageHeight)
    if not math.isclose(actual_width, expected_page_size[0], abs_tol=tolerance) or \
       not math.isclose(actual_height, expected_page_size[1], abs_tol=tolerance):
        result_text += f"Page size is {actual_width:.2f} cm x {actual_height:.2f} cm (A4 is 21 cm x 29.7 cm)\n"
    return result_text


def check_font_and_size(doc, expected_font="Times New Roman", expected_size=14, exclude_after=None):
    result_text = ""
    content_started = False

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        if not text or text in ['\x07', '\x0c']:  # Skip empty/special characters
            continue

        if not content_started:
            if "ЗМІСТ" in text.upper():
                content_started = True
            continue

        # If we encounter the "ДОДАТКИ" section, we stop checking
        if exclude_after and exclude_after in text.upper():
            break


        font = paragraph.Range.Font

        # Skip paragraphs with absurd font sizes
        if font.Size == 9999999.0:
            # result_text += f"Skipping paragraph with abnormal font size: {text}\n"
            # print(f"Skipping paragraph with abnormal font size: {text}")
            continue

        # Check the font name
        if font.Name != expected_font:
            result_text += f"Incorrect font: {font.Name} in paragraph: {text}\n"
            # print(f"Incorrect font: {font.Name} in paragraph: {text}")

        # Check the font size
        if font.Size != expected_size:
            result_text += f"Incorrect font size: {font.Size} pt in paragraph: {text}\n"
            # print(f"Incorrect font size: {font.Size} pt in paragraph: {text}")
    return result_text

def check_interline_spacing(doc, expected_spacing=1.5):
    found_toc = False  # Flag to skip everything after "ДОДАТКИ"
    title_page_checked = False  # Flag to skip the title page if present
    result_text = ""
    content_started = False

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        # Skip empty paragraphs
        if not text:
            continue

        if not content_started:
            if "ЗМІСТ" in text.upper():
                content_started = True
            continue

        # Detect "ДОДАТКИ" section and stop checking after it
        if text == "ДОДАТКИ":
            found_toc = True
            continue

        if found_toc:
            continue  # Skip checking everything after "ДОДАТКИ"

        # Skip paragraphs that belong to tables
        if paragraph.Range.Tables.Count > 0:
            continue  # Ignore text inside tables

        # Get the actual line spacing
        actual_spacing = round(paragraph.Format.LineSpacing / 12, 2)  # Convert points to relative spacing
        page_number = paragraph.Range.Information(3)

        # Check if spacing is incorrect
        if actual_spacing != expected_spacing:
            result_text = f"Incorrect interline spacing on page {page_number}: '{text}' (should be {expected_spacing}, found {actual_spacing})\n"
    return result_text

def check_full_caps_bold(paragraph):

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
    result_text = ""
    for paragraph in doc.Paragraphs:
        paragraph_format = paragraph.Format
        text = paragraph.Range.Text.strip()
        left_indent = round(paragraph_format.LeftIndent / 28.35, 2)  # Convert from points to cm
        first_line_indent = round(paragraph_format.FirstLineIndent / 28.35, 2) # Convert from points to cm
        list_type = paragraph.Range.ListFormat.ListType
        # Skip if it's not a list
        if list_type == 0:
            continue
        # Check if it's a heading
        is_heading = text in headers or (paragraph.Range.Font.Bold == True and paragraph.Range.Text.isupper())
        # For List Type 3
        if list_type == 3:
            if not is_heading:  # Exclude headings from this indent check
                if left_indent != 1.75 or first_line_indent != -0.5:
                    result_text += (f"Incorrect indents in List Type 3 (non-heading) paragraph: {text}\n"
                                    f"Left Indent: {left_indent:.2f} cm, First Line Indent: {first_line_indent:.2f} cm\n")
        # For List Type 4
        elif list_type == 4:
            if left_indent != 2.25 or first_line_indent != -0.45:
                result_text += (f"Incorrect indents in List Type 4 paragraph: {text}\n"
                                f"Left Indent: {left_indent:.2f} cm, First Line Indent: {first_line_indent:.2f} cm\n")
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
                    result_text += f"Incorrect spacing after manually typed list marker in paragraph: {text}\n"
                break
    return result_text

def check_table_format(doc):
    result_text = ""
    paragraphs = list(doc.Paragraphs)

    table_heading_pattern = re.compile(r"^Таблиця\s+\d+(\.\d+)*$")
    continuation_pattern = re.compile(r"^Продовження табл\. (\d+(\.\d+)*)$")

    for i, paragraph in enumerate(paragraphs):
        text = paragraph.Range.Text.strip()

        # Check for valid table heading (e.g., "Таблиця 1", "Таблиця 1.1", "Таблиця 1.1.2")
        if table_heading_pattern.match(text):
            if round(paragraph.Range.ParagraphFormat.RightIndent / 28.35, 2) != 0.25:
                result_text += f"Incorrect right indent for table number: '{text}' (should be 0.25 cm).\n"

            if i + 1 < len(paragraphs):
                next_paragraph = paragraphs[i + 1]
                next_text = next_paragraph.Range.Text.strip()
                if next_paragraph.Range.ParagraphFormat.Alignment != 1:
                    result_text += f"Incorrect alignment for table name: '{next_text}' (should be centered).\n"

        elif continuation_pattern.match(text):
            match = continuation_pattern.match(text)
            table_number = match.group(1)
            expected_text = f"Продовження табл. {table_number}"
            if text != expected_text:
                result_text += f"Incorrect continuation format: '{text}' (expected '{expected_text}')\n"

    # Table formatting checks
    for idx, table in enumerate(doc.Tables, start=1):
        checked_cells = set()
        try:
            for row in table.Rows:
                for cell in row.Cells:
                    cell_text = ""
                    for para in cell.Range.Paragraphs:
                        raw_text = para.Range.Text
                        cleaned = re.sub(r'[\x00-\x1F\x7F]', '', raw_text).strip()
                        cleaned = re.sub(r'\s+', ' ', cleaned)
                        if cleaned:
                            cell_text += cleaned + " "
                    cell_text = cell_text.strip()
                    if cell_text and cell_text not in checked_cells:
                        checked_cells.add(cell_text)

                        font = cell.Range.Font
                        if font.Name != "Times New Roman" or font.Size != 14:
                            result_text += (
                                f"Incorrect font or size in Table {idx} cell: '{cell_text}' "
                                f"(Font: {font.Name}, Size: {font.Size})\n"
                            )
        except Exception:
            result_text += f"Skipping Table {idx} due to merged cell issue.\n"

    return result_text


def check_table_page_count(doc):
    table_info = {}
    result_text = ""
    for i, table in enumerate(doc.Tables):
        try:
            # Check for tables with no rows or malformed tables
            if table.Rows.Count == 0:
                result_text += f"Table {i+1} has no rows. Skipping it.\n"
                table_info[i + 1] = (None, None)
                continue
            first_row = table.Rows[1]
            last_row = table.Rows.Last
            start_page = first_row.Range.Information(3)
            last_row_page = last_row.Range.Information(3)
            if last_row_page != start_page:
                table_info[i + 1] = (start_page, last_row_page)
            else:
                table_info[i + 1] = (start_page, start_page)

        except Exception as e:
            # Handle the case of vertically merged cells or any other error
            if "Cannot access individual rows" in str(e):
                result_text += f"Table {i+1} has vertically merged cells. Skipping page count check.\n"
            table_info[i + 1] = (None, None)
            continue
    # Display results for tables that span multiple pages
    for idx, (start, end) in table_info.items():
        if start != end and start is not None:  # Only print if it spans multiple pages
            result_text += f"Table {idx} spans multiple pages ({start} → {end}).\n"
    return result_text

def check_images_and_captions(doc):
    result_text = ""
    for shape in doc.InlineShapes:
        shape_range = shape.Range
        image_paragraph = shape_range.Paragraphs(1)
        # Get the page number where the image is located
        page_number = shape_range.Information(3)  # 3 corresponds to wdActiveEndPageNumber
        # Check if the image is centered
        if image_paragraph.Range.ParagraphFormat.Alignment != 1:
            result_text += f"Image on page {page_number} is not centered.\n"
        # Find the next valid paragraph (skip empty ones)
        next_para = image_paragraph
        while next_para and (not next_para.Range.Text.strip() or next_para.Range.Text.strip() == "/"):
            next_para = next_para.Next()

        if next_para:
            caption_text = next_para.Range.Text.strip()
            normalized_caption = caption_text.lower()
            if normalized_caption.startswith("рис."):

                if next_para.Range.ParagraphFormat.Alignment != 1:
                    result_text += f"Incorrect alignment for caption on page {page_number}: '{caption_text}' (should be centered).\n"
                # Check if caption is bold (should not be)
                if next_para.Range.Font.Bold:
                    result_text += f"Caption is incorrectly bold: '{caption_text}'\n"
                # Check if 'Рис.' is wrongly capitalized
                if caption_text.startswith("Рис.") and not caption_text[0].isupper():
                    result_text += f"Caption contains incorrectly capitalized 'Рис.': '{caption_text}'\n"
                # Check if the rest of the caption text is in full uppercase
                rest_of_caption = caption_text[caption_text.find("Рис.") + 4:].strip()
                if rest_of_caption.isupper():
                    result_text += f"Caption contains full uppercase text: '{caption_text}'\n"
        else:
            result_text += "Warning: No caption found after the image."
            print("Warning: No caption found after the image.")
    return result_text

def check_centered_items_indents_in_document(doc):
    result_text = ""

    for paragraph in doc.Paragraphs:
        # Skip paragraphs inside tables
        if paragraph.Range.Tables.Count > 0:
            continue
        if paragraph.Range.InlineShapes.Count > 0:
            if paragraph.Range.ParagraphFormat.Alignment == 1:
                if paragraph.Range.ParagraphFormat.LeftIndent != 0:
                    result_text += (
                        f"Image on page {paragraph.Range.Information(3)} has incorrect left indent: "
                        f"{round(paragraph.Range.ParagraphFormat.LeftIndent, 2)}\n"
                    )
                if paragraph.Range.ParagraphFormat.RightIndent != 0:
                    result_text += (
                        f"Image on page {paragraph.Range.Information(3)} has incorrect right indent: "
                        f"{round(paragraph.Range.ParagraphFormat.RightIndent, 2)}\n"
                    )
        elif paragraph.Range.ParagraphFormat.Alignment == 1:
            text = paragraph.Range.Text.strip()
            if not text or text in ['\x07', '\x0c']:
                continue
            if paragraph.Range.ParagraphFormat.LeftIndent != 0:
                result_text += (
                    f"Centered paragraph: '{text}' on page {paragraph.Range.Information(3)} has incorrect left indent: "
                    f"{round(paragraph.Range.ParagraphFormat.LeftIndent, 2)}\n"
                )
            if paragraph.Range.ParagraphFormat.RightIndent != 0:
                result_text += (
                    f"Centered paragraph: '{text}' on page {paragraph.Range.Information(3)} has incorrect right indent: "
                    f"{round(paragraph.Range.ParagraphFormat.RightIndent, 2)}\n"
                )

    return result_text

def clean_topic_name(topic, to_upper=False, to_lower=False):
    cleaned_topic = ''.join([i for i in topic if not i.isdigit()]).replace('.', '').replace('\t', '').strip()
    if to_upper:
        return cleaned_topic.upper()
    elif to_lower:
        return cleaned_topic.lower()
    return cleaned_topic