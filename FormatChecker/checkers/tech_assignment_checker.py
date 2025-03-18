import win32com.client as win32
import FormatChecker.checkers.doc_utils as doc_utils
import re


def check_topics(doc, topics):
    content_page_done = False
    main_content_started = False

    cleaned_topics = [doc_utils.clean_topic_name(topic, to_upper=True) for topic in topics]

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        if not text:
            continue

        cleaned_text = doc_utils.clean_topic_name(text, to_upper=True)

        if not content_page_done:
            if "ЗМІСТ" in text.upper():
                content_page_done = True
            continue

        if not main_content_started:
            if cleaned_text in cleaned_topics:
                main_content_started = True
            continue

        if re.search(r"[\d\s]+$", text):
            continue

        for topic in cleaned_topics:
            if topic == cleaned_text:
                result = doc_utils.check_full_caps_bold(paragraph)
                if not result:
                    print(f"Incorrect formatting for topic: {text}")
                break

def extract_topics_from_toc(doc, to_upper=False):
    """Extract topics from the Table of Contents (TOC) in the document."""
    topics = []

    # Try extracting using Word's TOC method first
    if doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)
        for toc_entry in toc.Range.Paragraphs:
            text = toc_entry.Range.Text.strip()
            if text:
                topic_text = doc_utils.clean_topic_name(text, to_upper)
                if topic_text:  # Ensure it's not empty after cleaning
                    topics.append(topic_text)

    if topics:
        return topics  # Return if standard TOC extraction worked

    # If no TOC found, try extracting topics manually
    for para in doc.Paragraphs:
        if "ЗМІСТ" in para.Range.Text.strip():
            print("Found 'ЗМІСТ'. Extracting topics...")
            current = para.Range.Next(Unit=3)  # Move to next paragraph

            while current and current.Text.strip():
                topic_text = doc_utils.clean_topic_name(current.Text.strip(), to_upper)

                # **Filter out empty or non-topic lines (like dots, page numbers)**
                if topic_text and not re.match(r'^[\d.\s]*$', topic_text):
                    topics.append(topic_text)

                current = current.Next(Unit=3)  # Move to next paragraph

            break  # Stop searching after first "ЗМІСТ"

    return topics

def get_paragraph_indents(paragraph):
    """Get the left and right indents of a paragraph in cm."""
    left_indent = round(paragraph.Format.LeftIndent / 28.35, 2)  # Convert from points to cm
    right_indent = round(paragraph.Format.RightIndent / 28.35, 2)

    # If indents return 0, try using FirstLineIndent (if needed)
    if left_indent == 0:
        left_indent = round(paragraph.Format.FirstLineIndent / 28.35, 2)

    return left_indent, right_indent

def check_project_stages_topic(doc, topics):
    """Check that all rows in 'ЕТАПИ ПРОЄКТУВАННЯ' section have the same left indent (0 cm or 1.25 cm)."""

    # Normalize topic names (remove numbers)
    cleaned_topics = [doc_utils.clean_topic_name(topic) for topic in topics]

    # Find the index of 'ЕТАПИ ПРОЄКТУВАННЯ' (ignore numbering)
    stage_index = next((i for i, topic in enumerate(cleaned_topics) if "ЕТАПИ ПРОЄКТУВАННЯ" in topic), None)

    if stage_index is None:
        print("Topic 'ЕТАПИ ПРОЄКТУВАННЯ' not found.")
        return

    # Identify the next topic dynamically
    next_topic = cleaned_topics[stage_index + 1] if stage_index + 1 < len(cleaned_topics) else None

    in_section = False
    expected_left_indent = None  # Will store the first detected left indent

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()
        cleaned_text = doc_utils.clean_topic_name(text)

        # Detect the start of the section
        if "ЕТАПИ ПРОЄКТУВАННЯ" in cleaned_text:
            in_section = True
            continue

        # Stop checking if we reach the next topic
        if next_topic and next_topic in cleaned_text:
            break

        # Skip empty rows
        if in_section and not text:
            continue

        # Process indentation for relevant paragraphs
        if in_section:
            left_indent, right_indent = get_paragraph_indents(paragraph)

            # Set expected left indent based on the first row
            if expected_left_indent is None:
                if left_indent in [0.00, 1.25]:  # Only allow 0 or 1.25
                    expected_left_indent = left_indent
                else:
                    print(f"Invalid left indent in first row: {text} ({left_indent:.2f} cm)")
                    return

            # Ensure all rows have the same left indent (either 0 or 1.25)
            if left_indent != expected_left_indent:
                print(f"Inconsistent left indent in row: {text}")
                print(f"Left Indent: {left_indent:.2f} cm (should be {expected_left_indent:.2f} cm)")

            # Right indent should always be 0.00 cm
            if right_indent != 0.00:
                print(f"Incorrect right indent in row: {text}")
                print(f"Right Indent: {right_indent:.2f} cm (should be 0.00 cm)")

def check_alignment(file_path):
    # Start Word application
    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep Word application hidden during processing
    # Open the document
    doc = word_app.Documents.Open(file_path)

    # Perform checks
    doc_utils.check_page_attributes(doc)  # Check the margins
    doc_utils.check_font_and_size(doc)  # Check font and size

    # Extract topics from the Table of Contents (TOC)
    topics = extract_topics_from_toc(doc, to_upper=True)  # Set to_upper=True to convert to full caps
    # Check the main text against the cleaned topics
    check_topics(doc, topics)
    doc_utils.check_list_formatting(doc, topics)

    check_project_stages_topic(doc, topics)


    doc.Close()
    word_app.Quit()