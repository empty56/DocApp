import win32com.client as win32
import FormatChecker.checkers.doc_utils as doc_utils
import FormatChecker.checkers.ai_utils as ai_utils
import re


def extract_main_part_topics(doc):
    topics = {"main_topics": [], "subtopics": []}
    flag = False  # Flag to track if we are concatenating rows without page numbers
    temp_str = ""  # Temporary string to hold concatenated rows
    current_topic = ""  # Variable to accumulate multi-line topics

    # Ensure the document has a Table of Contents
    if doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)

        for toc_entry in toc.Range.Paragraphs:
            full_text = toc_entry.Range.Text.strip()  # Full text with numbers
            cleaned_text = doc_utils.clean_topic_name(full_text, to_upper=False)  # Cleaned topic name without numbers

            if full_text:  # Make sure full_text is not empty

                # Check if row has a number at the end
                if re.search(r'\d{1,2}$', full_text):  # Matches a number at the end of the line
                    if flag:  # If flag is active, finalize the temporary string
                        full_text = temp_str + " " + full_text  # Concatenate temporary string with the current full text
                        cleaned_text = doc_utils.clean_topic_name(full_text, to_upper=False)  # Clean the concatenated text
                        temp_str = ""  # Reset temporary string
                        flag = False  # Reset the flag

                    # Check if it's a main topic: fully capitalized and no numbers
                    if cleaned_text.isupper() and not re.search(r'\d', full_text):
                        topics["main_topics"].append(cleaned_text)
                    # Check if it's a subtopic: contains numbering like "1.2.", "1.12.", etc.
                    elif re.match(r'^\d{1,2}\.\d{1,2}\.\s', full_text):  # Matches "1.2.", "2.3.", "3.6."
                        topics["subtopics"].append(cleaned_text)
                    # Otherwise, consider it a main topic
                    else:
                        topics["main_topics"].append(cleaned_text)
                else:
                    if not flag:  # If flag is not active, start concatenating
                        temp_str = full_text
                        flag = True  # Activate the flag
                    else:  # If flag is active, continue concatenating
                        temp_str += " " + full_text

    return topics

def check_topics(doc, topics):
    toc_done = False  # Flag to skip the ToC
    main_content_started = False  # Flag to start checking after ToC

    cleaned_main_topics = [doc_utils.clean_topic_name(topic, to_upper=True) for topic in topics["main_topics"]]
    cleaned_subtopics = [doc_utils.clean_topic_name(topic, to_lower=True) for topic in topics["subtopics"]]

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        if not text:
            continue

        cleaned_text = doc_utils.clean_topic_name(text)

        # Detect and skip the Table of Contents
        if not toc_done:
            if "ЗМІСТ" in text.upper():
                toc_done = True
            continue

        if not main_content_started:
            if cleaned_text.upper() in cleaned_main_topics or cleaned_text.lower() in cleaned_subtopics:
                main_content_started = True
            continue

        if re.search(r"[\d\s]+$", text):  # Ignore lines that are just numbers
            continue

        # Extract the actual subtopic text (remove leading numbers, spaces, and tabs)
        subtopic_match = re.match(r"[\d.]+\s*(.*)", text)
        subtopic_text = subtopic_match.group(1) if subtopic_match else text

        # Now we're in the main content, check topic formatting
        if cleaned_text.upper() in cleaned_main_topics:
            result = doc_utils.check_full_caps_bold(paragraph)
            if not result:
                print(f"Incorrect formatting for main topic: {text}")

        elif cleaned_text.lower() in cleaned_subtopics:
            is_bold = paragraph.Range.Font.Bold == -1  # Check for bold (Word uses -1 for bold)

            # Ensure correct bold for subtopics
            if not is_bold:
                print(f"Incorrect formatting for subtopic: {text} (should be bold)")

            # **Fix: Check capitalization only on the extracted subtopic text**
            if subtopic_text != subtopic_text.capitalize():
                print(f"Incorrect capitalization for subtopic: {text} (should start with a capital letter)")

def check_alignment(file_path):
    # Start Word application
    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep Word application hidden during processing
    doc = word_app.Documents.Open(file_path)

    # doc_utils.check_page_attributes(doc)
    # doc_utils.check_font_and_size(doc, exclude_after="ДОДАТКИ")

    # topics = extract_main_part_toc(doc)
    # check_topics(doc, topics)

    # doc_utils.check_table_format(doc)

    # doc_utils.get_table_page_count(doc)

    # doc_utils.check_images_and_captions(doc)

    # doc_utils.check_interline_spacing(doc)

    # doc_utils.check_centered_items_indents_in_document(doc)

    ai_utils.check_document_spelling(doc)

    doc.Close()
    word_app.Quit()