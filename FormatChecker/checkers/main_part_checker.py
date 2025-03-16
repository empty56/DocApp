import win32com.client as win32
import FormatChecker.checkers.doc_utils as doc_utils
import re

def clean_topic_name(topic, to_upper=False, to_lower=False):
    cleaned_topic = ''.join([i for i in topic if not i.isdigit()]).replace('.', '').replace('\t', '').strip()

    # Convert to uppercase if needed
    if to_upper:
        return cleaned_topic.upper()

    elif to_lower:
        return cleaned_topic.lower()

    return cleaned_topic

def extract_main_part_toc(doc):
    topics = {"main_topics": [], "subtopics": []}
    flag = False  # Flag to track if we are concatenating rows without page numbers
    temp_str = ""  # Temporary string to hold concatenated rows
    current_topic = ""  # Variable to accumulate multi-line topics

    # Ensure the document has a Table of Contents
    if doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)

        for toc_entry in toc.Range.Paragraphs:
            full_text = toc_entry.Range.Text.strip()  # Full text with numbers
            cleaned_text = clean_topic_name(full_text, to_upper=False)  # Cleaned topic name without numbers

            if full_text:  # Make sure full_text is not empty

                # Check if row has a number at the end
                if re.search(r'\d{1,2}$', full_text):  # Matches a number at the end of the line
                    if flag:  # If flag is active, finalize the temporary string
                        full_text = temp_str + " " + full_text  # Concatenate temporary string with the current full text
                        cleaned_text = clean_topic_name(full_text, to_upper=False)  # Clean the concatenated text
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
    content_page_done = False  # Flag to skip the ToC
    main_content_started = False  # Flag to start checking after ToC

    cleaned_main_topics = [clean_topic_name(topic, to_upper=True) for topic in topics["main_topics"]]
    cleaned_subtopics = [clean_topic_name(topic, to_upper=False).lower() for topic in topics["subtopics"]]

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()

        if not text:
            continue

        cleaned_text = clean_topic_name(text)

        # Detect and skip the Table of Contents
        if not content_page_done:
            if "ЗМІСТ" in text.upper():
                content_page_done = True
            continue

        if not main_content_started:
            if cleaned_text.upper() in cleaned_main_topics or cleaned_text.lower() in cleaned_subtopics:
                main_content_started = True
            continue

        if re.search(r"[\d\s]+$", text):  # Ignore lines that are just numbers
            continue

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

            # Check that the text is not fully capitalized and starts with a capital letter
            if text != text.capitalize():
                print(f"Incorrect capitalization for subtopic: {text} (should start with a capital letter)")

            # Check that it's not in italics
            is_italic = paragraph.Range.Font.Italic == -1  # Word uses -1 for italic text
            if is_italic:
                print(f"Incorrect formatting for subtopic: {text} (should not be italic)")

def check_alignment(file_path):
    # Start Word application
    word_app = win32.Dispatch('Word.Application')
    word_app.Visible = False  # Keep Word application hidden during processing
    # Open the document
    doc = word_app.Documents.Open(file_path)
    # Access page setup (which includes margins)
    doc_utils.check_page_attributes(doc)

    # font_issues = doc_utils.check_font_and_size(doc)  # Check font and size MODIFY
    # if font_issues:
    #     for issue in font_issues:
    #         print(issue)

    topics = extract_main_part_toc(doc)
    check_topics(doc, topics)

    doc.Close()
    word_app.Quit()