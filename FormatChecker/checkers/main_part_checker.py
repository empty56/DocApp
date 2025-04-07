import FormatChecker.checkers.doc_utils as doc_utils
import re


def extract_main_part_topics(doc):
    topics = {"main_topics": [], "subtopics": []}
    flag = False  # Flag to track if we are concatenating rows without page numbers
    temp_str = ""  # Temporary string to hold concatenated rows

    if doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)

        for toc_entry in toc.Range.Paragraphs:
            full_text = toc_entry.Range.Text.strip()
            cleaned_text = doc_utils.clean_topic_name(full_text, to_upper=False)

            if full_text:
                if re.search(r'\d{1,2}$', full_text):  # Check if the row has a number at the end
                    if flag:
                        full_text = temp_str + " " + full_text  # Concatenate with previous row
                        cleaned_text = doc_utils.clean_topic_name(full_text, to_upper=False)
                        temp_str = ""
                        flag = False

                    # Categorize topics
                    if cleaned_text.isupper() and not re.search(r'\d', full_text):
                        topics["main_topics"].append(cleaned_text)
                    elif re.match(r'^\d{1,2}\.\d{1,2}\.\s', full_text):  # Subtopics like "1.2."
                        topics["subtopics"].append(cleaned_text)
                    else:
                        topics["main_topics"].append(cleaned_text)
                else:
                    temp_str = full_text if not flag else temp_str + " " + full_text
                    flag = True

    return topics

def check_topics(doc, topics):
    toc_done = False  # Flag to skip the ToC
    main_content_started = False  # Flag to start checking after ToC

    cleaned_main_topics = [doc_utils.clean_topic_name(topic, to_upper=True) for topic in topics["main_topics"]]
    cleaned_subtopics = [doc_utils.clean_topic_name(topic, to_lower=True) for topic in topics["subtopics"]]

    result_text = ""
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
                result_text += f"Incorrect formatting for main topic: {text}\n"
                # print(f"Incorrect formatting for main topic: {text}")

        elif cleaned_text.lower() in cleaned_subtopics:
            is_bold = paragraph.Range.Font.Bold == -1  # Check for bold (Word uses -1 for bold)

            # Ensure correct bold for subtopics
            if not is_bold:
                result_text += f"Incorrect formatting for subtopic: {text} (should be bold)\n"
                # print(f"Incorrect formatting for subtopic: {text} (should be bold)")

            # **Fix: Check capitalization only on the extracted subtopic text**
            if subtopic_text != subtopic_text.capitalize():
                result_text += f"Incorrect capitalization for subtopic: {text} (should start with a capital letter)\n"
                # print(f"Incorrect capitalization for subtopic: {text} (should start with a capital letter)")
    return result_text

def check_formatting(doc):
    result_text = ""
    result_text += doc_utils.check_page_attributes(doc)
    result_text += doc_utils.check_font_and_size(doc, exclude_after="ДОДАТКИ")

    topics = extract_main_part_topics(doc)
    result_text += check_topics(doc, topics)

    result_text += doc_utils.check_table_format(doc)
    result_text += doc_utils.check_table_page_count(doc)
    result_text += doc_utils.check_images_and_captions(doc)
    result_text += doc_utils.check_interline_spacing(doc)
    result_text += doc_utils.check_centered_items_indents_in_document(doc)
    return result_text