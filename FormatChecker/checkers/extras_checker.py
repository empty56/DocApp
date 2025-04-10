import FormatChecker.checkers.doc_utils as doc_utils
import re

def check_topics(doc, topics):
    content_page_done = False
    main_content_started = False
    result_text = ""
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
                    result_text += f"Incorrect formatting for topic: {text}\n"
                    # print(f"Incorrect formatting for topic: {text}")
                break
    return result_text

def extract_topics_from_toc(doc, to_upper=False):
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
            # print("Found 'ЗМІСТ'. Extracting topics...")
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
    left_indent = round(paragraph.Format.LeftIndent / 28.35, 2)  # Convert from points to cm
    right_indent = round(paragraph.Format.RightIndent / 28.35, 2)

    # If indents return 0, try using FirstLineIndent (if needed)
    if left_indent == 0:
        left_indent = round(paragraph.Format.FirstLineIndent / 28.35, 2)

    return left_indent, right_indent

def check_project_stages_topic(doc, topics):
    # Normalize topic names (remove numbers)
    result_text = ""
    cleaned_topics = [doc_utils.clean_topic_name(topic) for topic in topics]

    # Find the index of 'ЕТАПИ ПРОЄКТУВАННЯ' (ignore numbering)
    stage_index = next((i for i, topic in enumerate(cleaned_topics) if "ЕТАПИ ПРОЄКТУВАННЯ" in topic), None)

    if stage_index is None:
        result_text += "Topic 'ЕТАПИ ПРОЄКТУВАННЯ' not present\n"
        # print("Topic 'ЕТАПИ ПРОЄКТУВАННЯ' not found.")
        return result_text

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

            if expected_left_indent is None:
                if left_indent in [0.00, 1.25]:  # Only allow 0 or 1.25
                    expected_left_indent = left_indent

            # Ensure all rows have the same left indent (either 0 or 1.25)
            if left_indent != expected_left_indent:
                result_text += f"Left Indent: '{text}' ({left_indent:.2f} cm). All left indents have to be same and either 0.00 or 1.25 cm\n"
                # print(f"Left Indent: '{text}' ({left_indent:.2f} cm). All left indents have to be same and either 0.00 or 1.25 cm")

            # Right indent should always be 0.00 cm
            if right_indent != 0.00:
                result_text += f"Right Indent: {text} ({right_indent:.2f} cm) (should be 0.00 cm)\n"
                # print(f"Right Indent: {text} ({right_indent:.2f} cm) (should be 0.00 cm)")
    return result_text

def check_formatting(doc):
    topics = extract_topics_from_toc(doc, to_upper=True)
    checks = [doc_utils.check_page_attributes(doc),
              doc_utils.check_font_and_size(doc),
              check_topics(doc, topics),
              doc_utils.check_list_formatting(doc, topics),
              check_project_stages_topic(doc, topics),
              doc_utils.check_interline_spacing(doc),
              doc_utils.check_centered_items_indents_in_document(doc),
              doc_utils.check_table_format(doc),
              doc_utils.check_table_page_count(doc)]
    result = [item for item in checks if item]
    return result