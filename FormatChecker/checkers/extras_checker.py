import FormatChecker.checkers.doc_utils as doc_utils
import re


def check_topics(doc, topics):
    results = []
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
                    results.append(f"Incorrect formatting for topic: {text}")
                break
    return results


def extract_topics_from_toc(doc, to_upper=False):
    topics = []

    if doc.TablesOfContents.Count > 0:
        toc = doc.TablesOfContents(1)
        for toc_entry in toc.Range.Paragraphs:
            text = toc_entry.Range.Text.strip()
            if text:
                topic_text = doc_utils.clean_topic_name(text, to_upper)
                if topic_text:
                    topics.append(topic_text)

    if topics:
        return topics

    for para in doc.Paragraphs:
        if "ЗМІСТ" in para.Range.Text.strip():
            current = para.Range.Next(Unit=3)

            while current and current.Text.strip():
                topic_text = doc_utils.clean_topic_name(current.Text.strip(), to_upper)

                if topic_text and not re.match(r'^[\d.\s]*$', topic_text):
                    topics.append(topic_text)

                current = current.Next(Unit=3)

            break

    return topics


def get_paragraph_indents(paragraph):
    left_indent = round(paragraph.Format.LeftIndent / 28.35, 2)
    right_indent = round(paragraph.Format.RightIndent / 28.35, 2)

    if left_indent == 0:
        left_indent = round(paragraph.Format.FirstLineIndent / 28.35, 2)

    return left_indent, right_indent


def check_project_stages_topic(doc, topics):
    results = []
    cleaned_topics = [doc_utils.clean_topic_name(topic) for topic in topics]

    stage_index = next((i for i, topic in enumerate(cleaned_topics) if "ЕТАПИ ПРОЄКТУВАННЯ" in topic), None)

    if stage_index is None:
        results.append("Topic 'ЕТАПИ ПРОЄКТУВАННЯ' not found.")
        return results

    next_topic = cleaned_topics[stage_index + 1] if stage_index + 1 < len(cleaned_topics) else None

    in_section = False
    expected_left_indent = None

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()
        cleaned_text = doc_utils.clean_topic_name(text)

        if "ЕТАПИ ПРОЄКТУВАННЯ" in cleaned_text:
            in_section = True
            continue

        if next_topic and next_topic in cleaned_text:
            break

        if in_section and not text:
            continue

        if in_section:
            left_indent, right_indent = get_paragraph_indents(paragraph)

            if expected_left_indent is None:
                if left_indent in [0.00, 1.25]:
                    expected_left_indent = left_indent
                else:
                    results.append(f"Invalid left indent in first row: {text} ({left_indent:.2f} cm)")
                    return results

            if left_indent != expected_left_indent:
                results.append(
                    f"Inconsistent left indent in row: {text} (Found: {left_indent:.2f} cm, Expected: {expected_left_indent:.2f} cm)")

            if right_indent != 0.00:
                results.append(
                    f"Incorrect right indent in row: {text} (Found: {right_indent:.2f} cm, Expected: 0.00 cm)")

    return results


def check_formatting(doc):
    results = {
        "page_attributes": doc_utils.check_page_attributes(doc),
        "font_and_size": doc_utils.check_font_and_size(doc),
        "topics": [],
        "list_formatting": [],
        "project_stages": [],
        "spacing": [],
        "centered_items": []
    }

    topics = extract_topics_from_toc(doc, to_upper=True)

    results["topics"] = check_topics(doc, topics)
    results["list_formatting"] = doc_utils.check_list_formatting(doc, topics)
    results["project_stages"] = check_project_stages_topic(doc, topics)
    results["spacing"] = doc_utils.check_interline_spacing(doc)
    results["centered_items"] = doc_utils.check_centered_items_indents_in_document(doc)

    return results
