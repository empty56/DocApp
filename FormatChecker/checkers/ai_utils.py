import requests
import re
from rapidfuzz import fuzz

LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"

SIMILARITY_THRESHOLD = 85

def extract_abbreviations(text):
    return set(re.findall(r"\b[А-ЯЇЄҐ]{2,}\b", text))  # Matches 2+ uppercase Ukrainian letters

def is_similar(word1, word2):
    similarity = fuzz.ratio(word1.lower(), word2.lower())
    return similarity >= SIMILARITY_THRESHOLD

def extract_word_from_brackets(text):
    match = re.search(r"«(.+?)»", text)
    return match.group(1) if match else None

def check_spelling(text, page_number, exception_words, lang="uk"):
    params = {
        "text": text,
        "language": lang,
    }
    result_text = ""
    response = requests.post(LANGUAGETOOL_URL, data=params)

    if response.status_code == 200:
        matches = response.json().get("matches", [])

        if not matches:
            return ""  # No issues found

        abbreviations = extract_abbreviations(text)  # Get document-specific abbreviations

        for match in matches:
            rule_desc = match.get("message", "Unknown issue")
            error_word = match["context"]["text"][match["offset"]:match["offset"] + match["length"]].strip()

            # 🔹 Extract word from brackets in rule_desc
            suggested_word = extract_word_from_brackets(rule_desc)

            # If no word in brackets, fallback to LanguageTool's suggested correction
            if not suggested_word:
                suggested_replacements = match.get("replacements", [])
                if suggested_replacements:
                    suggested_word = suggested_replacements[0]["value"]
                else:
                    suggested_word = error_word  # Fallback to original if no suggestion

            if error_word in abbreviations or suggested_word in abbreviations:
                continue

            if not error_word:
                continue

            if error_word in exception_words:
                continue

            if any(is_similar(suggested_word, exception) for exception in exception_words):
                continue

            # Ignore capitalization mistakes after `;`**
            if match["rule"]["id"] == "UPPERCASE_SENTENCE_START":
                before_offset = text[:match["offset"]].strip()
                if before_offset and before_offset[-1] == ";":
                    continue  # Ignore capitalization mistake if previous sentence ends with ';'

                if error_word.lower() == error_word:  # Word starts in lowercase (e.g., a verb)
                    continue

            if (rule_desc == "Знайдено потенційну орфографічну помилку." or rule_desc== "Це слово є жаргонним") and not error_word.strip():
                continue  # Skip if the error word is empty

            sentence_part = ' '.join(text.split()[:5]) + "..." if len(text.split()) > 5 else text

            result_text += f"Issue on page {page_number}: {rule_desc} → '{error_word}' (suggested: {suggested_word}) in sentence: {sentence_part}\n"
    else:
        result_text += f"Error: Unable to reach LanguageTool API (status: {response.status_code})"

    return result_text

def check_document_spelling(doc, exception_words):
    in_appendices = False
    content_started = False
    result_text = []
    for paragraph in doc.Paragraphs:
        try:
            text = paragraph.Range.Text.strip()
            page_num = paragraph.Range.Information(3)  # Page number from Range.Information(3)

            if not text:
                continue  # Skip empty rows

            if not content_started:
                if "ЗМІСТ" in text.upper():
                    content_started = True
                continue

            if text.upper().strip() == "ДОДАТКИ":
                in_appendices = True
                continue

            if in_appendices:
                continue  # Skip everything after "ДОДАТКИ"

            if paragraph.Range.Tables.Count > 0:
                continue

            checked_row = check_spelling(text, page_num, exception_words)
            if checked_row:
                result_text.append(checked_row)

        except Exception as e:
            print(e)
            continue

    return result_text if result_text else ["No grammar errors found"]