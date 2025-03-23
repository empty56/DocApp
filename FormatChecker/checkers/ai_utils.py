import requests
import re
from rapidfuzz import fuzz

LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"

EXCEPTION_WORDS = {"вебзастосунок", "вебзастосунку", "вебзастосунки", "ЗДО",
                   "КПІ", "Формалізування", "Кросплатформеність", "існуючих"}

SIMILARITY_THRESHOLD = 85

def extract_abbreviations(text):
    return set(re.findall(r"\b[А-ЯЇЄҐ]{2,}\b", text))  # Matches 2+ uppercase Ukrainian letters

def get_similarity(word1, word2):
    similarity = fuzz.ratio(word1.lower(), word2.lower())
    return similarity >= SIMILARITY_THRESHOLD

def extract_word_from_brackets(text):
    match = re.search(r"«(.+?)»", text)
    return match.group(1) if match else None

def check_spelling(text, page_number, lang="uk"):
    params = {
        "text": text,
        "language": lang,
    }

    response = requests.post(LANGUAGETOOL_URL, data=params)

    if response.status_code == 200:
        matches = response.json().get("matches", [])

        if not matches:
            return None  # No issues found

        abbreviations = extract_abbreviations(text)  # Get document-specific abbreviations

        for match in matches:
            rule_desc = match.get("message", "Unknown issue")
            error_word = match["context"]["text"][match["offset"]:match["offset"] + match["length"]]

            # Extract word from brackets in rule_desc
            suggested_word = extract_word_from_brackets(rule_desc)

            # If no word in brackets, fallback to LanguageTool's suggested correction
            if not suggested_word:
                suggested_replacements = match.get("replacements", [])
                if suggested_replacements:
                    suggested_word = suggested_replacements[0]["value"]
                else:
                    suggested_word = error_word  # Fallback to original if no suggestion

            # Skip detected abbreviations**
            if error_word in abbreviations or suggested_word in abbreviations:
                continue

            # Skip if error_word is in EXCEPTION_WORDS**
            if error_word in EXCEPTION_WORDS:
                continue

            # Skip if suggested_word is similar to an EXCEPTION_WORD**
            if any(get_similarity(suggested_word, exception) for exception in EXCEPTION_WORDS):
                continue

            # RESTORED:** Ignore capitalization mistakes after `;`
            if match["rule"]["id"] == "UPPERCASE_SENTENCE_START":
                before_offset = text[:match["offset"]].strip()
                if before_offset and before_offset[-1] == ";":
                    continue  # Ignore capitalization mistake if previous sentence ends with ';'

                # Additional check to avoid false positives for lowercase verbs**
                if error_word.lower() == error_word:  # Word starts in lowercase (e.g., a verb)
                    continue

            # Skip "Знайдено потенційну орфографічну помилку" with empty error word**
            if rule_desc == "Знайдено потенційну орфографічну помилку." and not error_word.strip():
                continue  # Skip if the error word is empty

            # Output only page number and the first 5 words of the sentence
            sentence_part = ' '.join(text.split()[:5]) + "..." if len(text.split()) > 5 else text
            print(f"Grammar issue on page {page_number}: {rule_desc} → '{error_word}' (suggested: {suggested_word}) in sentence: {sentence_part}")

    else:
        print(f"Error: Unable to reach LanguageTool API (status: {response.status_code})")


def check_document_spelling(doc):
    in_appendices = False
    content_started = False

    for paragraph in doc.Paragraphs:
        text = paragraph.Range.Text.strip()
        page_num = paragraph.Range.Information(3)  # Page number from Range.Information(3)

        if not text:
            continue  # Skip empty rows

        # if "ЗМІСТ" in text.upper():  # Skip "ЗМІСТ" section
        #     continue

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

            # Pass page number and text to check_spelling function
        check_spelling(text, page_num)






