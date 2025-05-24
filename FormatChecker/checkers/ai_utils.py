import requests
import re
import Levenshtein

LANGUAGETOOL_URL = "https://api.languagetool.org/v2/check"

SIMILARITY_THRESHOLD = 85

def extract_abbreviations(text):
    return set(re.findall(r"\b[А-ЯЇЄҐ]{2,}\b", text))

def clean_word(word):
    # Remove common non-alphabetic characters (e.g. dashes, punctuation, spaces)
    return re.sub(r"[^\wа-яА-ЯїЇєЄґҐ]", "", word).lower()

def is_similar(word1, word2):
    word1 = clean_word(word1.lower())
    word2 = clean_word(word2.lower())
    dist = Levenshtein.distance(word1, word2)
    max_len = max(len(word1), len(word2))

    if max_len <= 4:
        max_dist = 1
    elif max_len <= 7:
        max_dist = 2
    else:
        max_dist = 3
    return dist <= max_dist

def extract_word_from_brackets(text):
    match = re.search(r"«(.+?)»", text)
    return match.group(1) if match else None

def check_spelling(text, page_number, exception_words, lang="uk"):
    response = requests.post(LANGUAGETOOL_URL, data={"text": text, "language": lang})
    if response.status_code != 200:
        return f"Error: Unable to reach LanguageTool API (status: {response.status_code})"

    matches = response.json().get("matches", [])
    if not matches:
        return ""

    abbreviations = extract_abbreviations(text)
    result_lines = []

    for match in matches:
        rule_id = match.get("rule", {}).get("id", "")
        rule_desc = match.get("message", "Unknown issue")
        ctx = match.get("context", {})
        error_word = ctx.get("text", "")[match.get("offset", 0):match.get("offset", 0) + match.get("length", 0)].strip()

        # Determine suggested word
        suggested_word = extract_word_from_brackets(rule_desc)
        if not suggested_word:
            suggested_word = match.get("replacements", [{}])[0].get("value", error_word)

        # Skip checks
        if not error_word:
            continue
        if error_word in abbreviations or suggested_word in abbreviations:
            continue
        if error_word in exception_words or any(is_similar(suggested_word, ex) for ex in exception_words):
            continue
        if rule_id == "UPPERCASE_SENTENCE_START":
            before = text[:match.get("offset", 0)].strip()
            if before.endswith(";") or error_word.islower():
                continue
        if rule_desc in {"Знайдено потенційну орфографічну помилку.", "Це слово є жаргонним"} and not error_word:
            continue

        snippet = ' '.join(text.split()[:5]) + "..." if len(text.split()) > 5 else text
        result_lines.append(f"Issue on page {page_number}: {rule_desc} → '{error_word}' (suggested: {suggested_word}) in sentence: {snippet}")

    return "\n".join(result_lines)

def check_document_spelling(doc, exception_words):
    content_started = False
    result_text = []
    for paragraph in doc.Paragraphs:
        try:
            text = paragraph.Range.Text.strip()
            page_num = paragraph.Range.Information(3)  # Page number from Range.Information(3)

            if not text:
                continue

            elif paragraph.Range.Tables.Count > 0:
                continue

            elif not content_started:
                if "ЗМІСТ" in text.upper():
                    content_started = True
                continue

            if text.upper().strip() == "ДОДАТКИ":
                break

            checked_row = check_spelling(text, page_num, exception_words)
            if checked_row:
                result_text.append(checked_row)

        except Exception as e:
            print(e)
            continue

    return result_text if result_text else ["No grammar errors found"]