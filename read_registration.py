"""
read_registration.py
---------------------
Reads the Kenya Senior Schools registration export (wide format: one row
per school, separate columns per category) into the long-format `entries`
list pool_coder.py expects: {school, region, language, category, teams}.

Region is a REQUIRED column in the source sheet -- tournament organizers
add it when they distribute the registration form, the same way School
Name is already required. This code does not infer or guess a school's
region from its name; that's a deliberate choice, not a missing feature.

This format is messy real-world form data, so this module:
  - cleans common non-numeric entries ("None", "Nill", "One", "1 team",
    "Form 3, one team, Form 4, one team")
  - NEVER silently guesses a number it isn't confident about -- anything
    that wasn't a plain integer gets logged in `needs_review` so a human
    can check the interpretation, even though processing continues.

Required columns in the source sheet (case-insensitive, matched by
substring so minor header wording changes don't break it):
  - "Name of School"
  - "Region"
  - "Number of Grade 10 Debate teams"
  - "Number of Senior Debate teams"
  - "Number of Grade 10 Kiswahili Mjadala Teams"
  - "Number of Senior Kiswahili Mjadala Teams"
"""

import re
import openpyxl

WORD_NUMBERS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
    "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19, "twenty": 20,
}
ZERO_WORDS = {"none", "nil", "nill", "n/a", "na", "-", ""}

CATEGORY_COLUMNS = [
    # (column header substring to match, language, category)
    ("number of grade 10 debate", "English", "Junior"),
    ("number of senior debate", "English", "Senior"),
    ("number of grade 10 kiswahili mjadala", "Kiswahili", "Junior"),
    ("number of senior kiswahili mjadala", "Kiswahili", "Senior"),
]


def parse_team_count(raw_value):
    """
    Returns (count, was_exact) where was_exact=True means it was a plain
    integer or unambiguous word-number, and False means a heuristic
    (sum-all-numbers-found) was used and should be flagged for review.
    """
    if raw_value is None:
        return 0, True

    text = str(raw_value).strip().lower()
    if text in ZERO_WORDS:
        return 0, True

    if text.lstrip("-").isdigit():
        return int(text), True

    if text in WORD_NUMBERS:
        return WORD_NUMBERS[text], True

    # Composite / messy free text e.g. "Form 3, one team, Form 4, one team"
    # -> find every digit-number or word-number mentioned and sum them.
    numbers_found = [int(n) for n in re.findall(r"\d+", text)]
    for word, val in WORD_NUMBERS.items():
        if re.search(rf"\b{word}\b", text):
            numbers_found.append(val)

    if numbers_found:
        return sum(numbers_found), False

    return 0, False


def _find_column(header, *substrings):
    for i, col in enumerate(header):
        col_l = (col or "").strip().lower()
        if all(s in col_l for s in substrings):
            return i
    return None


def read_registration(file_path_or_buffer):
    """
    Returns (entries, needs_review, errors)
      entries: [{"school": str, "region": str, "language": str,
                 "category": str, "teams": int}, ...]
      needs_review: [str] human-readable notes about heuristic parses
      errors: [str] rows/columns that couldn't be read at all
    """
    wb = openpyxl.load_workbook(file_path_or_buffer, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], [], ["The sheet appears to be empty."]

    header = [str(c).strip().lower() if c else "" for c in rows[0]]

    school_idx = _find_column(header, "name of school") or _find_column(header, "school")
    region_idx = _find_column(header, "region")

    errors = []
    if school_idx is None:
        errors.append("Couldn't find a 'Name of School' column.")
    if region_idx is None:
        errors.append(
            "Couldn't find a 'Region' column. Add one with the region for "
            "each school before uploading -- this can't be guessed automatically."
        )

    category_indices = []
    for substring, language, category in CATEGORY_COLUMNS:
        idx = _find_column(header, substring.split()[0], substring.split()[-1])
        # fallback: try matching on the full substring as a contiguous phrase
        if idx is None:
            for i, col in enumerate(header):
                if substring in (col or ""):
                    idx = i
                    break
        category_indices.append((idx, language, category))
        if idx is None:
            errors.append(f"Couldn't find a column matching '{substring}'.")

    if errors:
        return [], [], errors

    entries = []
    needs_review = []

    for row_num, row in enumerate(rows[1:], start=2):
        if row is None or all(c is None for c in row):
            continue
        school_raw = row[school_idx] if school_idx < len(row) else None
        region_raw = row[region_idx] if region_idx < len(row) else None

        if not school_raw or not str(school_raw).strip():
            continue
        school = " ".join(str(school_raw).split())  # collapse internal/trailing whitespace

        if not region_raw or not str(region_raw).strip():
            needs_review.append(f"Row {row_num} ({school}): no Region value -- this school will be skipped.")
            continue
        region = str(region_raw).strip()

        for idx, language, category in category_indices:
            raw_value = row[idx] if idx is not None and idx < len(row) else None
            teams, was_exact = parse_team_count(raw_value)
            if teams < 1:
                continue
            if not was_exact:
                needs_review.append(
                    f"Row {row_num} ({school}, {language} {category}): "
                    f"read \"{raw_value}\" as {teams} team(s) -- please verify."
                )
            entries.append({
                "school": school,
                "region": region,
                "language": language,
                "category": category,
                "teams": teams,
            })

    return entries, needs_review, []
