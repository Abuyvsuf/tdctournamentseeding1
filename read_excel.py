"""
read_excel.py
--------------
Reads a buyer's Excel sheet into the `entries` format pool_coder.py expects.

Expected columns (header row, any order, case-insensitive):
    School      - school name
    Category    - e.g. "Grade 10", "Senior"
    Teams       - number of teams that school is bringing in that category

Extra columns are ignored. Blank rows are skipped.
"""

import openpyxl


def read_entries_from_excel(file_path_or_buffer):
    wb = openpyxl.load_workbook(file_path_or_buffer, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], ["The sheet appears to be empty."]

    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    try:
        school_idx = header.index("school")
        category_idx = header.index("category")
        teams_idx = header.index("teams")
    except ValueError:
        return [], [
            "Couldn't find columns named 'School', 'Category', and 'Teams' "
            "in the first row. Found: " + ", ".join(h for h in header if h)
        ]

    entries = []
    errors = []
    for row_num, row in enumerate(rows[1:], start=2):
        if row is None or all(c is None for c in row):
            continue
        school = row[school_idx] if school_idx < len(row) else None
        category = row[category_idx] if category_idx < len(row) else None
        teams_raw = row[teams_idx] if teams_idx < len(row) else None

        if school is None or category is None or teams_raw is None:
            errors.append(f"Row {row_num}: missing School, Category, or Teams value — skipped.")
            continue
        try:
            teams = int(teams_raw)
        except (TypeError, ValueError):
            errors.append(f"Row {row_num}: '{teams_raw}' isn't a whole number — skipped.")
            continue
        if teams < 1:
            errors.append(f"Row {row_num}: team count must be at least 1 — skipped.")
            continue

        entries.append({
            "school": str(school).strip(),
            "category": str(category).strip(),
            "teams": teams,
        })

    return entries, errors
