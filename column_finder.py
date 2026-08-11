"""
Shared utilities for locating and extracting columns from raw Salesforce/
Power BI Excel exports.

The old approach across these cleaners was to find one "first" column by
exact name, then grab everything to its right up to some assumed blank gap
or "last" column. That breaks constantly because real exports vary: gaps
move, column names get renamed slightly between report versions, extra
columns get added or removed.

This module instead finds each needed column independently, anywhere in
the header row, by keyword -- so column order, gaps, and minor naming
differences (case, "Mailing" vs "Billing" vs no prefix, etc.) don't matter.
"""

import re

import pandas as pd


def normalize_header_value(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    # Collapse any run of whitespace (including double spaces, which show
    # up a lot in these Salesforce exports) down to a single space.
    return re.sub(r"\s+", " ", value)


def find_column(normalized_row, *keywords, max_length=60):
    """Return the index of the first cell in this row whose (lowercased)
    text contains any of the given keywords. Cells longer than max_length
    are skipped -- this avoids matching Salesforce's long filter-description
    sentences (e.g. 'Primary Campaign Source equals "X","Y","Z"') which can
    coincidentally contain the same keywords as the real header."""
    for column_number, value in enumerate(normalized_row):
        if len(value) > max_length:
            continue
        value_lower = value.lower()
        for keyword in keywords:
            if keyword in value_lower:
                return column_number
    return None


def find_header_row(raw_df, column_specs):
    """Locate the header row and the column index for each field this
    cleaner needs.

    column_specs: dict of {output_name: (keywords_tuple, required_bool)}
    e.g. {"City": (("city",), True), "Notes": (("notes",), False)}

    Returns (row_number, {output_name: column_index_or_None}).
    Raises ValueError with a specific, actionable message if any required
    field can't be found in any row.

    A row only counts as "the header row" if every REQUIRED field is
    found in it -- optional fields are matched within that same row if
    present, but never block header-row detection on their own.
    """
    best_attempt = None  # (row_number, found_dict, missing_list) with fewest missing

    for row_number, row in raw_df.iterrows():
        normalized_row = [normalize_header_value(v) for v in row]

        found = {}
        for output_name, (keywords, required) in column_specs.items():
            found[output_name] = find_column(normalized_row, *keywords)

        missing_required = [
            name for name, (_, required) in column_specs.items()
            if required and found[name] is None
        ]

        if not missing_required:
            return row_number, found

        if best_attempt is None or len(missing_required) < len(best_attempt[2]):
            best_attempt = (row_number, found, missing_required)

    if best_attempt is not None:
        row_number, found, missing_required = best_attempt
        found_names = [name for name, col in found.items() if col is not None]
        raise ValueError(
            f"Could not find a header row with all required columns. "
            f"Closest match was row {row_number + 1}, missing: "
            f"{', '.join(missing_required)}. "
            f"Columns it did find: {', '.join(found_names) if found_names else 'none'}."
        )

    raise ValueError("The uploaded file appears to be empty.")


def extract_dataframe(raw_df, header_row_number, columns):
    """Given a row number and a {output_name: column_index_or_None} dict,
    build a DataFrame with clean, renamed columns. Fields whose index is
    None (optional and not found) become an all-blank column so downstream
    code can rely on the column always existing."""
    data = raw_df.iloc[header_row_number + 1:].reset_index(drop=True)

    df = pd.DataFrame({
        name: (data.iloc[:, col_index] if col_index is not None else pd.NA)
        for name, col_index in columns.items()
    })

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=0, how="all")
    df = df.reset_index(drop=True)

    return df


def clean_export(input_path, column_specs, sheet_name=0):
    """Convenience wrapper: read an Excel file and return a cleanly-named
    DataFrame using the column_specs described in find_header_row()."""
    raw_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)
    header_row_number, columns = find_header_row(raw_df, column_specs)
    return extract_dataframe(raw_df, header_row_number, columns)


def find_header_bounds(raw_df, anchor_keywords, max_blank_run=1):
    """For cases that need to preserve an entire contiguous range of
    columns (not just a fixed set of named fields) -- e.g. a raw cash or
    bulk-distribution export where every column should pass through to
    the output. Finds the header row by an anchor column (matched by
    keyword), then scans rightward from there, tolerating up to
    max_blank_run consecutive blank cells before deciding the row has
    ended.

    This replaces the old pattern of hardcoding which specific column
    index is allowed to be blank -- that breaks the moment a different
    export has its gap in a different place (this exact bug has caused
    real production failures).

    Returns (header_row_number, first_column_number, last_column_number).
    """
    for row_number, row in raw_df.iterrows():
        normalized_row = [normalize_header_value(v) for v in row]

        first_column_number = find_column(normalized_row, *anchor_keywords)
        if first_column_number is None:
            continue

        last_column_number = first_column_number
        blank_run = 0

        for column_number in range(first_column_number, len(normalized_row)):
            if normalized_row[column_number] == "":
                blank_run += 1
                if blank_run > max_blank_run:
                    break
                continue

            blank_run = 0
            last_column_number = column_number

        return row_number, first_column_number, last_column_number

    raise ValueError(
        f"Could not find a header row containing a column matching: "
        f"{', '.join(anchor_keywords)}."
    )


def clean_export_range(input_path, anchor_keywords, max_blank_run=1, sheet_name=0):
    """Convenience wrapper around find_header_bounds() -- reads the file,
    locates the header row/range by anchor keyword, and returns a cleanly
    named DataFrame with every column in that range preserved."""
    raw_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    header_row_number, first_col, last_col = find_header_bounds(
        raw_df, anchor_keywords, max_blank_run=max_blank_run
    )

    df = raw_df.iloc[header_row_number:, first_col:last_col + 1].reset_index(drop=True)
    df.columns = [normalize_header_value(v) for v in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=0, how="all")
    df = df.reset_index(drop=True)

    return df


def detect_file_type(input_path, type_signatures, sheet_name=0):
    """For multi-file cleaners that currently rely on filename matching:
    identify which known file type an upload actually is, by checking
    which type's required columns are present in its header row.

    type_signatures: dict of {type_name: column_specs} (same shape as
    find_header_row's column_specs, but every field should be required
    here since this is used purely for identification).

    Returns the type_name that matches, or None if no type's required
    columns are all present (ambiguous/unrecognized file).
    """
    raw_df = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

    matches = []
    for type_name, column_specs in type_signatures.items():
        try:
            find_header_row(raw_df, column_specs)
            matches.append(type_name)
        except ValueError:
            continue

    if len(matches) == 1:
        return matches[0]
    return None  # none matched, or more than one matched (ambiguous)
