from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import geography_maps


def normalize_header_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_column(normalized_row, *keywords):
    """Return the index of the first cell in this header row whose
    (lowercased) text contains any of the given keywords. Real header
    cells are short labels; long sentences (like the Salesforce filter
    description, which can also contain these words) are skipped."""
    for column_number, value in enumerate(normalized_row):
        if len(value) > 60:
            continue
        value_lower = value.lower()
        for keyword in keywords:
            if keyword in value_lower:
                return column_number
    return None


def find_header_row(df):
    """Locate the row containing the Campaign Source column and return
    that row number plus the column index for every field we need,
    regardless of ordering or gaps between them."""
    for row_number, row in df.iterrows():
        normalized_row = [normalize_header_value(v) for v in row]

        campaign_col = find_column(normalized_row, "campaign source")
        if campaign_col is None:
            continue

        amount_col = find_column(normalized_row, "payment amount", "amount received")
        date_col = find_column(normalized_row, "close date", "payment date")
        city_col = find_column(normalized_row, "city")
        state_col = find_column(normalized_row, "state", "province")
        zip_col = find_column(normalized_row, "zip", "postal")

        found = {
            "Campaign Source": campaign_col,
            "Payment Amount": amount_col,
            "Date": date_col,
            "City": city_col,
            "State": state_col,
            "Zipcode": zip_col,
        }

        missing = [name for name, col in found.items() if col is None]
        if missing:
            raise ValueError(
                f"Found the header row (Campaign Source at column {campaign_col}) "
                f"but could not find a column for: {', '.join(missing)}. "
                f"Header row contents: {normalized_row}"
            )

        return row_number, found

    raise ValueError(
        "Could not find a header row containing a 'Campaign Source' column."
    )


def clean_geography_and_time_dataframe(input_path):
    input_path = Path(input_path)

    raw = pd.read_excel(input_path, sheet_name=0, header=None)

    header_row_number, columns = find_header_row(raw)

    data = raw.iloc[header_row_number + 1:].reset_index(drop=True)

    df = pd.DataFrame({
        name: data.iloc[:, col_index]
        for name, col_index in columns.items()
    })

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=0, how="all")
    df = df.reset_index(drop=True)

    return df


def save_geography_and_time_workbook(df, output_path):
    output_path = Path(output_path)

    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("Geography and Time Cleaner")

    df_to_save = df.copy()
    df_to_save = df_to_save.astype(object)
    df_to_save = df_to_save.where(pd.notna(df_to_save), None)

    for row in dataframe_to_rows(df_to_save, index=False, header=True):
        worksheet.append(row)

    workbook.save(output_path)


def clean_multiple_geography_and_time_imports(input_paths, output_path):
    output_path = Path(output_path)

    cleaned_dataframes = [clean_geography_and_time_dataframe(p) for p in input_paths]

    if len(cleaned_dataframes) == 0:
        raise ValueError("No files were provided.")

    df = pd.concat(cleaned_dataframes, ignore_index=True)

    # "Campaign Source" is the Salesforce grouped-report column -- only
    # populated on the first row of each campaign group, blank on the rest.
    df["Campaign Source"] = df["Campaign Source"].ffill()

    df["Zipcode"] = df["Zipcode"].astype(str).str.extract(r"(\d{1,5})")[0].str.zfill(5)
    df["State"] = df["State"].astype(str).str.upper().str.strip()

    df = df.apply(geography_maps.hq_states, axis=1)

    df["Region"] = df["State"].map(geography_maps.state_to_region)
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)
    df["Field Office"] = df["State"].map(geography_maps.state_to_field_office)
    df = df.apply(lambda row: geography_maps.assign_field_office(row, True), axis=1)
    df["Chapter"] = df.apply(geography_maps.assign_chapter, axis=1)
    df["Region"] = df["Region"].str.upper()

    df = geography_maps._add_date_columns(df)

    df["Payment Amount"] = pd.to_numeric(df["Payment Amount"], errors="coerce").fillna(0)

    df = df.rename(columns={"Region Number": "RSN"})

    column_order = [
        "Year", "Quarter", "Month Name", "Month Number", "Date",
        "Region", "RSN", "Chapter", "State", "Field Office",
        "City", "Zipcode", "Campaign Source", "Payment Amount",
    ]
    existing_columns = [c for c in column_order if c in df.columns]
    df = df[existing_columns]

    save_geography_and_time_workbook(df, output_path)
