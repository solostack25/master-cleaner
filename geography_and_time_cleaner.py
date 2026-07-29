from pathlib import Path
import re

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import geography_maps

FIRST_HEADER = "Primary Campaign Source ↑"
LAST_HEADER = "Mailing Zip/Postal Code"


def normalize_header_value(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    return re.sub(r"\s+", " ", value)


def find_header_bounds(df):
    first_target = normalize_header_value(FIRST_HEADER)
    last_target = normalize_header_value(LAST_HEADER)

    for row_number, row in df.iterrows():
        normalized_row = [normalize_header_value(v) for v in row]

        if first_target in normalized_row:
            first_column_number = normalized_row.index(first_target)
            if last_target in normalized_row[first_column_number:]:
                last_column_number = normalized_row.index(last_target, first_column_number)
                return row_number, first_column_number, last_column_number

    raise ValueError(
        f"Could not find a header row with columns '{FIRST_HEADER}' through '{LAST_HEADER}'."
    )


def clean_geography_and_time_dataframe(input_path):
    input_path = Path(input_path)

    df = pd.read_excel(input_path, sheet_name=0, header=None)

    header_row_number, first_column_number, last_column_number = find_header_bounds(df)

    df = df.iloc[header_row_number:, first_column_number:last_column_number + 1]
    df = df.reset_index(drop=True)

    df.columns = [normalize_header_value(value) for value in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

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

    df = df.loc[:, df.columns != ""]

    df = df.rename(columns={
        "Primary Campaign Source ↑": "Campaign Source",
        "Payment Amount Received": "Payment Amount",
        "Close Date": "Date",
        "Mailing City": "City",
        "Mailing State/Province": "State",
        "Mailing Zip/Postal Code": "Zipcode",
    })

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
