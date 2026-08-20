from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import geography_maps
import column_finder as cf

COLUMN_SPECS = {
    "Payment Amount": (("payment amount", "amount received", "amount"), True),
    "Date": (("close date", "payment date", "date"), True),
    "City": (("city",), True),
    "State": (("state", "province"), True),
    "Zipcode": (("zip", "postal"), True),
    "Campaign Source": (("campaign source", "subhead", "campaign name", "campaign"), False),
}


def clean_geography_and_time_dataframe(input_path):
    raw_df = pd.read_excel(input_path, sheet_name=0, header=None)
    header_row_number, mapped_columns = cf.find_header_row(raw_df, COLUMN_SPECS)

    # Which column index maps to which canonical name we need (Amount,
    # Date, City, State, Zipcode, Campaign Source).
    canonical_by_index = {
        col_index: name for name, col_index in mapped_columns.items()
        if col_index is not None
    }

    header_row = raw_df.iloc[header_row_number]
    data = raw_df.iloc[header_row_number + 1:].reset_index(drop=True)

    # Keep EVERY non-blank column from the header row -- not just the ones
    # this cleaner needs for geography/time assignment. Anything extra
    # (donor name, email, phone, contact ID, etc.) passes straight
    # through to the output untouched.
    columns = {}
    for col_index, raw_header_value in enumerate(header_row):
        header_text = cf.normalize_header_value(raw_header_value)

        if col_index in canonical_by_index:
            output_name = canonical_by_index[col_index]
        elif header_text != "":
            output_name = header_text
        else:
            continue  # truly blank spacer column, not needed and not named

        if output_name in columns:
            continue  # keep the first occurrence if a name repeats

        columns[output_name] = data.iloc[:, col_index]

    df = pd.DataFrame(columns)
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

    # Exclude any Total/Subtotal rows Salesforce may have left in the
    # export -- these show up as literal "Total"/"Subtotal" text in one
    # of the group-by columns instead of real data.
    exclude_values = {"total", "subtotal", "grand total"}
    for column in ["Campaign Source", "City", "State"]:
        if column in df.columns:
            is_total_row = df[column].astype(str).str.strip().str.lower().isin(exclude_values)
            df = df[~is_total_row]
    df = df.reset_index(drop=True)

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
    standard_columns = [c for c in column_order if c in df.columns]
    extra_columns = [c for c in df.columns if c not in column_order]
    df = df[standard_columns + extra_columns]

    save_geography_and_time_workbook(df, output_path)
