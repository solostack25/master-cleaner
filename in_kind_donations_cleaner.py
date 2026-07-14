from pathlib import Path

import pandas as pd
import geography_maps


COLUMNS_TO_KEEP = 9


def _cell_text(value) -> str:
    """
    Convert a value into clean comparable text.
    """
    if pd.isna(value):
        return ""

    return str(value).strip()


def _find_header_row(input_path: Path) -> int:
    """
    Find the row where the real table headers begin.

    Returns a zero-based row index for pandas.
    Example:
        Excel row 11 becomes pandas header index 10.
    """
    preview_df = pd.read_excel(
        input_path,
        header=None,
        dtype=str,
        nrows=30
    )

    for index, row in preview_df.iterrows():
        row_text = " | ".join(_cell_text(value).lower() for value in row)

        has_office_column = "icna relief office" in row_text
        has_program_column = "program" in row_text
        has_date_received_column = "date received" in row_text
        has_inkind_id_column = "inkind id" in row_text

        if (
            has_office_column
            and has_program_column
            and has_date_received_column
            and has_inkind_id_column
        ):
            return index

    raise ValueError(
        "Could not find the In-Kind Donations header row. "
        "Expected headers like ICNA Relief Office, Program, Date Received, and InKind ID."
    )


def in_kind_special_geography(row):
    office_account_name = row["ICNA Relief Office: Account Name  ↑"]

    if "kansas city" in office_account_name.lower():
        row['Region'] = 'South Central'
        row['Region Number'] = 5
        row['Field Office'] = "Kansas City Office"
        row["State"] = 'KS'
    elif "tampa" in office_account_name.lower():
        row['Region'] = 'Southeast'
        row['Region Number'] = 6
        row['Field Office'] = "Tampa Office"
    elif "bay area" in office_account_name.lower():
        row['Field Office'] = "Bay Area Office"
    elif "fullerton" in office_account_name.lower():
        row['Field Office'] = "Los Angeles Office"
    elif "sacramento" in office_account_name.lower():
        row['Field Office'] = "Sacramento Office"
    elif "san diego" in office_account_name.lower():
        row['Field Office'] = "San Diego Office"
    elif "walnut" in office_account_name.lower():
        row['Field Office'] = "Los Angeles Office"
    elif "central florida" in office_account_name.lower():
        row['Field Office'] = "Orlando Office"
    elif "dallas" in office_account_name.lower():
        row['Field Office'] = "Dallas Office"
    elif "fort worth" in office_account_name.lower():
        row['Field Office'] = "Dallas Office"
    elif "houston" in office_account_name.lower():
        row['Field Office'] = "Houston Office"
    elif "austin" in office_account_name.lower():
        row['Field Office'] = "Austin Office"
    elif "alexandria" in office_account_name.lower():
        row['Field Office'] = "Alexandria Office"
    elif "richmond" in office_account_name.lower():
        row['Field Office'] = "Richmond Office"
    elif row['State'] not in geography_maps.state_to_region:
        row['Field Office'] = "Unassigned"
        row['State'] = "HQ"
        row["Region Number"] = 8
        row['Region'] = "Unassigned"

    return row

def in_kind_donations_zero_value_rows(df):
    new_dfs = []

    programs = [
        "Hunger Prevention",
        "General",
        "Refugee Services",
        "Transitional Housing",
        "Muslim Family Services",
        "Health Services",
        "Back 2 School",
        "FATE",
        "Disaster Relief"
    ]

    value_columns = [
        "TOTAL VALUE OF ARTICLES",
        "TYPES OF ARTICLES",
        "QUANTITY OF ARTICLES",
    ]

    # Make sure column names match.
    df.columns = df.columns.map(lambda col: str(col).strip().upper())

    # Copy the 732-row zero value geography/month dataframe.
    base_df = geography_maps.zero_value_df.copy()
    base_df.columns = base_df.columns.map(lambda col: str(col).strip().upper())

    # Get years that already exist in the data.
    years = sorted(df["YEAR"].dropna().astype(int).unique())

    for year in years:
        for program in programs:
            temp_df = base_df.copy()

            # Use the year from the actual uploaded data.
            temp_df["YEAR"] = year

            # If you want the DATE year to match the current year,
            # keep this part. It uses the month already in base_df.
            temp_df["DATE"] = pd.to_datetime(temp_df["DATE"], errors="coerce")
            temp_df["DATE"] = temp_df["DATE"].apply(
                lambda date_value: pd.Timestamp(
                    year=year,
                    month=date_value.month,
                    day=1
                ).date()
                if pd.notna(date_value)
                else date_value
            )

            # Set the program.
            temp_df["PROGRAM  ↑"] = program

            # Set the value columns to zero.
            for value_column in value_columns:
                temp_df[value_column] = 0

            # Make sure temp_df has every column that df has.
            for column in df.columns:
                if column not in temp_df.columns:
                    temp_df[column] = ""

            # Keep the same column order as df.
            temp_df = temp_df[df.columns]

            new_dfs.append(temp_df)

    if new_dfs:
        new_df = pd.concat(new_dfs, ignore_index=True)
        df = pd.concat([df, new_df], ignore_index=True)

    return df


def clean_in_kind_donations(input_path: Path, output_path: Path):
    input_path = Path(input_path)
    output_path = Path(output_path)

    header_row_index = _find_header_row(input_path)

    # Read the spreadsheet again, this time using the real header row.
    df = pd.read_excel(
        input_path,
        header=header_row_index
    )

    # Step 3: delete original column A.
    df = df.iloc[:, 1:]

    # Step 4: delete the current third column.
    # After deleting original column A, this removes the empty third column.
    df = df.drop(df.columns[2], axis=1)

    # Step 5: fill forward columns A and B.
    # These should now be the first two columns in the cleaned DataFrame.
    df.iloc[:, 0] = df.iloc[:, 0].ffill()
    df.iloc[:, 1] = df.iloc[:, 1].ffill()

    # Step 6: keep only the first 9 columns.
    df = df.iloc[:, :COLUMNS_TO_KEEP]

    df.rename(columns={"Date Received": "Date", "Total Value": "Total Value of Articles",
                       "Total Items": "Types of Articles", "Total Quantity of Items": "Quantity of Articles"},
              inplace=True)

    # Step 8 and 9: rename Date Received and add Year/Quarter/Month Name.
    df = geography_maps._add_date_columns(df)

    # Step 10: Geography

    df["State"] = df["ICNA Relief Office: Account Name  ↑"].str[:2]
    df['Region'] = df['State'].map(geography_maps.state_to_region)
    # Assign RSN
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)

    

    # Assign Field Office
    df["Field Office"] = df["State"].map(geography_maps.state_to_field_office)
    df = df.apply(in_kind_special_geography, axis=1)

    df["Chapter"] = df.apply(geography_maps.assign_chapter, axis=1)
    df['Region'] = df['Region'].str.upper()

    first_columns = ["Year", "Quarter", "Month Name", "Month Number", "Date", "Region", "Chapter", "State",
                     "Field Office", "Region Number"]

    remaining_columns = [
        column for column in df.columns
        if column not in first_columns
    ]

    df = df[first_columns + remaining_columns]

    df.columns = df.columns.str.upper()

    # Step 8: save as Excel.
    df.to_excel(output_path, index=False)

def clean_multiple_in_kind_donations(input_paths, output_path):
    cleaned_dataframes = []
    temporary_output_paths = []

    for file_number, input_path in enumerate(input_paths, start=1):
        temporary_output_path = Path(output_path).parent / f"{Path(output_path).stem}_temporary_cleaned_{file_number}.xlsx"

        # Use the existing single-file cleaner.
        clean_in_kind_donations(input_path, temporary_output_path)

        # Read the cleaned file back into pandas.
        cleaned_df = pd.read_excel(temporary_output_path)
        cleaned_dataframes.append(cleaned_df)

        temporary_output_paths.append(temporary_output_path)

    combined_df = pd.concat(cleaned_dataframes, ignore_index=True)

    # adds zero value rows
    combined_df = in_kind_donations_zero_value_rows(combined_df)

    # Capitalizes all regions
    combined_df["REGION"] = combined_df["REGION"].str.upper()

    combined_df.to_excel(output_path, index=False)
    geography_maps.format_currency_column(output_path, "TOTAL VALUE OF ARTICLES")

    # Delete temporary cleaned files.
    for temporary_output_path in temporary_output_paths:
        try:
            temporary_output_path.unlink()
        except Exception:
            pass