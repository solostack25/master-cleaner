from pathlib import Path
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import geography_maps
import column_finder as cf

HUNGER_PREVENTION_FILENAME_LOOKUP = {
    "reg_clients": "registered_clients",
    "registered_clients": "registered_clients",
    "bulk": "bulk",
    "cash": "cash",
    "in_kind": "in_kind",
    "irfas": "irfas"
}


REQUIRED_HUNGER_PREVENTION_FILES = [
    "registered_clients",
    "bulk",
    "cash",
    "in_kind",
    "irfas"
]


HUNGER_PREVENTION_DISPLAY_NAMES = {
    "registered_clients": "Registered Clients",
    "bulk": "Bulk",
    "cash": "Cash",
    "in_kind": "In-Kind",
    "irfas": "IRFAS"
}

def normalize_hunger_prevention_filename(original_filename):
    filename = Path(original_filename).name.strip().lower()

    file_stem = Path(filename).stem.strip()
    file_extension = Path(filename).suffix.lower()

    file_stem = file_stem.replace(" ", "_")
    file_stem = file_stem.replace("-", "_")

    return file_stem, file_extension


def parse_hunger_prevention_filename(original_filename):
    file_stem, file_extension = normalize_hunger_prevention_filename(original_filename)

    if file_extension != ".xlsx":
        raise ValueError(
            f"Hunger Prevention files must be .xlsx files. Invalid file: {original_filename}"
        )

    if file_stem not in HUNGER_PREVENTION_FILENAME_LOOKUP:
        raise ValueError(
            f"Could not recognize Hunger Prevention file: {original_filename}. "
            "Expected one of: reg_clients.xlsx, registered_clients.xlsx, bulk.xlsx, "
            "cash.xlsx, in-kind.xlsx, or irfas.xlsx."
        )

    return HUNGER_PREVENTION_FILENAME_LOOKUP[file_stem]

def normalize_header_value(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    # Replace weird spacing/newlines with one normal space
    value = re.sub(r"\s+", " ", value)

    return value

def clean_dataframe(input_path, first_column_name):
    """Kept as a thin wrapper for backwards compatibility with existing
    call sites -- first_column_name is now used as a keyword search
    (case/whitespace-insensitive substring match) rather than an exact
    string match, and the header-range detection tolerates blank gaps
    anywhere instead of assuming a specific column position is the one
    that's allowed to be empty (that assumption caused real failures)."""
    anchor_keyword = normalize_header_value(first_column_name).lower()
    # Strip the trailing arrow character some of these Salesforce columns
    # have (e.g. "State  ↑") since it's not a reliable keyword to match on.
    anchor_keyword = anchor_keyword.replace("↑", "").replace("↓", "").strip()

    return cf.clean_export_range(input_path, (anchor_keyword,), max_blank_run=2)


def validate_hunger_prevention_uploads(input_files, input_filenames):
    recognized_files = {}

    for file_num, file_info in enumerate(input_files):
        original_filename = input_filenames[file_num]

        file_type = parse_hunger_prevention_filename(original_filename)

        if file_type in recognized_files:
            display_name = HUNGER_PREVENTION_DISPLAY_NAMES[file_type]

            raise ValueError(
                f"Duplicate Hunger Prevention file found for {display_name}. "
                f"Only upload one {display_name} file."
            )

        recognized_files[file_type] = file_info

    missing_file_types = []

    for required_file_type in REQUIRED_HUNGER_PREVENTION_FILES:
        if required_file_type not in recognized_files:
            missing_file_types.append(required_file_type)

    if len(missing_file_types) > 0:
        missing_display_names = []

        for file_type in missing_file_types:
            missing_display_names.append(HUNGER_PREVENTION_DISPLAY_NAMES[file_type])

        raise ValueError(
            "Missing required Hunger Prevention file(s): "
            + ", ".join(missing_display_names)
        )

    return recognized_files

def save_hunger_prevention_starter_workbook(df1, output_path):
    output_path = Path(output_path)

    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("Hunger Prevention Data")

    df_to_save = df1.copy()
    df_to_save = df_to_save.astype(object)
    df_to_save = df_to_save.where(pd.notna(df_to_save), None)

    for row in dataframe_to_rows(df_to_save, index=False, header=True):
        worksheet.append(row)

    workbook.save(output_path)

def zero_value_rows(df):
    new_dfs = []

    distribution_types = ["IRFAS", "In-Kind", "Cash", "Registered Clients", "Bulk"]

    first_columns = ["Year", "Quarter", "Month Name", "Month Number", "Date", "Region", "Chapter", "State",
                     "Field Office", "Region Number", "City", "Zipcode", "Location", "Distribution Type"]

    remaining_columns = [
        column for column in df.columns
        if column not in first_columns
    ]

    # Make sure column names match.
    df.columns = df.columns.map(lambda col: str(col).strip().upper())

    # Copy the 732-row zero value geography/month dataframe.
    base_df = geography_maps.zero_value_df.copy()
    base_df.columns = base_df.columns.map(lambda col: str(col).strip().upper())

    # Get years that already exist in the data.
    years = sorted(df["YEAR"].dropna().astype(int).unique())

    for year in years:
        for distribution_type in distribution_types:
            temp_df = base_df.copy()

            # Use the year from the actual uploaded data.
            temp_df["YEAR"] = year

            # If you want the DATE year to match the current year,
            # keep this part. It uses the month already in base_df.
            temp_df["DATE"] = pd.to_datetime(temp_df["DATE"], errors="coerce")
            temp_df["DATE"] = temp_df["DATE"].apply(lambda date_value: pd.Timestamp(year=year, month=date_value.month,
                    day=1).date()
                if pd.notna(date_value)
                else date_value
            )

            # Set the distribution type.
            temp_df["DISTRIBUTION TYPE"] = distribution_type

            # Set the value columns to zero.
            for value_column in remaining_columns:
                temp_df[value_column.upper()] = 0

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



def clean_multiple_hunger_prevention_imports(input_files, input_filenames, output_path):
    recognized_files = validate_hunger_prevention_uploads(input_files, input_filenames)

    for file_name, file_info in recognized_files.items():
        print(file_name, file_info)
    registered_clients_df = clean_dataframe(recognized_files["registered_clients"], "State  ↑")
    bulk_df = clean_dataframe(recognized_files["bulk"], 'ICNA Relief Office: Account Name  ↑')
    cash_df = clean_dataframe(recognized_files["cash"], 'Primary Campaign Source  ↑')
    in_kind_df = pd.read_excel(recognized_files["in_kind"], sheet_name=0)
    irfas_df = pd.read_excel(recognized_files["irfas"], sheet_name=0)

    # change names of columns
    registered_clients_df = registered_clients_df.rename(columns={"State ↑": "State", "Date Distributed ↑": "Date",
                                                                  "Office Location ↓": "Location"})
    bulk_df = bulk_df.rename(columns={"Date Distributed": "Date",
                                      "ICNA Relief Office: Account Name ↑": "Location",
                                      "ICNA Relief Office: Billing State/Province": "State",
                                      "ICNA Relief Office: Billing City": "City",
                                      "ICNA Relief Office: Billing Zip/Postal Code": "Zipcode"})
    cash_df = cash_df.rename(columns={"Mailing State/Province": "State", "Mailing City": "City",
                                      "Mailing Zip/Postal Code": "Zipcode", "Close Date": "Date"})
    in_kind_df = in_kind_df.rename(columns={"VALUE": "TOTAL VALUE OF ARTICLES", "DATE RECEIVED": "DATE"})
    irfas_df = irfas_df.rename(columns={"MONTH": "MONTH NAME", "MSN": "MONTH NUMBER", "VALUE": "TOTAL VALUE OF ARTICLES",
                                        "RSN": "REGION NUMBER", "DECISION DATE": "DATE", "LOCATION": "CITY"})

    # Forward Fills
    print(registered_clients_df.columns)
    print(bulk_df.columns)
    print(cash_df.columns)
    registered_clients_df["State"] = registered_clients_df["State"].ffill()
    registered_clients_df["Location"] = registered_clients_df["Location"].ffill()
    bulk_df["Location"] = bulk_df["Location"].ffill()

    # get the state and Zipcode columns in propper order
    registered_clients_df["State"] = registered_clients_df["State"].str.upper()
    bulk_df["State"] = bulk_df["State"].str.upper()
    cash_df["State"] = cash_df["State"].str.upper()

    bulk_df["Zipcode"] = bulk_df["Zipcode"].astype(str).str.extract(r'(\d{1,5})')[0].str.zfill(5)
    cash_df["Zipcode"] = cash_df["Zipcode"].astype(str).str.extract(r'(\d{1,5})')[0].str.zfill(5)

    registered_clients_df = registered_clients_df.apply(geography_maps.hq_states, axis=1)
    bulk_df = bulk_df.apply(geography_maps.hq_states, axis=1)
    cash_df = cash_df.apply(geography_maps.hq_states, axis=1)

    # Assign Field Office
    registered_clients_df["Field Office"] = registered_clients_df["State"].map(geography_maps.state_to_field_office)
    bulk_df["Field Office"] = bulk_df["State"].map(geography_maps.state_to_field_office)
    cash_df["Field Office"] = cash_df["State"].map(geography_maps.state_to_field_office)

    # For Registered Clients Only, Assigns the Field Office For 5 Multi Field Office States
    multi_field_office_states_location_to_field_office = {"CA - Bay Area": "Bay Area Office",
                                                          "CA - Fullerton": "Los Angeles Office",
                                                          "CA - Walnut": "Los Angeles Office",
                                                          "CA - San Diego": "San Diego Office",
                                                          "CA - Sacramento": "Sacramento Office",
                                                          "FL - South Florida": "Miami Office",
                                                          "FL - Orlando": "Orlando Office",
                                                          "FL - Tampa": "Tampa Office",
                                                          "MO - Kansas City": "Kansas City Office",
                                                          "MO - St. Louis": "St. Louis Office",
                                                          "TX - Houston": "Houston Office",
                                                          "TX - Austin": "Austin Office",
                                                          "TX - Dallas": "Dallas Office",
                                                          "TX - Oak Cliff Dallas": "Dallas Office",
                                                          "VA - Alexandria": "Alexandria Office",
                                                          "VA - Richmond": "Richmond Office"}

    registered_clients_df["Field Office"] = registered_clients_df["Field Office"].fillna(
        registered_clients_df["Location"].map(multi_field_office_states_location_to_field_office))

    # Add " Office" to the end of each value
    registered_clients_df['Location'] = registered_clients_df['Location'].astype(str) + ' Office'

    registered_clients_df.loc[registered_clients_df["Field Office"] == "Kansas City Office", "State"] = "KS"


    # Assign Field Office
    bulk_df = bulk_df.apply(lambda row: geography_maps.assign_field_office(row, True), axis=1)
    cash_df = cash_df.apply(lambda row: geography_maps.assign_field_office(row, False), axis=1)

    registered_clients_df["Region"] = registered_clients_df["State"].map(geography_maps.state_to_region)
    bulk_df["Region"] = bulk_df["State"].map(geography_maps.state_to_region)
    cash_df["Region"] = cash_df["State"].map(geography_maps.state_to_region)
    in_kind_df.columns = in_kind_df.columns.str.title()
    irfas_df.columns = irfas_df.columns.str.title()
    in_kind_df["Region"] = in_kind_df["Region"].str.title()

    # Assign RSN
    registered_clients_df["Region Number"] = registered_clients_df["Region"].map(geography_maps.region_to_rsn)
    bulk_df["Region Number"] = bulk_df["Region"].map(geography_maps.region_to_rsn)
    cash_df["Region Number"] = cash_df["Region"].map(geography_maps.region_to_rsn)
    in_kind_df["Region Number"] = in_kind_df["Region"].map(geography_maps.region_to_rsn)

    # Assign Chapter
    bulk_df["Chapter"] = bulk_df.apply(geography_maps.assign_chapter, axis=1)
    cash_df["Chapter"] = cash_df.apply(geography_maps.assign_chapter, axis=1)
    in_kind_df["Chapter"] = in_kind_df.apply(geography_maps.assign_chapter, axis=1)
    registered_clients_df["Chapter"] = registered_clients_df.apply(geography_maps.assign_chapter, axis=1)

    # ALL UPPER CASE the Region
    bulk_df["Region"] = bulk_df["Region"].str.upper()
    cash_df["Region"] = cash_df["Region"].str.upper()
    in_kind_df["Region"] = in_kind_df["Region"].str.upper()
    registered_clients_df["Region"] = registered_clients_df["Region"].str.upper()

    # Add Date columns
    registered_clients_df = geography_maps._add_date_columns(registered_clients_df)
    bulk_df = geography_maps._add_date_columns(bulk_df)
    cash_df = geography_maps._add_date_columns(cash_df)
    cash_df = cash_df.drop(columns=["Primary Campaign Source ↑"])
    in_kind_df = in_kind_df.drop(columns=["Year"])
    in_kind_df = geography_maps._add_date_columns(in_kind_df)

    # Edits the Distribution Name Whitespace
    # Targets optional spaces, an optional dash, and spaces right before the final date
    bulk_df['Bulk Distribution Event: Bulk Distribution Name'] = bulk_df['Bulk Distribution Event: Bulk Distribution Name'].str.replace(r'\s*-?\s*\d{4}-\d{2}-\d{2}$', '', regex=True)

    # Only grabs the needed columns and rows
    # axis=1 targets columns; how='all' drops them only if EVERY value is missing
    # errors='ignore' since not every raw export leaves behind a blank spacer column
    bulk_df = bulk_df.drop(columns=[""], errors="ignore")
    cash_df = cash_df.drop(columns=[""], errors="ignore")

    years_in_data = cash_df["Year"].dropna().unique()
    in_kind_df = in_kind_df[in_kind_df["Year"].isin(years_in_data)]
    irfas_df = irfas_df[irfas_df["Year"].isin(years_in_data)]

    irfas_df = irfas_df[irfas_df["Program"] == "HP"]
    in_kind_df = in_kind_df[in_kind_df["Program"] == "Hunger Prevention"]

    in_kind_df = in_kind_df[["Year", "Quarter", "Month Name", "Month Number", "Region", "Chapter", "State",
                            "Field Office", "Region Number", "Location", "Total Value Of Articles"]]
    irfas_df = irfas_df[["Year", "Quarter", "Month Name", "Month Number", "Region", "Chapter", "State", "Field Office",
                         "Region Number", "City", "Amount Approved"]]

    # Distribution Type
    in_kind_df["Distribution Type"] = "In-Kind"
    irfas_df["Distribution Type"] = "IRFAS"
    registered_clients_df["Distribution Type"] = "Registered Clients"
    bulk_df["Distribution Type"] = "Bulk"
    cash_df["Distribution Type"] = "Cash"

    # Put all 5 of your DataFrames into a list
    list_of_dataframes = [registered_clients_df, bulk_df, cash_df, in_kind_df, irfas_df]

    # Combine them vertically
    hunger_prevention_df = pd.concat(list_of_dataframes, axis=0, ignore_index=True)

    # Capitalizes and orders columns
    first_columns = ["Year", "Quarter", "Month Name", "Month Number", "Date", "Region", "Chapter", "State",
                     "Field Office", "Region Number", "City", "Zipcode", "Location", "Distribution Type"]
    remaining_columns = [
        column for column in hunger_prevention_df.columns
        if column not in first_columns
    ]

    hunger_prevention_df = hunger_prevention_df[first_columns + remaining_columns]

    hunger_prevention_df = zero_value_rows(hunger_prevention_df)

    hunger_prevention_df.columns = hunger_prevention_df.columns.str.upper()

    save_hunger_prevention_starter_workbook(hunger_prevention_df, output_path)