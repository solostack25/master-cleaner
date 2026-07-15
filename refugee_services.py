from pathlib import Path
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import geography_maps


def _normalize_header(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    return re.sub(r"\s+", " ", value)


def read_rs_sheet(input_path, first_column_name, last_column_name):
    """Locate the header row by first/last column labels (normalized for
    whitespace) and return everything between them, inclusive. Unlike
    hunger_prevention.clean_dataframe, this doesn't assume a blank spacer
    column always sits at a fixed index -- it locates both ends by name."""
    raw = pd.read_excel(input_path, sheet_name=0, header=None)

    first_target = _normalize_header(first_column_name)
    last_target = _normalize_header(last_column_name)

    header_row_number = None
    first_col = None
    last_col = None

    for row_number, row in raw.iterrows():
        normalized_row = [_normalize_header(v) for v in row]
        if first_target in normalized_row:
            candidate_first = normalized_row.index(first_target)
            if last_target in normalized_row[candidate_first:]:
                header_row_number = row_number
                first_col = candidate_first
                last_col = normalized_row.index(last_target, candidate_first)
                break

    if header_row_number is None:
        raise ValueError(
            f"Could not find header row with columns '{first_column_name}' "
            f"through '{last_column_name}' in {input_path}"
        )

    df = raw.iloc[header_row_number:, first_col:last_col + 1].reset_index(drop=True)
    df.columns = [_normalize_header(v) for v in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(axis=0, how="all")
    df = df.reset_index(drop=True)

    return df

RS_FILENAME_LOOKUP = {
    "cash": "cash",
    "bulk": "bulk",
    "in_kind": "in_kind",
    "inkind": "in_kind",
    "irfas": "irfas",
    "case_management": "case_management",
    "case_mgmt": "case_management",
}

REQUIRED_RS_FILES = [
    "cash",
    "bulk",
    "in_kind",
    "irfas",
    "case_management",
]

RS_DISPLAY_NAMES = {
    "cash": "Cash",
    "bulk": "Bulk",
    "in_kind": "In-Kind",
    "irfas": "IRFAS",
    "case_management": "Case Management",
}

# Final output column order/names, matching the reference RS-PBI FINAL TABLE.
FINAL_COLUMNS = [
    "CATEGORY", "YEAR", "QUARTER", "MONTH", "MSN", "RSN", "REGION", "CHAPTER",
    "STATE", "FIELD OFFICE", "Client: Office Location  ↑",
    "Bulk Distribution Event: Bulk Distribution Name",
    "Total Value of Services Provided", "Number of Beneficiaries",
    "Service Name  ↑", "Client Country of Origin", "Client Ethnicity",
    "CASH", "INKIND", "TOTAL REV",
]

# Same convention used in hunger_prevention.py for the 5 states that get
# split into multiple field offices, matched against "State - Location" text.
MULTI_FIELD_OFFICE_LOCATION_TO_OFFICE = {
    "CA - Bay Area": "Bay Area Office",
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
    "VA - Richmond": "Richmond Office",
}

# Confirmed with Travis: for BULK reporting, the 4 individual CA offices get
# grouped into a simpler Northern/Southern CA split.
NORTHERN_CALIFORNIA_OFFICES = {"Sacramento Office", "Bay Area Office"}
SOUTHERN_CALIFORNIA_OFFICES = {"Los Angeles Office", "San Diego Office"}


def normalize_rs_filename(original_filename):
    filename = Path(original_filename).name.strip().lower()
    file_stem = Path(filename).stem.strip()
    file_extension = Path(filename).suffix.lower()

    file_stem = file_stem.replace(" ", "_")
    file_stem = file_stem.replace("-", "_")

    # Strip common prefixes/suffixes people tend to leave on exported files
    # (e.g. "rs_cash_-_travis_copy" -> "cash").
    for token in ["rs_", "_travis_copy", "_copy"]:
        file_stem = file_stem.replace(token, "")

    file_stem = file_stem.strip("_")

    return file_stem, file_extension


def parse_rs_filename(original_filename):
    file_stem, file_extension = normalize_rs_filename(original_filename)

    if file_extension != ".xlsx":
        raise ValueError(
            f"Refugee Services files must be .xlsx files. Invalid file: {original_filename}"
        )

    if file_stem not in RS_FILENAME_LOOKUP:
        raise ValueError(
            f"Could not recognize Refugee Services file: {original_filename}. "
            "Expected filenames containing one of: cash, bulk, in_kind, irfas, "
            "case_management."
        )

    return RS_FILENAME_LOOKUP[file_stem]


def validate_rs_uploads(input_files, input_filenames):
    recognized_files = {}

    for file_num, file_info in enumerate(input_files):
        original_filename = input_filenames[file_num]
        file_type = parse_rs_filename(original_filename)

        if file_type in recognized_files:
            display_name = RS_DISPLAY_NAMES[file_type]
            raise ValueError(
                f"Duplicate Refugee Services file found for {display_name}. "
                f"Only upload one {display_name} file."
            )

        recognized_files[file_type] = file_info

    missing_file_types = [
        file_type for file_type in REQUIRED_RS_FILES
        if file_type not in recognized_files
    ]

    if missing_file_types:
        missing_display_names = [RS_DISPLAY_NAMES[ft] for ft in missing_file_types]
        raise ValueError(
            "Missing required Refugee Services file(s): "
            + ", ".join(missing_display_names)
        )

    return recognized_files


def _assign_office_from_location(location_value):
    """Given a 'CA - Bay Area' / 'AZ - Phoenix' style string, return the
    (state, field_office) tuple using the same rules registered_clients_df
    uses in hunger_prevention.py."""
    if not isinstance(location_value, str) or " - " not in location_value:
        return None, None

    state = location_value.split(" - ")[0].strip().upper()

    if location_value.strip() in MULTI_FIELD_OFFICE_LOCATION_TO_OFFICE:
        field_office = MULTI_FIELD_OFFICE_LOCATION_TO_OFFICE[location_value.strip()]
    else:
        field_office = geography_maps.state_to_field_office.get(state)

    return state, field_office


def _apply_northern_southern_ca_split(field_office):
    if field_office in NORTHERN_CALIFORNIA_OFFICES:
        return "Northern California"
    if field_office in SOUTHERN_CALIFORNIA_OFFICES:
        return "Southern California"
    return field_office


def clean_case_management(input_path):
    df = read_rs_sheet(input_path, "Delivery Date ↑", "# of Beneficiaries")

    df["Client: Office Location ↑"] = df["Client: Office Location ↑"].ffill()
    df["Delivery Date ↑"] = df["Delivery Date ↑"].ffill()
    df["Service Name ↑"] = df["Service Name ↑"].ffill()

    # "Delivery Date ↑" can be a single date or a "start - end" week range;
    # use the first date in either case.
    first_date = df["Delivery Date ↑"].astype(str).str.split(" - ").str[0]
    df["Date"] = pd.to_datetime(first_date, errors="coerce")

    parsed = df["Client: Office Location ↑"].apply(_assign_office_from_location)
    df["State"] = parsed.apply(lambda pair: pair[0])
    df["Field Office"] = parsed.apply(lambda pair: pair[1])

    df["Region"] = df["State"].map(geography_maps.state_to_region)
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)
    df["Chapter"] = df.apply(
        lambda row: geography_maps.assign_chapter(
            {"Region": row["Region"], "Field Office": row["Field Office"]}
        ),
        axis=1,
    )

    df = geography_maps._add_date_columns(df)

    df["Client Country of Origin"] = df["Client Country of Origin"].fillna("Not Provided")
    df["Client Ethnicity"] = df["Client Ethnicity"].fillna("Not Provided")
    df["Total Value of Services Provided"] = pd.to_numeric(
        df["Total Value of Services Provided"], errors="coerce"
    ).fillna(0)
    df["# of Beneficiaries"] = pd.to_numeric(
        df["# of Beneficiaries"], errors="coerce"
    ).fillna(0)

    out = pd.DataFrame({
        "CATEGORY": "MAIN",
        "YEAR": df["Year"],
        "QUARTER": df["Quarter"],
        "MONTH": df["Month Name"],
        "MSN": df["Month Number"],
        "RSN": df["Region Number"],
        "REGION": df["Region"],
        "CHAPTER": df["Chapter"],
        "STATE": df["State"],
        "FIELD OFFICE": df["Field Office"],
        "Client: Office Location  ↑": df["Client: Office Location ↑"],
        "Bulk Distribution Event: Bulk Distribution Name": pd.NA,
        "Total Value of Services Provided": df["Total Value of Services Provided"],
        "Number of Beneficiaries": df["# of Beneficiaries"],
        "Service Name  ↑": df["Service Name ↑"],
        "Client Country of Origin": df["Client Country of Origin"],
        "Client Ethnicity": df["Client Ethnicity"],
        "CASH": pd.NA,
        "INKIND": pd.NA,
        "TOTAL REV": 0.0,
    })

    return out


def clean_cash(input_path):
    df = read_rs_sheet(input_path, "Primary Campaign Source ↑", "Billing Zip/Postal Code")

    df = df.rename(columns={
        "Billing State/Province": "State",
        "Billing City": "City",
        "Billing Zip/Postal Code": "Zipcode",
        "Close Date": "Date",
    })

    df["State"] = df["State"].astype(str).str.upper()
    df["Zipcode"] = df["Zipcode"].astype(str).str.extract(r"(\d{1,5})")[0].str.zfill(5)

    df = df.apply(geography_maps.hq_states, axis=1)
    df["Field Office"] = df["State"].map(geography_maps.state_to_field_office)
    df = df.apply(lambda row: geography_maps.assign_field_office(row, False), axis=1)

    df["Region"] = df["State"].map(geography_maps.state_to_region)
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)
    df["Chapter"] = df.apply(geography_maps.assign_chapter, axis=1)

    df = geography_maps._add_date_columns(df)

    df["Payment Amount Received"] = pd.to_numeric(
        df["Payment Amount Received"], errors="coerce"
    ).fillna(0)

    out = pd.DataFrame({
        "CATEGORY": "CASH",
        "YEAR": df["Year"],
        "QUARTER": df["Quarter"],
        "MONTH": df["Month Name"],
        "MSN": df["Month Number"],
        "RSN": df["Region Number"],
        "REGION": df["Region"],
        "CHAPTER": pd.NA,
        "STATE": df["State"],
        "FIELD OFFICE": df["Field Office"],
        "Client: Office Location  ↑": pd.NA,
        "Bulk Distribution Event: Bulk Distribution Name": pd.NA,
        "Total Value of Services Provided": 0.0,
        "Number of Beneficiaries": 0.0,
        "Service Name  ↑": pd.NA,
        "Client Country of Origin": pd.NA,
        "Client Ethnicity": pd.NA,
        "CASH": df["Payment Amount Received"],
        "INKIND": pd.NA,
        "TOTAL REV": df["Payment Amount Received"],
    })

    return out


def clean_bulk(input_path):
    df = read_rs_sheet(input_path, "ICNA Relief Office: Account Name ↑", "No of People Served")

    df = df[~df["ICNA Relief Office: Account Name ↑"].isin(["Subtotal", "Total"])]
    df["ICNA Relief Office: Account Name ↑"] = (
        df["ICNA Relief Office: Account Name ↑"].ffill()
    )

    df["State"] = df["ICNA Relief Office: Account Name ↑"].astype(str).str.split(" - ").str[0].str.upper()
    df["Field Office"] = df["State"].map(geography_maps.state_to_field_office)

    # Override with the 5-state multi-office logic when the account name
    # matches one of the known sub-locations (prefix match, since these
    # names carry extra trailing words like "Office"/"Food Pantry").
    for location_key, office in MULTI_FIELD_OFFICE_LOCATION_TO_OFFICE.items():
        matches = df["ICNA Relief Office: Account Name ↑"].astype(str).str.startswith(location_key)
        df.loc[matches, "Field Office"] = office

    # Confirmed grouping: BULK reports CA as Northern/Southern CA rather
    # than the 4 individual offices.
    df["Field Office"] = df["Field Office"].apply(_apply_northern_southern_ca_split)

    df["Region"] = df["State"].map(geography_maps.state_to_region)
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)
    df["Chapter"] = df.apply(
        lambda row: geography_maps.assign_chapter(
            {"Region": row["Region"], "Field Office": row["Field Office"]}
        ),
        axis=1,
    )

    df = df.rename(columns={"Date Distributed": "Date"})
    df = geography_maps._add_date_columns(df)

    df["No of People Served"] = pd.to_numeric(
        df["No of People Served"], errors="coerce"
    ).fillna(0)

    out = pd.DataFrame({
        "CATEGORY": "BULK",
        "YEAR": df["Year"],
        "QUARTER": df["Quarter"],
        "MONTH": df["Month Name"],
        "MSN": df["Month Number"],
        "RSN": df["Region Number"],
        "REGION": df["Region"],
        "CHAPTER": df["Chapter"],
        "STATE": df["State"],
        "FIELD OFFICE": df["Field Office"],
        "Client: Office Location  ↑": df["ICNA Relief Office: Account Name ↑"],
        "Bulk Distribution Event: Bulk Distribution Name": "Bulk Household Items",
        "Total Value of Services Provided": pd.NA,
        "Number of Beneficiaries": df["No of People Served"],
        "Service Name  ↑": "Bulk Household Items",
        "Client Country of Origin": "Not Provided",
        "Client Ethnicity": "Not Provided",
        "CASH": pd.NA,
        "INKIND": pd.NA,
        "TOTAL REV": 0.0,
    })

    return out


def clean_inkind(input_path):
    df = pd.read_excel(input_path, sheet_name=0)
    df.columns = df.columns.str.upper()

    df = df[df["PROGRAM  ↑"] == "Refugee Services"].copy()

    df["TOTAL VALUE OF ARTICLES"] = pd.to_numeric(
        df["TOTAL VALUE OF ARTICLES"], errors="coerce"
    ).fillna(0)

    out = pd.DataFrame({
        "CATEGORY": "INKIND",
        "YEAR": df["YEAR"],
        "QUARTER": df["QUARTER"],
        "MONTH": df["MONTH NAME"],
        "MSN": df["MONTH NUMBER"],
        "RSN": df["REGION NUMBER"],
        "REGION": df["REGION"],
        "CHAPTER": pd.NA,
        "STATE": df["STATE"],
        "FIELD OFFICE": df["FIELD OFFICE"],
        "Client: Office Location  ↑": pd.NA,
        "Bulk Distribution Event: Bulk Distribution Name": pd.NA,
        "Total Value of Services Provided": 0.0,
        "Number of Beneficiaries": 0.0,
        "Service Name  ↑": pd.NA,
        "Client Country of Origin": pd.NA,
        "Client Ethnicity": pd.NA,
        "CASH": pd.NA,
        "INKIND": df["TOTAL VALUE OF ARTICLES"],
        "TOTAL REV": df["TOTAL VALUE OF ARTICLES"],
    })

    return out


def clean_irfas(input_path):
    df = pd.read_excel(input_path, sheet_name=0)

    df = df[df["Program"].isin(["RSCE", "Refugee Services & Community Empowerment"])].copy()

    df["AMOUNT APPROVED"] = pd.to_numeric(df["AMOUNT APPROVED"], errors="coerce").fillna(0)

    out = pd.DataFrame({
        "CATEGORY": "IRFAS",
        "YEAR": df["YEAR"],
        "QUARTER": df["QUARTER"],
        "MONTH": df["MONTH"],
        "MSN": df["MSN"],
        "RSN": df["RSN"],
        "REGION": df["REGION"],
        "CHAPTER": df["CHAPTER"],
        "STATE": df["STATE"],
        "FIELD OFFICE": df["FIELD OFFICE"],
        "Client: Office Location  ↑": df["FIELD OFFICE"],
        "Bulk Distribution Event: Bulk Distribution Name": pd.NA,
        "Total Value of Services Provided": df["AMOUNT APPROVED"],
        "Number of Beneficiaries": pd.NA,
        "Service Name  ↑": "IRFAS-Financial Assistance",
        "Client Country of Origin": "Not Provided",
        "Client Ethnicity": "Not Provided",
        "CASH": 0.0,
        "INKIND": 0.0,
        "TOTAL REV": 0.0,
    })

    return out


def save_rs_starter_workbook(df, output_path):
    output_path = Path(output_path)

    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("COMBINED DATA")

    df_to_save = df.copy()
    df_to_save = df_to_save.astype(object)
    df_to_save = df_to_save.where(pd.notna(df_to_save), None)

    for row in dataframe_to_rows(df_to_save, index=False, header=True):
        worksheet.append(row)

    workbook.save(output_path)


def clean_multiple_refugee_services_imports(input_files, input_filenames, output_path):
    recognized_files = validate_rs_uploads(input_files, input_filenames)

    main_df = clean_case_management(recognized_files["case_management"])
    cash_df = clean_cash(recognized_files["cash"])
    bulk_df = clean_bulk(recognized_files["bulk"])
    inkind_df = clean_inkind(recognized_files["in_kind"])
    irfas_df = clean_irfas(recognized_files["irfas"])

    combined_df = pd.concat(
        [main_df, cash_df, bulk_df, inkind_df, irfas_df],
        axis=0,
        ignore_index=True,
    )

    combined_df = combined_df[FINAL_COLUMNS]

    save_rs_starter_workbook(combined_df, output_path)
