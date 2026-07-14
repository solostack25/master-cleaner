import pandas as pd
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Dictionary for Region assignment
state_to_region = {
    "VA": "DMV", "MD": "DMV", "DC": "DMV", "WV": "DMV",
    "IL": "Midwest 1", "MI": "Midwest 1",
    "ND": "Midwest 2", "SD": "Midwest 2", "NE": "Midwest 2", "MN": "Midwest 2",
    "IA": "Midwest 2", "MO": "Midwest 2", "WI": "Midwest 2", "IN": "Midwest 2",
    "OH": "Midwest 2", "KY": "Midwest 2",
    "TN": "Southeast", "NC": "Southeast", "SC": "Southeast", "GA": "Southeast",
    "AL": "Southeast", "MS": "Southeast", "LA": "Southeast", "FL": "Southeast",
    "AR": "South Central", "TX": "South Central", "NM": "South Central",
    "OK": "South Central", "KS": "South Central",
    "WA": "West", "OR": "West", "CA": "West", "ID": "West", "MT": "West",
    "WY": "West", "CO": "West", "NV": "West", "UT": "West", "AZ": "West",
    "AK": "West", "HI": "West",
    "ME": "Northeast", "VT": "Northeast", "NH": "Northeast", "CT": "Northeast",
    "MA": "Northeast", "NY": "Northeast", "NJ": "Northeast", "DE": "Northeast",
    "PA": "Northeast", "RI": "Northeast", 'HQ': 'Unassigned'
}

# Dictionary for RSN assignment based on Region
region_to_rsn = {
    "DMV": 1,
    "Midwest 1": 2,
    "Midwest 2": 3,
    "Northeast": 4,
    "South Central": 5,
    "Southeast": 6,
    "West": 7,
    "Unassigned": 8
}

# Dictionary for Field Office assignment based on State
state_to_field_office = {
    "ME": "Maine-Statewide", "VT": "Vermont-Statewide", "NH": "New Hampshire-Statewide",
    "MA": "Boston Office", "CT": "Connecticut-Statewide", "RI": "Rhode Island-Statewide",
    "NY": "New York Office", "NJ": "New Jersey Office", "DE": "Delaware-Statewide",
    "PA": "Philadelphia Office", "MD": "Baltimore Office", "DC": "District of Columbia",
    "WV": "West Virginia-Statewide", "IL": "Chicago Office", "MI": "Detroit Office",
    "ND": "North Dakota-Statewide", "SD": "South Dakota-Statewide", "NE": "Nebraska-Statewide",
    "MN": "Minneapolis Office", "IA": "Cedar Rapids Office",
    # "MO": "St. Louis Office",
    "WI": "Wisconsin-Statewide", "IN": "Indianapolis Office", "OH": "Ohio-Statewide",
    "KY": "Kentucky-Statewide", "TN": "Memphis Office", "NC": "Durham Office",
    "SC": "Charleston Office", "GA": "Atlanta Office", "AL": "Alabama-Statewide",
    "MS": "Mississippi-Statewide", "LA": "New Orleans Office", "AR": "Arkansas-Statewide",
    "NM": "New Mexico-Statewide", "OK": "Oklahoma City Office", "KS": "Kansas City Office",
    "WA": "Seattle Office", "OR": "Oregon-Statewide", "ID": "Idaho-Statewide",
    "MT": "Montana-Statewide", "WY": "Wyoming-Statewide", "CO": "Denver Office",
    "NV": "Nevada-Statewide", "UT": "Utah-Statewide", "AZ": "Phoenix Office",
    "AK": "Alaska-Statewide", "HI": "Hawaii-Statewide", 'HQ': 'Unassigned'
}

split_field_office_to_state = {
    "Alexandria Office": "VA",
    "Richmond Office": "VA",

    "Sacramento Office": "CA",
    "Bay Area Office": "CA",
    "Los Angeles Office": "CA",
    "San Diego Office": "CA",

    "Houston Office": "TX",
    "Dallas Office": "TX",
    "Austin Office": "TX",

    "St. Louis Office": "MO",
    "Kansas City Office": "KS",

    "Miami Office": "FL",
    "Tampa Office": "FL",
    "Orlando Office": "FL",
}



# reads the data that assigns each city in CA, TX, FL, VA, MO to a Field Office
state_splits = pd.read_csv("state_splits_cities.csv")


# coverts the Zip Code column in state_splits to a string
# Ensure all ZIP codes are strings and properly formatted
state_splits['Zip Code'] = (
    state_splits['Zip Code']
    .fillna("")  # Replace NaN with empty string (or use "00000" if needed)
    .astype(str)  # Convert to string
    .str.split('.').str[0]  # Remove decimal if present (caused by float conversion)
    .str.zfill(5)  # Ensure 5-digit format
)

chapter_map = {
    "Northeast": {
        "New England": ["Maine-Statewide", "Vermont-Statewide", "Boston Office", "New Hampshire-Statewide",
                        "Rhode Island-Statewide"],
        "default": "Northeast Other"
    },
    "DMV": {
        "Virginia": ["Richmond Office", "Alexandria Office"],
        "default": "DMV Other"
    },
    "West": {
        "California": ["Sacramento Office", "Bay Area Office", "Los Angeles Office", "San Diego Office"],
        "default": "West Other"
    },
    "South Central": {
        "Texas": ["Houston Office", "Austin Office", "Dallas Office"],
        "default": "South Central Other"
    },
    "Southeast": {
        "Florida": ["Orlando Office", "South Florida Office", "Tampa Office"],
        "default": "Southeast Other"
    },
    "Unassigned": {
        "default": "National"
    }
}


def assign_chapter(row):
    region = row['Region']
    office = row['Field Office']

    region_map = chapter_map.get(region)
    if region_map:
        for chapter, offices in region_map.items():
            if chapter != "default" and office in offices:
                return chapter
        return region_map["default"]
    else:
        return region  # fallback for any other region


zero_value_df = pd.read_csv("zero_values_2026_starter.csv")

states_of_interest = ['CA', 'TX', 'FL', 'VA', 'MO']

# creates a function to assign a Field Office to rows that correspond to the following 5 states: CA, TX, FL, VA, MO
def assign_field_office(row, zip_there):
    states_of_interest = {'CA', 'TX', 'FL', 'VA', 'MO'}  # Use a set for faster lookup

    if row['State'] in states_of_interest:
        city = row['City']
        state = row['State']
        zip_code = row['Zipcode']

        # Normalize city name
        if isinstance(city, str):
            city = city.title().strip()
        else:
            city = 'Unknown'

        if not state_splits[
            (state_splits['Main City'] == city) & (state_splits['STATE'] == state)
        ].empty:
            field_office = state_splits.loc[
                (state_splits['Main City'] == city) & (state_splits['STATE'] == state),
                "ASSIGNED TO"
            ].values[0]
        elif city == 'St Louis':
            field_office = 'St. Louis Office'
        elif city == 'St Petersburg':
            field_office = 'Miami Office'
        elif zip_there and not state_splits[
            (state_splits['Zip Code'] == zip_code) & (state_splits['STATE'] == state)
        ].empty and zip_code != "00000":
            field_office = state_splits.loc[
                (state_splits['Zip Code'] == zip_code) & (state_splits['STATE'] == state),
                "ASSIGNED TO"
            ].values[0]
        else:
            # Default based on state
            state_defaults = {
                'TX': 'Dallas Office',
                'FL': 'Miami Office',
                'CA': 'Los Angeles Office',
                'VA': 'Alexandria Office',
                'MO': 'St. Louis Office'
            }
            field_office = state_defaults.get(state, 'Unknown')
            print(f"No reference for {city}, {state} {zip_code}\nAutomatically assigned to {field_office}\n")

        # Assign Region and RSN if Kansas City Office
        if field_office == 'Kansas City Office':
            row['Region'] = 'South Central'
            row['Region Number'] = 5
            row['State'] = 'KS'

        row['Field Office'] = field_office

    return row

# A function to assign unrecognized states to HQ
def hq_states(row):
    state = row["State"]
    if (state not in state_to_region.keys()) and (state not in states_of_interest):
        row['State'] = 'HQ'
    return row

def _add_date_columns(df):
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    df["Year"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
    df["Month Number"] = pd.Series(pd.NA, index=df.index, dtype="Int64")
    df["Month Name"] = ""
    df["Quarter"] = ""

    valid_date_rows = df["Date"].notna()

    df.loc[valid_date_rows, "Year"] = (
        df.loc[valid_date_rows, "Date"].dt.year.astype("Int64")
    )

    df.loc[valid_date_rows, "Month Number"] = (
        df.loc[valid_date_rows, "Date"].dt.month.astype("Int64")
    )

    df.loc[valid_date_rows, "Month Name"] = (
        df.loc[valid_date_rows, "Date"].dt.month_name()
    )

    quarter_number = (
        ((df.loc[valid_date_rows, "Month Number"] - 1) // 3) + 1
    )

    df.loc[valid_date_rows, "Quarter"] = (
        "Q" + quarter_number.astype(str)
    )

    return df

def save_excel_light(df, output_path):
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("Cleaned Volunteers")

    df_to_save = df.copy()
    df_to_save = df_to_save.astype(object)
    df_to_save = df_to_save.where(pd.notna(df_to_save), None)

    for row in dataframe_to_rows(df_to_save, index=False, header=True):
        worksheet.append(row)

    workbook.save(output_path)

def format_currency_column(output_path, currency_column):
    workbook = load_workbook(output_path)
    worksheet = workbook.active

    headers = [cell.value for cell in worksheet[1]]

    currency_column_number = None

    for index, header in enumerate(headers, start=1):
        if str(header).strip().upper() == currency_column.strip().upper():
            currency_column_number = index
            break

    if currency_column_number is None:
        workbook.save(output_path)
        return

    for row_number in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=row_number, column=currency_column_number)

        value = cell.value

        if value is None or str(value).strip() == "":
            continue

        # If it is already a number, keep it.
        if isinstance(value, (int, float)):
            numeric_value = value
        else:
            cleaned_value = (
                str(value)
                .strip()
                .replace("$", "")
                .replace(",", "")
            )

            # Handle accounting-style negatives like ($1,250.00)
            if cleaned_value.startswith("(") and cleaned_value.endswith(")"):
                cleaned_value = "-" + cleaned_value[1:-1]

            try:
                numeric_value = float(cleaned_value)
            except ValueError:
                # If it cannot be converted, leave it as-is.
                continue

        # Replace the cell value with the numeric version.
        cell.value = numeric_value

        # Format display as dollars.
        cell.number_format = '$#,##0.00'

    workbook.save(output_path)

program_assigner_df = pd.read_csv("Word List.csv")
# Build mapping dictionary
mapping = dict(zip(program_assigner_df["Keyword"], program_assigner_df["Program"]))