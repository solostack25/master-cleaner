from pathlib import Path

import pandas as pd
import geography_maps
import re


def volunteered(df, year):
    temp = (
        df[(df["Year"] == year) & (df["Contact ID (Distinct)"] != "(Anonymous)")]
        .groupby("Contact ID (Distinct)")
        .size()
        .gt(0)
        .astype(int)
        .rename(f"Volunteered_{year}")
    )
    return temp

def retained(row, year_next, year_prev):
    if row[f"Volunteered_{year_prev}"] == 1 and row[f"Volunteered_{year_next}"] == 1:
        return 1
    else:
        return 0

def retained_id(row):
    if row["Retained"] == 1:
        return row["Contact ID (Distinct)"]
    else:
        return None

def clean_volunteers(input_path, output_path):
    print("STARTING VOLUNTEER CLEANER", flush=True)
    last_year_df = pd.read_csv('2025 Volunteering Data.csv')
    input_path = Path(input_path)
    output_path = Path(output_path)

    input_df = pd.read_excel(
        input_path
    )

    df = pd.concat([input_df, last_year_df], ignore_index=True)

    print("After concat:", df.shape, flush=True)
    print("Memory MB:", df.memory_usage(deep=True).sum() / 1024 / 1024, flush=True)

    del input_df
    del last_year_df

    # Geography Stuff and Date Stuff

    df.rename(columns={"Home Address - State": "State", "Home Address - City": "City",
                       "Home Address - PostalCode": "Zipcode", "User ID": "Contact ID (Distinct)",
                       "Ethnicity": "ESN", "Event Start": "Date"}, inplace=True)

    # Makes sure the State column is completely Capitalized
    df["State"] = df["State"].str.upper()

    df = df.apply(geography_maps.hq_states, axis=1)

    df["Region"] = df["State"].map(geography_maps.state_to_region)

    # Assign RSN
    df["Region Number"] = df["Region"].map(geography_maps.region_to_rsn)

    # Assign Field Office
    df["Field Office"] = df["State"].map(geography_maps.state_to_field_office)

    # Changes ZIP Code column to string, fixes leading zeroes, and takes only the first 5 digits
    df["Zipcode"] = df["Zipcode"].astype(str).str.extract(r'(\d{1,5})')[0].str.zfill(5)

    df = df.apply(
        lambda row: geography_maps.assign_field_office(row, True),
        axis=1
    )

    df["Chapter"] = df.apply(geography_maps.assign_chapter, axis=1)
    df['Region'] = df['Region'].str.upper()
    df = geography_maps._add_date_columns(df)

    # rename columns
    df.rename(columns={'Date': 'Event Start'}, inplace=True)

    df["Waiver 1 Signed"] = df["Waiver 1 Signed"].map({True: "Yes", False: "No"}).fillna("No Information Provided")
    df["Gender"] = df["Gender"].fillna("Not Provided")

    df["Contact ID (Total)"] = df["Contact ID (Distinct)"]

    df["Volunteering Value"] = pd.to_numeric(df["Hours"], errors="coerce") * 34.79

    # Program Assigner
    program_assigner_df = pd.read_csv("Word List.csv")

    # Build mapping dictionary
    mapping = dict(zip(program_assigner_df["Keyword"], program_assigner_df["Program"]))

    # Default program
    df["Program"] = "General"

    for keyword, program in mapping.items():
        pattern = r"\b{}\b".format(re.escape(keyword))

        # FIRST: check Event Name
        mask_name = df["Event Name"].str.contains(pattern, case=False, na=False, regex=True)
        df.loc[(df["Program"] == "General") & mask_name, "Program"] = program

        # SECOND: check Event Group only where still unassigned
        mask_group = df["Event Group"].str.contains(pattern, case=False, na=False, regex=True)
        df.loc[(df["Program"] == "General") & mask_group, "Program"] = program

        # THIRD: check Lineage only where still unassigned
        mask_lineage = df["Event Group Lineage"].str.contains(pattern, case=False, na=False, regex=True)
        df.loc[(df["Program"] == "General") & mask_lineage, "Program"] = program

    # Define age group columns
    age_groups = ["0-17", "18-45", "46-64", "65 or older", "Age Not Available"]

    # Initialize
    for col in age_groups:
        df[col] = 0

    # Assign groups
    df.loc[(df["Age"] >= 0) & (df["Age"] <= 17), "0-17"] = 1
    df.loc[(df["Age"] >= 18) & (df["Age"] <= 45), "18-45"] = 1
    df.loc[(df["Age"] >= 46) & (df["Age"] <= 64), "46-64"] = 1
    df.loc[df["Age"] >= 65, "65 or older"] = 1

    # Not Available
    df.loc[
        df["Age"].isna() | (df["Age"] < 0),
        "Age Not Available"
    ] = 1

    ethnicity_map = {1: "Asian", 2: "Black or African-American", 3: "Hispanic or Latino",
                     4: "Middle Eastern or North African", 5: "Native American or Indigenous", 6: "Pacific Islander",
                     7: "White or Caucasian", 8: "Two or More"}

    df["Ethnicity"] = df["ESN"].map(ethnicity_map).fillna("No Information Provided")

    program_map = {"Back2School": 1, "Disaster Relief Services": 2, "Health Services": 3, "Hunger Prevention": 4,
                   "Muslim Family Services": 5, "FATE": 5.5, "Refugee Services & Community Empowerment": 6,
                   "Transitional Housing": 7, "Outreach": 8, "General": 9}

    program_list = list(program_map.keys())

    df["PSN"] = df["Program"].map(program_map)

    df = df[["City", "State", "Zipcode", "Field Office", "Region", "Region Number", "Chapter", "Full Name - FirstName",
             "Full Name - LastName",
             "Email", "Gender", "Age", "0-17", "18-45", "46-64", "65 or older", "Age Not Available", "Ethnicity", "ESN",
             "Event Start", "Month Name", "Month Number", "Year", "Quarter", "Event Name", "Event Group", "Event Group Lineage",
             "Program", "PSN", "Waiver 1 Signed", "Contact ID (Distinct)", "Contact ID (Total)", "Hours",
             "Volunteering Value"]]

    years_in_data = pd.to_numeric(df["Year"], errors="coerce")
    years_in_data = years_in_data.dropna()

    year_next = int(years_in_data.max())
    year_prev = year_next - 1

    # ----------------------
    # STEP 1 — Base volunteer table (Contact ID-based)
    # ----------------------

    vol_ids = df["Contact ID (Distinct)"].unique()
    vol = pd.DataFrame({"Contact ID (Distinct)": vol_ids})

    vol = vol.merge(volunteered(df, year_prev), on="Contact ID (Distinct)", how="left")
    vol = vol.merge(volunteered(df, year_next), on="Contact ID (Distinct)", how="left")


    # Fill missing with 0
    vol = vol.fillna(0)

    # ----------------------
    # STEP 3 — Program participation flags (skip anonymous)
    # ----------------------

    prev_df = (
        df[(df["Year"] == year_prev) & (df["Contact ID (Distinct)"] != "(Anonymous)")]
        .groupby(["Contact ID (Distinct)", "Program"])
        .size()
        .reset_index(name="count")
    )

    next_df = (
        df[(df["Year"] == year_next) & (df["Contact ID (Distinct)"] != "(Anonymous)")]
        .groupby(["Contact ID (Distinct)", "Program"])
        .size()
        .reset_index(name="count")
    )

    for program in program_list:
        col_prev = f"{program} {year_prev}"
        col_next = f"{program} {year_next}"
        col_retained = f"{program} Retained"

        # participation flag for prev year
        vol[col_prev] = vol["Contact ID (Distinct)"].isin(
            prev_df[prev_df["Program"] == program]["Contact ID (Distinct)"]
        ).astype(int)

        # participation flag for next year
        vol[col_next] = vol["Contact ID (Distinct)"].isin(
            next_df[next_df["Program"] == program]["Contact ID (Distinct)"]
        ).astype(int)

        # retention = participated both years
        vol[col_retained] = (
                (vol[col_prev] == 1) & (vol[col_next] == 1)
        ).astype(int)

    # ----------------------
    # STEP 4 — FORCE ALL "ANONYMOUS" ROWS TO ZERO
    # ----------------------

    anon_mask = vol["Contact ID (Distinct)"] == "(Anonymous)"
    cols_to_zero = vol.columns.difference(["Contact ID (Distinct)"])

    vol.loc[anon_mask, cols_to_zero] = 0

    # ----------------------
    # STEP 5 — MERGE BACK INTO ORIGINAL df
    # ----------------------

    df = df.merge(vol, on="Contact ID (Distinct)", how="left")

    print("After retention merge:", df.shape, flush=True)
    print("Memory MB:", df.memory_usage(deep=True).sum() / 1024 / 1024, flush=True)

    df["Retained"] = (
            (df[f"Volunteered_{year_prev}"] == 1)
            & (df[f"Volunteered_{year_next}"] == 1)
    ).astype(int)

    cols_to_move = [f"Volunteered_{year_prev}", f"Volunteered_{year_next}", "Retained"]

    df = df[[c for c in df.columns if c not in cols_to_move] + cols_to_move]

    df["Retained ID"] = None

    retained_rows = df["Retained"] == 1

    df.loc[retained_rows, "Retained ID"] = df.loc[
        retained_rows,
        "Contact ID (Distinct)"
    ]

    df = df[df["Year"] == year_next]

    df = df[[
        "State", "Full Name - FirstName", "Full Name - LastName",
        "Email", "Event Name", "Program", "PSN", "Hours", "Contact ID (Distinct)", "Contact ID (Total)", "City",
        "Zipcode", "Event Start", "Year", "Month Name", "Month Number", "Region", "Region Number", "Field Office",
        "Volunteering Value", "Chapter", "Gender", "Age", "0-17", "18-45", "46-64", "65 or older", "Age Not Available",
        "Ethnicity", "ESN", "Quarter", "Event Group", "Event Group Lineage", "Waiver 1 Signed",
        f"Back2School {year_prev}", f"Back2School {year_next}", "Back2School Retained",
        f"Disaster Relief Services {year_prev}", f"Disaster Relief Services {year_next}",
        "Disaster Relief Services Retained",
        f"Health Services {year_prev}", f"Health Services {year_next}", "Health Services Retained",
        f"Hunger Prevention {year_prev}", f"Hunger Prevention {year_next}", "Hunger Prevention Retained",
        f"Muslim Family Services {year_prev}", f"Muslim Family Services {year_next}", "Muslim Family Services Retained",
        f"FATE {year_prev}", f"FATE {year_next}", "FATE Retained",
        f"Refugee Services & Community Empowerment {year_prev}",
        f"Refugee Services & Community Empowerment {year_next}", "Refugee Services & Community Empowerment Retained",
        f"Transitional Housing {year_prev}", f"Transitional Housing {year_next}", "Transitional Housing Retained",
        f"Outreach {year_prev}", f"Outreach {year_next}", "Outreach Retained",
        f"General {year_prev}", f"General {year_next}",
        "General Retained",
        f"Volunteered_{year_prev}", f"Volunteered_{year_next}", "Retained", "Retained ID"
    ]]

    first_columns = ["Year", "Quarter", "Month Name", "Month Number", "Event Start", "Region", "Chapter", "State",
                     "Field Office", "Region Number"]

    remaining_columns = [
        column for column in df.columns
        if column not in first_columns
    ]
    df = df[first_columns + remaining_columns]
    # Capitalizes All Columns
    df.columns = df.columns.str.upper()

    print("Before output:", df.shape, flush=True)
    print("Memory MB:", df.memory_usage(deep=True).sum() / 1024 / 1024, flush=True)

    print("About to save Excel output", flush=True)

    geography_maps.save_excel_light(df, output_path)

    print("Finished saving Excel output", flush=True)

    print("About to format currency columns", flush=True)

    geography_maps.format_currency_column(
        output_path,
        "VOLUNTEERING VALUE"
    )

    print("Finished formatting currency columns", flush=True)