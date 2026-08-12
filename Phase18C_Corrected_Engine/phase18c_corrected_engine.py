import os
import pandas as pd

# ============================================================
# PHASE 18C
# MIKA SALES INTELLIGENCE
# CORRECTED MULTI-YEAR INTELLIGENCE ENGINE
# ============================================================

print("=" * 75)
print("PHASE 18C - MIKA SALES INTELLIGENCE")
print("CORRECTED MULTI-YEAR INTELLIGENCE ENGINE")
print("=" * 75)

# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = r"C:\Users\admin\Desktop\Ogutu"

SOURCE_FILE = os.path.join(
    BASE_DIR,
    "1. Sell-In Region Wise_IAL - DO NOT EDIT.xlsm"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "Phase18_MIKA_Sales_Intelligence",
    "Phase18C_Corrected_Engine"
)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "Phase18C_MIKA_Corrected_Intelligence.xlsx"
)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# YEAR MAPPING
# ============================================================

YEAR_COLUMNS = {
    2023: "Value Exc. VAT",
    2024: "Value Exc. VAT.1",
    2025: "Value Exc. VAT.2",
    2026: "Value Exc. VAT.3",
}

DEFAULT_YEAR = 2026

print("\nYEAR MAPPING")

for year, column in YEAR_COLUMNS.items():
    print(f"{year} -> {column}")

print(f"\nDefault reporting year: {DEFAULT_YEAR}")

# ============================================================
# CHECK SOURCE FILE
# ============================================================

if not os.path.exists(SOURCE_FILE):
    raise FileNotFoundError(
        f"Source file not found:\n{SOURCE_FILE}"
    )

print("\nSource file found successfully.")

# ============================================================
# LOAD SOURCE
# ============================================================

print("\nLoading original MIKA source workbook...")

source = pd.read_excel(
    SOURCE_FILE,
    sheet_name="Sell-in_RegionTownAreaCustomer",
    header=7
)

print(f"Rows loaded: {len(source)}")

# ============================================================
# NORMALISE COLUMN NAMES
# ============================================================

source.columns = [
    str(c).strip()
    for c in source.columns
]

required_columns = [
    "Selling Type",
    "Region Name",
    "Value Exc. VAT",
    "Value Exc. VAT.1",
    "Value Exc. VAT.2",
    "Value Exc. VAT.3",
]

missing_columns = [
    column
    for column in required_columns
    if column not in source.columns
]

if missing_columns:
    raise ValueError(
        "Missing required columns:\n"
        + "\n".join(missing_columns)
    )

# ============================================================
# CLEAN TEXT
# ============================================================

source["Selling Type"] = (
    source["Selling Type"]
    .fillna("")
    .astype(str)
    .str.strip()
)

source["Region Name"] = (
    source["Region Name"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ============================================================
# NUMERIC CONVERSION
# ============================================================

for column in YEAR_COLUMNS.values():

    source[column] = pd.to_numeric(
        source[column],
        errors="coerce"
    ).fillna(0)

# ============================================================
# IDENTIFY GRAND TOTAL
# ============================================================

grand_total_mask = (
    source["Selling Type"]
    .str.upper()
    .eq("GRAND TOTAL")
)

grand_total_rows = source.loc[
    grand_total_mask
].copy()

if len(grand_total_rows) != 1:
    raise ValueError(
        f"Expected exactly one Grand Total row. "
        f"Found {len(grand_total_rows)}."
    )

grand_total_row = grand_total_rows.iloc[0]

# ============================================================
# OFFICIAL SOURCE TOTALS
# ============================================================

source_totals = []

for year, column in YEAR_COLUMNS.items():

    official_total = float(
        grand_total_row[column]
    )

    source_totals.append({
        "Reporting Year": year,
        "Source Grand Total": official_total,
        "Source Value Column": column,
    })

source_totals_df = pd.DataFrame(
    source_totals
)

# ============================================================
# SOURCE STRUCTURE
# ============================================================

summary_mask = (
    source["Selling Type"]
    .str.upper()
    .isin([
        "LOCAL SALES TOTAL",
        "GRAND TOTAL",
    ])
)

detail = source.loc[
    ~summary_mask
].copy()

print("\nSOURCE STRUCTURE")

print(
    f"Total source rows: {len(source)}"
)

print(
    f"Summary rows excluded: {summary_mask.sum()}"
)

print(
    f"Detail rows retained: {len(detail)}"
)

print("\nSelling types retained:")

for value in sorted(
    detail["Selling Type"].unique()
):

    print(f"  - {value}")

# ============================================================
# STOCKIST STATUS
# ============================================================

stockist_columns = [
    "Stockist /Customer Chain Name",
    "Stockist Branch SAP Code",
    "Stockist SAP Code",
]

existing_stockist_columns = [
    column
    for column in stockist_columns
    if column in detail.columns
]

if not existing_stockist_columns:

    detail["Stockist Data Status"] = (
        "Missing Stockist"
    )

else:

    stockist_present = pd.Series(
        False,
        index=detail.index
    )

    for column in existing_stockist_columns:

        values = (
            detail[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        valid_values = (
            values.ne("")
            & values.str.lower().ne("nan")
            & values.str.lower().ne("none")
        )

        stockist_present = (
            stockist_present
            | valid_values
        )

    detail["Stockist Data Status"] = (
        stockist_present.map({
            True: "Identified Stockist",
            False: "Missing Stockist",
        })
    )

# ============================================================
# BUILD MULTI-YEAR DATASET
# ============================================================

print("\nBuilding multi-year transaction dataset...")

year_frames = []

for year, value_column in YEAR_COLUMNS.items():

    temp = detail.copy()

    temp["Reporting Year"] = year

    temp["Sales"] = pd.to_numeric(
        temp[value_column],
        errors="coerce"
    ).fillna(0)

    temp["Source Value Column"] = value_column

    year_frames.append(temp)

multi_year = pd.concat(
    year_frames,
    ignore_index=True
)

print(
    f"Multi-year transaction rows: "
    f"{len(multi_year):,}"
)

# ============================================================
# CALCULATED TOTALS
# ============================================================

calculated_totals = (
    multi_year
    .groupby(
        "Reporting Year",
        as_index=False
    )["Sales"]
    .sum()
    .rename(
        columns={
            "Sales": "Calculated Transaction Total"
        }
    )
)

# ============================================================
# YEAR RECONCILIATION
# ============================================================

reconciliation = (
    source_totals_df
    .merge(
        calculated_totals,
        on="Reporting Year",
        how="left"
    )
)

reconciliation["Difference"] = (
    reconciliation["Calculated Transaction Total"]
    - reconciliation["Source Grand Total"]
)

reconciliation["Absolute Difference"] = (
    reconciliation["Difference"].abs()
)

reconciliation["Status"] = (
    reconciliation["Absolute Difference"]
    .apply(
        lambda value:
        "PASS"
        if value < 1
        else "FAIL"
    )
)

print("\n" + "=" * 75)
print("YEAR RECONCILIATION")
print("=" * 75)

for _, row in reconciliation.iterrows():

    print(
        f"{int(row['Reporting Year'])}: "
        f"Source = KSh "
        f"{row['Source Grand Total']:,.2f} | "
        f"Calculated = KSh "
        f"{row['Calculated Transaction Total']:,.2f} | "
        f"Difference = KSh "
        f"{row['Difference']:,.2f} | "
        f"{row['Status']}"
    )

# ============================================================
# YOY GROWTH
# ============================================================

annual = reconciliation[
    [
        "Reporting Year",
        "Source Grand Total",
        "Calculated Transaction Total",
        "Difference",
        "Status",
    ]
].copy()

annual = annual.sort_values(
    "Reporting Year"
)

annual["YoY Growth %"] = (
    annual["Source Grand Total"]
    .pct_change()
    * 100
)

# ============================================================
# HELPER FUNCTION
# ============================================================

def official_total(year):

    values = source_totals_df.loc[
        source_totals_df["Reporting Year"] == year,
        "Source Grand Total"
    ]

    if values.empty:
        raise ValueError(
            f"No official source total found for {year}"
        )

    return float(values.iloc[0])


def get_year_data(year):

    return multi_year.loc[
        multi_year["Reporting Year"] == year
    ].copy()

# ============================================================
# REGIONAL ANALYSIS
# ============================================================

def regional_analysis(year):

    data = get_year_data(year)

    data = data.loc[
        data["Region Name"].ne("")
    ].copy()

    result = (
        data
        .groupby(
            "Region Name",
            as_index=False
        )["Sales"]
        .sum()
    )

    total = official_total(year)

    result["Sales %"] = (
        result["Sales"]
        / total
        * 100
        if total != 0
        else 0
    )

    records = (
        data
        .groupby("Region Name")
        .size()
        .reset_index(
            name="Rows"
        )
    )

    result = result.merge(
        records,
        on="Region Name",
        how="left"
    )

    result.insert(
        0,
        "Reporting Year",
        year
    )

    result = result[
        [
            "Reporting Year",
            "Region Name",
            "Rows",
            "Sales",
            "Sales %",
        ]
    ]

    return result.sort_values(
        "Sales",
        ascending=False
    )

# ============================================================
# SELLING TYPE ANALYSIS
# ============================================================

def selling_type_analysis(year):

    data = get_year_data(year)

    result = (
        data
        .groupby(
            "Selling Type",
            as_index=False
        )["Sales"]
        .sum()
    )

    total = official_total(year)

    result["Sales %"] = (
        result["Sales"]
        / total
        * 100
        if total != 0
        else 0
    )

    records = (
        data
        .groupby("Selling Type")
        .size()
        .reset_index(
            name="Rows"
        )
    )

    result = result.merge(
        records,
        on="Selling Type",
        how="left"
    )

    result.insert(
        0,
        "Reporting Year",
        year
    )

    result = result[
        [
            "Reporting Year",
            "Selling Type",
            "Rows",
            "Sales",
            "Sales %",
        ]
    ]

    return result.sort_values(
        "Sales",
        ascending=False
    )

# ============================================================
# STOCKIST ANALYSIS
# ============================================================

def stockist_analysis(year):

    data = get_year_data(year)

    result = (
        data
        .groupby(
            "Stockist Data Status",
            as_index=False
        )["Sales"]
        .sum()
    )

    total = official_total(year)

    result["Sales %"] = (
        result["Sales"]
        / total
        * 100
        if total != 0
        else 0
    )

    records = (
        data
        .groupby("Stockist Data Status")
        .size()
        .reset_index(
            name="Rows"
        )
    )

    result = result.merge(
        records,
        on="Stockist Data Status",
        how="left"
    )

    result.insert(
        0,
        "Reporting Year",
        year
    )

    result = result[
        [
            "Reporting Year",
            "Stockist Data Status",
            "Rows",
            "Sales",
            "Sales %",
        ]
    ]

    return result.sort_values(
        "Sales",
        ascending=False
    )

# ============================================================
# MISSING STOCKIST BY REGION
# ============================================================

def missing_stockist_region_analysis(year):

    data = get_year_data(year)

    data = data.loc[
        data["Stockist Data Status"]
        == "Missing Stockist"
    ].copy()

    data = data.loc[
        data["Region Name"].ne("")
    ].copy()

    result = (
        data
        .groupby(
            "Region Name",
            as_index=False
        )["Sales"]
        .sum()
        .rename(
            columns={
                "Sales":
                "Missing Stockist Sales"
            }
        )
    )

    total_missing = result[
        "Missing Stockist Sales"
    ].sum()

    result["Sales %"] = (
        result["Missing Stockist Sales"]
        / total_missing
        * 100
        if total_missing != 0
        else 0
    )

    records = (
        data
        .groupby("Region Name")
        .size()
        .reset_index(
            name="Missing Rows"
        )
    )

    result = result.merge(
        records,
        on="Region Name",
        how="left"
    )

    result.insert(
        0,
        "Reporting Year",
        year
    )

    result = result[
        [
            "Reporting Year",
            "Region Name",
            "Missing Rows",
            "Missing Stockist Sales",
            "Sales %",
        ]
    ]

    return result.sort_values(
        "Missing Stockist Sales",
        ascending=False
    )

# ============================================================
# ZONE ANALYSIS
# ============================================================

def zone_analysis(year):

    data = get_year_data(year)

    if "Zone Name" not in data.columns:

        return pd.DataFrame(
            columns=[
                "Reporting Year",
                "Region Name",
                "Zone Name",
                "Rows",
                "Sales",
                "Sales %",
            ]
        )

    data["Zone Name"] = (
        data["Zone Name"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    data = data.loc[
        data["Zone Name"].ne("")
    ].copy()

    result = (
        data
        .groupby(
            [
                "Region Name",
                "Zone Name",
            ],
            as_index=False
        )["Sales"]
        .sum()
    )

    total = official_total(year)

    result["Sales %"] = (
        result["Sales"]
        / total
        * 100
        if total != 0
        else 0
    )

    records = (
        data
        .groupby(
            [
                "Region Name",
                "Zone Name",
            ]
        )
        .size()
        .reset_index(
            name="Rows"
        )
    )

    result = result.merge(
        records,
        on=[
            "Region Name",
            "Zone Name",
        ],
        how="left"
    )

    result.insert(
        0,
        "Reporting Year",
        year
    )

    result = result[
        [
            "Reporting Year",
            "Region Name",
            "Zone Name",
            "Rows",
            "Sales",
            "Sales %",
        ]
    ]

    return result.sort_values(
        "Sales",
        ascending=False
    )

# ============================================================
# BUILD ANALYTICAL TABLES
# ============================================================

print("\nBuilding analytical tables...")

regions_list = []
selling_types_list = []
stockist_list = []
missing_regions_list = []
zones_list = []

for year in YEAR_COLUMNS:

    print(f"Processing {year}...")

    regions_list.append(
        regional_analysis(year)
    )

    selling_types_list.append(
        selling_type_analysis(year)
    )

    stockist_list.append(
        stockist_analysis(year)
    )

    missing_regions_list.append(
        missing_stockist_region_analysis(year)
    )

    zones_list.append(
        zone_analysis(year)
    )

regions_all = pd.concat(
    regions_list,
    ignore_index=True
)

selling_types_all = pd.concat(
    selling_types_list,
    ignore_index=True
)

stockist_all = pd.concat(
    stockist_list,
    ignore_index=True
)

missing_regions_all = pd.concat(
    missing_regions_list,
    ignore_index=True
)

zones_all = pd.concat(
    zones_list,
    ignore_index=True
)

# ============================================================
# EXECUTIVE SUMMARY
# ============================================================

executive_rows = []

for year in YEAR_COLUMNS:

    data = get_year_data(year)

    total = official_total(year)

    local_sales = data.loc[
        data["Selling Type"]
        == "LOCAL SALES",
        "Sales"
    ].sum()

    missing_sales = data.loc[
        data["Stockist Data Status"]
        == "Missing Stockist",
        "Sales"
    ].sum()

    identified_sales = data.loc[
        data["Stockist Data Status"]
        == "Identified Stockist",
        "Sales"
    ].sum()

    missing_pct = (
        missing_sales
        / total
        * 100
        if total != 0
        else 0
    )

    identified_pct = (
        identified_sales
        / total
        * 100
        if total != 0
        else 0
    )

    region_table = regional_analysis(
        year
    )

    if not region_table.empty:

        top_region = region_table.iloc[0]

        top_region_name = (
            top_region["Region Name"]
        )

        top_region_sales = float(
            top_region["Sales"]
        )

        top_region_share = float(
            top_region["Sales %"]
        )

    else:

        top_region_name = ""
        top_region_sales = 0
        top_region_share = 0

    selling_table = selling_type_analysis(
        year
    )

    if not selling_table.empty:

        top_type = selling_table.iloc[0]

        top_selling_type = (
            top_type["Selling Type"]
        )

        top_selling_type_sales = float(
            top_type["Sales"]
        )

        top_selling_type_share = float(
            top_type["Sales %"]
        )

    else:

        top_selling_type = ""
        top_selling_type_sales = 0
        top_selling_type_share = 0

    executive_rows.append({

        "Reporting Year": year,

        "Total Sales": total,

        "Local Sales": local_sales,

        "Missing Stockist Sales":
            missing_sales,

        "Missing Stockist %":
            missing_pct,

        "Identified Stockist Sales":
            identified_sales,

        "Identified Stockist %":
            identified_pct,

        "Top Region":
            top_region_name,

        "Top Region Sales":
            top_region_sales,

        "Top Region Share %":
            top_region_share,

        "Top Selling Type":
            top_selling_type,

        "Top Selling Type Sales":
            top_selling_type_sales,

        "Top Selling Type Share %":
            top_selling_type_share,
    })

executive_summary = pd.DataFrame(
    executive_rows
)

# ============================================================
# CURRENT 2026 SUMMARY
# ============================================================

current_summary = executive_summary.loc[
    executive_summary["Reporting Year"]
    == DEFAULT_YEAR
].copy()

# ============================================================
# DATA DICTIONARY
# ============================================================

data_dictionary = pd.DataFrame({

    "Field": [
        "Reporting Year",
        "Sales",
        "Selling Type",
        "Region Name",
        "Zone Name",
        "Stockist Data Status",
        "Source Value Column",
    ],

    "Meaning": [
        "Year represented by the sales value",
        "Sales value for the reporting year",
        "Sales channel/category",
        "Sales region or territory",
        "Sales zone",
        "Whether stockist/customer information is available",
        "Original Excel column used for that year",
    ]
})

# ============================================================
# FINAL VALIDATION
# ============================================================

failed = reconciliation.loc[
    reconciliation["Status"] != "PASS"
]

print("\n" + "=" * 75)
print("FINAL YEAR VALIDATION")
print("=" * 75)

print(
    reconciliation[
        [
            "Reporting Year",
            "Source Value Column",
            "Source Grand Total",
            "Calculated Transaction Total",
            "Difference",
            "Status",
        ]
    ].to_string(index=False)
)

if not failed.empty:

    print("\nVALIDATION FAILED.")

    raise ValueError(
        "One or more years failed reconciliation. "
        "Dashboard generation must stop."
    )

print(
    "\nALL FOUR YEARS RECONCILE SUCCESSFULLY."
)

# ============================================================
# WRITE OUTPUT
# ============================================================

print("\nWriting Phase 18C workbook...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    annual.to_excel(
        writer,
        sheet_name="Annual Sales",
        index=False
    )

    executive_summary.to_excel(
        writer,
        sheet_name="Executive Summary",
        index=False
    )

    current_summary.to_excel(
        writer,
        sheet_name="Current 2026 Summary",
        index=False
    )

    regions_all.to_excel(
        writer,
        sheet_name="Sales by Region",
        index=False
    )

    selling_types_all.to_excel(
        writer,
        sheet_name="Selling Type",
        index=False
    )

    stockist_all.to_excel(
        writer,
        sheet_name="Stockist Status",
        index=False
    )

    missing_regions_all.to_excel(
        writer,
        sheet_name="Missing Stockist Region",
        index=False
    )

    zones_all.to_excel(
        writer,
        sheet_name="Sales by Zone",
        index=False
    )

    multi_year.to_excel(
        writer,
        sheet_name="Clean Multi Year Data",
        index=False
    )

    validation = reconciliation.copy()

    validation.to_excel(
        writer,
        sheet_name="Year Validation",
        index=False
    )

    source_totals_df.to_excel(
        writer,
        sheet_name="Source Totals",
        index=False
    )

    data_dictionary.to_excel(
        writer,
        sheet_name="Data Dictionary",
        index=False
    )

# ============================================================
# FINAL REPORT
# ============================================================

print("\n" + "=" * 75)
print("PHASE 18C COMPLETE")
print("=" * 75)

print("\nOFFICIAL SOURCE TOTALS")

for _, row in source_totals_df.iterrows():

    print(
        f"{int(row['Reporting Year'])}: "
        f"KSh {row['Source Grand Total']:,.2f}"
    )

print("\nYOY GROWTH")

for _, row in annual.iterrows():

    year = int(
        row["Reporting Year"]
    )

    yoy_value = row["YoY Growth %"]

    if pd.isna(yoy_value):

        print(
            f"{year}: N/A"
        )

    else:

        print(
            f"{year}: "
            f"{yoy_value:.2f}%"
        )

print("\nCURRENT REPORTING YEAR")

current = executive_summary.loc[
    executive_summary["Reporting Year"]
    == DEFAULT_YEAR
].iloc[0]

print(
    f"Year: {DEFAULT_YEAR}"
)

print(
    f"Total Sales: "
    f"KSh {current['Total Sales']:,.2f}"
)

print(
    f"Local Sales: "
    f"KSh {current['Local Sales']:,.2f}"
)

print(
    f"Missing Stockist Sales: "
    f"KSh {current['Missing Stockist Sales']:,.2f}"
)

print(
    f"Missing Stockist %: "
    f"{current['Missing Stockist %']:.2f}%"
)

print(
    f"Identified Stockist Sales: "
    f"KSh {current['Identified Stockist Sales']:,.2f}"
)

print(
    f"Identified Stockist %: "
    f"{current['Identified Stockist %']:.2f}%"
)

print(
    f"Top Region: "
    f"{current['Top Region']}"
)

print(
    f"Top Region Sales: "
    f"KSh {current['Top Region Sales']:,.2f}"
)

print(
    f"Top Selling Type: "
    f"{current['Top Selling Type']}"
)

print(
    f"Top Selling Type Sales: "
    f"KSh {current['Top Selling Type Sales']:,.2f}"
)

print("\nOUTPUT FILE:")
print(OUTPUT_FILE)

print("\nSTATUS:")
print(
    "READY FOR PHASE 18 DASHBOARD INTEGRATION"
)

