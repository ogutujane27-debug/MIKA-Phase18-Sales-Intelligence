import os
import pandas as pd

# ============================================================
# PHASE 18 - MIKA SALES INTELLIGENCE
# YEAR & RECONCILIATION ANALYSIS
# ============================================================

BASE_DIR = r"C:\Users\admin\Desktop\Ogutu"

SOURCE_FILE = os.path.join(
    BASE_DIR,
    "1. Sell-In Region Wise_IAL - DO NOT EDIT.xlsm"
)

PHASE16_FILE = os.path.join(
    BASE_DIR,
    "Phase16_MIKA_Sales_Intelligence_Engine.xlsx"
)

OUTPUT_DIR = os.path.join(
    BASE_DIR,
    "Phase18_MIKA_Sales_Intelligence"
)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "Phase18_MIKA_Year_Reconciliation.xlsx"
)

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("KSh", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def money(value):
    return f"KSh {float(value):,.2f}"


def pct(value):
    return f"{float(value):.2f}%"


# ============================================================
# CHECK FILES
# ============================================================

print("=" * 70)
print("PHASE 18 - MIKA SALES INTELLIGENCE")
print("YEAR & RECONCILIATION ANALYSIS")
print("=" * 70)

if not os.path.exists(SOURCE_FILE):
    raise FileNotFoundError(
        f"Source workbook not found:\n{SOURCE_FILE}"
    )

if not os.path.exists(PHASE16_FILE):
    raise FileNotFoundError(
        f"Phase 16 workbook not found:\n{PHASE16_FILE}"
    )

print("\nSource files found successfully.")


# ============================================================
# LOAD ORIGINAL SOURCE
# ============================================================

print("\nLoading original MIKA source workbook...")

source = pd.read_excel(
    SOURCE_FILE,
    sheet_name="Sell-in_RegionTownAreaCustomer",
    header=7
)

print(f"Source rows loaded: {len(source):,}")


# ============================================================
# IDENTIFY YEAR COLUMNS
# ============================================================

year_columns = {
    2023: "Value Exc. VAT",
    2024: "Value Exc. VAT.1",
    2025: "Value Exc. VAT.2",
    2026: "Value Exc. VAT.3",
}


for year, column in year_columns.items():

    if column not in source.columns:
        raise ValueError(
            f"Column for {year} not found: {column}"
        )

    source[column] = clean_number(
        source[column]
    )


# ============================================================
# ANNUAL SALES
# ============================================================

annual_rows = []

for year, column in year_columns.items():

    sales = source[column].sum()

    annual_rows.append(
        {
            "Year": year,
            "Sales": sales,
        }
    )


annual_sales = pd.DataFrame(
    annual_rows
)


# ============================================================
# YOY ANALYSIS
# ============================================================

annual_sales["Previous Year Sales"] = (
    annual_sales["Sales"].shift(1)
)

annual_sales["YoY Change"] = (
    annual_sales["Sales"]
    - annual_sales["Previous Year Sales"]
)

annual_sales["YoY Growth %"] = (
    annual_sales["YoY Change"]
    / annual_sales["Previous Year Sales"]
    * 100
)

annual_sales["YoY Growth %"] = (
    annual_sales["YoY Growth %"]
    .fillna(0)
)


# ============================================================
# PHASE 16 COMPARISON
# ============================================================

phase16 = pd.read_excel(
    PHASE16_FILE,
    sheet_name="Clean Transactions"
)

phase16_value_column = "Value Exc. VAT"

if phase16_value_column not in phase16.columns:

    raise ValueError(
        "Phase 16 Value Exc. VAT column not found."
    )


phase16[phase16_value_column] = clean_number(
    phase16[phase16_value_column]
)


phase16_sales = phase16[
    phase16_value_column
].sum()


source_2026_sales = annual_sales.loc[
    annual_sales["Year"] == 2026,
    "Sales"
].iloc[0]


variance = (
    source_2026_sales
    - phase16_sales
)


variance_pct = (
    variance
    / source_2026_sales
    * 100
    if source_2026_sales != 0
    else 0
)


reconciliation = pd.DataFrame(
    {
        "Metric": [
            "Original Source 2026 Sales",
            "Phase 16 Processed Sales",
            "Variance",
            "Variance % of Source",
        ],
        "Value": [
            source_2026_sales,
            phase16_sales,
            variance,
            variance_pct,
        ],
    }
)


# ============================================================
# SOURCE REGION ANALYSIS
# ============================================================

if "Region Name" in source.columns:

    region_columns = [
        "Region Name",
        "Value Exc. VAT",
        "Value Exc. VAT.1",
        "Value Exc. VAT.2",
        "Value Exc. VAT.3",
    ]

    region_source = source[
        region_columns
    ].copy()

    region_source = (
        region_source
        .groupby("Region Name", dropna=False)
        .sum(numeric_only=True)
        .reset_index()
    )

else:

    region_source = pd.DataFrame()


# ============================================================
# REGION YEAR ANALYSIS
# ============================================================

region_year_rows = []

if not region_source.empty:

    for _, row in region_source.iterrows():

        region = row["Region Name"]

        for year, column in year_columns.items():

            region_year_rows.append(
                {
                    "Region": region,
                    "Year": year,
                    "Sales": row[column],
                }
            )


region_year = pd.DataFrame(
    region_year_rows
)


# ============================================================
# REGION 2026 RANKING
# ============================================================

if not region_source.empty:

    region_2026 = region_source[
        [
            "Region Name",
            "Value Exc. VAT.3"
        ]
    ].copy()

    region_2026 = region_2026.rename(
        columns={
            "Region Name": "Region",
            "Value Exc. VAT.3": "Sales 2026",
        }
    )

    region_2026 = region_2026.sort_values(
        "Sales 2026",
        ascending=False
    )

    region_2026["Share of 2026 %"] = (
        region_2026["Sales 2026"]
        / source_2026_sales
        * 100
    )

else:

    region_2026 = pd.DataFrame()


# ============================================================
# SELLING TYPE ANALYSIS
# ============================================================

if "Selling Type" in source.columns:

    selling_columns = [
        "Selling Type",
        "Value Exc. VAT",
        "Value Exc. VAT.1",
        "Value Exc. VAT.2",
        "Value Exc. VAT.3",
    ]

    selling_source = source[
        selling_columns
    ].copy()

    selling_source = (
        selling_source
        .groupby("Selling Type", dropna=False)
        .sum(numeric_only=True)
        .reset_index()
    )

else:

    selling_source = pd.DataFrame()


# ============================================================
# SELLING TYPE 2026
# ============================================================

if not selling_source.empty:

    selling_2026 = selling_source[
        [
            "Selling Type",
            "Value Exc. VAT.3"
        ]
    ].copy()

    selling_2026 = selling_2026.rename(
        columns={
            "Value Exc. VAT.3": "Sales 2026"
        }
    )

    selling_2026 = selling_2026.sort_values(
        "Sales 2026",
        ascending=False
    )

    selling_2026["Share of 2026 %"] = (
        selling_2026["Sales 2026"]
        / source_2026_sales
        * 100
    )

else:

    selling_2026 = pd.DataFrame()


# ============================================================
# EXECUTIVE ANSWERS
# ============================================================

executive_answers = pd.DataFrame(
    {
        "Question": [
            "What were total sales in 2023?",
            "What were total sales in 2024?",
            "What were total sales in 2025?",
            "What were total sales in 2026?",
            "What was 2024 YoY growth?",
            "What was 2025 YoY growth?",
            "What was 2026 YoY growth?",
            "What does Phase 16 report as sales?",
            "What is the difference between source 2026 and Phase 16?",
            "What percentage of source 2026 sales is represented by Phase 16?",
        ],
        "Answer": [
            money(annual_sales.loc[
                annual_sales["Year"] == 2023,
                "Sales"
            ].iloc[0]),

            money(annual_sales.loc[
                annual_sales["Year"] == 2024,
                "Sales"
            ].iloc[0]),

            money(annual_sales.loc[
                annual_sales["Year"] == 2025,
                "Sales"
            ].iloc[0]),

            money(source_2026_sales),

            pct(annual_sales.loc[
                annual_sales["Year"] == 2024,
                "YoY Growth %"
            ].iloc[0]),

            pct(annual_sales.loc[
                annual_sales["Year"] == 2025,
                "YoY Growth %"
            ].iloc[0]),

            pct(annual_sales.loc[
                annual_sales["Year"] == 2026,
                "YoY Growth %"
            ].iloc[0]),

            money(phase16_sales),

            money(variance),

            pct(
                phase16_sales
                / source_2026_sales
                * 100
            ),
        ],
    }
)


# ============================================================
# MANAGEMENT CONCLUSION
# ============================================================

management_conclusion = pd.DataFrame(
    {
        "Area": [
            "Annual Sales Trend",
            "2026 Source Sales",
            "Phase 16 Sales",
            "Reconciliation Issue",
            "Management Interpretation",
        ],
        "Conclusion": [
            (
                "Sales increased from 2023 to 2025, "
                "while 2026 is lower than 2025 based on "
                "the available source-period data."
            ),

            (
                f"The original source reports "
                f"{money(source_2026_sales)} for 2026."
            ),

            (
                f"Phase 16 currently reports "
                f"{money(phase16_sales)}."
            ),

            (
                f"The difference is {money(variance)}, "
                f"meaning Phase 16 represents approximately "
                f"{pct(phase16_sales / source_2026_sales * 100)} "
                f"of the source 2026 value."
            ),

            (
                "The difference must be investigated before "
                "the Phase 16 processed figure is presented "
                "as the complete 2026 sales total."
            ),
        ],
    }
)


# ============================================================
# SOURCE STRUCTURE
# ============================================================

source_structure = pd.DataFrame(
    {
        "Metric": [
            "Source Workbook",
            "Source Sheet",
            "Rows Loaded",
            "2023 Sales Column",
            "2024 Sales Column",
            "2025 Sales Column",
            "2026 Sales Column",
        ],
        "Value": [
            os.path.basename(SOURCE_FILE),
            "Sell-in_RegionTownAreaCustomer",
            len(source),
            "Value Exc. VAT",
            "Value Exc. VAT.1",
            "Value Exc. VAT.2",
            "Value Exc. VAT.3",
        ],
    }
)


# ============================================================
# WRITE OUTPUT
# ============================================================

print("\nWriting Phase 18 workbook...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    annual_sales.to_excel(
        writer,
        sheet_name="Annual Sales",
        index=False
    )

    reconciliation.to_excel(
        writer,
        sheet_name="Reconciliation",
        index=False
    )

    region_year.to_excel(
        writer,
        sheet_name="Region Year Analysis",
        index=False
    )

    region_2026.to_excel(
        writer,
        sheet_name="2026 Region Ranking",
        index=False
    )

    selling_2026.to_excel(
        writer,
        sheet_name="2026 Selling Type",
        index=False
    )

    executive_answers.to_excel(
        writer,
        sheet_name="Executive Answers",
        index=False
    )

    management_conclusion.to_excel(
        writer,
        sheet_name="Management Conclusion",
        index=False
    )

    source_structure.to_excel(
        writer,
        sheet_name="Source Structure",
        index=False
    )

    source.to_excel(
        writer,
        sheet_name="Source Data",
        index=False
    )


# ============================================================
# FINAL REPORT
# ============================================================

print("\n" + "=" * 70)
print("PHASE 18 COMPLETE")
print("=" * 70)

print("\nANNUAL SALES")

for _, row in annual_sales.iterrows():

    print(
        f"{int(row['Year'])}: "
        f"{money(row['Sales'])}"
    )


print("\nYOY GROWTH")

for _, row in annual_sales.iterrows():

    if int(row["Year"]) == 2023:
        continue

    print(
        f"{int(row['Year'])}: "
        f"{pct(row['YoY Growth %'])}"
    )


print("\nRECONCILIATION")

print(
    f"Source 2026:     {money(source_2026_sales)}"
)

print(
    f"Phase 16:        {money(phase16_sales)}"
)

print(
    f"Difference:      {money(variance)}"
)

print(
    f"Phase16 Coverage: "
    f"{pct(phase16_sales / source_2026_sales * 100)}"
)

print("\nOUTPUT FILE:")

print(OUTPUT_FILE)

print("\nPhase 18 analysis completed successfully.")