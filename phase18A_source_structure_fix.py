import os
import pandas as pd

# ============================================================
# PHASE 18A - MIKA SALES INTELLIGENCE
# SOURCE STRUCTURE & RECONCILIATION FIX
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
    "Phase18A_Source_Structure_Fix.xlsx"
)

# ============================================================
# HEADER
# ============================================================

print("=" * 70)
print("PHASE 18A - MIKA SALES INTELLIGENCE")
print("SOURCE STRUCTURE & RECONCILIATION FIX")
print("=" * 70)
print()

# ============================================================
# CHECK FILES
# ============================================================

if not os.path.exists(SOURCE_FILE):
    raise FileNotFoundError(
        f"Source workbook not found:\n{SOURCE_FILE}"
    )

if not os.path.exists(PHASE16_FILE):
    raise FileNotFoundError(
        f"Phase 16 workbook not found:\n{PHASE16_FILE}"
    )

os.makedirs(OUTPUT_DIR, exist_ok=True)

print("Source files found successfully.")
print()

# ============================================================
# LOAD ORIGINAL SOURCE
# ============================================================

print("Loading original MIKA source workbook...")

source = pd.read_excel(
    SOURCE_FILE,
    sheet_name="Sell-in_RegionTownAreaCustomer",
    header=7
)

print(f"Source rows loaded: {len(source)}")
print()

# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

source.columns = (
    source.columns
    .astype(str)
    .str.strip()
)

# ============================================================
# IDENTIFY YEAR SALES COLUMNS
# ============================================================

YEAR_COLUMNS = {
    2023: "Value Exc. VAT",
    2024: "Value Exc. VAT.1",
    2025: "Value Exc. VAT.2",
    2026: "Value Exc. VAT.3",
}

# ============================================================
# CONVERT SALES COLUMNS TO NUMERIC
# ============================================================

for year, column in YEAR_COLUMNS.items():

    if column not in source.columns:
        raise KeyError(
            f"Required column for {year} not found: {column}"
        )

    source[column] = pd.to_numeric(
        source[column],
        errors="coerce"
    ).fillna(0)

# ============================================================
# CLEAN TEXT COLUMNS
# ============================================================

TEXT_COLUMNS = [
    "Selling Type",
    "Region Name",
    "Zone Name",
    "Town/Area Name",
    "Road / Street Name",
    "Stockist /Customer Chain Name",
]

for column in TEXT_COLUMNS:

    if column in source.columns:

        source[column] = (
            source[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

# ============================================================
# REMOVE GRAND TOTAL / SUMMARY ROWS
# ============================================================

print("Cleaning summary rows...")

summary_words = [
    "GRAND TOTAL",
    "TOTAL",
    "SUBTOTAL",
]

summary_mask = pd.Series(
    False,
    index=source.index
)

for column in [
    "Selling Type",
    "Region Name",
    "Zone Name",
    "Town/Area Name",
]:

    if column in source.columns:

        values = (
            source[column]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.strip()
        )

        for word in summary_words:

            summary_mask = (
                summary_mask
                | values.eq(word)
                | values.str.startswith(
                    word + " ",
                    na=False
                )
            )

summary_rows_removed = int(summary_mask.sum())

clean_source = source.loc[
    ~summary_mask
].copy()

print(
    f"Summary rows removed: {summary_rows_removed}"
)

# ============================================================
# REMOVE ROWS WITHOUT A REGION
# ============================================================

blank_region_mask = (
    clean_source["Region Name"]
    .fillna("")
    .astype(str)
    .str.strip()
    .eq("")
)

blank_region_rows = int(
    blank_region_mask.sum()
)

clean_source = clean_source.loc[
    ~blank_region_mask
].copy()

print(
    f"Rows with blank Region removed: {blank_region_rows}"
)

# ============================================================
# REMOVE ROWS WITHOUT SELLING TYPE
# ============================================================

blank_type_mask = (
    clean_source["Selling Type"]
    .fillna("")
    .astype(str)
    .str.strip()
    .eq("")
)

blank_type_rows = int(
    blank_type_mask.sum()
)

clean_source = clean_source.loc[
    ~blank_type_mask
].copy()

print(
    f"Rows with blank Selling Type removed: {blank_type_rows}"
)

# ============================================================
# 2026 SOURCE TOTAL
# ============================================================

source_2026 = clean_source[
    YEAR_COLUMNS[2026]
].sum()

print()
print(
    f"Clean 2026 source sales: "
    f"KSh {source_2026:,.2f}"
)

# ============================================================
# REGIONAL ANALYSIS
# ============================================================

regional_2026 = (
    clean_source
    .groupby(
        "Region Name",
        as_index=False
    )
    .agg(
        Sales_2026=(
            YEAR_COLUMNS[2026],
            "sum"
        ),
        Records=(
            YEAR_COLUMNS[2026],
            "count"
        )
    )
    .sort_values(
        "Sales_2026",
        ascending=False
    )
)

regional_2026["Sales %"] = (
    regional_2026["Sales_2026"]
    / source_2026
    * 100
)

regional_2026 = regional_2026.reset_index(
    drop=True
)

# ============================================================
# SELLING TYPE ANALYSIS
# ============================================================

selling_2026 = (
    clean_source
    .groupby(
        "Selling Type",
        as_index=False
    )
    .agg(
        Sales_2026=(
            YEAR_COLUMNS[2026],
            "sum"
        ),
        Records=(
            YEAR_COLUMNS[2026],
            "count"
        )
    )
    .sort_values(
        "Sales_2026",
        ascending=False
    )
)

selling_2026["Sales %"] = (
    selling_2026["Sales_2026"]
    / source_2026
    * 100
)

selling_2026 = selling_2026.reset_index(
    drop=True
)

# ============================================================
# YEAR TOTALS
# ============================================================

annual_sales = []

for year, column in YEAR_COLUMNS.items():

    total = clean_source[column].sum()

    annual_sales.append(
        {
            "Year": year,
            "Sales": total
        }
    )

annual_sales = pd.DataFrame(
    annual_sales
)

# ============================================================
# YOY GROWTH
# ============================================================

annual_sales["YoY Growth %"] = (
    annual_sales["Sales"]
    .pct_change()
    * 100
)

# ============================================================
# LOAD PHASE 16
# ============================================================

phase16 = pd.read_excel(
    PHASE16_FILE,
    sheet_name="Clean Transactions"
)

phase16.columns = (
    phase16.columns
    .astype(str)
    .str.strip()
)

phase16_value_column = "Value Exc. VAT"

if phase16_value_column not in phase16.columns:

    raise KeyError(
        "Phase 16 Value Exc. VAT column not found."
    )

phase16[
    phase16_value_column
] = pd.to_numeric(
    phase16[phase16_value_column],
    errors="coerce"
).fillna(0)

phase16_sales = phase16[
    phase16_value_column
].sum()

# ============================================================
# RECONCILIATION
# ============================================================

variance = (
    source_2026
    - phase16_sales
)

coverage = (
    phase16_sales
    / source_2026
    * 100
    if source_2026 != 0
    else 0
)

missing_coverage = (
    100 - coverage
)

reconciliation = pd.DataFrame(
    {
        "Metric": [
            "2026 Original Source Sales",
            "Phase 16 Processed Sales",
            "Variance",
            "Phase 16 Coverage %",
            "Unrepresented Source %",
        ],
        "Value": [
            source_2026,
            phase16_sales,
            variance,
            coverage,
            missing_coverage,
        ]
    }
)

# ============================================================
# REGIONAL RECONCILIATION
# ============================================================

regional_total = regional_2026[
    "Sales_2026"
].sum()

regional_difference = (
    regional_total
    - source_2026
)

regional_reconciliation = pd.DataFrame(
    {
        "Metric": [
            "Clean Source 2026",
            "Regional Total",
            "Difference",
        ],
        "Value": [
            source_2026,
            regional_total,
            regional_difference,
        ]
    }
)

# ============================================================
# SELLING TYPE RECONCILIATION
# ============================================================

selling_total = selling_2026[
    "Sales_2026"
].sum()

selling_difference = (
    selling_total
    - source_2026
)

selling_reconciliation = pd.DataFrame(
    {
        "Metric": [
            "Clean Source 2026",
            "Selling Type Total",
            "Difference",
        ],
        "Value": [
            source_2026,
            selling_total,
            selling_difference,
        ]
    }
)

# ============================================================
# DATA CLEANING SUMMARY
# ============================================================

cleaning_summary = pd.DataFrame(
    {
        "Metric": [
            "Original Source Rows",
            "Summary Rows Removed",
            "Blank Region Rows Removed",
            "Blank Selling Type Rows Removed",
            "Final Clean Rows",
        ],
        "Value": [
            len(source),
            summary_rows_removed,
            blank_region_rows,
            blank_type_rows,
            len(clean_source),
        ]
    }
)

# ============================================================
# MANAGEMENT ANSWERS
# ============================================================

if not regional_2026.empty:

    top_region = regional_2026.iloc[0]

    top_region_name = top_region[
        "Region Name"
    ]

    top_region_sales = top_region[
        "Sales_2026"
    ]

    top_region_pct = top_region[
        "Sales %"
    ]

else:

    top_region_name = "N/A"
    top_region_sales = 0
    top_region_pct = 0


if not selling_2026.empty:

    top_type = selling_2026.iloc[0]

    top_type_name = top_type[
        "Selling Type"
    ]

    top_type_sales = top_type[
        "Sales_2026"
    ]

    top_type_pct = top_type[
        "Sales %"
    ]

else:

    top_type_name = "N/A"
    top_type_sales = 0
    top_type_pct = 0


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
            "Which region has the highest 2026 sales?",
            "Which selling type has the highest 2026 sales?",
        ],
        "Answer": [
            f"KSh {annual_sales.loc[annual_sales['Year'] == 2023, 'Sales'].iloc[0]:,.2f}",
            f"KSh {annual_sales.loc[annual_sales['Year'] == 2024, 'Sales'].iloc[0]:,.2f}",
            f"KSh {annual_sales.loc[annual_sales['Year'] == 2025, 'Sales'].iloc[0]:,.2f}",
            f"KSh {annual_sales.loc[annual_sales['Year'] == 2026, 'Sales'].iloc[0]:,.2f}",
            f"{annual_sales.loc[annual_sales['Year'] == 2024, 'YoY Growth %'].iloc[0]:.2f}%",
            f"{annual_sales.loc[annual_sales['Year'] == 2025, 'YoY Growth %'].iloc[0]:.2f}%",
            f"{annual_sales.loc[annual_sales['Year'] == 2026, 'YoY Growth %'].iloc[0]:.2f}%",
            f"KSh {phase16_sales:,.2f}",
            f"KSh {variance:,.2f}",
            f"{coverage:.2f}%",
            f"{top_region_name} — KSh {top_region_sales:,.2f} ({top_region_pct:.2f}%)",
            f"{top_type_name} — KSh {top_type_sales:,.2f} ({top_type_pct:.2f}%)",
        ]
    }
)

# ============================================================
# PRINT RESULTS
# ============================================================

print()
print("=" * 70)
print("CLEANED 2026 REGIONAL ANALYSIS")
print("=" * 70)

print(
    regional_2026.to_string(
        index=False
    )
)

print()
print("=" * 70)
print("CLEANED 2026 SELLING TYPE ANALYSIS")
print("=" * 70)

print(
    selling_2026.to_string(
        index=False
    )
)

print()
print("=" * 70)
print("RECONCILIATION")
print("=" * 70)

print(
    f"Source 2026:     KSh {source_2026:,.2f}"
)

print(
    f"Phase 16:        KSh {phase16_sales:,.2f}"
)

print(
    f"Difference:      KSh {variance:,.2f}"
)

print(
    f"Phase16 Coverage: {coverage:.2f}%"
)

print()
print(
    f"Highest Region: {top_region_name} "
    f"— KSh {top_region_sales:,.2f} "
    f"({top_region_pct:.2f}%)"
)

print(
    f"Highest Selling Type: {top_type_name} "
    f"— KSh {top_type_sales:,.2f} "
    f"({top_type_pct:.2f}%)"
)

# ============================================================
# WRITE OUTPUT
# ============================================================

print()
print("Writing Phase 18A workbook...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    annual_sales.to_excel(
        writer,
        sheet_name="Annual Sales",
        index=False
    )

    regional_2026.to_excel(
        writer,
        sheet_name="2026 Sales by Region",
        index=False
    )

    selling_2026.to_excel(
        writer,
        sheet_name="2026 Selling Type",
        index=False
    )

    reconciliation.to_excel(
        writer,
        sheet_name="Reconciliation",
        index=False
    )

    regional_reconciliation.to_excel(
        writer,
        sheet_name="Regional Reconciliation",
        index=False
    )

    selling_reconciliation.to_excel(
        writer,
        sheet_name="Selling Reconciliation",
        index=False
    )

    cleaning_summary.to_excel(
        writer,
        sheet_name="Cleaning Summary",
        index=False
    )

    executive_answers.to_excel(
        writer,
        sheet_name="Executive Answers",
        index=False
    )

    clean_source.to_excel(
        writer,
        sheet_name="Clean Source",
        index=False
    )

# ============================================================
# COMPLETE
# ============================================================

print()
print("=" * 70)
print("PHASE 18A COMPLETE")
print("=" * 70)

print()
print("OUTPUT FILE:")
print(OUTPUT_FILE)

print()
print("Next step:")
print(
    "Run the Phase 18A validation commands before updating "
    "the Streamlit dashboard."
)