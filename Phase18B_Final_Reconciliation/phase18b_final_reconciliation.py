import os
import pandas as pd

# ============================================================
# PHASE 18B - MIKA SALES INTELLIGENCE
# FINAL SOURCE RECONCILIATION
# ============================================================

print("=" * 70)
print("PHASE 18B - MIKA SALES INTELLIGENCE")
print("FINAL SOURCE RECONCILIATION")
print("=" * 70)

# ============================================================
# FILE CONFIGURATION
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
    "Phase18_MIKA_Sales_Intelligence",
    "Phase18B_Final_Reconciliation"
)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "Phase18B_MIKA_Final_Reconciliation.xlsx"
)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# CHECK FILES
# ============================================================

if not os.path.exists(SOURCE_FILE):
    raise FileNotFoundError(f"Source file not found:\n{SOURCE_FILE}")

if not os.path.exists(PHASE16_FILE):
    raise FileNotFoundError(f"Phase 16 file not found:\n{PHASE16_FILE}")

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

print(f"Source rows loaded: {len(source)}")

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
    "Value Exc. VAT.3"
]

missing = [
    c for c in required_columns
    if c not in source.columns
]

if missing:
    raise ValueError(
        f"Missing required source columns: {missing}"
    )

# ============================================================
# NUMERIC CONVERSION
# ============================================================

year_columns = {
    2023: "Value Exc. VAT",
    2024: "Value Exc. VAT.1",
    2025: "Value Exc. VAT.2",
    2026: "Value Exc. VAT.3"
}

for year, column in year_columns.items():
    source[column] = pd.to_numeric(
        source[column],
        errors="coerce"
    ).fillna(0)

# ============================================================
# CLEAN TEXT FIELDS
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
# IDENTIFY GRAND TOTAL
# ============================================================

grand_total_rows = source[
    source["Selling Type"].str.upper().eq("GRAND TOTAL")
]

if len(grand_total_rows) != 1:
    raise ValueError(
        "Expected exactly one GRAND TOTAL row, "
        f"but found {len(grand_total_rows)}."
    )

grand_total_row = grand_total_rows.iloc[0]

source_2023 = float(grand_total_row["Value Exc. VAT"])
source_2024 = float(grand_total_row["Value Exc. VAT.1"])
source_2025 = float(grand_total_row["Value Exc. VAT.2"])
source_2026 = float(grand_total_row["Value Exc. VAT.3"])

print("\nOFFICIAL SOURCE GRAND TOTALS")

print(f"2023: KSh {source_2023:,.2f}")
print(f"2024: KSh {source_2024:,.2f}")
print(f"2025: KSh {source_2025:,.2f}")
print(f"2026: KSh {source_2026:,.2f}")

# ============================================================
# YOY GROWTH
# ============================================================

def yoy(current, previous):
    if previous == 0:
        return None

    return ((current - previous) / previous) * 100


yoy_2024 = yoy(source_2024, source_2023)
yoy_2025 = yoy(source_2025, source_2024)
yoy_2026 = yoy(source_2026, source_2025)

print("\nYOY GROWTH")

print(f"2024: {yoy_2024:.2f}%")
print(f"2025: {yoy_2025:.2f}%")
print(f"2026: {yoy_2026:.2f}%")

# ============================================================
# REMOVE SUMMARY ROWS
# ============================================================

summary_mask = (
    source["Selling Type"]
    .str.upper()
    .isin([
        "LOCAL SALES TOTAL",
        "GRAND TOTAL"
    ])
)

summary_rows_removed = int(summary_mask.sum())

detail = source.loc[
    ~summary_mask
].copy()

# ============================================================
# REMOVE BLANK REGIONS
# ============================================================

blank_region_mask = (
    detail["Region Name"]
    .eq("")
)

blank_region_removed = int(blank_region_mask.sum())

detail = detail.loc[
    ~blank_region_mask
].copy()

print("\nSOURCE CLEANING")

print(
    f"Summary rows removed: {summary_rows_removed}"
)

print(
    f"Rows with blank Region removed: {blank_region_removed}"
)

print(
    f"Clean detail rows: {len(detail)}"
)

# ============================================================
# 2026 REGIONAL ANALYSIS
# ============================================================

region_2026 = (
    detail
    .groupby("Region Name", as_index=False)
    ["Value Exc. VAT.3"]
    .sum()
    .rename(
        columns={
            "Value Exc. VAT.3": "Sales_2026"
        }
    )
)

region_2026 = region_2026.sort_values(
    "Sales_2026",
    ascending=False
)

region_2026["Sales %"] = (
    region_2026["Sales_2026"]
    / source_2026
    * 100
)

region_2026["Records"] = (
    detail
    .groupby("Region Name")
    .size()
    .reindex(region_2026["Region Name"])
    .values
)

region_2026 = region_2026[
    [
        "Region Name",
        "Sales_2026",
        "Records",
        "Sales %"
    ]
]

# ============================================================
# SELLING TYPE ANALYSIS
# ============================================================

selling_type_2026 = (
    detail
    .groupby("Selling Type", as_index=False)
    ["Value Exc. VAT.3"]
    .sum()
    .rename(
        columns={
            "Value Exc. VAT.3": "Sales_2026"
        }
    )
)

selling_type_2026 = selling_type_2026.sort_values(
    "Sales_2026",
    ascending=False
)

selling_type_2026["Sales %"] = (
    selling_type_2026["Sales_2026"]
    / source_2026
    * 100
)

selling_type_2026["Records"] = (
    detail
    .groupby("Selling Type")
    .size()
    .reindex(selling_type_2026["Selling Type"])
    .values
)

selling_type_2026 = selling_type_2026[
    [
        "Selling Type",
        "Sales_2026",
        "Records",
        "Sales %"
    ]
]

# ============================================================
# YEARLY SALES TABLE
# ============================================================

annual_sales = pd.DataFrame({
    "Year": [2023, 2024, 2025, 2026],
    "Sales": [
        source_2023,
        source_2024,
        source_2025,
        source_2026
    ]
})

annual_sales["YoY Growth %"] = [
    None,
    yoy_2024,
    yoy_2025,
    yoy_2026
]

# ============================================================
# LOAD PHASE 16
# ============================================================

print("\nLoading Phase 16...")

phase16 = pd.read_excel(
    PHASE16_FILE,
    sheet_name="Clean Transactions"
)

phase16_sales = pd.to_numeric(
    phase16["Value Exc. VAT"],
    errors="coerce"
).fillna(0).sum()

# ============================================================
# RECONCILIATION
# ============================================================

difference = phase16_sales - source_2026

coverage = (
    phase16_sales / source_2026 * 100
    if source_2026 != 0
    else None
)

source_minus_phase16 = source_2026 - phase16_sales

print("\nRECONCILIATION")

print(
    f"Source 2026:     KSh {source_2026:,.2f}"
)

print(
    f"Phase 16:        KSh {phase16_sales:,.2f}"
)

print(
    f"Difference:      KSh {difference:,.2f}"
)

print(
    f"Phase16 / Source: {coverage:.2f}%"
)

# ============================================================
# RECONCILIATION STATUS
# ============================================================

if abs(difference) < 1:
    reconciliation_status = "PASS"
elif phase16_sales > source_2026:
    reconciliation_status = (
        "SCOPE / PERIOD MISMATCH - PHASE 16 EXCEEDS SOURCE"
    )
else:
    reconciliation_status = (
        "SCOPE / PERIOD MISMATCH - PHASE 16 BELOW SOURCE"
    )

# ============================================================
# TOP REGION
# ============================================================

if not region_2026.empty:

    top_region = region_2026.iloc[0]

    highest_region = top_region["Region Name"]
    highest_region_sales = float(
        top_region["Sales_2026"]
    )
    highest_region_share = float(
        top_region["Sales %"]
    )

else:

    highest_region = None
    highest_region_sales = 0
    highest_region_share = 0

# ============================================================
# TOP SELLING TYPE
# ============================================================

if not selling_type_2026.empty:

    top_type = selling_type_2026.iloc[0]

    highest_selling_type = top_type["Selling Type"]
    highest_selling_type_sales = float(
        top_type["Sales_2026"]
    )
    highest_selling_type_share = float(
        top_type["Sales %"]
    )

else:

    highest_selling_type = None
    highest_selling_type_sales = 0
    highest_selling_type_share = 0

print("\n2026 ANALYSIS")

print(
    f"Highest Region: {highest_region} — "
    f"KSh {highest_region_sales:,.2f} "
    f"({highest_region_share:.2f}%)"
)

print(
    f"Highest Selling Type: {highest_selling_type} — "
    f"KSh {highest_selling_type_sales:,.2f} "
    f"({highest_selling_type_share:.2f}%)"
)

print(
    f"\nReconciliation Status: "
    f"{reconciliation_status}"
)

# ============================================================
# EXECUTIVE ANSWERS
# ============================================================

executive_answers = pd.DataFrame({
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
        "What is the reconciliation status?"
    ],

    "Answer": [
        source_2023,
        source_2024,
        source_2025,
        source_2026,
        yoy_2024,
        yoy_2025,
        yoy_2026,
        phase16_sales,
        difference,
        coverage,
        highest_region,
        highest_selling_type,
        reconciliation_status
    ]
})

# ============================================================
# MANAGEMENT FINDING
# ============================================================

if phase16_sales > source_2026:

    management_finding = (
        "Phase 16 sales exceed the comparable 2026 Grand Total "
        "in the original source workbook. This indicates that "
        "the Phase 16 dataset and the source 2026 figure should "
        "not currently be treated as directly comparable. "
        "The reporting period, source scope, filters, and "
        "processing rules should be reconciled before using "
        "either figure as the definitive 2026 management total."
    )

elif phase16_sales < source_2026:

    management_finding = (
        "Phase 16 sales are below the comparable 2026 Grand "
        "Total in the original source workbook. The difference "
        "should be investigated to determine whether records "
        "were excluded during processing or whether the two "
        "datasets represent different reporting scopes."
    )

else:

    management_finding = (
        "Phase 16 reconciles to the 2026 Grand Total in the "
        "original source workbook."
    )

management_summary = pd.DataFrame({
    "Management Item": [
        "Official 2026 Source Sales",
        "Phase 16 Sales",
        "Difference",
        "Phase 16 / Source",
        "Reconciliation Status",
        "Highest Region",
        "Highest Region Sales",
        "Highest Region Share",
        "Highest Selling Type",
        "Highest Selling Type Sales",
        "Highest Selling Type Share",
        "Management Finding"
    ],

    "Value": [
        source_2026,
        phase16_sales,
        difference,
        coverage,
        reconciliation_status,
        highest_region,
        highest_region_sales,
        highest_region_share,
        highest_selling_type,
        highest_selling_type_sales,
        highest_selling_type_share,
        management_finding
    ]
})

# ============================================================
# SOURCE STRUCTURE TABLE
# ============================================================

source_structure = pd.DataFrame({
    "Row Type": [
        "Detail Rows",
        "LOCAL SALES Total",
        "Grand Total"
    ],

    "Purpose": [
        "Individual source records used for detailed analysis",
        "Subtotal and excluded from detail aggregation",
        "Official overall source total"
    ],

    "Treatment": [
        "Included",
        "Excluded from detail aggregation",
        "Used as official source total"
    ]
})

# ============================================================
# VALIDATION
# ============================================================

region_sum = region_2026["Sales_2026"].sum()

selling_type_sum = selling_type_2026["Sales_2026"].sum()

region_difference = region_sum - source_2026

selling_type_difference = selling_type_sum - source_2026

validation = pd.DataFrame({
    "Validation": [
        "Official 2026 Grand Total",
        "Sum of regional detail sales",
        "Regional difference",
        "Sum of selling type detail sales",
        "Selling type difference",
        "Phase 16 sales",
        "Phase 16 vs source difference"
    ],

    "Value": [
        source_2026,
        region_sum,
        region_difference,
        selling_type_sum,
        selling_type_difference,
        phase16_sales,
        difference
    ]
})

# ============================================================
# WRITE OUTPUT
# ============================================================

print("\nWriting Phase 18B workbook...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    annual_sales.to_excel(
        writer,
        sheet_name="Annual Sales",
        index=False
    )

    region_2026.to_excel(
        writer,
        sheet_name="2026 Sales by Region",
        index=False
    )

    selling_type_2026.to_excel(
        writer,
        sheet_name="2026 Selling Type",
        index=False
    )

    executive_answers.to_excel(
        writer,
        sheet_name="Executive Answers",
        index=False
    )

    management_summary.to_excel(
        writer,
        sheet_name="Management Summary",
        index=False
    )

    validation.to_excel(
        writer,
        sheet_name="Validation",
        index=False
    )

    source_structure.to_excel(
        writer,
        sheet_name="Source Structure",
        index=False
    )

    detail.to_excel(
        writer,
        sheet_name="Clean Source Detail",
        index=False
    )

# ============================================================
# COMPLETION
# ============================================================

print("\n" + "=" * 70)
print("PHASE 18B COMPLETE")
print("=" * 70)

print("\nOFFICIAL SOURCE TOTALS")

print(f"2023: KSh {source_2023:,.2f}")
print(f"2024: KSh {source_2024:,.2f}")
print(f"2025: KSh {source_2025:,.2f}")
print(f"2026: KSh {source_2026:,.2f}")

print("\nYOY GROWTH")

print(f"2024: {yoy_2024:.2f}%")
print(f"2025: {yoy_2025:.2f}%")
print(f"2026: {yoy_2026:.2f}%")

print("\nRECONCILIATION")

print(f"Source 2026:     KSh {source_2026:,.2f}")
print(f"Phase 16:        KSh {phase16_sales:,.2f}")
print(f"Difference:      KSh {difference:,.2f}")
print(f"Phase16 / Source: {coverage:.2f}%")

print("\nSTATUS")
print(reconciliation_status)

print("\nOUTPUT FILE:")
print(OUTPUT_FILE)

print("\nNext step:")
print("Run the validation command before updating the dashboard.")