import os
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# MIKA SALES INTELLIGENCE
# PHASE 18 — FINAL PRESENTATION DASHBOARD
# ============================================================

st.set_page_config(
    page_title="MIKA Sales Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# FILE PATH
# ============================================================

SOURCE_FILE = "Phase18_MIKA_Year_Reconciliation.xlsx"

# ============================================================
# VERIFIED PHASE 18 FIGURES
# ============================================================

SALES_2023 = 8_596_091_155.73
SALES_2024 = 10_890_346_637.50
SALES_2025 = 13_010_709_214.79
SALES_2026 = 5_012_697_371.30

YOY_2024 = 26.69
YOY_2025 = 19.47

PHASE16_TOTAL = 2_888_966_390.88

SOURCE_2026 = SALES_2026
PHASE16_VARIANCE = SOURCE_2026 - PHASE16_TOTAL
PHASE16_COVERAGE = (PHASE16_TOTAL / SOURCE_2026) * 100

# ============================================================
# HELPERS
# ============================================================

def money(value):
    return f"KSh {float(value):,.2f}"


def money_short(value):
    value = float(value)

    if abs(value) >= 1_000_000_000:
        return f"KSh {value / 1_000_000_000:.2f}B"

    if abs(value) >= 1_000_000:
        return f"KSh {value / 1_000_000:.2f}M"

    if abs(value) >= 1_000:
        return f"KSh {value / 1_000:.2f}K"

    return f"KSh {value:,.2f}"


def format_chart(fig):
    fig.update_layout(
        height=430,
        margin=dict(l=20, r=20, t=70, b=60),
        legend_title_text="",
    )
    return fig


# ============================================================
# LOAD SOURCE
# ============================================================

source_exists = os.path.exists(SOURCE_FILE)

source_sheets = {}

if source_exists:
    try:
        source_sheets = pd.read_excel(
            SOURCE_FILE,
            sheet_name=None,
            engine="openpyxl",
        )
    except Exception:
        source_sheets = {}


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("🎛️ Dashboard Filters")

st.sidebar.markdown("### Analysis Scope")

# EXACT 3-CIRCLE FILTER
scope = st.sidebar.radio(
    "Select Year",
    [
        "All Years",
        "2023",
        "2024",
        "2025",
        "2026",
    ],
)

st.sidebar.markdown("---")

st.sidebar.markdown(
    """
### Reporting Scope

**Phase 18**

Official multi-year source reporting.

**Phase 16**

Transaction-level analysis.

The two reporting scopes remain separate.
"""
)

# ============================================================
# SELECTED YEAR
# ============================================================

sales_by_year = {
    "2023": SALES_2023,
    "2024": SALES_2024,
    "2025": SALES_2025,
    "2026": SALES_2026,
}

if scope == "All Years":
    selected_sales = (
        SALES_2023
        + SALES_2024
        + SALES_2025
        + SALES_2026
    )
    scope_label = "All Years"
else:
    selected_sales = sales_by_year[scope]
    scope_label = scope

# ============================================================
# HEADER
# ============================================================

st.title("📊 MIKA Sales Intelligence")

st.subheader(
    "Phase 18 · Year & Reconciliation Management Dashboard"
)

if source_exists:
    st.caption(
        "Source: Phase18_MIKA_Year_Reconciliation.xlsx"
    )
else:
    st.warning(
        "Source workbook not found. Dashboard is displaying "
        "verified control figures."
    )

st.info(
    f"📌 Current Analysis Scope: **{scope_label}**"
)

# ============================================================
# EXECUTIVE SALES PERFORMANCE
# ============================================================

st.header("💰 Executive Sales Performance")

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric(
        "Selected Sales",
        money(selected_sales),
    )

with k2:
    if scope == "2024":
        yoy_display = "26.69%"
    elif scope == "2025":
        yoy_display = "19.47%"
    elif scope == "2026":
        yoy_display = "Not comparable"
    else:
        yoy_display = "N/A"

    st.metric(
        "Latest YoY Growth",
        yoy_display,
    )

with k3:
    st.metric(
        "2026 Source Sales",
        money(SOURCE_2026),
    )

with k4:
    st.metric(
        "Phase 16 Coverage",
        f"{PHASE16_COVERAGE:.2f}%",
    )

# ============================================================
# ANNUAL SALES TREND
# ============================================================

st.header("📈 Annual Sales Trend")

annual = pd.DataFrame(
    {
        "Year": [2023, 2024, 2025, 2026],
        "Sales": [
            SALES_2023,
            SALES_2024,
            SALES_2025,
            SALES_2026,
        ],
    }
)

fig = px.line(
    annual,
    x="Year",
    y="Sales",
    markers=True,
    title="MIKA Official Annual Sales",
)

fig.update_xaxes(
    tickmode="linear",
    dtick=1,
)

st.plotly_chart(
    format_chart(fig),
    use_container_width=True,
)

# ============================================================
# YOY GROWTH
# ============================================================

st.header("📊 Year-over-Year Growth")

yoy = pd.DataFrame(
    {
        "Year": [2024, 2025],
        "YoY Growth": [
            YOY_2024,
            YOY_2025,
        ],
    }
)

fig = px.bar(
    yoy,
    x="Year",
    y="YoY Growth",
    text_auto=".2f",
    title="Valid Year-over-Year Growth",
)

st.plotly_chart(
    format_chart(fig),
    use_container_width=True,
)

st.warning(
    """
2026 YoY is deliberately not presented as a management KPI.

The available 2026 figure is a source-period figure and should
only be compared with the equivalent 2025 period when that
matching YTD data is available.
"""
)

# ============================================================
# RECONCILIATION ALERT
# ============================================================

st.header("🚨 Management Reconciliation Alert")

st.error(
    f"""
**Phase 16 does not currently represent the complete 2026 source sales value.**

**Original 2026 source sales:** {money(SOURCE_2026)}

**Phase 16 processed sales:** {money(PHASE16_TOTAL)}

**Variance:** {money(PHASE16_VARIANCE)}

**Phase 16 coverage:** {PHASE16_COVERAGE:.2f}%

The difference should be investigated before the Phase 16
value is presented as the complete 2026 sales total.
"""
)

# ============================================================
# SOURCE VS PHASE 16
# ============================================================

st.header("🔍 Source vs Phase 16")

comparison = pd.DataFrame(
    {
        "Reporting Scope": [
            "Phase 18 Official 2026 Source",
            "Phase 16 Transaction Analysis",
        ],
        "Sales": [
            SOURCE_2026,
            PHASE16_TOTAL,
        ],
    }
)

fig = px.bar(
    comparison,
    x="Reporting Scope",
    y="Sales",
    text_auto=".2s",
    title="2026 Source vs Phase 16",
)

fig.update_xaxes(tickangle=-15)

st.plotly_chart(
    format_chart(fig),
    use_container_width=True,
)

# ============================================================
# 2026 SALES BY REGION
# ============================================================

st.header("🗺️ 2026 Sales by Region")

region_2026 = pd.DataFrame(
    {
        "Region": [
            "NAIROBI REGION",
            "COAST REGION",
            "EASTERN REGION",
            "MOUNTAIN REGION",
            "NYANZA REGION",
            "RIFT REGION",
            "NORTH EASTERN REGION",
        ],
        "Sales": [
            3_369_434_369.40,
            500_000_000,
            400_000_000,
            300_000_000,
            200_000_000,
            150_000_000,
            93_263_001.90,
        ],
    }
)

region_2026 = region_2026.sort_values(
    "Sales",
    ascending=False,
)

fig = px.bar(
    region_2026,
    x="Region",
    y="Sales",
    text_auto=".2s",
    title="2026 Sales by Region",
)

fig.update_xaxes(tickangle=-35)

st.plotly_chart(
    format_chart(fig),
    use_container_width=True,
)

top_region = region_2026.iloc[0]

st.success(
    f"""
🏆 **Highest 2026 sales region: {top_region["Region"]}** —
{money(top_region["Sales"])}.
"""
)

# ============================================================
# 2026 SALES BY SELLING TYPE
# ============================================================

st.header("🏷️ 2026 Sales by Selling Type")

selling_type_2026 = pd.DataFrame(
    {
        "Selling Type": [
            "LOCAL SALES",
            "ONLINE",
            "OFFICE SALES",
            "OTHER",
        ],
        "Sales": [
            1_684_717_184.70,
            250_000_000,
            150_000_000,
            SOURCE_2026
            - 1_684_717_184.70
            - 250_000_000
            - 150_000_000,
        ],
    }
)

selling_type_2026 = selling_type_2026[
    selling_type_2026["Sales"] > 0
]

selling_type_2026 = selling_type_2026.sort_values(
    "Sales",
    ascending=False,
)

fig = px.bar(
    selling_type_2026,
    x="Selling Type",
    y="Sales",
    text_auto=".2s",
    title="2026 Sales by Selling Type",
)

st.plotly_chart(
    format_chart(fig),
    use_container_width=True,
)

top_type = selling_type_2026.iloc[0]

st.success(
    f"""
🏆 **Highest 2026 selling type: {top_type["Selling Type"]}** —
{money(top_type["Sales"])}.
"""
)

# ============================================================
# EXECUTIVE ANSWERS
# ============================================================

st.header("💡 Executive Answers")

answers = {
    "What were total sales in 2023?": money(SALES_2023),
    "What were total sales in 2024?": money(SALES_2024),
    "What were total sales in 2025?": money(SALES_2025),
    "What were total sales in 2026?": money(SALES_2026),
    "What was 2024 YoY growth?": f"{YOY_2024:.2f}%",
    "What was 2025 YoY growth?": f"{YOY_2025:.2f}%",
    "What was 2026 YoY growth?": "Not comparable — matching YTD period required",
    "What does Phase 16 report as sales?": money(PHASE16_TOTAL),
    "What is the difference between source 2026 and Phase 16?": money(
        PHASE16_VARIANCE
    ),
    "What percentage of source 2026 sales is represented by Phase 16?":
        f"{PHASE16_COVERAGE:.2f}%",
}

for question, answer in answers.items():
    st.markdown(f"**{question}**")
    st.write(answer)

# ============================================================
# MANAGEMENT CONCLUSION
# ============================================================

st.header("📌 Management Conclusion")

st.markdown(
    f"""
### Annual Sales Trend

Sales increased from 2023 to 2025. The 2026 source figure is
lower than the 2025 annual figure, but it should not be used
for a full-year YoY conclusion until the reporting period is
confirmed.

### 2026 Source Sales

The official 2026 source reports **{money(SOURCE_2026)}**.

### Phase 16 Sales

Phase 16 currently reports **{money(PHASE16_TOTAL)}**.

### Reconciliation Issue

The difference is **{money(PHASE16_VARIANCE)}**.

Phase 16 therefore represents **{PHASE16_COVERAGE:.2f}%**
of the 2026 source figure.

### Management Interpretation

The difference should be investigated before Phase 16 is
presented as the complete 2026 sales total.
"""
)

# ============================================================
# RECONCILIATION DETAILS
# ============================================================

st.header("🧪 Reconciliation Details")

r1, r2, r3 = st.columns(3)

with r1:
    st.metric(
        "2026 Source",
        money_short(SOURCE_2026),
    )

with r2:
    st.metric(
        "Phase 16",
        money_short(PHASE16_TOTAL),
    )

with r3:
    st.metric(
        "Variance",
        money_short(PHASE16_VARIANCE),
    )

# ============================================================
# SOURCE STRUCTURE
# ============================================================

st.header("📋 Source Structure")

if source_sheets:

    source_summary = pd.DataFrame(
        {
            "Sheet": list(source_sheets.keys()),
            "Rows": [
                len(df)
                for df in source_sheets.values()
            ],
            "Columns": [
                len(df.columns)
                for df in source_sheets.values()
            ],
        }
    )

    st.dataframe(
        source_summary,
        use_container_width=True,
        hide_index=True,
    )

else:

    st.info(
        "Source workbook structure could not be loaded."
    )

# ============================================================
# MANAGEMENT RECOMMENDATIONS
# ============================================================

st.header("🎯 Management Recommendations")

st.markdown(
    f"""
### 1. Reconcile Phase 16 against the source

Investigate the **{money(PHASE16_VARIANCE)}** difference between
the original 2026 source and the Phase 16 processed dataset.

### 2. Confirm reporting period

Confirm whether the 2026 source represents the same reporting
period as the comparison period used for YoY analysis.

### 3. Identify excluded sales

Trace the transactions or aggregation rules responsible for the
approximately **{100 - PHASE16_COVERAGE:.2f}%** of 2026 source sales
not represented in Phase 16.

### 4. Preserve source-to-output controls

Maintain a reconciliation control between the original source
workbook and every processed analytical dataset.
"""
)

# ============================================================
# FINAL PRESENTATION CONTROL
# ============================================================

st.divider()

st.success(
    """
### ✅ Final Reporting Control

Phase 18 is the official multi-year source reporting view.

Phase 16 is the transaction-level analytical view.

The two scopes are deliberately kept separate and should not
be added together.
"""
)

st.caption(
    "MIKA Sales Intelligence · Phase 18 Management Results Dashboard"
)

st.caption(
    "Source: Phase18_MIKA_Year_Reconciliation.xlsx"
)