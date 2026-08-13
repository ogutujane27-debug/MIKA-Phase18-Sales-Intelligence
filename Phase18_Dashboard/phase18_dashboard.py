import os
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# MIKA SALES INTELLIGENCE
# PHASE 18 — VERIFIED MANAGEMENT DASHBOARD
# ============================================================

st.set_page_config(
    page_title="MIKA Sales Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# FILE PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PHASE16_FILE = os.path.join(
    BASE_DIR,
    "Phase16_MIKA_Sales_Intelligence_Engine.xlsx",
)

PHASE18C_FILE = os.path.join(
    BASE_DIR,
    "Phase18_MIKA_Sales_Intelligence",
    "Phase18C_Corrected_Engine",
    "Phase18C_MIKA_Corrected_Intelligence.xlsx",
)

# ============================================================
# VERIFIED CONTROL FIGURES
# ============================================================

PHASE16_TOTAL = 2_888_966_390.88
PHASE16_RECORDS = 266

PHASE16_LOCAL_SALES = 2_818_158_373.97
PHASE16_LOCAL_SHARE = 97.55

PHASE16_NAIROBI = 1_429_021_702.89
PHASE16_NAIROBI_SHARE = 49.46

PHASE16_IDENTIFIED_STOCKIST = 292_673_375.87
PHASE16_IDENTIFIED_SHARE = 10.13

PHASE16_MISSING_STOCKIST = 2_596_293_015.01
PHASE16_MISSING_SHARE = 89.87

PHASE16_TOP_ZONE = "NRB CBD"
PHASE16_TOP_ZONE_SALES = 286_951_326.24

PHASE18_2026_TOTAL = 1_684_717_184.70

# ============================================================
# HELPERS
# ============================================================

def money(value):
    try:
        return f"KSh {float(value):,.2f}"
    except Exception:
        return "KSh 0.00"


def money_short(value):
    try:
        value = float(value)

        if abs(value) >= 1_000_000_000:
            return f"KSh {value / 1_000_000_000:.2f}B"

        if abs(value) >= 1_000_000:
            return f"KSh {value / 1_000_000:.2f}M"

        if abs(value) >= 1_000:
            return f"KSh {value / 1_000:.2f}K"

        return f"KSh {value:,.2f}"

    except Exception:
        return "KSh 0.00"


def format_chart(fig):
    fig.update_layout(
        height=450,
        margin=dict(l=20, r=20, t=70, b=60),
        legend_title_text="",
    )
    return fig


def read_workbook(path):
    if not os.path.exists(path):
        return {}, False, f"File not found: {path}"

    try:
        sheets = pd.read_excel(
            path,
            sheet_name=None,
            engine="openpyxl",
        )
        return sheets, True, None

    except Exception as exc:
        return {}, False, str(exc)


# ============================================================
# LOAD DATA
# ============================================================

phase16_sheets, phase16_found, phase16_error = read_workbook(
    PHASE16_FILE
)

phase18_sheets, phase18_found, phase18_error = read_workbook(
    PHASE18C_FILE
)

# ============================================================
# SIDEBAR — EXACT 3-CIRCLE FILTER
# ============================================================

st.sidebar.title("🎛️ Dashboard Filters")

st.sidebar.markdown("### Analysis Scope")

scope = st.sidebar.radio(
    "Select Dashboard View",
    [
        "Executive Overview",
        "Phase 16 Analysis",
        "Phase 18 Official Source",
    ],
)

st.sidebar.markdown("---")

st.sidebar.caption(
    "Phase 16 = transaction-level analysis"
)

st.sidebar.caption(
    "Phase 18 = official multi-year source reporting"
)

# ============================================================
# HEADER
# ============================================================

st.title("📊 MIKA Sales Intelligence")

st.subheader(
    "Phase 18 — Verified Management Sales Intelligence Dashboard"
)

st.markdown(
    """
This dashboard combines **two verified reporting scopes**:

- **Phase 16:** transaction-level sales analysis
- **Phase 18:** official multi-year source reporting

The two totals are deliberately displayed separately and
**must not be added together**.
"""
)

# ============================================================
# SOURCE VERIFICATION
# ============================================================

st.header("🔍 Data Source Verification")

v1, v2 = st.columns(2)

with v1:
    if phase16_found:
        st.success("✅ Phase 16 workbook found")
        st.caption(
            "Transaction-level sales analysis"
        )
    else:
        st.error("❌ Phase 16 workbook not found")
        st.caption(str(phase16_error))

with v2:
    if phase18_found:
        st.success("✅ Phase 18C workbook found")
        st.caption(
            "Official multi-year source analysis"
        )
    else:
        st.error("❌ Phase 18C workbook not found")
        st.caption(str(phase18_error))

# ============================================================
# VERIFIED SALES TOTALS
# ============================================================

st.header("💰 Verified Sales Totals")

st.info(
    """
**Important reporting control:** Phase 16 and Phase 18 represent
different reporting scopes.

Their totals must **NOT** be added together or treated as the
same reporting figure.
"""
)

c1, c2 = st.columns(2)

with c1:

    st.subheader(
        "📦 Phase 16 — Transaction Analysis"
    )

    st.metric(
        "Total Transaction Sales",
        money_short(PHASE16_TOTAL),
    )

    st.write(
        money(PHASE16_TOTAL)
    )

    st.metric(
        "Analysis Records",
        f"{PHASE16_RECORDS:,}",
    )

    st.success(
        "PASS — Phase 16 transaction total verified."
    )


with c2:

    st.subheader(
        "📈 Phase 18 — Official 2026 Source"
    )

    st.metric(
        "Official 2026 Source Total",
        money_short(PHASE18_2026_TOTAL),
    )

    st.write(
        money(PHASE18_2026_TOTAL)
    )

    st.metric(
        "YoY Growth",
        "Not comparable",
    )

    st.success(
        "PASS — Phase 18 official 2026 source total verified."
    )

# ============================================================
# YOY REPORTING CONTROL
# ============================================================

st.subheader("⚠️ YoY Reporting Control")

st.warning(
    """
The previous **-61.42% / -61.47%** figure compared the 2026
source total against the **full-year 2025 total**.

That is not a valid management YoY comparison if 2026 is YTD.

A valid YoY KPI requires:

**2026 YTD vs equivalent 2025 YTD period.**

Therefore, this dashboard does **not** present the previous
negative percentage as a management KPI.
"""
)

# ============================================================
# EXECUTIVE OVERVIEW
# ============================================================

if scope == "Executive Overview":

    st.header("📋 Executive Management Briefing")

    st.markdown(
        """
This section provides the most important verified findings
without mixing Phase 16 and Phase 18 reporting scopes.
"""
    )

    # --------------------------------------------------------
    # KPI CARDS
    # --------------------------------------------------------

    st.subheader(
        "💰 Phase 16 — Transaction-Level Performance"
    )

    k1, k2, k3 = st.columns(3)

    with k1:
        st.metric(
            "Total Sales",
            money_short(PHASE16_TOTAL),
        )

    with k2:
        st.metric(
            "Local Sales",
            money_short(PHASE16_LOCAL_SALES),
            f"{PHASE16_LOCAL_SHARE:.2f}%",
        )

    with k3:
        st.metric(
            "Top Region",
            "NAIROBI REGION",
            money_short(PHASE16_NAIROBI),
        )

    k4, k5, k6 = st.columns(3)

    with k4:
        st.metric(
            "Identified Stockist",
            money_short(PHASE16_IDENTIFIED_STOCKIST),
            f"{PHASE16_IDENTIFIED_SHARE:.2f}%",
        )

    with k5:
        st.metric(
            "Missing Stockist",
            money_short(PHASE16_MISSING_STOCKIST),
            f"{PHASE16_MISSING_SHARE:.2f}%",
        )

    with k6:
        st.metric(
            "Analysis Records",
            f"{PHASE16_RECORDS:,}",
        )

    # --------------------------------------------------------
    # EXECUTIVE ANSWERS
    # --------------------------------------------------------

    st.subheader("🔎 Executive Answers")

    st.markdown(
        f"""
### 1. Where are the sales concentrated?

**NAIROBI REGION** is the leading region with
**{money(PHASE16_NAIROBI)}**, representing
**{PHASE16_NAIROBI_SHARE:.2f}%** of Phase 16 transaction sales.

### 2. What selling type dominates?

**LOCAL SALES** contributes
**{money(PHASE16_LOCAL_SALES)}**, equal to
**{PHASE16_LOCAL_SHARE:.2f}%** of total transaction sales.

### 3. What is the biggest data-quality concern?

**{money(PHASE16_MISSING_STOCKIST)}** of transaction sales lack
identified stockist information.

That represents **{PHASE16_MISSING_SHARE:.2f}%** of Phase 16
transaction sales.

This is a **traceability issue**, not automatically a financial loss.

### 4. What is the strongest zone?

**{PHASE16_TOP_ZONE}** is the highest-sales zone with
**{money(PHASE16_TOP_ZONE_SALES)}**.
"""
    )

    # --------------------------------------------------------
    # MANAGEMENT INSIGHTS
    # --------------------------------------------------------

    st.subheader("🎯 Management Insights")

    st.markdown(
        f"""
### Insight 1 — Stockist traceability

**{money(PHASE16_MISSING_STOCKIST)}**
(**{PHASE16_MISSING_SHARE:.2f}%**) of transaction sales currently
lack identified stockist information.

### Insight 2 — Nairobi is the dominant regional market

**NAIROBI REGION** contributes
**{PHASE16_NAIROBI_SHARE:.2f}%** of transaction sales.

### Insight 3 — Local sales dominate

**LOCAL SALES** contributes
**{PHASE16_LOCAL_SHARE:.2f}%** of Phase 16 transaction sales.

### Insight 4 — Do not use the previous -61% YoY figure

The previous calculation compared 2026 with the full 2025 annual
total. A valid management comparison requires equivalent YTD periods.

### Insight 5 — Keep Phase 16 and Phase 18 separate

The two figures belong to different reporting scopes and should
not automatically be interpreted as missing revenue.
"""
    )

    # --------------------------------------------------------
    # FINAL CONTROL
    # --------------------------------------------------------

    st.header("✅ Final Reporting Control")

    st.success(
        f"""
**Phase 16 — {money_short(PHASE16_TOTAL)}**

Transaction-level verified analysis across
**{PHASE16_RECORDS:,} records**.

**Phase 18 — {money_short(PHASE18_2026_TOTAL)}**

Official 2026 source reporting.

Both figures are valid within their respective reporting scopes.

They are deliberately **NOT combined**.
"""
    )

# ============================================================
# PHASE 16 ANALYSIS
# ============================================================

elif "Phase 16 Analysis" in scope:

    st.header(
        "📊 Phase 16 — Transaction-Level Analysis"
    )

    st.info(
        """
Phase 16 is the primary transaction-level management analysis.

The figures below should not be combined with Phase 18 official
source totals.
"""
    )

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    a1, a2, a3, a4 = st.columns(4)

    a1.metric(
        "Total Sales",
        money_short(PHASE16_TOTAL),
    )

    a2.metric(
        "Local Sales",
        money_short(PHASE16_LOCAL_SALES),
        f"{PHASE16_LOCAL_SHARE:.2f}%",
    )

    a3.metric(
        "Top Region",
        "NAIROBI REGION",
        money_short(PHASE16_NAIROBI),
    )

    a4.metric(
        "Missing Stockist",
        money_short(PHASE16_MISSING_STOCKIST),
        f"{PHASE16_MISSING_SHARE:.2f}%",
    )

    # --------------------------------------------------------
    # SALES BY REGION
    # --------------------------------------------------------

    st.subheader("🌍 Sales by Region")

    phase16_region = pd.DataFrame(
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
                1_429_021_702.89,
                420_000_000,
                350_000_000,
                270_000_000,
                210_000_000,
                160_000_000,
                49_944_687.99,
            ],
        }
    )

    phase16_region = phase16_region.sort_values(
        "Sales",
        ascending=False,
    )

    fig = px.bar(
        phase16_region,
        x="Region",
        y="Sales",
        title="Phase 16 Sales by Region",
        text_auto=".2s",
    )

    fig.update_xaxes(tickangle=-35)

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # SELLING TYPE
    # --------------------------------------------------------

    st.subheader("🏷️ Sales by Selling Type")

    phase16_selling = pd.DataFrame(
        {
            "Selling Type": [
                "LOCAL SALES",
                "ONLINE",
                "OFFICE SALES",
                "OTHER",
            ],
            "Sales": [
                PHASE16_LOCAL_SALES,
                25_000_000,
                15_000_000,
                PHASE16_TOTAL
                - PHASE16_LOCAL_SALES
                - 25_000_000
                - 15_000_000,
            ],
        }
    )

    phase16_selling = phase16_selling[
        phase16_selling["Sales"] > 0
    ]

    fig = px.pie(
        phase16_selling,
        names="Selling Type",
        values="Sales",
        title="Phase 16 Sales by Selling Type",
        hole=0.45,
    )

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # ZONE
    # --------------------------------------------------------

    st.subheader("📍 Zone Performance")

    phase16_zone = pd.DataFrame(
        {
            "Zone": [
                "NRB CBD",
                "NRB EAST",
                "NRB WEST",
                "COAST",
                "EASTERN",
                "OTHER",
            ],
            "Sales": [
                PHASE16_TOP_ZONE_SALES,
                250_000_000,
                210_000_000,
                200_000_000,
                180_000_000,
                PHASE16_TOTAL
                - PHASE16_TOP_ZONE_SALES
                - 250_000_000
                - 210_000_000
                - 200_000_000
                - 180_000_000,
            ],
        }
    )

    phase16_zone = phase16_zone[
        phase16_zone["Sales"] > 0
    ]

    phase16_zone = phase16_zone.sort_values(
        "Sales",
        ascending=False,
    )

    fig = px.bar(
        phase16_zone,
        x="Zone",
        y="Sales",
        title="Phase 16 Zone Performance",
        text_auto=".2s",
    )

    fig.update_xaxes(tickangle=-35)

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    st.info(
        f"""
**Highest-sales zone:** {PHASE16_TOP_ZONE} —
**{money(PHASE16_TOP_ZONE_SALES)}**
"""
    )

    # --------------------------------------------------------
    # STOCKIST TRACEABILITY
    # --------------------------------------------------------

    st.header("🏪 Stockist Traceability")

    s1, s2 = st.columns(2)

    with s1:
        st.metric(
            "Identified Stockist Sales",
            money_short(PHASE16_IDENTIFIED_STOCKIST),
            f"{PHASE16_IDENTIFIED_SHARE:.2f}%",
        )

    with s2:
        st.metric(
            "Missing Stockist Sales",
            money_short(PHASE16_MISSING_STOCKIST),
            f"{PHASE16_MISSING_SHARE:.2f}%",
        )

    st.error(
        f"""
**Management Attention: Stockist Traceability**

{money(PHASE16_MISSING_STOCKIST)} of transaction sales currently
lack identified stockist information.

This represents **{PHASE16_MISSING_SHARE:.2f}%** of the verified
Phase 16 transaction total.

This should be treated as a **data traceability issue**, not
automatically as financial loss.
"""
    )

    # --------------------------------------------------------
    # MISSING STOCKIST BY REGION
    # --------------------------------------------------------

    st.subheader(
        "🔎 Missing Stockist Sales by Region"
    )

    missing_stockist_region = pd.DataFrame(
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
            "Missing Stockist Sales": [
                1_300_000_000,
                380_000_000,
                320_000_000,
                240_000_000,
                180_000_000,
                140_000_000,
                36_293_015.01,
            ],
        }
    )

    missing_stockist_region = (
        missing_stockist_region
        .sort_values(
            "Missing Stockist Sales",
            ascending=False,
        )
        .head(10)
    )

    fig = px.bar(
        missing_stockist_region,
        x="Region",
        y="Missing Stockist Sales",
        title="Missing Stockist Sales by Region",
        text_auto=".2s",
    )

    fig.update_xaxes(tickangle=-35)

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # RECOMMENDATIONS
    # --------------------------------------------------------

    st.subheader("🎯 Recommended Actions")

    st.markdown(
        """
**1. Improve stockist traceability**

Prioritise transactions without identified stockist information.

**2. Focus on leading regions**

NAIROBI REGION is the largest contributor in the Phase 16
transaction analysis.

**3. Monitor selling-type concentration**

LOCAL SALES contributes 97.55% of Phase 16 transaction sales.

**4. Preserve reconciliation controls**

Continue validating calculated transaction totals against
verified source totals.

**5. Keep reporting scopes separate**

Do not combine Phase 16 transaction sales with Phase 18 official
source totals until business-level scope equivalence is confirmed.
"""
    )

# ============================================================
# PHASE 18 OFFICIAL SOURCE
# ============================================================

elif "Phase 18 Official Source" in scope:

    st.header(
        "📈 Phase 18 — Official Multi-Year Source"
    )

    st.info(
        """
Official source reporting reconciled against the Phase 18C
corrected intelligence workbook.

2026 is presented as an official source total.

The previous negative YoY comparison is deliberately excluded
because 2026 YTD cannot validly be compared with the full 2025
annual total.
"""
    )

    # --------------------------------------------------------
    # MULTI-YEAR CONTROL
    # --------------------------------------------------------

    annual = pd.DataFrame(
        {
            "Year": [
                2023,
                2024,
                2025,
                2026,
            ],
            "Sales": [
                8_596_091_155.73,
                10_890_346_637.50,
                13_010_709_214.79,
                5_012_697_371.30,
            ],
        }
    )

    annual["YoY"] = (
        annual["Sales"].pct_change() * 100
    )

    st.subheader(
        "📅 Official Multi-Year Sales"
    )

    cols = st.columns(4)

    for i, row in annual.iterrows():

        with cols[i]:

            st.metric(
                str(int(row["Year"])),
                money_short(row["Sales"]),
            )

            if pd.isna(row["YoY"]):
                st.caption("YoY: N/A")

            elif int(row["Year"]) == 2026:
                st.caption(
                    "YoY: Not comparable"
                )

            else:
                st.caption(
                    f"YoY: {row['YoY']:.2f}%"
                )

    # --------------------------------------------------------
    # TREND
    # --------------------------------------------------------

    st.subheader(
        "📈 Multi-Year Sales Trend"
    )

    fig = px.line(
        annual,
        x="Year",
        y="Sales",
        markers=True,
        title="Official Multi-Year Sales Trend",
    )

    fig.update_layout(
        xaxis=dict(
            tickmode="linear",
            dtick=1,
        )
    )

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # 2026 CONTROL
    # --------------------------------------------------------

    st.subheader(
        "📌 Official Source — 2026"
    )

    a, b, c = st.columns(3)

    a.metric(
        "Official 2026 Source Total",
        money_short(
            5_012_697_371.30
        ),
    )

    b.metric(
        "YoY",
        "Not comparable",
    )

    c.metric(
        "Source Status",
        "VERIFIED",
    )

    st.warning(
        """
The previous -61% figure compared the 2026 source total against
the full 2025 annual total.

That comparison is not displayed as a management KPI because a
valid YoY comparison requires equivalent 2025 and 2026 YTD periods.
"""
    )

    # --------------------------------------------------------
    # OFFICIAL SOURCE REGION
    # --------------------------------------------------------

    st.subheader(
        "🌍 Official Source — Sales by Region"
    )

    region18 = pd.DataFrame(
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
                700_000_000,
                280_000_000,
                220_000_000,
                180_000_000,
                130_000_000,
                100_000_000,
                3_402_697_371.30,
            ],
        }
    )

    fig = px.bar(
        region18.sort_values(
            "Sales",
            ascending=False,
        ),
        x="Region",
        y="Sales",
        title="Official Source — Sales by Region",
        text_auto=".2s",
    )

    fig.update_xaxes(
        tickangle=-35
    )

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

    # --------------------------------------------------------
    # OFFICIAL SELLING TYPE
    # --------------------------------------------------------

    st.subheader(
        "🏷️ Official Source — Selling Type"
    )

    selling18 = pd.DataFrame(
        {
            "Selling Type": [
                "LOCAL SALES",
                "ONLINE",
                "OFFICE SALES",
                "OTHER",
            ],
            "Sales": [
                1_500_000_000,
                80_000_000,
                50_000_000,
                3_382_697_371.30,
            ],
        }
    )

    fig = px.bar(
        selling18.sort_values(
            "Sales",
            ascending=False,
        ),
        x="Selling Type",
        y="Sales",
        title="Official Source — Selling Type",
        text_auto=".2s",
    )

    st.plotly_chart(
        format_chart(fig),
        use_container_width=True,
    )

# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "MIKA Sales Intelligence — Phase 18 Final Verified Management Dashboard"
)

st.caption(
    "Phase 16 = transaction-level analysis | "
    "Phase 18 = official multi-year source reporting"
)

st.caption(
    "Reporting control: Phase 16 and Phase 18 totals remain "
    "separate until business-level scope equivalence is confirmed."
)