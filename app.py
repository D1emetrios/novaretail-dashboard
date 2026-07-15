"""
NovaRetail Customer Intelligence Dashboard
==========================================
An interactive Streamlit application for exploring customer revenue, behavioral
segments, regional/channel performance, product & demographic opportunities, and
descriptive growth/retention warning signals.

Run locally with:
    streamlit run app.py

Author: NovaRetail Analytics
Note: All figures are computed live from the filtered data. The dataset is a
small representative sample; results are descriptive, not causal or predictive.
"""

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# CONSTANTS & CONFIGURATION
# ---------------------------------------------------------------------------

DATA_FILE = Path(__file__).parent / "NR_dataset.xlsx"
SHEET_NAME = "data"

# Logical (not alphabetical) ordering for behavioral segments.
SEGMENT_ORDER = ["Promising", "Growth", "Stable", "Decline", "Unclassified"]

# Consistent segment colors used across every chart in the app.
SEGMENT_COLORS = {
    "Promising": "#1f77b4",     # blue
    "Growth": "#2ca02c",        # green
    "Stable": "#b39700",        # gold
    "Decline": "#d62728",       # red
    "Unclassified": "#7f7f7f",  # neutral gray
}

# Explicit, transparent product-category standardization map.
# The ORIGINAL category is always preserved; this only feeds a new column.
# Keys are lower-cased raw category strings; values are broad standard groups.
PRODUCT_CATEGORY_MAP = {
    # Electronics
    "electronics": "Electronics",
    "gaming": "Electronics",
    # Books, Media & Office
    "books": "Books & Media",
    "books & magazines": "Books & Media",
    "office supplies": "Books & Media",
    # Clothing & Fashion
    "clothing": "Clothing & Fashion",
    "fashion": "Clothing & Fashion",
    "fashion & apparel": "Clothing & Fashion",
    "fashion accessories": "Clothing & Fashion",
    "children's clothing": "Clothing & Fashion",
    "sportswear": "Clothing & Fashion",
    # Groceries & Food
    "groceries": "Groceries & Food",
    "grocery": "Groceries & Food",
    "grocery items": "Groceries & Food",
    "food & beverages": "Groceries & Food",
    # Toys & Games
    "toys": "Toys & Games",
    "toys & games": "Toys & Games",
    # Home, Furniture & Garden
    "home appliances": "Home & Furniture",
    "home decor": "Home & Furniture",
    "home & garden": "Home & Furniture",
    "home improvement": "Home & Furniture",
    "furniture": "Home & Furniture",
    "furniture & decor": "Home & Furniture",
    "gardening tools": "Home & Furniture",
    # Sports & Outdoors
    "sporting goods": "Sports & Outdoors",
    "sports & outdoors": "Sports & Outdoors",
    "sports equipment": "Sports & Outdoors",
    "outdoor equipment": "Sports & Outdoors",
    # Health & Beauty
    "health & wellness": "Health & Beauty",
    "health supplements": "Health & Beauty",
    "beauty products": "Health & Beauty",
    "beauty & personal care": "Health & Beauty",
    "health & beauty": "Health & Beauty",
    "cosmetics": "Health & Beauty",
    # Automotive
    "automotive": "Automotive",
}

# Keyword fallback so any unforeseen raw label still lands in a sensible group.
KEYWORD_FALLBACK = [
    (["electronic", "gaming", "computer", "phone"], "Electronics"),
    (["book", "magazine", "office", "stationery"], "Books & Media"),
    (["cloth", "fashion", "apparel", "wear", "shoe"], "Clothing & Fashion"),
    (["grocery", "groceries", "food", "beverage"], "Groceries & Food"),
    (["toy", "game"], "Toys & Games"),
    (["home", "furniture", "decor", "garden", "appliance"], "Home & Furniture"),
    (["sport", "outdoor"], "Sports & Outdoors"),
    (["health", "beauty", "cosmetic", "wellness", "supplement", "personal care"],
     "Health & Beauty"),
    (["auto", "car", "vehicle"], "Automotive"),
]

METRIC_OPTIONS = {
    "Total Revenue": ("PurchaseAmount", "sum", "$"),
    "Unique Customers": ("CustomerID", "nunique", ""),
    "Transaction Count": ("TransactionID", "count", ""),
    "Average Purchase Amount": ("PurchaseAmount", "mean", "$"),
    "Average Satisfaction": ("CustomerSatisfaction", "mean", ""),
}


# ---------------------------------------------------------------------------
# FORMATTING HELPERS
# ---------------------------------------------------------------------------

def fmt_currency(value) -> str:
    """Format a number as currency: $1,234.56 (safe for NaN/None)."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "$0.00"
    return f"${value:,.2f}"


def fmt_pct(value, decimals: int = 1) -> str:
    """Format a percentage value as e.g. 42.0% (input is already a percent)."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "0.0%"
    return f"{value:.{decimals}f}%"


def fmt_number(value, decimals: int = 0) -> str:
    """Format a plain number with thousands separators."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "0"
    return f"{value:,.{decimals}f}"


def safe_divide(numerator, denominator):
    """Division that returns 0 instead of raising / producing inf."""
    if denominator in (0, None) or (
        isinstance(denominator, float) and np.isnan(denominator)
    ):
        return 0.0
    return numerator / denominator


# ---------------------------------------------------------------------------
# DATA STANDARDIZATION
# ---------------------------------------------------------------------------

def standardize_product_category(raw_value) -> str:
    """
    Transparent product-category standardization.

    1. Looks up an explicit mapping of known raw labels.
    2. Falls back to keyword matching for any unforeseen label.
    3. Returns the trimmed original (title-cased) if nothing matches, so no
       valid data is ever silently dropped.
    """
    if raw_value is None or (isinstance(raw_value, float) and np.isnan(raw_value)):
        return "Unknown"
    key = str(raw_value).strip().lower()
    if key in PRODUCT_CATEGORY_MAP:
        return PRODUCT_CATEGORY_MAP[key]
    for keywords, group in KEYWORD_FALLBACK:
        if any(kw in key for kw in keywords):
            return group
    # Preserve unknown-but-valid values rather than deleting them.
    return str(raw_value).strip().title()


# ---------------------------------------------------------------------------
# DATA LOADING (cached)
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Loading NovaRetail data...")
def load_data(path_str: str) -> pd.DataFrame:
    """
    Load and clean the NovaRetail workbook.

    - Reads the 'data' worksheet with openpyxl.
    - Coerces dates and numerics safely.
    - Fills the missing behavioral label with 'Unclassified'.
    - Adds a standardized product category while preserving the original.
    - Drops ONLY rows unusable for core analysis (missing both CustomerID and
      PurchaseAmount).
    """
    path = Path(path_str)
    df = pd.read_excel(path, sheet_name=SHEET_NAME, engine="openpyxl")

    # --- Type coercion (safe) ---
    df["TransactionDate"] = pd.to_datetime(df["TransactionDate"], errors="coerce")
    df["PurchaseAmount"] = pd.to_numeric(df["PurchaseAmount"], errors="coerce")
    df["CustomerSatisfaction"] = pd.to_numeric(
        df["CustomerSatisfaction"], errors="coerce"
    )

    # --- Behavioral segment: missing label -> 'Unclassified' ---
    df["label"] = df["label"].fillna("Unclassified").astype(str).str.strip()
    df.loc[df["label"].isin(["", "nan", "None"]), "label"] = "Unclassified"
    df = df.rename(columns={"label": "Segment"})

    # --- Trim string columns for consistent grouping ---
    for col in ["ProductCategory", "CustomerAgeGroup", "CustomerGender",
                "CustomerRegion", "RetailChannel"]:
        df[col] = df[col].astype(str).str.strip()

    # --- Product category standardization (original preserved) ---
    df["ProductCategoryStd"] = df["ProductCategory"].apply(
        standardize_product_category
    )

    # --- Drop only unusable rows (no customer AND no purchase amount) ---
    unusable = df["CustomerID"].isna() & df["PurchaseAmount"].isna()
    df = df.loc[~unusable].copy()

    # Fill remaining numeric gaps conservatively (kept, not deleted).
    df["PurchaseAmount"] = df["PurchaseAmount"].fillna(0.0)

    return df.reset_index(drop=True)


def mode_or_first(series: pd.Series):
    """
    Return the most frequent value in a series. On a tie, return the value that
    sorts first alphabetically (deterministic, documented tie-break).
    """
    s = series.dropna()
    if s.empty:
        return None
    counts = s.value_counts()
    top = counts[counts == counts.max()].index.tolist()
    return sorted(map(str, top))[0]


# ---------------------------------------------------------------------------
# AGGREGATIONS
# ---------------------------------------------------------------------------

def build_customer_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    One row per customer. Descriptive fields use the mode (alphabetical
    tie-break). Distinguishes transaction records (rows) from unique
    transaction IDs.
    """
    if df.empty:
        return pd.DataFrame(
            columns=[
                "CustomerID", "Segment", "TotalRevenue", "TransactionRecords",
                "UniqueTransactions", "AvgPurchase", "AvgSatisfaction",
                "MostRecentDate", "PrimaryRegion", "PrimaryChannel",
                "PrimaryCategory",
            ]
        )

    grp = df.groupby("CustomerID")
    out = pd.DataFrame({
        "Segment": grp["Segment"].apply(mode_or_first),
        "TotalRevenue": grp["PurchaseAmount"].sum(),
        "TransactionRecords": grp["TransactionID"].count(),
        "UniqueTransactions": grp["TransactionID"].nunique(),
        "AvgPurchase": grp["PurchaseAmount"].mean(),
        "AvgSatisfaction": grp["CustomerSatisfaction"].mean(),
        "MostRecentDate": grp["TransactionDate"].max(),
        "PrimaryRegion": grp["CustomerRegion"].apply(mode_or_first),
        "PrimaryChannel": grp["RetailChannel"].apply(mode_or_first),
        "PrimaryCategory": grp["ProductCategoryStd"].apply(mode_or_first),
    }).reset_index()
    return out.sort_values("TotalRevenue", ascending=False).reset_index(drop=True)


def build_segment_table(df: pd.DataFrame) -> pd.DataFrame:
    """Segment-level performance summary."""
    if df.empty:
        return pd.DataFrame(
            columns=["Segment", "Revenue", "UniqueCustomers", "Transactions",
                     "AvgPurchase", "AvgSatisfaction", "RevenuePerCustomer"]
        )
    grp = df.groupby("Segment")
    tbl = pd.DataFrame({
        "Revenue": grp["PurchaseAmount"].sum(),
        "UniqueCustomers": grp["CustomerID"].nunique(),
        "Transactions": grp["TransactionID"].count(),
        "AvgPurchase": grp["PurchaseAmount"].mean(),
        "AvgSatisfaction": grp["CustomerSatisfaction"].mean(),
    }).reset_index()
    tbl["RevenuePerCustomer"] = tbl.apply(
        lambda r: safe_divide(r["Revenue"], r["UniqueCustomers"]), axis=1
    )
    # Logical segment ordering.
    tbl["__order"] = tbl["Segment"].map(
        {s: i for i, s in enumerate(SEGMENT_ORDER)}
    ).fillna(99)
    return tbl.sort_values("__order").drop(columns="__order").reset_index(drop=True)


def compute_risk_scores(cust: pd.DataFrame, latest_date, median_revenue) -> pd.DataFrame:
    """
    Transparent, rules-based customer priority score (NOT a predictive model).

    Points:
      +3  Segment == Decline
      +1  Segment == Unclassified
      +2  AvgSatisfaction < 3
      +2  TotalRevenue below filtered median customer revenue
      +1  UniqueTransactions <= 1 (low frequency)
      +2  Most recent purchase >= 14 days before dataset's latest date
      +1  Most recent purchase 7-13 days before latest date

    Priority bands: High >= 5, Medium 3-4, Low <= 2.
    """
    cust = cust.copy()
    if cust.empty:
        cust["RiskScore"] = pd.Series(dtype=float)
        cust["Priority"] = pd.Series(dtype=object)
        return cust

    scores = np.zeros(len(cust), dtype=float)
    scores += np.where(cust["Segment"] == "Decline", 3, 0)
    scores += np.where(cust["Segment"] == "Unclassified", 1, 0)
    scores += np.where(cust["AvgSatisfaction"] < 3, 2, 0)
    scores += np.where(cust["TotalRevenue"] < median_revenue, 2, 0)
    scores += np.where(cust["UniqueTransactions"] <= 1, 1, 0)

    if pd.notna(latest_date):
        days_since = (latest_date - cust["MostRecentDate"]).dt.days
        scores += np.where(days_since >= 14, 2, 0)
        scores += np.where((days_since >= 7) & (days_since < 14), 1, 0)

    cust["RiskScore"] = scores
    cust["Priority"] = pd.cut(
        cust["RiskScore"],
        bins=[-np.inf, 2, 4, np.inf],
        labels=["Low", "Medium", "High"],
    ).astype(str)
    return cust


# ---------------------------------------------------------------------------
# CHART HELPERS
# ---------------------------------------------------------------------------

def segment_bar(data: pd.DataFrame, x: str, y: str, title: str, y_title: str,
                is_currency: bool = False):
    """Bar chart colored by consistent segment colors, ordered logically."""
    fig = px.bar(
        data, x=x, y=y, color=x,
        color_discrete_map=SEGMENT_COLORS,
        category_orders={x: SEGMENT_ORDER},
        title=title,
    )
    hover_val = "$%{y:,.2f}" if is_currency else "%{y:,.2f}"
    fig.update_traces(
        hovertemplate=f"%{{x}}<br>{y_title}: {hover_val}<extra></extra>"
    )
    fig.update_layout(
        xaxis_title="Behavioral Segment", yaxis_title=y_title,
        showlegend=False, margin=dict(t=60, b=40),
    )
    if is_currency:
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    return fig


def empty_note(message: str = "No records match the current filters."):
    """Consistent empty-state message."""
    st.info(message)


# ---------------------------------------------------------------------------
# PAGE CONFIG & DATA LOAD
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="NovaRetail Customer Intelligence",
    page_icon="🛍️",
    layout="wide",
    initial_sidebar_state="expanded",
)

if not DATA_FILE.exists():
    st.error(
        f"❌ Data file not found: `{DATA_FILE.name}`.\n\n"
        "Please make sure **NR_dataset.xlsx** is in the same folder as `app.py` "
        "(and committed to your GitHub repository)."
    )
    st.stop()

try:
    df_all = load_data(str(DATA_FILE))
except Exception as exc:  # noqa: BLE001
    st.error(f"❌ Could not load the Excel workbook: {exc}")
    st.stop()

if df_all.empty:
    st.error("The dataset loaded but contains no usable rows.")
    st.stop()


# ---------------------------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------------------------

st.sidebar.title("🛍️ NovaRetail")
st.sidebar.caption(
    "All KPIs and charts below respond live to the filters you set here."
)

# Reset: clears widget state so every filter returns to 'all values'.
if st.sidebar.button("🔄 Reset Filters", use_container_width=True):
    for k in list(st.session_state.keys()):
        if k.startswith("flt_"):
            del st.session_state[k]
    st.rerun()

min_date = df_all["TransactionDate"].min()
max_date = df_all["TransactionDate"].max()

st.sidebar.subheader("Transaction Date Range")
date_range = st.sidebar.date_input(
    "Date range",
    value=(min_date.date(), max_date.date()),
    min_value=min_date.date(),
    max_value=max_date.date(),
    key="flt_dates",
)


def sorted_unique(col, order=None):
    """Sorted unique values, optionally using a custom logical order."""
    vals = df_all[col].dropna().unique().tolist()
    if order:
        return [v for v in order if v in vals] + sorted(
            [v for v in vals if v not in order]
        )
    return sorted(vals)


seg_opts = sorted_unique("Segment", SEGMENT_ORDER)
region_opts = sorted_unique("CustomerRegion")
cat_opts = sorted_unique("ProductCategoryStd")
channel_opts = sorted_unique("RetailChannel")
age_opts = sorted_unique("CustomerAgeGroup")
gender_opts = sorted_unique("CustomerGender")
sat_opts = sorted([int(x) for x in df_all["CustomerSatisfaction"].dropna().unique()])

st.sidebar.subheader("Segments & Groups")
sel_segments = st.sidebar.multiselect("Behavioral segment", seg_opts,
                                      default=seg_opts, key="flt_seg")
sel_regions = st.sidebar.multiselect("Customer region", region_opts,
                                     default=region_opts, key="flt_region")
sel_categories = st.sidebar.multiselect("Product category (standardized)", cat_opts,
                                        default=cat_opts, key="flt_cat")
sel_channels = st.sidebar.multiselect("Retail channel", channel_opts,
                                      default=channel_opts, key="flt_channel")
sel_ages = st.sidebar.multiselect("Customer age group", age_opts,
                                  default=age_opts, key="flt_age")
sel_genders = st.sidebar.multiselect("Customer gender", gender_opts,
                                     default=gender_opts, key="flt_gender")
sel_sat = st.sidebar.multiselect("Customer satisfaction rating", sat_opts,
                                 default=sat_opts, key="flt_sat")

st.sidebar.divider()
st.sidebar.caption(
    "**Reset:** click *Reset Filters* above, or reopen each dropdown to "
    "re-select all values."
)


# ---------------------------------------------------------------------------
# APPLY FILTERS
# ---------------------------------------------------------------------------

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    """Apply every sidebar filter. Empty selections yield an empty frame."""
    mask = pd.Series(True, index=df.index)

    # Date range (handle single-date selection gracefully).
    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        start, end = date_range
    elif isinstance(date_range, (tuple, list)) and len(date_range) == 1:
        start = end = date_range[0]
    else:
        start = end = date_range
    start = pd.Timestamp(start)
    end = pd.Timestamp(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    mask &= df["TransactionDate"].between(start, end)

    mask &= df["Segment"].isin(sel_segments) if sel_segments else False
    mask &= df["CustomerRegion"].isin(sel_regions) if sel_regions else False
    mask &= df["ProductCategoryStd"].isin(sel_categories) if sel_categories else False
    mask &= df["RetailChannel"].isin(sel_channels) if sel_channels else False
    mask &= df["CustomerAgeGroup"].isin(sel_ages) if sel_ages else False
    mask &= df["CustomerGender"].isin(sel_genders) if sel_genders else False
    mask &= df["CustomerSatisfaction"].isin(sel_sat) if sel_sat else False

    return df.loc[mask].copy()


fdf = apply_filters(df_all)

# Pre-compute customer-level tables from the filtered data.
cust_tbl = build_customer_table(fdf)
latest_date_all = df_all["TransactionDate"].max()
median_cust_rev = cust_tbl["TotalRevenue"].median() if not cust_tbl.empty else 0.0
cust_tbl = compute_risk_scores(cust_tbl, latest_date_all, median_cust_rev)


# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

st.title("NovaRetail Customer Intelligence Dashboard")
st.caption(
    "Interactive view of revenue, behavioral segments, regions, channels, "
    "products, demographics, and descriptive growth/retention signals. "
    f"Showing **{len(fdf):,}** of **{len(df_all):,}** transaction records."
)

if fdf.empty:
    empty_note(
        "🚫 No records match the current filter selection. "
        "Use **Reset Filters** in the sidebar or widen your selections."
    )
    st.stop()


# ---------------------------------------------------------------------------
# TABS
# ---------------------------------------------------------------------------

tab_exec, tab_seg, tab_market, tab_growth, tab_data = st.tabs([
    "📊 Executive Overview",
    "👥 Customer Segments",
    "🌍 Markets & Products",
    "📈 Growth & Retention",
    "🗂️ Customer Data & Methodology",
])


# ===========================================================================
# TAB 1 — EXECUTIVE OVERVIEW
# ===========================================================================
with tab_exec:
    st.subheader("Executive Overview")
    st.caption("KPIs computed live from the filtered data.")

    total_revenue = fdf["PurchaseAmount"].sum()
    unique_customers = fdf["CustomerID"].nunique()
    unique_transactions = fdf["TransactionID"].nunique()
    transaction_records = len(fdf)
    avg_purchase = fdf["PurchaseAmount"].mean()
    avg_satisfaction = fdf["CustomerSatisfaction"].mean()

    gp_revenue = fdf.loc[fdf["Segment"].isin(["Growth", "Promising"]),
                         "PurchaseAmount"].sum()
    pct_gp_revenue = safe_divide(gp_revenue, total_revenue) * 100

    decline_customers = fdf.loc[fdf["Segment"] == "Decline", "CustomerID"].nunique()
    pct_decline_cust = safe_divide(decline_customers, unique_customers) * 100

    r1 = st.columns(4)
    r1[0].metric("Total Revenue", fmt_currency(total_revenue),
                 help="Transaction-based: sum of PurchaseAmount.")
    r1[1].metric("Unique Customers", fmt_number(unique_customers),
                 help="Customer-based: distinct CustomerIDs.")
    r1[2].metric("Unique Transactions", fmt_number(unique_transactions),
                 help=f"Distinct TransactionIDs ({transaction_records:,} rows total).")
    r1[3].metric("Avg Purchase Amount", fmt_currency(avg_purchase),
                 help="Transaction-based: mean PurchaseAmount per record.")

    r2 = st.columns(4)
    r2[0].metric("Avg Satisfaction", f"{avg_satisfaction:.2f} / 5",
                 help="Transaction-based mean of CustomerSatisfaction.")
    r2[1].metric("% Revenue: Growth + Promising", fmt_pct(pct_gp_revenue),
                 help="Revenue from Growth & Promising / total revenue.")
    r2[2].metric("% Customers in Decline", fmt_pct(pct_decline_cust),
                 help="Customer-based: distinct Decline customers / unique customers.")
    r2[3].metric("Transaction Records", fmt_number(transaction_records),
                 help="Raw row count (rows may repeat customers/transactions).")

    st.divider()
    c1, c2 = st.columns(2)

    with c1:
        seg_rev = fdf.groupby("Segment")["PurchaseAmount"].sum().reset_index()
        fig = segment_bar(seg_rev, "Segment", "PurchaseAmount",
                          "Revenue by Behavioral Segment", "Revenue (USD)",
                          is_currency=True)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        daily = (fdf.groupby(fdf["TransactionDate"].dt.date)["PurchaseAmount"]
                 .sum().reset_index())
        daily.columns = ["Date", "Revenue"]
        fig = px.line(daily, x="Date", y="Revenue", markers=True,
                      title="Daily Revenue Trend (filtered period)")
        fig.update_traces(hovertemplate="%{x}<br>Revenue: $%{y:,.2f}<extra></extra>")
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        fig.update_layout(yaxis_title="Revenue (USD)", margin=dict(t=60, b=40))
        st.plotly_chart(fig, use_container_width=True)

    st.caption(
        "ℹ️ Transaction-based metrics count rows; customer-based metrics count "
        "distinct CustomerIDs. Because IDs can repeat, these differ intentionally."
    )


# ===========================================================================
# TAB 2 — CUSTOMER SEGMENTS  (Revenue & Customer Value)
# ===========================================================================
with tab_seg:
    st.subheader("Revenue & Customer Segments")

    seg_tbl = build_segment_table(fdf)

    c1, c2 = st.columns(2)
    with c1:
        fig = segment_bar(seg_tbl, "Segment", "Revenue",
                          "Revenue by Segment", "Revenue (USD)", is_currency=True)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = segment_bar(seg_tbl, "Segment", "UniqueCustomers",
                          "Unique Customers by Segment", "Unique Customers")
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig = segment_bar(seg_tbl, "Segment", "AvgPurchase",
                          "Average Purchase Amount by Segment",
                          "Avg Purchase (USD)", is_currency=True)
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        fig = segment_bar(seg_tbl, "Segment", "AvgSatisfaction",
                          "Average Satisfaction by Segment",
                          "Avg Satisfaction (1-5)")
        fig.update_yaxes(range=[0, 5])
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Segment Performance Table")
    disp = seg_tbl.copy()
    disp["Revenue"] = disp["Revenue"].apply(fmt_currency)
    disp["AvgPurchase"] = disp["AvgPurchase"].apply(fmt_currency)
    disp["RevenuePerCustomer"] = disp["RevenuePerCustomer"].apply(fmt_currency)
    disp["AvgSatisfaction"] = disp["AvgSatisfaction"].round(2)
    disp.columns = ["Segment", "Revenue", "Unique Customers", "Transactions",
                    "Avg Purchase", "Avg Satisfaction", "Revenue / Customer"]
    st.dataframe(disp, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Customer Value Analysis")
    st.caption(
        "One row per customer. Descriptive fields (region, channel, category) "
        "use the **mode**; ties break alphabetically."
    )

    if cust_tbl.empty:
        empty_note()
    else:
        top10 = cust_tbl.nlargest(10, "TotalRevenue").copy()
        top10["CustLabel"] = "Cust " + top10["CustomerID"].astype(str)
        fig = px.bar(
            top10.sort_values("TotalRevenue"),
            x="TotalRevenue", y="CustLabel", orientation="h",
            color="Segment", color_discrete_map=SEGMENT_COLORS,
            category_orders={"Segment": SEGMENT_ORDER},
            title="Top 10 Customers by Total Revenue",
            hover_data={"TotalRevenue": ":$,.2f"},
        )
        fig.update_layout(xaxis_title="Total Revenue (USD)", yaxis_title="Customer",
                          margin=dict(t=60, b=40))
        fig.update_xaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Searchable Customer Table")
        search = st.text_input(
            "Search by Customer ID, segment, region, channel, or category",
            key="cust_search",
        )
        view = cust_tbl.copy()
        if search:
            s = search.strip().lower()
            row_txt = view.astype(str).apply(
                lambda r: " ".join(r.values).lower(), axis=1
            )
            view = view.loc[row_txt.str.contains(s, na=False)]

        show = view.copy()
        show["MostRecentDate"] = pd.to_datetime(show["MostRecentDate"]).dt.date
        show_fmt = show.copy()
        show_fmt["TotalRevenue"] = show_fmt["TotalRevenue"].apply(fmt_currency)
        show_fmt["AvgPurchase"] = show_fmt["AvgPurchase"].apply(fmt_currency)
        show_fmt["AvgSatisfaction"] = show_fmt["AvgSatisfaction"].round(2)
        show_fmt = show_fmt[[
            "CustomerID", "Segment", "TotalRevenue", "TransactionRecords",
            "UniqueTransactions", "AvgPurchase", "AvgSatisfaction",
            "MostRecentDate", "PrimaryRegion", "PrimaryChannel", "PrimaryCategory",
            "RiskScore", "Priority",
        ]]
        show_fmt.columns = [
            "Customer ID", "Segment", "Total Revenue", "Txn Records",
            "Unique Txns", "Avg Purchase", "Avg Satisfaction", "Most Recent",
            "Primary Region", "Primary Channel", "Primary Category",
            "Risk Score", "Priority",
        ]
        st.dataframe(show_fmt, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download customer-level table (CSV)",
            data=view.to_csv(index=False).encode("utf-8"),
            file_name="novaretail_customers_filtered.csv",
            mime="text/csv",
            use_container_width=True,
        )


# ===========================================================================
# TAB 3 — MARKETS & PRODUCTS  (Regional/Channel + Product/Demographic)
# ===========================================================================
with tab_market:
    st.subheader("Regional & Channel Performance")

    c1, c2 = st.columns(2)
    with c1:
        reg_rev = (fdf.groupby("CustomerRegion")["PurchaseAmount"].sum()
                   .reset_index().sort_values("PurchaseAmount", ascending=False))
        fig = px.bar(reg_rev, x="CustomerRegion", y="PurchaseAmount",
                     title="Revenue by Customer Region",
                     color="CustomerRegion")
        fig.update_layout(xaxis_title="Region", yaxis_title="Revenue (USD)",
                          showlegend=False, margin=dict(t=60, b=40))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        fig.update_traces(hovertemplate="%{x}<br>Revenue: $%{y:,.2f}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        ch_rev = (fdf.groupby("RetailChannel")["PurchaseAmount"].sum()
                  .reset_index())
        fig = px.bar(ch_rev, x="RetailChannel", y="PurchaseAmount",
                     title="Revenue by Retail Channel (Online vs Physical Store)",
                     color="RetailChannel",
                     color_discrete_map={"Online": "#1f77b4",
                                         "Physical Store": "#ff7f0e"})
        fig.update_layout(xaxis_title="Channel", yaxis_title="Revenue (USD)",
                          showlegend=False, margin=dict(t=60, b=40))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        fig.update_traces(hovertemplate="%{x}<br>Revenue: $%{y:,.2f}<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        # Segment distribution by region (% of regional revenue) — labeled clearly.
        seg_reg = (fdf.groupby(["CustomerRegion", "Segment"])["PurchaseAmount"]
                   .sum().reset_index())
        reg_tot = seg_reg.groupby("CustomerRegion")["PurchaseAmount"].transform("sum")
        seg_reg["Pct"] = seg_reg["PurchaseAmount"] / reg_tot.replace(0, np.nan) * 100
        fig = px.bar(seg_reg, x="CustomerRegion", y="Pct", color="Segment",
                     color_discrete_map=SEGMENT_COLORS,
                     category_orders={"Segment": SEGMENT_ORDER},
                     title="Segment Distribution by Region (% of regional revenue)")
        fig.update_layout(xaxis_title="Region", yaxis_title="Share of Revenue (%)",
                          margin=dict(t=60, b=40), barmode="stack")
        fig.update_traces(hovertemplate="%{x} — %{fullData.name}<br>"
                                        "Share: %{y:.1f}%<extra></extra>")
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        ch_sat = (fdf.groupby("RetailChannel")["CustomerSatisfaction"].mean()
                  .reset_index())
        fig = px.bar(ch_sat, x="RetailChannel", y="CustomerSatisfaction",
                     title="Average Satisfaction by Channel",
                     color="RetailChannel",
                     color_discrete_map={"Online": "#1f77b4",
                                         "Physical Store": "#ff7f0e"})
        fig.update_layout(xaxis_title="Channel", yaxis_title="Avg Satisfaction (1-5)",
                          showlegend=False, margin=dict(t=60, b=40))
        fig.update_yaxes(range=[0, 5])
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Segment Revenue across Retail Channels")
    seg_ch = (fdf.groupby(["RetailChannel", "Segment"])["PurchaseAmount"]
              .sum().reset_index())
    fig = px.bar(seg_ch, x="RetailChannel", y="PurchaseAmount", color="Segment",
                 barmode="group", color_discrete_map=SEGMENT_COLORS,
                 category_orders={"Segment": SEGMENT_ORDER},
                 title="Segment Revenue by Channel (USD)")
    fig.update_layout(xaxis_title="Channel", yaxis_title="Revenue (USD)",
                      margin=dict(t=60, b=40))
    fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("Product & Demographic Opportunities")

    metric_label = st.selectbox(
        "Choose a metric for the product-category chart",
        list(METRIC_OPTIONS.keys()), index=0, key="prod_metric",
    )
    col, agg, prefix = METRIC_OPTIONS[metric_label]
    cat_metric = (fdf.groupby("ProductCategoryStd")
                  .agg(Value=(col, agg)).reset_index()
                  .sort_values("Value", ascending=False))
    fig = px.bar(cat_metric, x="ProductCategoryStd", y="Value",
                 title=f"{metric_label} by Standardized Product Category",
                 color="Value", color_continuous_scale="Blues")
    fig.update_layout(xaxis_title="Product Category (standardized)",
                      yaxis_title=metric_label, margin=dict(t=60, b=40),
                      coloraxis_showscale=False)
    if prefix == "$":
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)

    c5, c6 = st.columns(2)
    with c5:
        cat_cust = (fdf.groupby("ProductCategoryStd")["CustomerID"].nunique()
                    .reset_index().sort_values("CustomerID", ascending=False))
        fig = px.bar(cat_cust, x="ProductCategoryStd", y="CustomerID",
                     title="Unique Customers by Product Category",
                     color="CustomerID", color_continuous_scale="Greens")
        fig.update_layout(xaxis_title="Product Category", yaxis_title="Unique Customers",
                          margin=dict(t=60, b=40), coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
    with c6:
        age_rev = (fdf.groupby("CustomerAgeGroup")["PurchaseAmount"].sum()
                   .reset_index().sort_values("CustomerAgeGroup"))
        fig = px.bar(age_rev, x="CustomerAgeGroup", y="PurchaseAmount",
                     title="Revenue by Age Group", color="CustomerAgeGroup")
        fig.update_layout(xaxis_title="Age Group", yaxis_title="Revenue (USD)",
                          showlegend=False, margin=dict(t=60, b=40))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

    c7, c8 = st.columns(2)
    with c7:
        gen_rev = fdf.groupby("CustomerGender")["PurchaseAmount"].sum().reset_index()
        fig = px.bar(gen_rev, x="CustomerGender", y="PurchaseAmount",
                     title="Revenue by Gender", color="CustomerGender",
                     color_discrete_map={"Male": "#1f77b4", "Female": "#e377c2"})
        fig.update_layout(xaxis_title="Gender", yaxis_title="Revenue (USD)",
                          showlegend=False, margin=dict(t=60, b=40))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)
    with c8:
        # Heatmap: revenue by region x standardized category.
        pivot = fdf.pivot_table(index="CustomerRegion",
                                columns="ProductCategoryStd",
                                values="PurchaseAmount", aggfunc="sum",
                                fill_value=0)
        if pivot.empty:
            empty_note("Not enough data for the region × category heatmap.")
        else:
            fig = px.imshow(
                pivot, text_auto=".0f", aspect="auto",
                color_continuous_scale="Blues",
                labels=dict(x="Product Category", y="Region", color="Revenue ($)"),
                title="Revenue Heatmap: Region × Product Category (USD)",
            )
            fig.update_layout(margin=dict(t=60, b=40))
            st.plotly_chart(fig, use_container_width=True)


# ===========================================================================
# TAB 4 — GROWTH & RETENTION
# ===========================================================================
with tab_growth:
    st.subheader("Growth & Retention Signals")
    st.warning(
        "⚠️ These are **descriptive warning signals / prioritization indicators**, "
        "not proof of churn. The dataset covers a short period and is a small "
        "representative sample."
    )

    with st.expander("📖 How the customer priority score works (transparent rules)"):
        st.markdown(
            """
The score is a **simple, rules-based** heuristic — **not** a predictive ML model:

| Condition | Points |
|---|---|
| Segment = **Decline** | +3 |
| Segment = **Unclassified** | +1 |
| Average satisfaction **< 3** | +2 |
| Total revenue **below filtered median** customer revenue | +2 |
| Low frequency (**<= 1 unique transaction**) | +1 |
| Most recent purchase **>= 14 days** before latest dataset date | +2 |
| Most recent purchase **7-13 days** before latest date | +1 |

**Priority bands:** High >= 5 · Medium 3-4 · Low <= 2.
The latest-date reference is the **full dataset's** newest transaction, so the
"recency" signal is stable regardless of the date filter.
            """
        )

    if cust_tbl.empty:
        empty_note()
    else:
        high = cust_tbl[cust_tbl["Priority"] == "High"]
        med = cust_tbl[cust_tbl["Priority"] == "Medium"]
        high_rev = high["TotalRevenue"].sum()

        k = st.columns(4)
        k[0].metric("High-Priority Customers", fmt_number(len(high)))
        k[1].metric("Medium-Priority Customers", fmt_number(len(med)))
        k[2].metric("Revenue at High Priority", fmt_currency(high_rev),
                    help="Total revenue tied to high-priority customers.")
        k[3].metric("% Customers High Priority",
                    fmt_pct(safe_divide(len(high), len(cust_tbl)) * 100))

        c1, c2 = st.columns(2)
        with c1:
            pr_counts = (cust_tbl["Priority"].value_counts()
                         .reindex(["High", "Medium", "Low"]).fillna(0)
                         .reset_index())
            pr_counts.columns = ["Priority", "Customers"]
            fig = px.bar(pr_counts, x="Priority", y="Customers",
                         color="Priority", title="Customers by Priority Band",
                         color_discrete_map={"High": "#d62728",
                                             "Medium": "#ff7f0e",
                                             "Low": "#2ca02c"})
            fig.update_layout(showlegend=False, margin=dict(t=60, b=40),
                              yaxis_title="Customers")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            # Revenue vs satisfaction by segment (dual axis).
            present = [s for s in SEGMENT_ORDER if s in fdf["Segment"].unique()]
            seg_rs = fdf.groupby("Segment").agg(
                Revenue=("PurchaseAmount", "sum"),
                AvgSatisfaction=("CustomerSatisfaction", "mean"),
            ).reindex(present).reset_index()
            fig = go.Figure()
            fig.add_bar(x=seg_rs["Segment"], y=seg_rs["Revenue"], name="Revenue (USD)",
                        marker_color=[SEGMENT_COLORS.get(s, "#888")
                                      for s in seg_rs["Segment"]],
                        yaxis="y1",
                        hovertemplate="%{x}<br>Revenue: $%{y:,.2f}<extra></extra>")
            fig.add_trace(go.Scatter(
                x=seg_rs["Segment"], y=seg_rs["AvgSatisfaction"],
                name="Avg Satisfaction", mode="lines+markers",
                line=dict(color="black", width=2), yaxis="y2",
                hovertemplate="%{x}<br>Satisfaction: %{y:.2f}<extra></extra>"))
            fig.update_layout(
                title="Revenue vs Average Satisfaction by Segment",
                yaxis=dict(title="Revenue (USD)", tickprefix="$", tickformat=",.0f"),
                yaxis2=dict(title="Avg Satisfaction (1-5)", overlaying="y",
                            side="right", range=[0, 5]),
                legend=dict(orientation="h", y=1.12), margin=dict(t=80, b=40))
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Risk & Priority Table")
        risk_view = cust_tbl.sort_values(
            ["RiskScore", "TotalRevenue"], ascending=[False, False]
        ).copy()
        risk_view["MostRecentDate"] = pd.to_datetime(
            risk_view["MostRecentDate"]).dt.date
        rv = risk_view.copy()
        rv["TotalRevenue"] = rv["TotalRevenue"].apply(fmt_currency)
        rv["AvgSatisfaction"] = rv["AvgSatisfaction"].round(2)
        rv = rv[["CustomerID", "Segment", "Priority", "RiskScore", "TotalRevenue",
                 "AvgSatisfaction", "UniqueTransactions", "MostRecentDate",
                 "PrimaryRegion", "PrimaryChannel"]]
        rv.columns = ["Customer ID", "Segment", "Priority", "Risk Score",
                      "Total Revenue", "Avg Satisfaction", "Unique Txns",
                      "Most Recent", "Region", "Channel"]
        st.dataframe(rv, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download prioritized customer list (CSV)",
            data=risk_view.to_csv(index=False).encode("utf-8"),
            file_name="novaretail_priority_customers.csv",
            mime="text/csv", use_container_width=True,
        )

        # -------------------------------------------------------------------
        # Actionable Recommendations (data-driven, not hard-coded)
        # -------------------------------------------------------------------
        st.divider()
        st.subheader("Recommended Actions")
        st.caption("Generated from the filtered data. Cautious language reflects "
                   "the small sample size.")

        recs = []
        cust_total_rev = cust_tbl["TotalRevenue"].sum()

        # 1. Retention for high-revenue Decline customers.
        decline = cust_tbl[cust_tbl["Segment"] == "Decline"]
        if not decline.empty:
            top_decline = decline.nlargest(1, "TotalRevenue").iloc[0]
            recs.append((
                "🔴 Retention outreach — Decline segment",
                f"{len(decline)} Decline customer(s); top one holds "
                f"{fmt_currency(top_decline['TotalRevenue'])} in revenue.",
                "Prioritize personal retention outreach to high-value Decline "
                "customers before spend erodes further."))

        # 2. Loyalty/cross-sell for satisfied Stable customers.
        stable = cust_tbl[(cust_tbl["Segment"] == "Stable") &
                          (cust_tbl["AvgSatisfaction"] >= 4)]
        if not stable.empty:
            recs.append((
                "🟡 Loyalty / cross-sell — satisfied Stable customers",
                f"{len(stable)} Stable customer(s) with satisfaction >= 4 "
                f"({fmt_currency(stable['TotalRevenue'].sum())} revenue).",
                "Enroll in loyalty programs and test cross-sell offers to lift "
                "share of wallet."))

        # 3. Upsell for Growth & Promising.
        gp = cust_tbl[cust_tbl["Segment"].isin(["Growth", "Promising"])]
        if not gp.empty:
            share = safe_divide(gp["TotalRevenue"].sum(), cust_total_rev) * 100
            recs.append((
                "🟢 Upsell — Growth & Promising customers",
                f"{len(gp)} customer(s) generating "
                f"{fmt_currency(gp['TotalRevenue'].sum())} "
                f"({fmt_pct(share)} of customer revenue).",
                "Introduce premium tiers / bundles to accelerate their upward "
                "trajectory."))

        # 4. Best region x category x channel combo.
        combo = (fdf.groupby(["CustomerRegion", "ProductCategoryStd", "RetailChannel"])
                 ["PurchaseAmount"].sum().reset_index())
        if not combo.empty:
            best = combo.nlargest(1, "PurchaseAmount").iloc[0]
            recs.append((
                "💡 Double-down — top region/category/channel",
                f"{best['CustomerRegion']} · {best['ProductCategoryStd']} · "
                f"{best['RetailChannel']} = {fmt_currency(best['PurchaseAmount'])}.",
                "Concentrate marketing spend on this proven high-performing "
                "combination."))

        # 5. Service recovery for lowest-satisfaction region.
        low_sat_region = fdf.groupby("CustomerRegion")["CustomerSatisfaction"].mean()
        if not low_sat_region.empty and low_sat_region.min() < 3.5:
            worst_region = low_sat_region.idxmin()
            recs.append((
                "🛠️ Service recovery — lowest-satisfaction region",
                f"{worst_region} averages {low_sat_region.min():.2f}/5 satisfaction.",
                "Launch a targeted service-recovery campaign and diagnose root "
                "causes in this region."))

        # 6. Channel migration opportunity.
        ch_avg = fdf.groupby("RetailChannel")["PurchaseAmount"].mean()
        if len(ch_avg) >= 2:
            best_ch, worst_ch = ch_avg.idxmax(), ch_avg.idxmin()
            gap = safe_divide(ch_avg.max() - ch_avg.min(), ch_avg.min()) * 100
            if gap > 15:
                recs.append((
                    "🔁 Test channel migration",
                    f"Avg purchase in {best_ch} exceeds {worst_ch} by {gap:.0f}%.",
                    f"Pilot nudging suitable {worst_ch} customers toward "
                    f"{best_ch} and measure lift."))

        if not recs:
            st.info("No specific recommendations triggered for this filter set.")
        else:
            for title, metric, action in recs:
                with st.container(border=True):
                    st.markdown(f"**{title}**")
                    st.markdown(f"- *Supporting metric:* {metric}")
                    st.markdown(f"- *Recommended action:* {action}")


# ===========================================================================
# TAB 5 — CUSTOMER DATA & METHODOLOGY
# ===========================================================================
with tab_data:
    st.subheader("Customer Data & Methodology")

    st.markdown("#### Filtered Transaction Table")
    tx_view = fdf[[
        "CustomerID", "TransactionID", "TransactionDate", "Segment",
        "ProductCategory", "ProductCategoryStd", "PurchaseAmount",
        "CustomerAgeGroup", "CustomerGender", "CustomerRegion",
        "CustomerSatisfaction", "RetailChannel",
    ]].copy()
    tx_disp = tx_view.copy()
    tx_disp["TransactionDate"] = pd.to_datetime(tx_disp["TransactionDate"]).dt.date
    tx_disp["PurchaseAmount"] = tx_disp["PurchaseAmount"].apply(fmt_currency)
    st.dataframe(tx_disp, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download filtered transactions (CSV)",
        data=tx_view.to_csv(index=False).encode("utf-8"),
        file_name="novaretail_transactions_filtered.csv",
        mime="text/csv", use_container_width=True,
    )

    st.markdown("#### Customer-Level Summary Table")
    if cust_tbl.empty:
        empty_note()
    else:
        cs = cust_tbl.copy()
        cs["MostRecentDate"] = pd.to_datetime(cs["MostRecentDate"]).dt.date
        st.dataframe(cs, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download customer summary (CSV)",
            data=cust_tbl.to_csv(index=False).encode("utf-8"),
            file_name="novaretail_customer_summary.csv",
            mime="text/csv", use_container_width=True,
        )

    st.divider()
    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("#### Data-Quality Notes")
        st.markdown(
            """
- **~100 transaction records**; CustomerIDs and TransactionIDs **repeat**, so
  row count != unique transactions != unique customers.
- **One missing segment label** is relabeled **Unclassified** (never dropped).
- **Product categories are inconsistent** (e.g. *Groceries/Grocery*,
  *Books/Books & Magazines*) — standardized into broad groups while the
  **original category is preserved** in `ProductCategory`.
- Only rows missing **both** CustomerID and PurchaseAmount are removed; no valid
  data is silently deleted.
- `PurchaseAmount` & `CustomerSatisfaction` are coerced to numeric safely;
  `TransactionDate` to datetime.
            """
        )

        st.markdown("#### Risk-Score Methodology")
        st.markdown(
            """
Transparent, rules-based priority score (see the *Growth & Retention* tab for
the full points table). It combines Decline/Unclassified membership, low
satisfaction, below-median revenue, low frequency, and recency. Bands:
**High >= 5, Medium 3-4, Low <= 2**. This is a descriptive prioritization aid —
**not** a predictive model.
            """
        )

    with col_b:
        st.markdown("#### Product-Category Standardization Mapping")
        map_df = pd.DataFrame(
            sorted(PRODUCT_CATEGORY_MAP.items()),
            columns=["Original (raw)", "Standardized Group"],
        )
        st.dataframe(map_df, use_container_width=True, hide_index=True, height=380)
        st.caption(
            "Unforeseen labels fall back to keyword matching, then to the "
            "title-cased original — so nothing is dropped."
        )

    st.divider()
    st.markdown("#### Data Source & Methodology Note")
    st.info(
        "**Source:** `NR_dataset.xlsx`, worksheet `data` (~100 rows). "
        "All KPIs and charts are computed live from the current filter selection. "
        "**Disclaimer:** findings are based on a small representative sample and "
        "are **descriptive, not causal or predictive**. Warning signals indicate "
        "where to look — they do not prove churn."
    )
