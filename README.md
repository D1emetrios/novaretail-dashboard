# NovaRetail Customer Intelligence Dashboard

An interactive Streamlit dashboard that helps NovaRetail leadership understand
customer revenue, behavioral segments, regional/channel performance, product and
demographic opportunities, and descriptive growth/retention warning signals.

## Business Objective

NovaRetail is a nationwide omnichannel retailer. This dashboard helps leadership
answer four questions:

1. Which customers and segments generate the most revenue?
2. Which segments show warning signs of decline or reduced engagement?
3. Which regions, product categories, demographic groups, and channels offer the
   best growth and retention opportunities?
4. Where should NovaRetail prioritize commercial and marketing investment?

## Dashboard Features

- **Executive Overview** — KPI cards (total revenue, unique customers, unique
  transactions, average purchase, average satisfaction, % revenue from Growth +
  Promising, % customers in Decline), plus revenue-by-segment and daily revenue
  trend. Each KPI is labeled transaction-based vs customer-based.
- **Customer Segments** — revenue, unique customers, average purchase, and
  average satisfaction by segment; a full segment performance table; a
  customer-level value analysis with a Top-10 chart, a searchable table, and a
  CSV download.
- **Markets & Products** — revenue by region and channel, segment distribution by
  region, average satisfaction by channel, segment revenue across channels,
  revenue/customers by product category (with a switchable metric), revenue by
  age and gender, and a region x category revenue heatmap.
- **Growth & Retention** — a transparent, rules-based customer priority score,
  high-priority KPIs, a priority table, a revenue-vs-satisfaction chart, a CSV
  download, and **data-driven recommended actions**.
- **Customer Data & Methodology** — filtered transaction table, customer-level
  summary, data-quality notes, risk-score methodology, the product-category
  standardization mapping, and CSV downloads.

All KPIs and charts respond live to the sidebar filters (date range, segment,
region, standardized product category, channel, age group, gender, and
satisfaction). A **Reset Filters** button restores all defaults, and the app
shows a clear message when filters return no records.

## Repository File Structure

```
novaretail-dashboard/
├── app.py             # Streamlit application (all logic)
├── requirements.txt   # Python dependencies
├── README.md          # This file
└── NR_dataset.xlsx    # Source data (worksheet: "data")
```

## Local Installation

Requires Python 3.9+.

```bash
# 1. Clone your repository
git clone https://github.com/<your-username>/novaretail-dashboard.git
cd novaretail-dashboard

# 2. (Optional) create a virtual environment
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

## Local Launch

```bash
streamlit run app.py
```

The app opens at http://localhost:8501.

## Upload to GitHub

```bash
git init
git add app.py requirements.txt README.md NR_dataset.xlsx
git commit -m "NovaRetail customer intelligence dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/novaretail-dashboard.git
git push -u origin main
```

Make sure the repository is **public** so Streamlit Community Cloud can access it.

## Deploy to Streamlit Community Cloud

1. Sign in at https://share.streamlit.io with your GitHub account.
2. Click **Create app** → **Deploy a public app from GitHub**.
3. Select your repository, the `main` branch, and set the main file to `app.py`.
4. Click **Deploy**. The first build installs `requirements.txt` and launches the
   app.
5. Copy the public URL Streamlit assigns and paste it below and into your
   Brightspace submission.

**Public app URL:** `TODO — paste your Streamlit Community Cloud URL here`

## Customer-Priority Methodology

The priority score is a **transparent, rules-based heuristic — not a predictive
machine-learning model.** Points accumulate per customer:

| Condition | Points |
|---|---|
| Segment = Decline | +3 |
| Segment = Unclassified | +1 |
| Average satisfaction < 3 | +2 |
| Total revenue below the filtered median customer revenue | +2 |
| Low frequency (<= 1 unique transaction) | +1 |
| Most recent purchase >= 14 days before the dataset's latest date | +2 |
| Most recent purchase 7-13 days before the latest date | +1 |

Bands: **High >= 5, Medium 3-4, Low <= 2.** Recency is measured against the full
dataset's latest transaction date so it stays stable across date filters.

## Data Notes & Disclaimer

- The workbook holds ~100 transaction records. CustomerIDs and TransactionIDs
  repeat, so row count, unique transactions, and unique customers are all
  different and reported separately.
- One record with a missing segment label is relabeled **Unclassified** (never
  deleted). Inconsistent product categories (e.g. *Groceries/Grocery*,
  *Books/Books & Magazines*) are standardized into broad groups while the
  original category is preserved.

**Disclaimer:** Findings are based on a small, representative sample covering a
short period. All indicators are **descriptive, not causal or predictive**.
Warning signals show where to investigate — they do not prove customer churn.
