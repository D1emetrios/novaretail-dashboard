# NovaRetail — Key Insights & Contribution Note

**Data source:** `NR_dataset.xlsx`, worksheet `data` (100 transaction records,
34 unique customers, 25 unique transaction IDs, Jan 15 – Feb 20 2023).
**Disclaimer:** This is a small, representative sample. All findings are
**descriptive, not causal or predictive** — they indicate where to look, not
proof of future behavior.

---

## Three Key Insights

### 1. A small set of high-potential customers drives the majority of revenue
The **Growth (39.1%)** and **Promising (31.8%)** segments together generate
**70.9% of total revenue** — $12,024 of $16,969 — while representing a minority
of the customer base. Their average satisfaction is also the highest in the
dataset (Growth 4.21/5, Promising 4.45/5).

**Business value:** These segments are the strongest candidates for upsell,
premium tiers, and loyalty investment. Protecting and expanding this group
should be the first priority for commercial spend, because revenue is
concentrated rather than evenly spread.

### 2. The Decline segment is a high-value retention risk, not a lost cause
Decline customers hold **$3,649.81 (21.5% of revenue)** across **19 unique
customers**, yet their average satisfaction is just **1.58/5** — by far the
lowest of any segment. This is a meaningful amount of revenue sitting on the
weakest satisfaction signal.

**Business value:** This combination (real revenue + very low satisfaction)
marks Decline as the top **service-recovery and retention** priority. Targeted
outreach to the highest-revenue Decline customers is where retention effort is
most likely to protect existing dollars. *(Caveat: the short time window means
this is a warning signal, not confirmed churn.)*

### 3. Physical Store drives larger baskets, led by Electronics in the West
Average purchase in **Physical Store is $190.31 vs $149.07 Online** — roughly
**28% higher** per transaction. The single strongest revenue pocket is
**West · Electronics · Physical Store at $2,609.95**, and Electronics leads the
top three region/channel combinations overall. By region, **West ($5,033)** and
**North ($4,991)** lead, ahead of South ($3,564) and East ($3,382).

**Business value:** There is a clear case to concentrate marketing and inventory
in high-performing Electronics + Physical Store pockets (starting with the West
and North regions), and to test nudging suitable Online shoppers toward the
higher-basket Physical Store channel.

---

## Contribution Note

I built an interactive customer-intelligence dashboard for NovaRetail using
Python, Streamlit, pandas, and Plotly, designed for deployment via GitHub and
Streamlit Community Cloud. I loaded and cleaned the ~100-record dataset — safely
coercing dates and numerics, relabeling the one missing behavioral segment as
*Unclassified*, and standardizing 35 inconsistent product-category labels into 9
broad groups while preserving the original category. The app is organized into
five tabs (Executive Overview, Customer Segments, Markets & Products, Growth &
Retention, and Customer Data & Methodology) with filter-responsive KPIs,
segment and customer-level analysis, regional/channel/product/demographic
charts, a transparent rules-based customer-priority score, and data-driven
recommended actions. Throughout, I distinguished transaction records from unique
transactions and unique customers, handled empty-filter and single-category
edge cases, and framed all warning signals as descriptive rather than
predictive given the small sample.

**Public app URL:** _[paste your Streamlit Community Cloud URL here]_
**GitHub repository:** _[paste your repository URL here]_
