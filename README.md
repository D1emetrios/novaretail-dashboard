# Yelp Fusion API — Indianapolis Restaurants

A small Python project that pulls restaurant data for Indianapolis from the
**official Yelp Fusion API** and writes the results to CSV. Review excerpts
are supported as an optional step — see the note on plan access below.

## What it collects

For every restaurant the Fusion `/businesses/search` endpoint returns:

- `id`, `name`, `rating`, `review_count`, `url`
- `address`, `city`, `state`, `zip_code`, `neighborhood`
- `latitude`, `longitude`
- `phone`, `price`, `categories` (semicolon-joined), `is_closed`

If `--reviews` is passed **and** the API key's plan grants access, the
`/businesses/{id}/reviews` endpoint is hit for each business to pull up to
3 review excerpts per restaurant (reviewer name, rating, text, timestamp).

## Project layout

```
yelp_api_project/
├── README.md
├── requirements.txt
├── .env.example
├── .gitignore
├── sample_reviews.csv          # placeholder rows for schema reference
└── yelp_scraper/
    ├── __init__.py
    ├── client.py               # Yelp Fusion API wrapper + pagination
    ├── exporter.py             # payload → flat CSV row shaping
    └── main.py                 # CLI entry point
```

Outputs are written to `./output/` by default:

- `output/indianapolis_restaurants.csv`
- `output/indianapolis_reviews.csv` (only if `--reviews` is used and access is granted)

## Setup

### 1. Get a Yelp API key

1. Create an account and app at <https://docs.developer.yelp.com/>.
2. Copy the private API key Yelp generates for your app.

### 2. Install

```bash
# From the project root
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure your key

```bash
cp .env.example .env
# then edit .env and paste your key into YELP_API_KEY=
```

Or export it directly in your shell:

```bash
export YELP_API_KEY="your_key_here"   # Windows CMD: set YELP_API_KEY=your_key_here
```

## Run

Businesses only (default):

```bash
python -m yelp_scraper.main
```

Businesses **and** review excerpts:

```bash
python -m yelp_scraper.main --reviews
```

Other useful flags:

```bash
python -m yelp_scraper.main \
    --location "Indianapolis, IN" \
    --term restaurants \
    --max-results 1000 \
    --output-dir output \
    --reviews \
    --verbose
```

When it finishes you'll have:

```
output/indianapolis_restaurants.csv
output/indianapolis_reviews.csv     # if --reviews succeeded
```

## Pagination & result limits

The Fusion `/businesses/search` endpoint has two hard limits baked in by
Yelp itself:

- `limit` per request: **50** max
- `offset + limit`: **1000** max

This project paginates `limit=50` at a time up to that 1000-result ceiling,
stopping earlier if Yelp's reported `total` is reached. For Indianapolis
restaurants you will typically hit the 1000 cap.

## Review access — plan-dependent

**The `/businesses/{id}/reviews` endpoint is not available on every Yelp API
plan.** If your plan doesn't include it, the API returns `403 Forbidden`
when you call the reviews endpoint.

This project handles that case cleanly:

- When `--reviews` is passed, the first 403 from the reviews endpoint flips
  review collection off for the rest of the run and logs a warning. The
  business CSV still writes normally and the command exits 0.
- `sample_reviews.csv` in the project root shows the exact schema the
  `indianapolis_reviews.csv` file would use, with clearly-marked
  `[PLACEHOLDER]` rows. Use it as a reference for grading/demo purposes
  when your plan doesn't grant review access — **do not** submit it as
  real scraped data.

If you later upgrade to a plan with review access, re-run with
`--reviews` and the real `output/indianapolis_reviews.csv` will populate.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `ERROR: YELP_API_KEY not set` | `.env` missing or key not exported |
| `403 Forbidden` on businesses | Key invalid, revoked, or out of quota |
| `403 Forbidden` on reviews | Plan does not include review access — expected, run without `--reviews` |
| Stops at ~1000 results | Yelp's hard cap on `offset + limit` |
| `429` retries in logs | Rate-limited; the client backs off and retries automatically |

## Zipping for submission

From the parent directory of `yelp_api_project`:

```bash
# macOS/Linux — excludes venv, caches, and any .env you created
zip -r yelp_api_project.zip yelp_api_project \
    -x "yelp_api_project/.venv/*" \
       "yelp_api_project/**/__pycache__/*" \
       "yelp_api_project/.env" \
       "yelp_api_project/output/*"
```

On Windows (PowerShell):

```powershell
Compress-Archive -Path yelp_api_project -DestinationPath yelp_api_project.zip
```

Then double-check the archive doesn't contain your `.env`:

```bash
unzip -l yelp_api_project.zip | grep -i env
# should show .env.example but NOT .env
```
