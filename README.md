# Service 2 — CSV / Excel Data Pipeline Automation

## The pitch
> "Your managers shouldn't be human calculators. I build custom
> Python-Pandas scripts that turn your chaotic raw data dumps into
> executive-ready dashboards instantly — saving your team 40+ hours
> every month."

## Who buys this
- E-commerce founders (messy Shopify/marketplace exports)
- Operations managers (multi-source spreadsheet consolidation)
- Small finance/ops teams without a dedicated analyst

## What's in this folder
| File | Purpose |
|---|---|
| `pipeline.py` | The full ingest → clean → validate → analyze → report pipeline |
| `generate_sample_data.py` | Creates a realistic messy demo CSV (run this first) |
| `sample_data/messy_sales_export.csv` | Example raw client export (typos, missing values, currency symbols, mixed date formats, duplicates, an outlier) |
| `executive_report.xlsx` | The polished output the pipeline produces |

## Try it
```bash
pip install pandas numpy openpyxl
python generate_sample_data.py    # builds a messy demo CSV
python pipeline.py                # produces executive_report.xlsx
```

## What clients are actually paying for
1. **A visible data-quality report** — every fix (missing values filled,
   duplicates removed, unparseable values) is logged and shown on its own
   sheet. Clients trust automation more when they can *see* what it did.
2. **Auto-recomputed fields** — e.g. revenue is recalculated from
   quantity × price whenever it's missing or inconsistent, so the client
   isn't trusting stale/incorrect numbers.
3. **Outliers flagged, not silently deleted** — a $50k sale might be real.
   The pipeline flags it for human review instead of guessing.
4. **A polished `.xlsx`, not a raw CSV** — a formatted summary tab with a
   native Excel chart, ready to forward straight to leadership.

## How to adapt this per client
1. Point `INPUT_FILE` at the client's real export.
2. Update `NUMERIC_COLUMNS`, `DATE_COLUMNS`, `KEY_GROUP_COLUMN`, and
   `REVENUE_COLUMN` in the `CONFIG` block to match their column names.
3. Add any client-specific business rules to `clean()` (e.g. standardizing
   product SKUs, mapping regional aliases, currency conversion).
4. For a recurring engagement, wrap `run()` in a scheduled job that
   watches a folder (e.g. a synced Google Drive/Dropbox export) and
   emails the finished `.xlsx` automatically every morning.

## Pricing angle
- **One-off cleanup job**: flat fee based on file size/complexity
  ($75–$300).
- **Recurring "reporting as a service"**: monthly retainer ($100–$400/mo)
  for a pipeline that runs on a schedule and delivers a fresh report
  automatically — this is the higher-value, stickier offer.
