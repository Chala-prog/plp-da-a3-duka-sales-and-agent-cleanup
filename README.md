# PLP Data Analytics · Week 3 — Duka Sales & Agent Cleanup

Two client engagements in one repository.

| Engagement | What | Main files |
|---|---|---|
| 1 · Duka Electronics | Explore H1-2026 sales (1,000 orders) and answer the Head of Sales' questions | `duka_sales_exploration.ipynb`, `excel/duka_branch_summary.xlsx`, `outputs/premium_online_coast_capital.csv` |
| 2 · Mobile-money agents | Audit, clean, merge and summarise an unreliable transaction extract | `cleaning_pipeline.py`, `agent_cleanup.ipynb`, `outputs/agent_transactions_clean.csv` |

## Repository layout
```
README.md                         this file
generate_data.py                  both generators: Duka (seed 2026) and agents (seed 44)
build_duka.py, make_pivot.py      rebuild Engagement 1 from duka_sales.csv (make_pivot.py adds the PivotTable via LibreOffice)
duka_sales_exploration.ipynb      Engagement 1 notebook (Executive Summary, Tasks 1-8)
excel/duka_branch_summary.xlsx    SUMIFS + XLOOKUP table and a PivotTable of revenue by branch and channel
excel/*_recalc_check.csv          values read back from the workbook, used by Task 8 to reconcile with pandas
outputs/premium_online_coast_capital.csv     Premium + Online orders, Nairobi and Mombasa only
cleaning_pipeline.py              clean_agent_data(txns, agents) -> (clean_df, log)
agent_cleanup.ipynb               Engagement 2 notebook (audit, validation, summaries, Data Quality Report)
outputs/agent_transactions_clean.csv         cleaned, merged agent transactions
requirements.txt
```

## Regenerate the data and run
```bash
pip install -r requirements.txt
python generate_data.py        # Duka (seed 2026) + agents (seed 44); writes duka_sales.csv, agents.csv, agent_transactions.csv
python build_duka.py           # Engagement 1: rebuilds notebook, Excel workbook (with real PivotTable) and premium CSV from duka_sales.csv
jupyter nbconvert --to notebook --execute --inplace agent_cleanup.ipynb   # Engagement 2 (imports cleaning_pipeline.py)
```
Do not change the seeds (Part A uses 2026, Part B uses 44). `build_duka.py` needs LibreOffice (`soffice`) only to create the PivotTable; `duka_sales_exploration.ipynb` can also be opened and run on its own in Jupyter.

### Run each notebook
1. `python generate_data.py` first, so the three CSV inputs exist.
2. **Engagement 1:** open `duka_sales_exploration.ipynb` and choose *Restart & Run All* (it reads `duka_sales.csv` and `excel/*_recalc_check.csv`, and writes `outputs/premium_online_coast_capital.csv`).
3. **Engagement 2:** open `agent_cleanup.ipynb` and choose *Restart & Run All* (it imports `cleaning_pipeline.py` and writes `outputs/agent_transactions_clean.csv`).

### Assumptions
- `price_band` (not defined by the data): Budget below KES 5,000, Mid KES 5,000 to 29,999, Premium KES 30,000 and above, so Premium means the Laptop.
- The `Regions` sheet used by the XLOOKUP maps Nairobi to Central, Mombasa to Coast, Kisumu to Western, and Nakuru and Eldoret to Rift Valley.
- Revenue = `quantity × unit_price`. The data runs from 1 Jan to 29 Jun 2026.

## Engagement 2 — before/after row count
| Stage | Rows |
|---|---|
| Raw `agent_transactions.csv` | **612** |
| After removing 12 duplicate `txn_id` | **600** |
| After merge with agent register (left join) | **600** (7 on unknown agent AG099) |

Also: 20 amounts imputed (median by type, flagged `amount_imputed`), 40 IQR outliers flagged (`is_outlier`) including five identical KES 250,000 values held for verification.

## Duka Executive Summary
- **Nairobi** is the top branch with KES 26,106,100 (36.4% of the KES 71,792,500 total revenue across 1,000 orders).
- **Laptop** is the top product by revenue at KES 41,535,000, while **Tablet** leads on units with 691 sold.
- Weekend orders are 27.7% of orders (277), averaging KES 71,972 per order vs KES 71,724 on weekdays.
- **April** was the best month (KES 13,518,300) and **June** the worst (KES 10,903,400), a 24.0% gap.
- 33 Premium online orders from Nairobi and Mombasa generated KES 5,850,000, i.e. 8.1% of revenue.
