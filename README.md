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
Do not change the seeds. `build_duka.py` needs LibreOffice (`soffice`) only to create the PivotTable; `duka_sales_exploration.ipynb` can also be opened and run on its own in Jupyter.

> **Note on Part A of `generate_data.py`:** it currently holds a clearly-marked stand-in Duka generator. Before submitting, paste the course Practicals generator (seed 2026) there unchanged and re-run `python build_duka.py`, then refresh the Executive Summary below.

## Engagement 2 — before/after row count
| Stage | Rows |
|---|---|
| Raw `agent_transactions.csv` | **612** |
| After removing 12 duplicate `txn_id` | **600** |
| After merge with agent register (left join) | **600** (7 on unknown agent AG099) |

Also: 20 amounts imputed (median by type, flagged `amount_imputed`), 40 IQR outliers flagged (`is_outlier`) including five identical KES 250,000 values held for verification.

## Duka Executive Summary
- **Nairobi** is the top branch with KES 22,925,800 (32.4% of the KES 70,678,600 total revenue across 1,000 orders).
- **Laptop** is the top product by revenue at KES 18,972,000, while **Headphones** leads on units with 608 sold.
- Weekend orders are 27.9% of orders (279), averaging KES 65,970 per order vs KES 72,501 on weekdays.
- **March** was the best month (KES 13,930,900) and **May** the worst (KES 9,832,700), a 41.7% gap.
- 71 Premium online orders from Nairobi and Mombasa generated KES 8,784,000, i.e. 12.4% of revenue.

> Numbers above come from the stand-in generator; paste the course generator into `generate_data.py` and re-run `python generate_data.py && python build_duka.py`.
