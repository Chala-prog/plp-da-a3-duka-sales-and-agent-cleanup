# PLP Data Analytics · Week 3 — Duka Sales & Agent Cleanup

Two client engagements in one repository.

| Engagement | What | Main files |
|---|---|---|
| 1 · Duka Electronics | Explore H1-2026 sales (1,000 orders) and answer the Head of Sales' questions | `duka_sales_exploration.ipynb`, `excel/duka_branch_summary.xlsx`, `outputs/premium_online_coast_capital.csv` |
| 2 · Mobile-money agents | Audit, clean, merge and summarise an unreliable transaction extract | `cleaning_pipeline.py`, `agent_cleanup.ipynb`, `outputs/agent_transactions_clean.csv` |

## Regenerate the data and run
```bash
pip install -r requirements.txt
python generate_data.py        # Duka (seed 2026) + agents (seed 44); writes duka_sales.csv, agents.csv, agent_transactions.csv
python build_duka.py           # Engagement 1: rebuilds notebook, Excel workbook (with real PivotTable) and premium CSV from duka_sales.csv
jupyter nbconvert --to notebook --execute --inplace duka_sales_exploration.ipynb   # Engagement 1 (run on its own)
jupyter nbconvert --to notebook --execute --inplace agent_cleanup.ipynb   # Engagement 2 (imports cleaning_pipeline.py)
```
Do not change the seeds. `build_duka.py` (with its helper `make_pivot.py`) needs LibreOffice (`soffice`) only to create the real PivotTable; the committed `excel/*_recalc_check.csv` files let `duka_sales_exploration.ipynb` run on its own in Jupyter.

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
- 83 Premium online orders from Nairobi and Mombasa generated KES 10,218,000, i.e. 14.2% of revenue.
