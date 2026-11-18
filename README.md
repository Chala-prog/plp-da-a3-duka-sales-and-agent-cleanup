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
This exploratory analysis evaluates 5,000 retail transactions recorded across 3 regional branches during fiscal year 2026. Online channels generated a total revenue of KES 12,450,000, representing 45% of total enterprise turnover. Premium-tier products achieved an average order value of KES 3,200, outperforming standard items by 25%. Weekend sales volume peaked on Saturdays with over 1,200 orders processed. Overall data completeness reached 98.5% prior to imputation.

> Numbers above come from the stand-in `duka_sales.csv`; re-run `python build_duka.py` after replacing it with the course file and paste the refreshed summary.
