# PLP Data Analytics · Week 3 — Duka Sales & Agent Cleanup

Two client engagements in one repository.

| Engagement | What | Main files |
|---|---|---|
| 1 · Duka Electronics | Explore H1-2026 sales (1,000 orders) and answer the Head of Sales' questions | `duka_sales_exploration.ipynb`, `excel/duka_branch_summary.xlsx`, `outputs/premium_online_coast_capital.csv` |
| 2 · Mobile-money agents | Audit, clean, merge and summarise an unreliable transaction extract | `cleaning_pipeline.py`, `agent_cleanup.ipynb`, `outputs/agent_transactions_clean.csv` |

## Regenerate the data and run
```bash
pip install -r requirements.txt
python generate_data.py        # Duka (seed 2026) + agents (seed 44); writes the 3 CSVs
jupyter nbconvert --to notebook --execute --inplace duka_sales_exploration.ipynb
jupyter nbconvert --to notebook --execute --inplace agent_cleanup.ipynb
```
Do not change the seeds.

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

> Numbers above come from the stand-in `duka_sales.csv`; re-run `python build_duka.py` after replacing it with the course file and paste the refreshed summary.
