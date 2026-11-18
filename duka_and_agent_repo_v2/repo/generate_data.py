"""Dataset generators for Week 3 assignment.

Part A  - Duka Electronics (seed 2026) -> duka_sales.csv
Part B  - Mobile money agent network (seed 44) -> agents.csv, agent_transactions.csv
"""
import numpy as np, pandas as pd

# ---------------------------------------------------------------------------
# PART A: Duka Electronics  (seed 2026)
# PASTE the Practicals-section generator here UNCHANGED (do not alter the seed).
# ---------------------------------------------------------------------------
# >>> STAND-IN (NOT the course generator). Replace this block with the Practicals-section code, unchanged. <<<
rng = np.random.default_rng(2026)
n = 1000
branches = ["Nairobi", "Mombasa", "Kisumu", "Nakuru", "Eldoret"]
prods = {"Laptop": 62000, "Smartphone": 28500, "Television": 45000, "Tablet": 24000,
         "Headphones": 6500, "Smartwatch": 15500, "Router": 7800, "Printer": 18500}
d = pd.DataFrame({
    "order_id": [f"D{i:04d}" for i in range(1, n + 1)],
    "order_date": (pd.Timestamp("2026-01-01") + pd.to_timedelta(rng.integers(0, 181, n), unit="D")).strftime("%Y-%m-%d"),
    "branch": rng.choice(branches, n, p=[.34, .22, .15, .15, .14]),
    "product": rng.choice(list(prods), n, p=[.10, .12, .08, .12, .22, .12, .14, .10]),
    "channel": rng.choice(["Online", "In-store"], n, p=[.4, .6]),
    "units": rng.integers(1, 6, n),
})
d["unit_price"] = d["product"].map(prods)
d.to_csv("duka_sales.csv", index=False)


# ---------------------------------------------------------------------------
# PART B: Agent network (seed 44) - exactly as supplied in the brief
# ---------------------------------------------------------------------------
rng = np.random.default_rng(44)
n = 600
agents = pd.DataFrame({
    "agent_id": [f"AG{i:03d}" for i in range(1, 21)],
    "county": rng.choice(["Nairobi", "Kiambu", "Nakuru", "Mombasa", "Kisumu"], 20),
    "onboarded": pd.date_range("2024-01-01", periods=20, freq="MS")
                   .strftime("%Y-%m-%d"),
})
txns = pd.DataFrame({
    "txn_id": [f"TX{i:05d}" for i in range(1, n + 1)],
    "agent_id": rng.choice(agents["agent_id"], n),
    "txn_type": rng.choice(["Deposit", "withdrawal", "DEPOSIT ",
                            "Withdrawal", "deposit"], n),
    "amount": rng.lognormal(7.5, 0.8, n).round(0).astype(str),
    "txn_date": (pd.to_datetime("2026-03-01")
                 + pd.to_timedelta(rng.integers(0, 60, n), unit="D"))
                .strftime("%Y-%m-%d"),
})
idx = rng.choice(n, 30, replace=False)
txns.loc[idx[:15], "amount"] = np.nan
txns.loc[idx[15:20], "amount"] = "N/A"
txns.loc[idx[20:25], "amount"] = "250000"
txns.loc[idx[25:], "txn_date"] = txns.loc[idx[25:], "txn_date"].str.replace("-", "/")
txns = pd.concat([txns, txns.sample(12, random_state=1)], ignore_index=True)
txns.loc[rng.choice(len(txns), 8, replace=False), "agent_id"] = "AG099"
agents.to_csv("agents.csv", index=False)
txns.to_csv("agent_transactions.csv", index=False)
