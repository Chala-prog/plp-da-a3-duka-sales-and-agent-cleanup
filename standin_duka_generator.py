"""STAND-IN Duka generator (NOT the course generator).
Only used so the pipeline can run before the real duka_sales.csv is available.
Replace duka_sales.csv with the real file and re-run build_duka.py."""
import numpy as np, pandas as pd
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
