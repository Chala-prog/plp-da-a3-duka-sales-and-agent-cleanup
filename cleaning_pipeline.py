"""Cleaning pipeline for the mobile-money agent transaction extract."""
import numpy as np
import pandas as pd

VALID_TYPES = {"deposit": "Deposit", "withdrawal": "Withdrawal"}
INVALID_STRINGS = {"", "n/a", "na", "null", "none", "nan", "-"}
EXTREME_MULTIPLE = 100   # an amount more than 100x the median is treated as a suspected error (is_extreme_outlier)


def load_raw(txn_path="agent_transactions.csv", agent_path="agents.csv"):
    """Read both files as text so nothing is silently coerced.
    Only truly empty cells become NaN; the text 'N/A' is preserved for the audit."""
    kw = dict(dtype=str, keep_default_na=False, na_values=[""])
    return pd.read_csv(txn_path, **kw), pd.read_csv(agent_path, **kw)


def _amount_numeric(s):
    """Return numeric amount; NaN for missing or non-numeric placeholders."""
    s = s.astype("string").str.strip()
    s = s.mask(s.str.lower().isin(INVALID_STRINGS))
    return pd.to_numeric(s, errors="coerce")


def _dates(s):
    """Normalise '/' separators then parse as ISO; unparseable -> NaT."""
    s = s.astype("string").str.strip().str.replace("/", "-", regex=False)
    return pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")


def _iqr_bounds(x, k=1.5):
    q1, q3 = x.quantile([0.25, 0.75])
    iqr = q3 - q1
    return q1 - k * iqr, q3 + k * iqr


def audit_data(txns, agents):
    """Data-quality audit on the RAW extract. Returns DataFrame:
    issue, dimension, column, rows_affected."""
    t = txns.copy()
    amt_raw = t["amount"].astype("string")
    amt_num = _amount_numeric(t["amount"])
    clean_dates = t["txn_date"].astype("string")
    typ = t["txn_type"].astype("string")
    norm_typ = typ.str.strip().str.lower()
    lo, hi = _iqr_bounds(amt_num.dropna())
    rows = [
        ("Fully duplicated rows (same txn_id and all fields)", "Uniqueness", "all columns",
         int(t.duplicated(keep="first").sum())),
        ("Repeated txn_id (includes duplicates whose agent_id differs)", "Uniqueness", "txn_id",
         int(t["txn_id"].duplicated(keep="first").sum())),
        ("txn_type inconsistent case / trailing whitespace", "Consistency", "txn_type",
         int((typ != typ.str.strip().str.title()).sum())),
        ("txn_type outside {Deposit, Withdrawal} after normalising", "Validity", "txn_type",
         int((~norm_typ.isin(VALID_TYPES) ).sum())),
        ("amount missing (empty cell)", "Completeness", "amount", int(amt_raw.isna().sum())),
        ("amount is text placeholder ('N/A')", "Validity", "amount",
         int(amt_raw.str.strip().str.lower().isin(INVALID_STRINGS - {""}).sum())),
        ("amount stored as text; numeric values converted", "Validity", "amount", int(amt_num.notna().sum())),
        ("amount is IQR outlier (> Q3+1.5*IQR)", "Accuracy", "amount", int((amt_num > hi).sum())),
        (f"amount more than {EXTREME_MULTIPLE}x the median (extreme / suspected error)", "Accuracy", "amount",
         int((amt_num > EXTREME_MULTIPLE * amt_num.median()).sum())),
        ("txn_date uses '/' instead of ISO '-' separator", "Consistency", "txn_date",
         int(clean_dates.str.contains("/", na=False).sum())),
        ("txn_date unparseable", "Validity", "txn_date", int(_dates(t["txn_date"]).isna().sum())),
        ("agent_id not in agent register (e.g. AG099)", "Integrity / Consistency", "agent_id",
         int((~t["agent_id"].isin(agents["agent_id"])).sum())),
        ("Duplicate agent_id in register", "Uniqueness", "agents.agent_id",
         int(agents["agent_id"].duplicated().sum())),
    ]
    return pd.DataFrame(rows, columns=["issue", "dimension", "column", "rows_affected"])


def clean_agent_data(txns, agents):
    """Clean the agent transaction extract and merge it with the agent register.

    Rules applied (all rule-based, no hard-coded rows, no valid row is ever dropped except true duplicates):
      1. Uniqueness   - drop repeated txn_id (keep first).
      2. Consistency  - strip/lower txn_type and map to {'Deposit','Withdrawal'}.
      3. Validity     - amount -> numeric ('N/A' etc. -> NaN); txn_date '/' -> '-' then datetime;
                        an unparseable date is kept as NaT and logged, never dropped.
      4. Accuracy     - `is_outlier`: IQR (1.5x) flag on observed amounts (includes ordinary high-value
                        transactions); `is_extreme_outlier`: amount > EXTREME_MULTIPLE x the median
                        (suspected errors, excluded from headline figures until verified).
      5. Completeness - missing/invalid amount imputed with the median of valid, non-outlier amounts
                        of the same txn_type; every imputed row flagged in `amount_imputed`.
      6. Integrity    - left-join to the register for `county`; unknown agents kept with
                        county='Unknown' and `unknown_agent`=True.

    Parameters
    ----------
    txns : DataFrame  raw transactions (txn_id, agent_id, txn_type, amount, txn_date)
    agents : DataFrame agent register (agent_id, county, onboarded)

    Returns
    -------
    (clean_df, log) : cleaned DataFrame and a log DataFrame (step, rule, rows_affected, rows_after).
    """
    log = []

    def note(step, rule, affected, df):
        log.append({"step": step, "rule": rule, "rows_affected": int(affected), "rows_after": len(df)})

    df = txns.copy()
    note("0 Load", "Raw rows received", 0, df)

    # 1. Uniqueness
    n_before = len(df)
    df = df.drop_duplicates(subset="txn_id", keep="first").reset_index(drop=True)
    note("1 Uniqueness", "Drop repeated txn_id (keep first)", n_before - len(df), df)

    # 2. Consistency: txn_type
    raw_type = df["txn_type"].astype("string")
    df["txn_type"] = raw_type.str.strip().str.lower().map(VALID_TYPES)
    bad = df["txn_type"].isna().sum()
    note("2 Consistency", "Standardise txn_type to Deposit/Withdrawal "
         f"(rows changed; {bad} unmappable left as NaN)",
         (raw_type != df["txn_type"]).sum(), df)

    # 3. Validity: types
    raw_amt_missing = df["amount"].isna() | df["amount"].astype("string").str.strip().str.lower().isin(INVALID_STRINGS)
    df["amount"] = _amount_numeric(df["amount"])
    note("3a Validity", "Convert amount to numeric (text placeholders -> NaN)",
         raw_amt_missing.sum(), df)
    slash = df["txn_date"].astype("string").str.contains("/", na=False)
    df["txn_date"] = _dates(df["txn_date"])
    note("3b Validity", "Parse txn_date; '/' separators converted", slash.sum(), df)
    note("3c Validity", "Unparseable txn_date kept as missing (rows are never dropped)",
         df["txn_date"].isna().sum(), df)

    # 4. Accuracy: outliers, computed on observed amounts only (before imputation)
    lo, hi = _iqr_bounds(df["amount"].dropna())
    df["is_outlier"] = ((df["amount"] < lo) | (df["amount"] > hi)).fillna(False)
    note("4a Accuracy", f"Flag IQR outliers (bounds {lo:,.0f} to {hi:,.0f})", df["is_outlier"].sum(), df)
    median_observed = df["amount"].median()
    df["is_extreme_outlier"] = (df["amount"] > EXTREME_MULTIPLE * median_observed).fillna(False)
    note("4b Accuracy", f"Flag extreme outliers (> {EXTREME_MULTIPLE}x median of {median_observed:,.0f}); excluded from headline figures",
         df["is_extreme_outlier"].sum(), df)

    # 5. Completeness: median imputation by txn_type, excluding outliers
    df["amount_imputed"] = df["amount"].isna()
    med = df.loc[~df["is_outlier"]].groupby("txn_type")["amount"].median()
    df["amount"] = df["amount"].fillna(df["txn_type"].map(med))
    note("5 Completeness", "Impute missing/invalid amount with per-type median (flagged)",
         df["amount_imputed"].sum(), df)

    # 6. Integrity: merge
    ag = agents.copy()
    ag["agent_id"] = ag["agent_id"].str.strip()
    ag = ag.drop_duplicates("agent_id")
    n_before = len(df)
    df = df.merge(ag[["agent_id", "county", "onboarded"]], on="agent_id", how="left", validate="m:1")
    assert len(df) == n_before, "merge changed row count"
    df["unknown_agent"] = df["county"].isna()
    df["county"] = df["county"].fillna("Unknown")
    note("6 Integrity", "Left-join agent register; unknown agents kept, county='Unknown'",
         df["unknown_agent"].sum(), df)

    df["amount"] = df["amount"].astype(float).round(2)
    df["is_outlier"] = df["is_outlier"].astype(bool)
    df["is_extreme_outlier"] = df["is_extreme_outlier"].astype(bool)
    df["week_start"] = df["txn_date"] - pd.to_timedelta(df["txn_date"].dt.weekday, unit="D")
    return df, pd.DataFrame(log)
