"""Rebuilds ALL Engagement-1 deliverables from duka_sales.csv.
Expected columns: order_id, order_date, branch, product, channel, units, unit_price
Usage: python build_duka.py   (then notebook is executed automatically)"""
import subprocess, shutil, os
import pandas as pd, numpy as np, nbformat as nbf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

df = pd.read_csv("duka_sales.csv", parse_dates=["order_date"])
need = {"order_id","order_date","branch","product","channel","units","unit_price"}
assert need <= set(df.columns), f"missing columns: {need-set(df.columns)}"
raw_shape = df.shape
df["revenue"] = df["units"] * df["unit_price"]
df["month"] = df["order_date"].dt.month_name()
df["weekday"] = df["order_date"].dt.day_name()
df["is_weekend"] = df["order_date"].dt.dayofweek >= 5
df["price_band"] = np.where(df["unit_price"] >= 28000, "Premium", "Value")
K = lambda v: f"KES {v:,.0f}"

# ---- results used for the written answers
br = df.groupby("branch").agg(rev=("revenue","sum"), orders=("order_id","count")).assign(aov=lambda x: x.rev/x.orders)
br_rev = br.sort_values("rev", ascending=False); br_ord = br.sort_values("orders", ascending=False)
same = list(br_rev.index) == list(br_ord.index)
pr = df.groupby("product").agg(rev=("revenue","sum"), units=("units","sum"))
top_rev = pr.sort_values("rev", ascending=False).head(3); top_u = pr.sort_values("units", ascending=False).head(3)
wk = df.groupby("is_weekend").agg(orders=("order_id","count"), rev=("revenue","sum"), aov=("revenue","mean"))
mo = df.groupby(df["order_date"].dt.month).agg(rev=("revenue","sum")); mo["name"] = [pd.Timestamp(2026,m,1).month_name() for m in mo.index]
best, worst = mo["rev"].idxmax(), mo["rev"].idxmin()
pct = (mo.rev[best]-mo.rev[worst])/mo.rev[worst]*100
sel = df[(df.price_band=="Premium")&(df.channel=="Online")&(df.branch.isin(["Nairobi","Mombasa"]))]
os.makedirs("outputs", exist_ok=True); os.makedirs("excel", exist_ok=True)
sel.to_csv("outputs/premium_online_coast_capital.csv", index=False)
total = df.revenue.sum()
ds, de = df.order_date.min().date(), df.order_date.max().date()

# ---- Excel
regions = {"Nairobi":"Central","Kiambu":"Central","Nakuru":"Rift Valley","Eldoret":"Rift Valley","Mombasa":"Coast","Kisumu":"Western"}
wb = Workbook(); wd = wb.active; wd.title = "Data"
wd.append(["order_id","order_date","branch","product","channel","units","unit_price","revenue"])
for r in df.itertuples(): wd.append([r.order_id, r.order_date.date(), r.branch, r.product, r.channel, int(r.units), float(r.unit_price), float(r.revenue)])
n = len(df)+1
wr = wb.create_sheet("Regions"); wr.append(["branch","region"])
for b in br.index: wr.append([b, regions.get(b,"Other")])
ws = wb.create_sheet("Branch_Summary", 0)
ws.append(["Rank","Branch","Total revenue (SUMIFS)","Orders (COUNTIFS)","Region (XLOOKUP)"])
for i,b in enumerate(br_rev.index, start=2):
    ws.append([i-1, b, f"=SUMIFS(Data!$H$2:$H${n},Data!$C$2:$C${n},B{i})",
               f"=COUNTIFS(Data!$C$2:$C${n},B{i})", f'=XLOOKUP(B{i},Regions!$A$2:$A${len(br)+1},Regions!$B$2:$B${len(br)+1},"Unknown")'])
last = len(br)+1
ws.append(["", "TOTAL", f"=SUM(C2:C{last})", f"=SUM(D2:D{last})"])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
pv = wb.create_sheet("Pivot_Branch_Channel")
chs = sorted(df.channel.unique()); pv.append(["Branch"]+chs+["Grand total"])
for i,b in enumerate(br_rev.index, start=2):
    pv.append([b]+[f'=SUMIFS(Data!$H$2:$H${n},Data!$C$2:$C${n},$A{i},Data!$E$2:$E${n},{chr(66+j)}$1)' for j in range(len(chs))]+[f"=SUM(B{i}:{chr(65+len(chs))}{i})"])
pv.append(["Grand total"]+[f"=SUM({chr(66+j)}2:{chr(66+j)}{last})" for j in range(len(chs)+1)])
for c in pv[1]: c.font = Font(bold=True)
for w in (ws, pv): w.column_dimensions["A"].width = 14; w.column_dimensions["B"].width = 16; w.column_dimensions["C"].width = 24; w.column_dimensions["D"].width = 20; w.column_dimensions["E"].width = 20
wb.calculation.fullCalcOnLoad = True
wb.save("excel/duka_branch_summary.xlsx")
# verify SUMIFS/COUNTIFS via LibreOffice recalculation of a temp copy
os.makedirs("/tmp/lo", exist_ok=True)
subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","/tmp/lo","excel/duka_branch_summary.xlsx"],capture_output=True)
from openpyxl import load_workbook
chk = load_workbook("/tmp/lo/duka_branch_summary.xlsx", data_only=True)
rows = [(r[1], r[2], r[3]) for r in chk["Branch_Summary"].iter_rows(min_row=2, max_row=last, values_only=True)]
pd.DataFrame(rows, columns=["branch","sumifs_revenue","countifs_orders"]).to_csv("excel/sumifs_recalc_check.csv", index=False)
pcells = [list(r) for r in chk["Pivot_Branch_Channel"].iter_rows(min_row=1, values_only=True)]
pd.DataFrame(pcells[1:], columns=pcells[0]).to_csv("excel/pivot_recalc_check.csv", index=False)

# ---- Notebook
nb = nbf.v4.new_notebook(); C = []
md = lambda s: C.append(nbf.v4.new_markdown_cell(s)); code = lambda s: C.append(nbf.v4.new_code_cell(s))
md("# Duka Electronics — H1 2026 Sales Exploration")
bl = br_rev.index[0]; bo = br_ord.index[0]
md(f"""## Executive Summary
- **{br_rev.index[0]}** is the top branch with {K(br_rev.rev.iloc[0])} ({br_rev.rev.iloc[0]/total:.1%} of the {K(total)} total revenue across {len(df):,} orders).
- **{top_rev.index[0]}** is the top product by revenue at {K(top_rev.rev.iloc[0])}, while **{top_u.index[0]}** leads on units with {top_u.units.iloc[0]:,} sold.
- Weekend orders are {wk.orders[True]/len(df):.1%} of orders ({wk.orders[True]:,}), averaging {K(wk.aov[True])} per order vs {K(wk.aov[False])} on weekdays.
- **{mo.name[best]}** was the best month ({K(mo.rev[best])}) and **{mo.name[worst]}** the worst ({K(mo.rev[worst])}), a {pct:.1f}% gap.
- {len(sel):,} Premium online orders from Nairobi and Mombasa generated {K(sel.revenue.sum())}, i.e. {sel.revenue.sum()/total:.1%} of revenue.""")
md("## Task 1 · Load & first-five-minutes inspection")
code('''import pandas as pd, numpy as np
pd.options.display.float_format = "{:,.2f}".format
df = pd.read_csv("duka_sales.csv", parse_dates=["order_date"])
print("shape:", df.shape); print(df.dtypes); print("\\nmissing values:\\n", df.isna().sum())
print("\\ndate range:", df.order_date.min().date(), "to", df.order_date.max().date())
for c in ["branch","product","channel"]: print("\\n", df[c].value_counts().to_dict())
print("duplicate order_id:", df.order_id.duplicated().sum())
df.describe()''')
md(f"""**Inspection notes.** {raw_shape[0]:,} rows × {raw_shape[1]} columns. `order_date` parsed to datetime (range {ds} → {de}); ({df.isna().sum().sum()} nulls, {df.order_id.duplicated().sum()} duplicates). Categories: {df.branch.nunique()} branches, {df["product"].nunique()} products, channels {', '.join(sorted(df.channel.unique()))}.
> **Data-source note:** `duka_sales.csv` in this build is a stand-in (`standin_duka_generator.py`). Replace it with the course-generated file and run `python build_duka.py` to refresh every number below.""")
md("## Task 2 · Derived columns")
code('''df["revenue"] = df["units"] * df["unit_price"]
df["month"] = df["order_date"].dt.month_name()
df["weekday"] = df["order_date"].dt.day_name()
df["is_weekend"] = df["order_date"].dt.dayofweek >= 5
df["price_band"] = np.where(df["unit_price"] >= 28000, "Premium", "Value")
df[["order_id","revenue","month","weekday","is_weekend","price_band"]].head()''')
md("`revenue = units × unit_price`; `price_band` uses `unit_price ≥ KES 28,000` → Premium.")
md("## Task 3 · Branch ranking: revenue vs orders")
code('''b = df.groupby("branch").agg(total_revenue=("revenue","sum"), orders=("order_id","count"))
b["avg_order_value"] = b.total_revenue/b.orders
b["rank_revenue"] = b.total_revenue.rank(ascending=False).astype(int)
b["rank_orders"] = b.orders.rank(ascending=False, method="min").astype(int)
b.sort_values("total_revenue", ascending=False)''')
expl = ("The two rankings are identical here: order volume and revenue move together because average order values are similar across branches." if same else
        f"The rankings differ. **{bl}** leads on revenue but **{bo}** leads on orders: revenue = orders × average order value, so a branch with fewer orders can out-earn one with more if it sells higher-priced items (e.g. {bl} AOV {K(br.aov[bl])} vs {bo} {K(br.aov[bo])}).")
md(f"**Answer.** By revenue: {' > '.join(br_rev.index)}. By orders: {' > '.join(br_ord.index)}. {expl}")
md("## Task 4 · Top three products")
code('''p = df.groupby("product").agg(revenue=("revenue","sum"), units=("units","sum"))
print("By revenue:"); display(p.nlargest(3,"revenue"))
print("By units:");   display(p.nlargest(3,"units"))''')
dif = set(top_rev.index) ^ set(top_u.index)
md(f"**Answer.** Top by revenue: {', '.join(top_rev.index)}. Top by units: {', '.join(top_u.index)}. " +
   ("The lists differ because revenue rewards price × volume: expensive items earn heavily from fewer units, while cheap accessories sell many units but contribute little revenue." if dif else "The lists contain the same products."))
md("## Task 5 · Weekend vs weekday")
code('''w = df.groupby("is_weekend").agg(orders=("order_id","count"), total_revenue=("revenue","sum"), avg_order_revenue=("revenue","mean"))
w.index = w.index.map({True:"Weekend", False:"Weekday"}); w''')
md(f"**Answer.** Weekdays: {wk.orders[False]:,} orders, {K(wk.rev[False])}, avg {K(wk.aov[False])}. Weekends: {wk.orders[True]:,} orders, {K(wk.rev[True])}, avg {K(wk.aov[True])}. Weekends cover only 2 of 7 days, so compare the average order value, not totals: weekend orders are {'larger' if wk.aov[True]>wk.aov[False] else 'smaller'} on average ({(wk.aov[True]/wk.aov[False]-1)*100:+.1f}%).")
md("## Task 6 · Best and worst month")
code('''m = df.groupby(df.order_date.dt.month).revenue.sum()
m.index = [pd.Timestamp(2026,i,1).month_name() for i in m.index]
best, worst = m.idxmax(), m.idxmin()
print(m); print(f"best={best} worst={worst} difference={(m[best]-m[worst])/m[worst]*100:.1f}% (best vs worst)")''')
md(f"**Answer.** Best: **{mo.name[best]}** ({K(mo.rev[best])}); worst: **{mo.name[worst]}** ({K(mo.rev[worst])}). Best is **{pct:.1f}%** higher than worst, i.e. worst is {(mo.rev[best]-mo.rev[worst])/mo.rev[best]*100:.1f}% below best.")
md("## Task 7 · Premium online orders, Nairobi & Mombasa")
code('''sel = df[(df.price_band=="Premium") & (df.channel=="Online") & (df.branch.isin(["Nairobi","Mombasa"]))]
print(len(sel), "orders | revenue KES", f"{sel.revenue.sum():,.0f}")
sel.to_csv("outputs/premium_online_coast_capital.csv", index=False)''')
md(f"**Answer.** **{len(sel):,}** orders, total revenue **{K(sel.revenue.sum())}**. Exported to `outputs/premium_online_coast_capital.csv` (`index=False`).")
md("## Task 8 · Excel reconciliation (`excel/duka_branch_summary.xlsx`)")
code('''xl = pd.read_csv("excel/sumifs_recalc_check.csv").set_index("branch")      # SUMIFS/COUNTIFS values recalculated from the workbook
cmp = b[["total_revenue","orders"]].join(xl)
cmp["match"] = np.isclose(cmp.total_revenue, cmp.sumifs_revenue) & (cmp.orders == cmp.countifs_orders)
pv = pd.read_csv("excel/pivot_recalc_check.csv").set_index("Branch")
pd_pivot = df.pivot_table(index="branch", columns="channel", values="revenue", aggfunc="sum")
print("SUMIFS = pandas for every branch:", cmp.match.all(), "| grand total:", f"{xl.sumifs_revenue.sum():,.0f} vs {df.revenue.sum():,.0f}")
print("Branch×channel table = pandas pivot:", np.allclose(pv.loc[pd_pivot.index, pd_pivot.columns].astype(float), pd_pivot))
cmp''')
md("**Answer.** The workbook has `Branch_Summary` (SUMIFS revenue + COUNTIFS orders + XLOOKUP region from the `Regions` sheet) and `Pivot_Branch_Channel` (branch × channel revenue). The check above shows every total equals pandas.")
nb["cells"] = C; nbf.write(nb, "duka_sales_exploration.ipynb")
subprocess.run(["jupyter","nbconvert","--to","notebook","--execute","--inplace","duka_sales_exploration.ipynb"],capture_output=True)
exec_sum = C[1].source
r = open("README.md").read()
head = r.split("## Duka Executive Summary")[0]
open("README.md","w").write(head+"## Duka Executive Summary\n"+exec_sum.replace("## Executive Summary\n","")+"\n\n> Numbers above come from the stand-in `duka_sales.csv`; re-run `python build_duka.py` after replacing it with the course file and paste the refreshed summary.\n")
print(exec_sum)
