"""Rebuilds ALL Engagement-1 deliverables from duka_sales.csv.
Expected columns: order_id, order_date, branch, product, channel, quantity, unit_price
Usage: python build_duka.py   (then notebook is executed automatically)"""
import subprocess, shutil, os, sys
import pandas as pd, numpy as np, nbformat as nbf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

STANDIN = ">>> STAND-IN" in open("generate_data.py").read()   # notes about stand-in data show only until the course generator is pasted
df = pd.read_csv("duka_sales.csv", parse_dates=["order_date"])
need = {"order_id","order_date","branch","product","channel","quantity","unit_price"}
assert need <= set(df.columns), f"missing columns: {need-set(df.columns)}"
raw_shape = df.shape
df["revenue"] = df["quantity"] * df["unit_price"]
df["month"] = df["order_date"].dt.month_name()
df["weekday"] = df["order_date"].dt.day_name()
df["is_weekend"] = df["order_date"].dt.dayofweek >= 5
df["price_band"] = pd.cut(df["unit_price"], bins=[0, 5000, 30000, np.inf], right=False,
                          labels=["Budget", "Mid", "Premium"]).astype(str)
K = lambda v: f"KES {v:,.0f}"

# ---- results used for the written answers
br = df.groupby("branch").agg(rev=("revenue","sum"), orders=("order_id","count")).assign(aov=lambda x: x.rev/x.orders)
br_rev = br.sort_values("rev", ascending=False); br_ord = br.sort_values("orders", ascending=False)
same = list(br_rev.index) == list(br_ord.index)
pr = df.groupby("product").agg(rev=("revenue","sum"), units=("quantity","sum"))
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
wd.append(["order_id","order_date","branch","product","channel","quantity","unit_price","revenue"])
for r in df.itertuples(): wd.append([r.order_id, r.order_date.date(), r.branch, r.product, r.channel, int(r.quantity), float(r.unit_price), float(r.revenue)])
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
for w in (ws,): w.column_dimensions["A"].width = 14; w.column_dimensions["B"].width = 16; w.column_dimensions["C"].width = 24; w.column_dimensions["D"].width = 20; w.column_dimensions["E"].width = 20
# 1) workbook with Data + Regions only -> LibreOffice adds a REAL PivotTable -> 2) openpyxl re-opens it (pivot preserved) and adds the formula sheet
tmp = Workbook(); t0 = tmp.active; t0.title = "Data"
for row in wd.iter_rows(values_only=True): t0.append(list(row))
tmp.save("/tmp/duka_base.xlsx")
subprocess.run(["python3","make_pivot.py","/tmp/duka_base.xlsx","/tmp/duka_pivot.xlsx",str(n)],check=True)
from openpyxl import load_workbook
pvwb = load_workbook("/tmp/duka_pivot.xlsx", data_only=True)
pcells = [[c for c in r] for r in pvwb["Pivot_Branch_Channel"].iter_rows(values_only=True)]
pcells = [r for r in pcells if any(x is not None for x in r)]
pd.DataFrame(pcells).to_csv("excel/pivot_recalc_check.csv", index=False, header=False)
wb2 = load_workbook("/tmp/duka_pivot.xlsx")
for name in ("Branch_Summary","Regions"):
    new = wb2.create_sheet(name, 0 if name=="Branch_Summary" else None)
    for row in wb[name].iter_rows():
        for c in row: 
            nc = new.cell(row=c.row, column=c.column, value=c.value)
            if c.font and c.font.bold: nc.font = Font(bold=True, color=c.font.color.rgb if c.font.color else None); 
            if c.fill and c.fill.fgColor and c.fill.fgColor.rgb not in (None,"00000000"): nc.fill = PatternFill("solid", fgColor=c.fill.fgColor.rgb)
    for k,v in wb[name].column_dimensions.items(): new.column_dimensions[k].width = v.width
wb = wb2
wb.calculation.fullCalcOnLoad = True
wb.save("excel/duka_branch_summary.xlsx")
print("sheets:", wb.sheetnames, "| pivots:", len(wb["Pivot_Branch_Channel"]._pivots))
# verify SUMIFS/COUNTIFS via LibreOffice recalculation of a temp copy
os.makedirs("/tmp/lo", exist_ok=True)
subprocess.run(["soffice","--headless","--convert-to","xlsx","--outdir","/tmp/lo","excel/duka_branch_summary.xlsx"],capture_output=True)
from openpyxl import load_workbook
chk = load_workbook("/tmp/lo/duka_branch_summary.xlsx", data_only=True)
rows = [(r[1], r[2], r[3]) for r in chk["Branch_Summary"].iter_rows(min_row=2, max_row=last, values_only=True)]
pd.DataFrame(rows, columns=["branch","sumifs_revenue","countifs_orders"]).to_csv("excel/sumifs_recalc_check.csv", index=False)

# ---- store the computed results as the cached values of the Branch_Summary formulas, so every viewer
# (GitHub tools, Python, LibreOffice older than 24.8 which has no XLOOKUP) shows numbers. Excel still
# recalculates everything on open because fullCalcOnLoad is set.
import zipfile, re
def add_cached_values(path, values):
    src = zipfile.ZipFile(path); tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet") and b"Total revenue (SUMIFS)" in data:
                xml = data.decode("utf-8")
                for ref, val in values.items():
                    text = isinstance(val, str)
                    pattern = re.compile(r'<c r="' + ref + r'"([^>]*)>(<f>.*?</f>)<v></v></c>', re.S)
                    def fill_in(m, text=text, val=val):
                        kind = ' t="str"' if text else ""
                        return '<c r="' + ref + '"' + m.group(1) + kind + ">" + m.group(2) + "<v>" + str(val) + "</v></c>"
                    xml, hits = pattern.subn(fill_in, xml)
                    assert hits == 1, ref
                data = xml.encode("utf-8")
            dst.writestr(item, data)
    src.close(); os.replace(tmp, path)

cached = {}
for i, b in enumerate(br_rev.index, start=2):
    cached[f"C{i}"] = int(br.rev[b]); cached[f"D{i}"] = int(br.orders[b]); cached[f"E{i}"] = regions.get(b, "Unknown")
cached[f"C{last+1}"] = int(br.rev.sum()); cached[f"D{last+1}"] = int(br.orders.sum())
add_cached_values("excel/duka_branch_summary.xlsx", cached)

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
note = "\n> **Data-source note:** Part A of `generate_data.py` is still the stand-in generator, so every number here is provisional." if STANDIN else ""
md(f"""**Inspection notes.** {raw_shape[0]:,} rows × {raw_shape[1]} columns. `order_date` parsed to datetime (range {ds} → {de}); ({df.isna().sum().sum()} nulls, {df.order_id.duplicated().sum()} duplicates). Categories: {df.branch.nunique()} branches, {df["product"].nunique()} products, channels {', '.join(sorted(df.channel.unique()))}.
{note}""")
md("## Task 2 · Derived columns")
code('''df["revenue"] = df["quantity"] * df["unit_price"]
df["month"] = df["order_date"].dt.month_name()
df["weekday"] = df["order_date"].dt.day_name()
df["is_weekend"] = df["order_date"].dt.dayofweek >= 5
df["price_band"] = pd.cut(df["unit_price"], bins=[0, 5000, 30000, np.inf], right=False,
                          labels=["Budget", "Mid", "Premium"]).astype(str)
df[["order_id","revenue","month","weekday","is_weekend","price_band"]].head()''')
prem = ", ".join(sorted(df.loc[df["price_band"]=="Premium","product"].unique()))
md(f"`revenue = quantity × unit_price`. `price_band` splits `unit_price` into **Budget** (below KES 5,000), **Mid** (KES 5,000 to 29,999) and **Premium** (KES 30,000 and above); Premium here means: {prem}.")
md("## Task 3 · Branch ranking: revenue vs orders")
code('''b = df.groupby("branch").agg(total_revenue=("revenue","sum"), orders=("order_id","count"))
b["avg_order_value"] = b.total_revenue/b.orders
b["rank_revenue"] = b.total_revenue.rank(ascending=False).astype(int)
b["rank_orders"] = b.orders.rank(ascending=False, method="min").astype(int)
b.sort_values("total_revenue", ascending=False)''')
r = list(br_rev.index)
swap = next(((r[a], r[b]) for a in range(len(r)) for b in range(a+1, len(r)) if br.orders[r[a]] < br.orders[r[b]]), None)
if same:
    expl = (f"The two rankings are identical: order volume and revenue move together because average order values are close across branches "
            f"({K(br.aov.min())} to {K(br.aov.max())}), so no branch wins on revenue through pricier baskets alone.")
else:
    lead = (f"**{bl}** is first on both lists, but the order changes further down." if bl == bo
            else f"**{bl}** leads on revenue but **{bo}** leads on orders.")
    x, y = swap
    expl = (f"The rankings differ. {lead} For example, **{x}** out-earns **{y}** ({K(br.rev[x])} vs {K(br.rev[y])}) "
            f"despite fewer orders ({br.orders[x]:,} vs {br.orders[y]:,}), because revenue = orders × average order value "
            f"and {x}'s average order is larger ({K(br.aov[x])} vs {K(br.aov[y])}).")
md(f"**Answer.** By revenue: {' > '.join(br_rev.index)}. By orders: {' > '.join(br_ord.index)}. {expl}")
md("## Task 4 · Top three products")
code('''p = df.groupby("product").agg(revenue=("revenue","sum"), units_sold=("quantity","sum"))
print("By revenue:"); display(p.nlargest(3,"revenue"))
print("By units sold:");   display(p.nlargest(3,"units_sold"))''')
dif = set(top_rev.index) ^ set(top_u.index)
if list(top_rev.index) == list(top_u.index):
    why4 = "The two lists are identical, so the best sellers by volume are also the best earners."
elif not dif:
    ru, uu = top_rev.index[0], top_u.index[0]
    why4 = (f"The same three products appear in both lists but in a different order: **{uu}** sells the most units ({top_u.units.iloc[0]:,}) "
            f"while **{ru}** earns the most revenue ({K(top_rev.rev.iloc[0])}) because revenue = units × price and {ru} costs "
            f"{K(df.loc[df['product']==ru,'unit_price'].iloc[0])} against {K(df.loc[df['product']==uu,'unit_price'].iloc[0])} for {uu}.")
else:
    why4 = "The lists differ because revenue rewards price × volume: expensive items earn heavily from fewer units, while cheap accessories sell many units but contribute little revenue."
md(f"**Answer.** Top by revenue: {', '.join(top_rev.index)}. Top by units: {', '.join(top_u.index)}. " + why4)
md("## Task 5 · Weekend vs weekday")
code('''w = df.groupby("is_weekend").agg(orders=("order_id","count"), total_revenue=("revenue","sum"), avg_order_revenue=("revenue","mean"))
w.index = w.index.map({True:"Weekend", False:"Weekday"}); w''')
gap5 = (wk.aov[True]/wk.aov[False]-1)*100
share5 = wk.orders[True]/len(df)*100
read5 = (f"average order values are virtually the same ({gap5:+.1f}%), and weekends hold {share5:.1f}% of orders against the {2/7*100:.1f}% you would expect from 2 of 7 days, so there is no real weekend effect."
         if abs(gap5) < 2 else
         f"weekend orders are {'larger' if gap5>0 else 'smaller'} on average ({gap5:+.1f}%).")
md(f"**Answer.** Weekdays: {wk.orders[False]:,} orders, {K(wk.rev[False])}, avg {K(wk.aov[False])}. Weekends: {wk.orders[True]:,} orders, {K(wk.rev[True])}, avg {K(wk.aov[True])}. Weekends cover only 2 of 7 days, so compare the average order value, not totals: {read5}")
md("## Task 6 · Best and worst month")
code('''m = df.groupby(df.order_date.dt.month).revenue.sum()
m.index = [pd.Timestamp(2026,i,1).month_name() for i in m.index]
best, worst = m.idxmax(), m.idxmin()
print(m); print(f"best={best} worst={worst} difference={(m[best]-m[worst])/m[worst]*100:.1f}% (best vs worst)")''')
md(f"**Answer.** Best: **{mo.name[best]}** ({K(mo.rev[best])}); worst: **{mo.name[worst]}** ({K(mo.rev[worst])}). Best is **{pct:.1f}%** higher than worst, i.e. worst is {(mo.rev[best]-mo.rev[worst])/mo.rev[best]*100:.1f}% below best." + (f" Caveat: the data ends on {de}, so {mo.name[worst] if pd.Timestamp(de).month==worst else pd.Timestamp(de).month_name()} is missing its last day(s)." if pd.Timestamp(de).day != pd.Timestamp(de).days_in_month else ""))
md("## Task 7 · Premium online orders, Nairobi & Mombasa")
code('''sel = df[(df.price_band=="Premium") & (df.channel=="Online") & (df.branch.isin(["Nairobi","Mombasa"]))]
print(len(sel), "orders | revenue KES", f"{sel.revenue.sum():,.0f}")
sel.to_csv("outputs/premium_online_coast_capital.csv", index=False)''')
md(f"**Answer.** **{len(sel):,}** orders, total revenue **{K(sel.revenue.sum())}**. Exported to `outputs/premium_online_coast_capital.csv` (`index=False`).")
md("## Task 8 · Excel reconciliation (`excel/duka_branch_summary.xlsx`)")
code('''xl = pd.read_csv("excel/sumifs_recalc_check.csv").set_index("branch")      # SUMIFS/COUNTIFS values recalculated from the workbook
cmp = b[["total_revenue","orders"]].join(xl)
cmp["match"] = np.isclose(cmp.total_revenue, cmp.sumifs_revenue) & (cmp.orders == cmp.countifs_orders)
raw = pd.read_csv("excel/pivot_recalc_check.csv", header=None)         # values of the real PivotTable in the workbook
hdr_i = raw.index[raw.iloc[:,0].astype(str).str.contains("branch", case=False)][-1]
pv = raw.iloc[hdr_i+1:].copy(); pv.columns = ["branch"]+list(raw.iloc[hdr_i,1:]); pv = pv.set_index("branch")
pd_pivot = df.pivot_table(index="branch", columns="channel", values="revenue", aggfunc="sum")
ok = np.allclose(pv.loc[pd_pivot.index, pd_pivot.columns].astype(float), pd_pivot)
tot = pv.loc[pv.index.str.contains("Total", case=False)].iloc[0].astype(float).iloc[-1]
print("SUMIFS = pandas for every branch:", cmp.match.all(), "| grand total:", f"{xl.sumifs_revenue.sum():,.0f} vs {df.revenue.sum():,.0f}")
print("PivotTable branch×channel = pandas pivot:", ok, "| pivot grand total:", f"{tot:,.0f}")
cmp''')
md("**Answer.** The workbook has `Branch_Summary` (SUMIFS revenue + COUNTIFS orders + XLOOKUP region from the `Regions` sheet) and `Pivot_Branch_Channel` (a real PivotTable: branch rows × channel columns, sum of revenue). The check above shows every total equals pandas.")
nb["cells"] = C; nbf.write(nb, "duka_sales_exploration.ipynb")
subprocess.run([sys.executable,"-m","jupyter","nbconvert","--to","notebook","--execute","--inplace","duka_sales_exploration.ipynb"],capture_output=True)
exec_sum = C[1].source
r = open("README.md").read()
head = r.split("## Duka Executive Summary")[0]
open("README.md","w").write(head+"## Duka Executive Summary\n"+exec_sum.replace("## Executive Summary\n","")+("\n\n> Numbers above come from the stand-in generator; paste the course generator into `generate_data.py` and re-run `python generate_data.py && python build_duka.py`.\n" if STANDIN else "\n"))
print(exec_sum)
