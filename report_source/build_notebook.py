"""Generate the (unexecuted) Case Study 1 notebook.

Usage:  python report_source/build_notebook.py
Writes submission/CSE464_Case_Study_1.ipynb. Execute it afterwards with
report_source/run_notebook.py (or Runtime > Run all in Google Colab).
"""
from pathlib import Path
import nbformat as nbf

ROOT = Path(__file__).resolve().parents[1]
OUT_NB = ROOT / "submission" / "CSE464_Case_Study_1.ipynb"

cells = []


def md(text):
    cells.append(nbf.v4.new_markdown_cell(text.strip("\n")))


def code(text):
    cells.append(nbf.v4.new_code_cell(text.strip("\n")))


# ---------------------------------------------------------------------------
md(r"""
# CSE464 Case Study 1 – Sales KPI Review for a UK Online Gift Retailer

**Course:** CSE464 – Introduction to Data Science & Big Data Analytics (Fall 2026)
**Student:** Kenan Saleh (20220702115)
**Dataset:** UCI Online Retail, ID 352 (Chen, 2015), DOI [10.24432/C5BW33](https://doi.org/10.24432/C5BW33)
**Selected issue:** 1(b)(ii) – sales figures / company KPIs are below expectations.

**Time-series question.** *Across complete calendar months, when did net invoiced sales fall most relative to the
preceding month, which countries accounted for the change, and what should management investigate first?*

"Fall most" = the largest absolute month-to-month decline in GBP. The percentage change is reported separately.
The preceding complete month is used only as a **historical benchmark** – the company's real targets, costs and
margins are not in the data, so invoiced sales are never treated as profit.

The notebook follows the nine steps of the course hand-out. It runs top to bottom in a fresh Google Colab
runtime (Runtime ▸ Run all). It needs no Drive mount; it downloads the data from UCI and writes charts and a
results file to a relative `outputs/` folder.
""")

code(r"""
# Setup: install the official UCI import package if it is missing (Colab does not ship it).
import importlib.util, subprocess, sys
if importlib.util.find_spec("ucimlrepo") is None:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "ucimlrepo"], check=True)

import json, platform
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

OUT = Path("outputs"); FIG = OUT / "figures"
FIG.mkdir(parents=True, exist_ok=True)
pd.set_option("display.width", 160); pd.set_option("display.max_columns", 30)
pd.set_option("display.float_format", lambda v: f"{v:,.2f}")

# Chart style: light print palette, thin marks, recessive grid
BLUE, BLUE_LIGHT, ORANGE, RED, GRAY = "#2a78d6", "#9ec5f4", "#eb6834", "#e34948", "#898781"
INK, INK2, GRID = "#0b0b0b", "#52514e", "#e1e0d9"
plt.rcParams.update({
    "figure.dpi": 110, "savefig.dpi": 200, "font.size": 10, "axes.titlesize": 11,
    "axes.edgecolor": "#c3c2b7", "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "axes.axisbelow": True,
    "axes.spines.top": False, "axes.spines.right": False, "legend.frameon": False,
})
def money(v, dp=0):
    # GBP with the sign in front: money(-1234.5) -> '−£1,235'
    return f"{'−' if v < 0 else ''}£{abs(v):,.{dp}f}"

gbp_k = FuncFormatter(lambda v, _: f"{'−' if v < 0 else ''}£{abs(v)/1000:,.0f}k")

R = {}  # every number used in the report is stored here and saved to outputs/results.json

def save(fig, name):
    fig.savefig(FIG / f"{name}.png", bbox_inches="tight", facecolor="white")
    plt.show()

print("Python", platform.python_version(), "| pandas", pd.__version__, "| NumPy", np.__version__,
      "| matplotlib", matplotlib.__version__)
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 1 – Load the data

The hand-out asks for the UC Irvine Python import, so the primary route is `fetch_ucirepo(id=352)`.
The loader checks that both ID fields (`InvoiceNo`, `StockCode`) and all six feature columns came back.
If the package fails or omits a field, it falls back to UCI's official Excel workbook
(`online+retail.zip`), and if the network is unavailable it reads a local copy of that workbook
(`data/Online Retail.xlsx`). No synthetic or sample data is ever used.
""")

code(r"""
EXPECTED = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]
WORKBOOK_URL = "https://archive.ics.uci.edu/static/public/352/online+retail.zip"
LOCAL_WORKBOOK = Path("data") / "Online Retail.xlsx"

def load_online_retail(use_ucimlrepo=True):
    # Returns (raw DataFrame, source description, metadata dict or None)
    if use_ucimlrepo:
        try:
            from ucimlrepo import fetch_ucirepo
            repo = fetch_ucirepo(id=352)
            print("Returned object keys:", list(repo.keys()))
            print("repo.data keys:     ", list(repo.data.keys()))
            print("ids columns:        ", list(repo.data.ids.columns))
            print("features columns:   ", list(repo.data.features.columns))
            print("targets:            ", repo.data.targets)
            original = repo.data.original
            missing = [c for c in EXPECTED if c not in original.columns]
            if missing:
                raise ValueError(f"fetch_ucirepo omitted essential fields: {missing}")
            return original[EXPECTED].copy(), f"fetch_ucirepo(id=352) -> {repo.metadata['data_url']}", repo.metadata
        except Exception as err:
            print("Official Python import failed, using the UCI workbook instead:", repr(err))
    try:
        import io, urllib.request, zipfile
        with urllib.request.urlopen(WORKBOOK_URL, timeout=180) as resp:
            archive = zipfile.ZipFile(io.BytesIO(resp.read()))
        xlsx_name = [n for n in archive.namelist() if n.lower().endswith(".xlsx")][0]
        LOCAL_WORKBOOK.parent.mkdir(exist_ok=True)
        LOCAL_WORKBOOK.write_bytes(archive.read(xlsx_name))
        source = f"UCI official workbook {WORKBOOK_URL} ({xlsx_name})"
    except Exception as err:
        print("Workbook download failed:", repr(err))
        if not LOCAL_WORKBOOK.exists():
            raise FileNotFoundError(f"No network access and no local copy at {LOCAL_WORKBOOK}. "
                                    "Download online+retail.zip from the UCI page and unzip it there.")
        source = f"local copy of the UCI workbook: {LOCAL_WORKBOOK}"
    # dtype=str keeps alphanumeric codes such as 'C536379' and '85123A' intact
    frame = pd.read_excel(LOCAL_WORKBOOK, dtype={"InvoiceNo": str, "StockCode": str})
    return frame[EXPECTED].copy(), source, None

raw, SOURCE, meta = load_online_retail()
# Keep the identifiers as text exactly as published (strip stray spaces only)
for col in ["InvoiceNo", "StockCode"]:
    raw[col] = raw[col].astype(str).str.strip()

if meta is not None:
    print("\nDataset:", meta["name"], "| UCI id:", meta["uci_id"], "| DOI:", meta["dataset_doi"])
    print("Characteristics:", meta["characteristics"], "| instances listed by UCI:", meta["num_instances"])
    print("Summary:", meta["additional_info"]["summary"])
print("\nSource used:", SOURCE)
print("Raw rows (invoice lines):", f"{len(raw):,}", "| columns:", list(raw.columns))
_dates = (raw["InvoiceDate"] if pd.api.types.is_datetime64_any_dtype(raw["InvoiceDate"])
          else pd.to_datetime(raw["InvoiceDate"], format="%m/%d/%Y %H:%M"))
print("Date range:", _dates.min(), "to", _dates.max())
R["source"] = SOURCE
R["raw_rows"] = int(len(raw))
R["date_min"], R["date_max"] = str(_dates.min()), str(_dates.max())
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 2 – Understand the dataset

One row is an **invoice line**: one product (`StockCode`) on one invoice (`InvoiceNo`). An invoice can contain many
lines, so line counts and invoice counts are always reported separately.
""")

code(r"""
print("Shape (rows, columns):", raw.shape)
display(raw.head(5))
print("\ndata.dtypes")
print(raw.dtypes)
""")

code(r"""
print("data.describe() – numeric columns as loaded")
display(raw.describe())
print("data.describe() – text columns")
display(raw.select_dtypes(exclude="number").describe())
""")

code(r"""
field_info = pd.DataFrame([
    ["InvoiceNo",   "Identifier (text)",      "Invoice number; a leading 'C' marks a cancellation/credit note"],
    ["StockCode",   "Identifier (text)",      "Product code; a few codes are services (POST, DOT, M, AMAZONFEE ...)"],
    ["Description", "Category (text)",        "Product name"],
    ["Quantity",    "Numeric measure (units)", "Units on the line; negative on credit notes"],
    ["InvoiceDate", "Date-time",              "Day and time the invoice was generated"],
    ["UnitPrice",   "Numeric measure (GBP)",  "Price per unit in pounds sterling"],
    ["CustomerID",  "Identifier (stored as a number)", "Customer number; blank for some transactions"],
    ["Country",     "Category (text)",        "Country where the customer resides"],
], columns=["Field", "Type", "Meaning (UCI variable information)"])
display(field_info)

missing = pd.DataFrame({"missing": raw.isna().sum(), "missing_%": (raw.isna().mean() * 100).round(2)})
display(missing)
lines_per_invoice = raw.groupby("InvoiceNo").size()
print(f"Invoice lines: {len(raw):,} | unique invoices: {raw['InvoiceNo'].nunique():,} | "
      f"unique products: {raw['StockCode'].nunique():,} | known customers: {raw['CustomerID'].nunique():,} | "
      f"countries: {raw['Country'].nunique()}")
print(f"Lines per invoice: median {lines_per_invoice.median():.0f}, mean {lines_per_invoice.mean():.1f}, max {lines_per_invoice.max():,}")
R["missing_customer"] = int(raw["CustomerID"].isna().sum())
R["missing_description"] = int(raw["Description"].isna().sum())
R["raw_invoices"] = int(raw["InvoiceNo"].nunique())
R["raw_countries"] = int(raw["Country"].nunique())
R["lines_per_invoice_median"] = float(lines_per_invoice.median())
""")

md(r"""
**Observations.** `InvoiceDate` is loaded as text and must be parsed. `CustomerID` is an identifier stored as a
float, so its mean/std in `describe()` have no meaning. `Quantity` and `UnitPrice` both have negative minimums,
which points to credit notes and adjustments (Step 3). UCI's metadata says "no missing values", but the actual
counts show missing `CustomerID` and `Description` values, so the real counts are used.
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 3 – Clean the data

Checks: repeated header, data types, missing values, repeated rows, cancellations, negative quantities, unusual
prices and non-merchandise codes. Every row gets a single, first-matching exclusion reason so that
**raw = retained + excluded** can be verified.
""")

code(r"""
df = raw.copy()
# 3.1 Repeated header rows (a header line accidentally stored as data)
header_rows = (df["InvoiceNo"] == "InvoiceNo").sum()
print("Repeated header rows found:", header_rows, "-> nothing to remove" if header_rows == 0 else "")
df = df[df["InvoiceNo"] != "InvoiceNo"].copy()

# 3.2 Parse types; codes stay as text
if not pd.api.types.is_datetime64_any_dtype(df["InvoiceDate"]):
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], format="%m/%d/%Y %H:%M")
df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")
df["UnitPrice"] = pd.to_numeric(df["UnitPrice"], errors="coerce")
df["CustomerID"] = pd.to_numeric(df["CustomerID"], errors="coerce").astype("Int64")
for col in ["Description", "Country"]:
    df[col] = df[col].astype("string").str.strip()
print("Unparseable dates / quantities / prices:",
      int(df["InvoiceDate"].isna().sum()), int(df["Quantity"].isna().sum()), int(df["UnitPrice"].isna().sum()))
print(df.dtypes)
""")

code(r"""
# 3.3 Missing values: where are they?
print("Missing Description rows by UnitPrice:", df.loc[df["Description"].isna(), "UnitPrice"].value_counts().to_dict())
print(f"Missing CustomerID rows: {df['CustomerID'].isna().sum():,} (kept if the sales fields are valid)")

# 3.4 Repeated rows: identical in all 8 fields
dup_mask = df.duplicated(keep="first")
print(f"Exact repeated rows (all 8 fields identical, excluding first copy): {dup_mask.sum():,}")
print("Kept: the data has no line number, so two identical lines on the same invoice can be a real double order."
      " Step 6A tests whether removing them changes the conclusion.")

# 3.5 Cancellations and negative quantities
df["IsCredit"] = df["InvoiceNo"].str.startswith("C")
prefix = df["InvoiceNo"].str[0].where(~df["InvoiceNo"].str[0].str.isdigit(), "digit")
print("\nInvoice number prefix counts:", prefix.value_counts().to_dict())
print("Credit-note lines:", f"{df['IsCredit'].sum():,}", "| of which Quantity > 0:", int((df["IsCredit"] & (df["Quantity"] > 0)).sum()))
neg_normal = (~df["IsCredit"]) & (df["Quantity"] < 0)
print(f"Non-credit lines with negative Quantity: {neg_normal.sum():,}; their UnitPrice values:",
      df.loc[neg_normal, "UnitPrice"].value_counts().to_dict())

# 3.6 Unusual prices
print(f"\nUnitPrice == 0: {(df['UnitPrice'] == 0).sum():,} | UnitPrice < 0: {(df['UnitPrice'] < 0).sum():,}")
display(df[df["UnitPrice"] < 0][["InvoiceNo", "StockCode", "Description", "Quantity", "UnitPrice", "InvoiceDate"]])
print("Highest unit prices:")
display(df.nlargest(5, "UnitPrice")[["InvoiceNo", "StockCode", "Description", "Quantity", "UnitPrice"]])

# 3.7 Extreme quantities: are the huge orders reversed by a matching credit note?
big = df[df["Quantity"].abs() >= 10000][["InvoiceNo", "StockCode", "Description", "Quantity", "UnitPrice", "InvoiceDate", "CustomerID"]]
big = big.assign(LineValue=big["Quantity"] * big["UnitPrice"])
display(big)
print("The two priced extreme orders (Jan 2011, Dec 2011) are each cancelled by a credit note for the same product, customer"
      " and quantity within minutes, so each pair nets to zero inside its month. Both sides are kept: removing only one"
      " side would distort net sales. The 12,540-unit line has a zero price and is excluded by rule R1 below.")
R["extreme_orders"] = [{"invoice": r.InvoiceNo, "qty": int(r.Quantity), "value": float(r.LineValue),
                        "date": str(r.InvoiceDate.date())} for r in big.itertuples()]
""")

code(r"""
# 3.8 Non-merchandise codes: postage, fees, manual entries, discounts, samples, bad-debt and gift vouchers
NON_MERCH = {"POST", "DOT", "C2", "M", "m", "D", "S", "BANK CHARGES", "AMAZONFEE", "CRUK", "B"}
is_non_merch = df["StockCode"].isin(NON_MERCH) | df["StockCode"].str.startswith("gift_0001")
df["LineValue"] = df["Quantity"] * df["UnitPrice"]
nm = (df[is_non_merch].groupby("StockCode")
        .agg(lines=("LineValue", "size"), value_gbp=("LineValue", "sum"), example=("Description", "first"))
        .sort_values("lines", ascending=False))
display(nm)

# 3.9 Eligibility rules (first failing rule is the exclusion reason)
valid_sign = ((~df["IsCredit"]) & (df["Quantity"] > 0)) | (df["IsCredit"] & (df["Quantity"] < 0))
df["Exclusion"] = np.select(
    [df["UnitPrice"].isna() | df["Quantity"].isna() | df["InvoiceDate"].isna(),
     df["UnitPrice"] <= 0,
     ~valid_sign,
     is_non_merch],
    ["R0 unparseable field",
     "R1 zero/negative unit price (stock adjustments, bad-debt write-offs)",
     "R2 quantity sign does not match invoice type",
     "R3 non-merchandise code (postage, fees, manual, discount, samples, vouchers)"],
    default="Retained")
recon = df["Exclusion"].value_counts().rename_axis("Rule").to_frame("Invoice lines")
display(recon)

eligible = df[df["Exclusion"] == "Retained"].copy()
n_raw, n_ret = len(raw), len(eligible)
n_exc = int((df["Exclusion"] != "Retained").sum()) + int(header_rows)
assert n_raw == n_ret + n_exc, "Row reconciliation failed"
print(f"Raw {n_raw:,} = retained {n_ret:,} + excluded {n_exc:,}  ✔")
print(f"Retained lines with missing CustomerID: {eligible['CustomerID'].isna().sum():,} "
      f"(GBP {eligible.loc[eligible['CustomerID'].isna(), 'LineValue'].sum():,.0f}) – kept for the aggregate KPI")
print(f"Retained credit-note lines: {eligible['IsCredit'].sum():,} – kept so that returns reduce net sales")
cred_by_rule = df.loc[df["IsCredit"], "Exclusion"].value_counts()
print(f"Credit-note lines: raw {df['IsCredit'].sum():,} = retained {cred_by_rule.get('Retained', 0):,} + excluded "
      f"{df['IsCredit'].sum() - cred_by_rule.get('Retained', 0):,}; by rule:", cred_by_rule.to_dict())
assert df["IsCredit"].sum() == cred_by_rule.sum()
R.update(retained_rows=n_ret, excluded_rows=n_exc,
         excluded_by_rule={k: int(v) for k, v in recon["Invoice lines"].items() if k != "Retained"},
         duplicate_rows=int(dup_mask.sum()), credit_lines_raw=int(df["IsCredit"].sum()),
         retained_credit_lines=int(eligible["IsCredit"].sum()),
         excluded_credit_lines=int(df["IsCredit"].sum() - eligible["IsCredit"].sum()),
         credit_excluded_by_rule={k[:2]: int(v) for k, v in cred_by_rule.items() if k != "Retained"},
         retained_missing_customer=int(eligible["CustomerID"].isna().sum()),
         zero_price=int((df["UnitPrice"] == 0).sum()), negative_price=int((df["UnitPrice"] < 0).sum()),
         neg_qty_noncredit=int(neg_normal.sum()), non_merch_rows=int((df["Exclusion"].str.startswith("R3")).sum()))
""")

md(r"""
**Eligibility rules (why).**
- **R1** – lines with a zero or negative price are not sales. All 1,454 lines with a missing description, and all
  non-credit lines with a negative quantity, have a price of 0 (they are stock adjustments). The two negative-price
  lines are "Adjust bad debt" accounting entries.
- **R2** – a normal invoice must have a positive quantity and a credit note a negative one. In this data no line
  breaks this rule once R1 is applied, but the rule is kept as a guard.
- **R3** – postage, carriage, bank and Amazon fees, manual entries, discounts, samples, bad-debt adjustments and
  gift vouchers are not merchandise sales. They are excluded from the KPI scope. Step 6A tests putting them back.
- **Kept:** credit notes (they reduce net sales once), lines with no `CustomerID` (the sale is still valid for
  the monthly total) and identical repeated lines (they cannot be proven to be errors).
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 4 – Create the selected DataFrame and analyse the target variable Y

**Y = monthly net invoiced sales (GBP)** for merchandise lines that pass R1–R3.
For each line, `LineValue = Quantity × UnitPrice` (signed: credit lines are negative).
Per calendar month (by `InvoiceDate`):

* positive sales = sum of positive line values
* credit magnitude = −(sum of negative line values) – each credit line counted once, in the month it was issued
* **net sales = positive sales − credit magnitude = sum of all signed line values**

Three levels of observation are used: invoice **line** (row), **invoice** (all lines with one `InvoiceNo`) and
**month** (all invoices in a calendar month).
""")

code(r"""
sales = eligible[["InvoiceNo", "StockCode", "InvoiceDate", "Country", "CustomerID",
                  "Quantity", "UnitPrice", "LineValue", "IsCredit"]].copy()
sales["Month"] = sales["InvoiceDate"].dt.to_period("M")
sales["Day"] = sales["InvoiceDate"].dt.normalize()
sales["Region"] = np.where(sales["Country"] == "United Kingdom", "United Kingdom", "International")

# A month is complete only if the data covers it from its first to its last day
cover_start, cover_end = sales["InvoiceDate"].min().normalize(), sales["InvoiceDate"].max()
def is_complete(p):
    return (p.start_time >= cover_start) and (p.end_time <= cover_end)

def monthly_table(frame):
    g = frame.groupby("Month")
    t = pd.DataFrame({
        "lines": g.size(),
        "sales_invoices": frame[~frame["IsCredit"]].groupby("Month")["InvoiceNo"].nunique(),
        "credit_invoices": frame[frame["IsCredit"]].groupby("Month")["InvoiceNo"].nunique(),
        "invoice_days": g["Day"].nunique(),
        "known_customers": g["CustomerID"].nunique(),
        "positive_sales": g["LineValue"].apply(lambda s: s[s > 0].sum()),
        "credit_magnitude": g["LineValue"].apply(lambda s: -s[s < 0].sum()),
        "net_sales": g["LineValue"].sum(),
    })
    t[["sales_invoices", "credit_invoices"]] = t[["sales_invoices", "credit_invoices"]].fillna(0).astype(int)
    t["net_per_invoice_day"] = t["net_sales"] / t["invoice_days"]
    t["complete"] = [is_complete(p) for p in t.index]
    return t

monthly = monthly_table(sales)
assert np.allclose(monthly["positive_sales"] - monthly["credit_magnitude"], monthly["net_sales"])
assert np.isclose(monthly["net_sales"].sum(), sales["LineValue"].sum())
assert monthly["lines"].sum() == len(sales)
print("Checks passed: positive − credit = net for every month; monthly totals = sum of eligible signed line values.")
print(f"Coverage: {cover_start.date()} to {cover_end}")
display(monthly)
""")

code(r"""
complete = monthly[monthly["complete"]]
print("Partial months (excluded from month-to-month ranking):", [str(p) for p in monthly.index[~monthly["complete"]]])
print(f"Observations: {len(sales):,} invoice lines -> {sales['InvoiceNo'].nunique():,} invoices "
      f"({sales.loc[~sales['IsCredit'], 'InvoiceNo'].nunique():,} sales invoices + "
      f"{sales.loc[sales['IsCredit'], 'InvoiceNo'].nunique():,} credit notes) -> {len(monthly)} months ({len(complete)} complete)")
y = complete["net_sales"]
print("\nY over the 12 complete months (GBP):")
print(y.describe().apply(lambda v: f"{v:,.0f}"))
print(f"Lowest: {y.idxmin()} £{y.min():,.0f} | Highest: {y.idxmax()} £{y.max():,.0f} | "
      f"coefficient of variation {y.std() / y.mean() * 100:.1f}%")
R.update(complete_months=[str(p) for p in complete.index], n_complete=int(len(complete)),
         y_mean=float(y.mean()), y_median=float(y.median()), y_min=float(y.min()), y_min_month=str(y.idxmin()),
         y_max=float(y.max()), y_max_month=str(y.idxmax()), y_cv=float(y.std() / y.mean() * 100),
         y_total_complete=float(y.sum()), eligible_invoices=int(sales["InvoiceNo"].nunique()),
         eligible_sales_invoices=int(sales.loc[~sales["IsCredit"], "InvoiceNo"].nunique()),
         eligible_credit_invoices=int(sales.loc[sales["IsCredit"], "InvoiceNo"].nunique()),
         monthly={str(k): {c: (float(v) if not isinstance(v, (bool, np.bool_)) else bool(v)) for c, v in row.items()}
                  for k, row in monthly.iterrows()})
monthly.to_csv(OUT / "monthly_kpi.csv")
""")

md(r"""
**Y at a glance.** Over the 12 complete months, monthly net sales range from about £0.48m (April 2011) to
£1.43m (November 2011). The strong rise from September to November gives Y a wide spread (CV ≈ 35%).
December 2011 ends on 9 December, so it is flagged as partial and is never compared with a full month.
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 5 – Plot the data: dependent (Y) and independent (X) variables

| Variable | Role | Type | Unit | Aggregation level |
|---|---|---|---|---|
| Net invoiced sales | **Y** (dependent) | continuous numeric | GBP | calendar month |
| Month (`InvoiceDate`) | X – time index | date-time → ordinal period | month | month |
| Country / Region (UK vs International) | X – grouping | nominal category | – | month × country |
| Sales invoices | X – volume driver | count | invoices | month |
| Invoice value | X – order size | continuous numeric | GBP | invoice |
| Invoice days | X – coverage | count | days with ≥1 eligible line | month |
| Credit magnitude | X – returns | continuous numeric | GBP | month |
""")

code(r"""
labels = [p.strftime("%b\n%Y") if p.month in (12, 1) else p.strftime("%b") for p in monthly.index]
x = np.arange(len(monthly))
colors = [GRAY if not c else BLUE_LIGHT for c in monthly["complete"]]
for i, p in enumerate(monthly.index):
    if str(p) in ("2011-03", "2011-04"):
        colors[i] = BLUE
fig, ax = plt.subplots(figsize=(8, 3.7))
bars = ax.bar(x, monthly["net_sales"], color=colors, width=0.7, edgecolor="white", linewidth=1)
bars[-1].set_hatch("///"); bars[-1].set_edgecolor("white")
ax.set_xticks(x, labels); ax.yaxis.set_major_formatter(gbp_k)
ax.set_ylabel("Net invoiced sales (GBP)"); ax.set_xlabel("Calendar month (Dec 2010 – Dec 2011)")
ax.set_title("Monthly net invoiced sales, merchandise lines", loc="left")
for i, p in enumerate(monthly.index):
    if str(p) in ("2011-03", "2011-04", "2011-11"):
        ax.text(i, monthly["net_sales"].iloc[i] + 15000, f"£{monthly['net_sales'].iloc[i]/1000:,.0f}k",
                ha="center", va="bottom", fontsize=9, color=INK)
ax.text(len(x) - 1, monthly["net_sales"].iloc[-1] + 15000, f"£{monthly['net_sales'].iloc[-1]/1000:,.0f}k\npartial:\n1–9 Dec",
        ha="center", va="bottom", fontsize=8.5, color=INK, linespacing=1.1)
ax.set_ylim(0, monthly["net_sales"].max() * 1.12)
save(fig, "fig1_monthly_net_sales")
m, a = monthly.loc["2011-03", "net_sales"], monthly.loc["2011-04", "net_sales"]
print(f"Interpretation: net sales fell from £{m:,.0f} in March 2011 to £{a:,.0f} in April 2011, the lowest complete month; "
      f"the peak was £{monthly.loc['2011-11', 'net_sales']:,.0f} in November 2011.")
""")

code(r"""
reg = (sales[sales["Month"].isin(complete.index)]
       .pivot_table(index="Month", columns="Region", values="LineValue", aggfunc="sum"))
fig, ax = plt.subplots(figsize=(9, 3.6))
xc = np.arange(len(reg))
for name, col in [("United Kingdom", BLUE), ("International", ORANGE)]:
    ax.plot(xc, reg[name], color=col, lw=2, marker="o", ms=4, label=name)
    ax.text(xc[-1] + 0.15, reg[name].iloc[-1], name, va="center", fontsize=9, color=INK)
ax.set_xticks(xc, [p.strftime("%b\n%Y") if p.month in (12, 1) else p.strftime("%b") for p in reg.index])
ax.yaxis.set_major_formatter(gbp_k); ax.set_xlim(-0.4, len(xc) + 1.2)
ax.set_ylabel("Net invoiced sales (GBP)"); ax.set_xlabel("Complete calendar month")
ax.set_title("Net sales by region, complete months (Dec 2010 – Nov 2011)", loc="left")
ax.legend(loc="upper left")
save(fig, "fig2_region_monthly")
reg["intl_share_%"] = reg["International"] / (reg["United Kingdom"] + reg["International"]) * 100
display(reg)
print(f"Interpretation: the UK is {100 - reg['intl_share_%'].mean():.0f}% of complete-month net sales on average. "
      f"International net sales were £{reg.loc['2011-03', 'International']:,.0f} in March 2011 and "
      f"£{reg.loc['2011-04', 'International']:,.0f} in April 2011, their lowest month.")
R.update(intl_share_avg=float(reg["intl_share_%"].mean()),
         intl_share_mar=float(reg.loc["2011-03", "intl_share_%"]), intl_share_apr=float(reg.loc["2011-04", "intl_share_%"]),
         intl_min_month=str(reg["International"].idxmin()))
""")

code(r"""
fig, ax = plt.subplots(figsize=(6.2, 3.8))
ax.scatter(complete["sales_invoices"], complete["net_sales"], s=40, color=BLUE, edgecolor="white", linewidth=1.5, zorder=3)
for p, row in complete.iterrows():
    ax.annotate(p.strftime("%b %y"), (row["sales_invoices"], row["net_sales"]), xytext=(5, 3),
                textcoords="offset points", fontsize=8, color=INK2)
ax.yaxis.set_major_formatter(gbp_k)
ax.set_xlabel("Sales invoices in the month (count)"); ax.set_ylabel("Net invoiced sales (GBP)")
ax.set_title("Monthly order volume (X) vs net sales (Y)", loc="left")
save(fig, "fig3_invoices_vs_sales")
r_inv = np.corrcoef(complete["sales_invoices"], complete["net_sales"])[0, 1]
print(f"Pearson correlation, sales invoices vs net sales (12 complete months): r = {r_inv:.2f}")
print(f"Interpretation: April 2011 had {complete.loc['2011-04', 'sales_invoices']:,} sales invoices vs "
      f"{complete.loc['2011-03', 'sales_invoices']:,} in March 2011; months with more orders have higher net sales.")
R["r_invoices_sales"] = float(r_inv)
""")

md(r"""
**Comments by variable type.**
*Time (ordinal):* the months form a short series (12 complete points), so seasonality cannot be separated from
trend with only one year. *Country (nominal):* one category, the United Kingdom, dominates, so international
movements look small in total but can be large in relative terms (region chart above). *Counts (discrete):* the number of
sales invoices moves almost one-for-one with net sales (scatter chart above), which shows that order volume drives Y.
*GBP measures (continuous, right-skewed):* invoice values have a long upper tail (Step 6B).
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 6 – Select features and apply three descriptive methods

| Field | Kept? | Why it matters for the decision |
|---|---|---|
| `InvoiceDate` → Month, Day | yes | Defines the time series, complete months and invoice days |
| `Quantity`, `UnitPrice` → `LineValue` | yes | Build Y (signed GBP value) |
| `InvoiceNo` (+ `C` prefix) | yes | Separates sales from credits; groups lines into orders (method B) |
| `Country` → Region | yes | Explains *where* the change came from (method C) |
| `CustomerID` | yes (drill-down only) | Shows whether a country's change is one account or many; blanks kept |
| `StockCode` | scope filter only | Used to remove non-merchandise codes (R3) |
| `Description` | no | Free text; not needed for a monthly sales KPI |

The three methods are **A** time-series comparison, **B** distribution analysis of order values and
**C** country contribution analysis.
""")

md(r"""
### Method A – Time-series comparison of adjacent complete months

*Why:* the management question is about change over time, so we compare each complete month with the month before it.
Changes are given in GBP (used for ranking) and in % (= change ÷ previous month). If the previous month is zero or
negative, the % is undefined and shown as "n/a".
""")

code(r"""
def mom_changes(mt):
    c = mt[mt["complete"]].copy()
    c["prev_month"] = [str(p - 1) if (p - 1) in c.index else None for p in c.index]
    prev = c["net_sales"].shift(1).where([(p - 1) in c.index for p in c.index])
    c["change_gbp"] = c["net_sales"] - prev
    c["change_pct"] = np.where(prev > 0, c["change_gbp"] / prev.where(prev > 0) * 100, np.nan)  # zero/negative base -> n/a
    return c

mom = mom_changes(monthly)
show = mom[["prev_month", "net_sales", "change_gbp", "change_pct", "invoice_days", "net_per_invoice_day"]].copy()
show["change_pct"] = show["change_pct"].map(lambda v: "n/a" if pd.isna(v) else f"{v:+.1f}%")
display(show)

ranked = mom.dropna(subset=["change_gbp"]).sort_values("change_gbp")
worst = ranked.index[0]; prev_p = worst - 1
print("Largest declines (GBP):")
display(ranked[["prev_month", "change_gbp", "change_pct"]].head(3))
print(f"Largest absolute decline: {prev_p.strftime('%B %Y')} -> {worst.strftime('%B %Y')}: "
      f"{money(mom.loc[worst, 'change_gbp'])} ({mom.loc[worst, 'change_pct']:+.1f}%)")
print(f"December 2011 (1–9 Dec only, net £{monthly.loc['2011-12', 'net_sales']:,.0f}) is partial and NOT compared with November.")
R.update(worst_month=str(worst), prev_month=str(prev_p),
         worst_change=float(mom.loc[worst, "change_gbp"]), worst_pct=float(mom.loc[worst, "change_pct"]),
         second_month=str(ranked.index[1]), second_change=float(ranked["change_gbp"].iloc[1]),
         second_pct=float(ranked["change_pct"].iloc[1]),
         n_declines=int((ranked["change_gbp"] < 0).sum()), n_comparisons=int(len(ranked)),
         mom={str(k): {"change_gbp": float(v["change_gbp"]), "change_pct": float(v["change_pct"])}
              for k, v in ranked.iterrows()})
""")

code(r"""
fig, ax = plt.subplots(figsize=(9, 3.6))
ch = mom.dropna(subset=["change_gbp"])
ax.bar(np.arange(len(ch)), ch["change_gbp"], color=[RED if v < 0 else BLUE for v in ch["change_gbp"]],
       width=0.65, edgecolor="white")
ax.axhline(0, color="#c3c2b7", lw=1)
ax.set_xticks(np.arange(len(ch)), [f"{(p - 1).strftime('%b')}→{p.strftime('%b')}" for p in ch.index], fontsize=8.5)
ax.yaxis.set_major_formatter(gbp_k)
ax.set_ylabel("Change vs previous month (GBP)"); ax.set_xlabel("Adjacent complete months, Dec 2010 – Nov 2011")
ax.set_title("Month-to-month change in net sales (red = decline)", loc="left")
i = list(ch.index).index(worst)
ax.annotate(f"{money(ch.loc[worst, 'change_gbp']/1000, 1)}k ({ch.loc[worst, 'change_pct']:+.1f}%)", (i, ch.loc[worst, "change_gbp"]),
            xytext=(0, -14), textcoords="offset points", ha="center", fontsize=9, color=INK)
ax.set_ylim(ch["change_gbp"].min() * 1.35, ch["change_gbp"].max() * 1.15)
save(fig, "fig4_mom_change")
""")

code(r"""
# A.2 Coverage check: did April have fewer invoicing days than March?
a, b = monthly.loc[prev_p], monthly.loc[worst]
days_eff = (b["invoice_days"] - a["invoice_days"]) * a["net_per_invoice_day"]
rate_eff = b["invoice_days"] * (b["net_per_invoice_day"] - a["net_per_invoice_day"])
assert np.isclose(days_eff + rate_eff, b["net_sales"] - a["net_sales"])
d_prev = sorted(sales.loc[sales["Month"] == prev_p, "Day"].dt.day.unique())
d_worst = sorted(sales.loc[sales["Month"] == worst, "Day"].dt.day.unique())
cal = pd.date_range(worst.start_time, worst.end_time.normalize(), freq="D")
no_inv = [d for d in cal if d.day not in d_worst]
print(f"Invoice days: {prev_p} {int(a['invoice_days'])}, {worst} {int(b['invoice_days'])}")
print("Dates in", worst, "with no eligible invoice:", [d.strftime("%a %d %b") for d in no_inv])
print("Weekdays present in the whole dataset:", sorted(sales["InvoiceDate"].dt.day_name().unique()))
print(f"Net sales per invoice day: £{a['net_per_invoice_day']:,.0f} -> £{b['net_per_invoice_day']:,.0f} "
      f"({(b['net_per_invoice_day'] / a['net_per_invoice_day'] - 1) * 100:+.1f}%)")
print(f"Arithmetic split of the {money(b['net_sales'] - a['net_sales'])} change: six fewer invoice days valued at March's net sales per invoice day {money(days_eff)} "
      f"({days_eff / (b['net_sales'] - a['net_sales']) * 100:.0f}%), lower sales per invoice day {money(rate_eff)} "
      f"({rate_eff / (b['net_sales'] - a['net_sales']) * 100:.0f}%)")
print(f"Positive sales £{a['positive_sales']:,.0f} -> £{b['positive_sales']:,.0f} ({money(b['positive_sales'] - a['positive_sales'])}); "
      f"credits £{a['credit_magnitude']:,.0f} -> £{b['credit_magnitude']:,.0f} (+{money(b['credit_magnitude'] - a['credit_magnitude'])})")
R.update(days_prev=int(a["invoice_days"]), days_worst=int(b["invoice_days"]),
         no_invoice_dates=[d.strftime("%Y-%m-%d") for d in no_inv],
         perday_prev=float(a["net_per_invoice_day"]), perday_worst=float(b["net_per_invoice_day"]),
         perday_pct=float((b["net_per_invoice_day"] / a["net_per_invoice_day"] - 1) * 100),
         days_effect=float(days_eff), rate_effect=float(rate_eff),
         pos_prev=float(a["positive_sales"]), pos_worst=float(b["positive_sales"]),
         cred_prev=float(a["credit_magnitude"]), cred_worst=float(b["credit_magnitude"]),
         inv_prev=int(a["sales_invoices"]), inv_worst=int(b["sales_invoices"]))

# A.3 Partial December 2011: compare like-for-like with the same dates one year earlier
lfl = {}
for yr in (2010, 2011):
    s = sales[(sales["InvoiceDate"] >= f"{yr}-12-01") & (sales["InvoiceDate"] < f"{yr}-12-10")]
    lfl[yr] = (s["LineValue"].sum(), s["Day"].nunique())
print(f"\nLike-for-like 1–9 December: 2010 £{lfl[2010][0]:,.0f} ({lfl[2010][1]} invoice days) vs "
      f"2011 £{lfl[2011][0]:,.0f} ({lfl[2011][1]} invoice days): {(lfl[2011][0] / lfl[2010][0] - 1) * 100:+.1f}%")
R.update(lfl_2010=float(lfl[2010][0]), lfl_2011=float(lfl[2011][0]), lfl_pct=float((lfl[2011][0] / lfl[2010][0] - 1) * 100))
""")

code(r"""
# A.4 Sensitivity: do material cleaning decisions change the answer?
def scope_result(mask, label):
    s = df[mask].copy()
    s["Month"] = s["InvoiceDate"].dt.to_period("M"); s["Day"] = s["InvoiceDate"].dt.normalize()
    t = monthly_table(s); mm = mom_changes(t).dropna(subset=["change_gbp"]).sort_values("change_gbp")
    return {"scope": label, "lines": len(s), "largest decline": f"{(mm.index[0] - 1).strftime('%b')}→{mm.index[0].strftime('%b %Y')}",
            "change_gbp": mm["change_gbp"].iloc[0], "change_pct": mm["change_pct"].iloc[0],
            "2nd largest": f"{(mm.index[1] - 1).strftime('%b')}→{mm.index[1].strftime('%b %Y')}",
            "gap_to_2nd_gbp": mm["change_gbp"].iloc[1] - mm["change_gbp"].iloc[0]}

base = df["Exclusion"] == "Retained"
sens = pd.DataFrame([
    scope_result(base, "Main scope (R1–R3)"),
    scope_result(base & ~dup_mask, "Main scope, repeated rows removed"),
    scope_result((df["UnitPrice"] > 0) & valid_sign, "Non-merchandise codes included"),
    scope_result(base & df["CustomerID"].notna(), "Only lines with a CustomerID"),
])
display(sens)
same = (sens["largest decline"] == sens["largest decline"].iloc[0]).all()
print("Largest decline identical under every scope:", same)
R["sensitivity"] = sens.to_dict(orient="records"); R["sensitivity_same"] = bool(same)
""")

md(r"""
**Method A result.** Among the 11 comparisons of adjacent complete months, the largest GBP decline is
**March → April 2011: −£197,681 (−29.0%)**, followed by December 2010 → January 2011 (−£179,917, −23.7%).
April also had only 21 invoice days against 27 in March: no eligible invoice is dated 22–25 April or 29–30 April 2011
(Saturdays never appear in the data). Net sales per invoice day fell by only 8.8% (£25,207 → £22,995).
Valuing the six missing days at March's net sales per invoice day gives −£151,241, which corresponds arithmetically
to about 77% of the decline; lower sales per invoice day account for the other −£46,441. The 22–25 April dates coincide
with the 2011 Easter weekend and 29 April 2011 was a bank holiday for the royal wedding, but the records do not show
whether the business was closed, so the company's own trading calendar should be checked first. The same March → April result appears under all four
cleaning scopes. With non-merchandise codes included, however, January is only about £1.1k behind, so the ranking
of the two biggest falls is close and depends on scope.
Partial December 2011 is compared only with 1–9 December 2010 (+9.7%), never with full November.

*Limitation:* a month-to-month comparison cannot separate seasonality from a real problem, because one year of
data has no earlier April for comparison. The invoice-day split is arithmetic, not proof of cause.
""")

md(r"""
### Method B – Distribution of order (invoice) values

*Why:* a monthly total can fall because there are **fewer orders** or because orders are **smaller**. The mean is
pulled up by a few very large wholesale orders, so the median, quartiles and the 95th percentile (P95) are used.
*Qualifying invoices:* non-credit invoices in the 12 complete months, valued as the sum of their eligible
merchandise lines (always positive).
""")

code(r"""
inv = (sales[~sales["IsCredit"] & sales["Month"].isin(complete.index)]
       .groupby("InvoiceNo").agg(value=("LineValue", "sum"), Month=("Month", "first"),
                                 months=("Month", "nunique"), Country=("Country", "first")))
assert (inv["months"] == 1).all()  # no invoice spans two months
p95_all = inv["value"].quantile(0.95)
def dist(v):
    q = v.quantile([0.25, 0.5, 0.75, 0.95])
    return pd.Series({"invoices": len(v), "Q1": q[0.25], "median": q[0.5], "Q3": q[0.75], "P95": q[0.95],
                      "mean": v.mean(), "max": v.max(),
                      "large_orders(>12m P95)": int((v > p95_all).sum()),
                      "large_order_value": v[v > p95_all].sum(),
                      "share_of_value_from_large_%": v[v > p95_all].sum() / v.sum() * 100})
overall = dist(inv["value"])
by_month = inv.groupby("Month")["value"].apply(dist).unstack()
print(f"All {len(inv):,} qualifying invoices, Dec 2010 – Nov 2011:")
print(overall.apply(lambda v: f"{v:,.2f}"))
display(by_month.loc[[prev_p, worst]])
R.update(inv_n=int(overall["invoices"]), inv_q1=float(overall["Q1"]), inv_median=float(overall["median"]),
         inv_q3=float(overall["Q3"]), inv_p95=float(overall["P95"]), inv_mean=float(overall["mean"]),
         inv_max=float(overall["max"]), inv_large_share=float(overall["share_of_value_from_large_%"]),
         inv_by_month={str(p): {k: float(v) for k, v in by_month.loc[p].items()} for p in (prev_p, worst)})
""")

code(r"""
fig, ax = plt.subplots(figsize=(8, 3.5))
data = [inv.loc[inv["Month"] == p, "value"].values for p in complete.index]
bp = ax.boxplot(data, whis=(5, 95), showfliers=False, patch_artist=True, widths=0.55,
                medianprops=dict(color=INK, lw=1.5), whiskerprops=dict(color=INK2), capprops=dict(color=INK2))
for patch, p in zip(bp["boxes"], complete.index):
    patch.set_facecolor(BLUE if p in (prev_p, worst) else BLUE_LIGHT); patch.set_edgecolor("white")
ax.set_yscale("log")
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"£{v:,.0f}"))
ax.set_xticks(range(1, len(complete) + 1), [p.strftime("%b\n%Y") if p.month in (12, 1) else p.strftime("%b") for p in complete.index])
ax.axhline(p95_all, color=ORANGE, lw=1.2, ls="--")
ax.text(len(complete) + 0.45, p95_all, f"12-month P95\n£{p95_all:,.0f}", va="center", fontsize=8, color=INK)
ax.set_xlim(0.4, len(complete) + 1.6)
ax.set_ylabel("Invoice value (GBP, log scale)"); ax.set_xlabel("Complete calendar month")
ax.set_title("Order value per sales invoice (box = Q1–Q3, line = median, whiskers = P5–P95)", loc="left", fontsize=10)
save(fig, "fig5_invoice_value_distribution")
bm = by_month
print(f"Interpretation: median order £{bm.loc[prev_p, 'median']:,.0f} in March vs £{bm.loc[worst, 'median']:,.0f} in April 2011 – "
      f"the typical order did not shrink; large orders above £{p95_all:,.0f} fell from {int(bm.loc[prev_p, 'large_orders(>12m P95)'])} "
      f"(£{bm.loc[prev_p, 'large_order_value']:,.0f}) to {int(bm.loc[worst, 'large_orders(>12m P95)'])} (£{bm.loc[worst, 'large_order_value']:,.0f}).")
diff = bm.loc[worst] - bm.loc[prev_p]
print(f"{prev_p} -> {worst} differences: sales invoices {diff['invoices']:+,.0f}, median order {diff['median']:+,.2f} GBP, "
      f"large orders {diff['large_orders(>12m P95)']:+,.0f}, large-order value {money(diff['large_order_value'])}, "
      f"invoice days {monthly.loc[worst, 'invoice_days'] - monthly.loc[prev_p, 'invoice_days']:+d}")
""")

md(r"""
**Method B result.** Across 18,957 qualifying sales invoices, the typical order is about £303 (median), half of
all orders lie between £153 and £484, and the top 5% start at £1,506. These large orders contribute 36.3% of
all positive sales value. In March and April 2011 the **median order hardly moved (£295 → £300)**,
but April had 205 fewer sales invoices (1,440 → 1,235). Large orders above the 12-month P95 also fell from 71 to
42, and their value fell by about £116k (£232,009 → £115,550). The fall therefore came from fewer orders and fewer big orders, not from
smaller typical orders.

*Limitation:* invoice values describe orders, not customers. One wholesaler can place several large invoices,
and credit notes are not netted against the order they reverse.
""")

md(r"""
### Method C – Country contribution to the largest decline

*Why:* management needs to know **where** the change came from. For each country:
contribution = net sales in April − net sales in March (signed GBP). Countries that appear in only one of the two
months are included with 0 for the missing month. `Unspecified` is shown as the unknown-country group.
All contributions must add up to the total monthly change.
""")

code(r"""
pair = sales[sales["Month"].isin([prev_p, worst])].copy()
pair["Country"] = pair["Country"].replace({"Unspecified": "Unspecified (unknown)"})
cc = pair.pivot_table(index="Country", columns="Month", values="LineValue", aggfunc="sum", fill_value=0.0)
cc.columns = [str(c) for c in cc.columns]
cc = cc.rename(columns={str(prev_p): "prev", str(worst): "worst"})
cc["contribution_gbp"] = cc["worst"] - cc["prev"]
cc["share_of_change_%"] = cc["contribution_gbp"] / cc["contribution_gbp"].sum() * 100
cc["in_one_month_only"] = (cc["prev"] == 0) | (cc["worst"] == 0)
cc = cc.sort_values("contribution_gbp")
total_change = monthly.loc[worst, "net_sales"] - monthly.loc[prev_p, "net_sales"]
assert np.isclose(cc["contribution_gbp"].sum(), total_change), "Country contributions do not reconcile"
print(f"Countries in either month: {len(cc)} | sum of contributions {money(cc['contribution_gbp'].sum(), 2)} "
      f"= total change {money(total_change, 2)}  ✔")
print("Largest downward contributions:"); display(cc.head(8))
print("Largest offsets (increases):"); display(cc.tail(5).iloc[::-1])
print("Countries present in only one of the two months:", ", ".join(cc.index[cc["in_one_month_only"]]))
cc.to_csv(OUT / "country_contributions.csv")
""")

code(r"""
# Region view, also per invoice day, because April had fewer invoice days
rg = pair.groupby(["Region", "Month"]).agg(net=("LineValue", "sum"), days=("Day", "nunique")).unstack("Month")
reg_tab = pd.DataFrame({
    "net_prev": rg[("net", prev_p)], "net_worst": rg[("net", worst)],
    "change_gbp": rg[("net", worst)] - rg[("net", prev_p)],
    "change_%": (rg[("net", worst)] / rg[("net", prev_p)] - 1) * 100,
    "share_of_change_%": (rg[("net", worst)] - rg[("net", prev_p)]) / total_change * 100,
    "per_day_prev": rg[("net", prev_p)] / monthly.loc[prev_p, "invoice_days"],
    "per_day_worst": rg[("net", worst)] / monthly.loc[worst, "invoice_days"]})
reg_tab["per_day_change_%"] = (reg_tab["per_day_worst"] / reg_tab["per_day_prev"] - 1) * 100
display(reg_tab)

# Drill-down: accounts behind the four largest international declines
top_intl = [c for c in cc.index if c != "United Kingdom"][:4]
acct = (pair[pair["Country"].isin(top_intl)]
        .assign(Customer=lambda d: d["CustomerID"].astype("string").fillna("unknown"))
        .pivot_table(index=["Country", "Customer"], columns="Month", values="LineValue", aggfunc="sum", fill_value=0.0))
acct.columns = ["prev", "worst"]; acct["change_gbp"] = acct["worst"] - acct["prev"]
acct = acct.sort_values("change_gbp")
display(acct.head(8))
top4 = cc.loc[top_intl, "contribution_gbp"].sum()
print(f"Top-4 international countries {top_intl}: {money(top4)} ({top4 / total_change * 100:.0f}% of the decline)")
R.update(total_change=float(total_change), n_countries_pair=int(len(cc)),
         contrib={c: float(v) for c, v in cc["contribution_gbp"].items()},
         contrib_prev={c: float(v) for c, v in cc["prev"].items()}, contrib_worst={c: float(v) for c, v in cc["worst"].items()},
         one_month_only=list(cc.index[cc["in_one_month_only"]]),
         region={r: {k: float(v) for k, v in reg_tab.loc[r].items()} for r in reg_tab.index},
         top_intl=top_intl, top_intl_sum=float(top4),
         top_accounts=[{"country": i[0], "customer": i[1], "prev": float(r["prev"]), "worst": float(r["worst"]),
                        "change": float(r["change_gbp"])} for i, r in acct.head(5).iterrows()])
""")

code(r"""
show_n = 8
down = cc.head(show_n); up = cc[cc["contribution_gbp"] > 0].tail(4)
rest = cc.drop(down.index.union(up.index))
plot = pd.concat([down["contribution_gbp"],
                  pd.Series({f"All other countries ({len(rest)})": rest["contribution_gbp"].sum()}),
                  up["contribution_gbp"]])
fig, ax = plt.subplots(figsize=(7.4, 4.3))
yy = np.arange(len(plot))[::-1]
ax.barh(yy, plot.values, color=[RED if v < 0 else BLUE for v in plot.values], height=0.65, edgecolor="white")
ax.set_yticks(yy, plot.index, fontsize=9)
ax.axvline(0, color="#c3c2b7", lw=1)
ax.xaxis.set_major_formatter(gbp_k)
for yv, v in zip(yy, plot.values):
    ax.text(v + (-2500 if v < 0 else 2500), yv, f"{money(v/1000, 1)}k", va="center", ha="right" if v < 0 else "left", fontsize=8.5, color=INK)
ax.set_xlim(plot.min() * 1.22, max(plot.max(), 0) + 25000)
ax.set_xlabel(f"Change in net sales, {prev_p.strftime('%B')} → {worst.strftime('%B %Y')} (GBP)")
ax.set_title(f"Country contributions to the {money(total_change/1000, 1)}k change (red = decline, blue = offset)", loc="left", fontsize=10)
save(fig, "fig6_country_contributions")
uk = reg_tab.loc["United Kingdom"]; it = reg_tab.loc["International"]
d0, d1 = int(monthly.loc[prev_p, "invoice_days"]), int(monthly.loc[worst, "invoice_days"])
for name, r in (("UK", uk), ("International", it)):
    print(f"{name}: whole month {money(r['net_prev'])} -> {money(r['net_worst'])} ({r['change_%']:+.1f}%); "
          f"per invoice day {money(r['net_prev'])}/{d0} = {money(r['per_day_prev'])} -> {money(r['net_worst'])}/{d1} = "
          f"{money(r['per_day_worst'])} ({r['per_day_change_%']:+.1f}%); share of total change {r['share_of_change_%']:.0f}%")
""")

md(r"""
**Method C result.** All 32 countries reconcile exactly to the −£197,681 change. The United Kingdom contributed
−£127,560 (65%): its whole-month net sales fell 22.7%, but per invoice day they changed by only −0.6%
(£563,152 / 27 days = £20,857 → £435,592 / 21 days = £20,742). International customers contributed −£70,122 (35%)
even though they were only 17% of March sales: their whole-month net sales fell 59.7%, and per invoice day they fell
48.2% (£4,349 → £2,253, same 27- and 21-day denominators). Four countries – the Netherlands (−£19,320), Australia (−£17,072),
EIRE (−£11,013) and France (−£9,851) – account for −£57,255, or 29% of the total decline. Most of this came from a few large accounts that ordered less in April
(e.g., Netherlands customer 14646: £21,327 → £2,887; Australia customer 12415: £16,558 → −£425).
Small offsets came from Japan (+£4,866), Singapore (+£2,119) and Hong Kong (+£1,540).

*Limitation:* a country contribution shows where the money moved, not why. One large account can drive a whole
country, and credit notes are dated when issued, not when the original sale happened.
""")

code(r"""
methods = pd.DataFrame([
    ["A. Time-series comparison", "Question is about change between adjacent complete months",
     f"{prev_p.strftime('%b')}→{worst.strftime('%b %Y')}: {money(R['worst_change'])} ({R['worst_pct']:+.1f}%); per invoice day {R['perday_pct']:+.1f}%",
     "Cannot separate seasonality from a real problem with one year of data"],
    ["B. Distribution analysis", "Separates 'fewer orders' from 'smaller orders'; robust to outliers",
     f"Median order £{R['inv_by_month'][str(prev_p)]['median']:,.0f}→£{R['inv_by_month'][str(worst)]['median']:,.0f}; "
     f"sales invoices {R['inv_prev']:,}→{R['inv_worst']:,}",
     "Invoice ≠ customer; credits not matched to original orders"],
    ["C. Country contribution", "Shows where the change came from; sums exactly to the total",
     f"UK {money(R['region']['United Kingdom']['change_gbp'])} (month {R['region']['United Kingdom']['change_%']:+.1f}%, "
     f"per invoice day {R['region']['United Kingdom']['per_day_change_%']:+.1f}%); International "
     f"{money(R['region']['International']['change_gbp'])} (month {R['region']['International']['change_%']:+.1f}%, "
     f"per invoice day {R['region']['International']['per_day_change_%']:+.1f}%)",
     "Shows where, not why; a few accounts can dominate a country"],
], columns=["Method", "Why chosen", "Computed result", "Limitation"])
display(methods)
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 7 – Executive summary and introduction

**Company and data.** The client is the unnamed UK-based, registered, non-store online retailer described by
UCI and by Chen, Sain & Guo (2012). It sells mainly unique all-occasion gifts, and many of its customers are
wholesalers. The dataset contains all of its transactions between 1 December 2010 and 9 December 2011
(541,909 invoice lines, 38 countries). Since the data is the company's own invoice log, it can show where sales
rose and fell. It cannot show the reasons.

**Executive summary.** Management believes sales KPIs are below target (issue 1(b)(ii)). Across 12 complete
months, the largest month-to-month fall in net invoiced merchandise sales was **March → April 2011: −£197,681
(−29.0%)**, from £680,582 to £482,901. April had six fewer days with recorded invoices (21 vs 27). Valued at March's
net sales per invoice day, those six days correspond arithmetically to about 77% of the decline (−£151,241); the records
do not show why no invoices were recorded on them. On the same two bases, UK net sales fell 22.7% for the whole month
but only 0.6% per invoice day, while international net sales fell 59.7% for the whole month and 48.2% per invoice day.
A few large accounts in the Netherlands, Australia, EIRE and France ordered less in April (−£57,255 together). The
typical order size did not change (median about £300).
**Proposed action:** the sales team should first check the April trading calendar and then review the international
wholesale accounts that ordered less in April. Net sales per invoice day for UK and international customers should be
tracked as a monthly KPI.
""")

code(r"""
key = pd.Series({
    "Raw invoice lines": f"{R['raw_rows']:,}",
    "Retained / excluded lines": f"{R['retained_rows']:,} / {R['excluded_rows']:,}",
    "Largest complete-month decline": f"{R['prev_month']} -> {R['worst_month']}: {money(R['worst_change'])} ({R['worst_pct']:+.1f}%)",
    "Invoice days (Mar / Apr)": f"{R['days_prev']} / {R['days_worst']}",
    "Net sales per invoice day change": f"{R['perday_pct']:+.1f}%",
    "UK change: GBP, whole-month %, per-invoice-day %": f"{money(R['region']['United Kingdom']['change_gbp'])}, {R['region']['United Kingdom']['change_%']:+.1f}%, {R['region']['United Kingdom']['per_day_change_%']:+.1f}%",
    "International change: GBP, whole-month %, per-invoice-day %": f"{money(R['region']['International']['change_gbp'])}, {R['region']['International']['change_%']:+.1f}%, {R['region']['International']['per_day_change_%']:+.1f}%",
    "Credit-note lines raw / retained": f"{R['credit_lines_raw']:,} / {R['retained_credit_lines']:,}",
    "Top-4 international countries": f"{', '.join(R['top_intl'])}: {money(R['top_intl_sum'])}",
    "Median order Mar / Apr": f"£{R['inv_by_month'][R['prev_month']]['median']:,.0f} / £{R['inv_by_month'][R['worst_month']]['median']:,.0f}",
})
print(key.to_string())
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 8 – Challenge / problem statement

**Problem.** Sales and KPIs are reported to be far below target, and management must decide within three weeks.
The real targets are not in the data, so this analysis uses the **preceding complete month as a historical
benchmark**. KPI = monthly net invoiced merchandise sales (GBP). Period = the 12 complete months from
December 2010 to November 2011 (December 2011 is partial).

**Why it matters.** If the company misreads a calendar effect as a demand problem, or the reverse, it will take the
wrong action. It could change products or prices when the decline is concentrated in a few accounts that ordered
less. Management therefore needs to know *where* sales were lower before it acts.

**What the records can establish.** When sales fell, by how much, in which countries and accounts, and whether the
cause was fewer orders, smaller orders or more credits. **What they cannot establish:** targets, margins or profit,
or why customers changed their behaviour. Those need input from the sales team and from customers.
""")

# ---------------------------------------------------------------------------
md(r"""
## Step 9 – Solution statement and implementation plan

**Completed analysis (this notebook).** Cleaned and reconciled the full UCI log, built the monthly net-sales KPI,
and applied three descriptive methods (time-series comparison, order-value distribution, country contribution).
All three point the same way: the March → April 2011 fall coincides with fewer days with recorded invoices and with
lower international wholesale orders. The median order size did not change.

**Priority recommendation.** *The Head of Sales, with the international account managers, should review the
international wholesale accounts that ordered less in April 2011, starting with the four largest country contributors
(Netherlands, Australia, EIRE, France). Before that, Finance should confirm the April trading calendar.* The timeline
below is a proposed management schedule, not a completed result.

| When | Who | What | Output |
|---|---|---|---|
| Week 1 | Finance + Sales operations | Confirm whether 22–25 and 29–30 April 2011 were non-trading days. List April credit notes (£33,317 vs £10,675 in March) | Calendar-adjusted April baseline |
| Week 2 | Head of Sales + account managers | Contact the accounts that ordered less (e.g., NL 14646, AU 12415, EIRE 14156/14911). Record why April orders were lower (stock, price, timing, other) | Account-by-account reasons |
| Week 3 | Executive management | Decide on recovery actions per account. Approve the KPI below | Decision memo |
| Monthly, from then on | Sales operations | Review the KPI on the first working day of each month | KPI dashboard |

**Metric to monitor:** net invoiced sales **per invoice day**, split UK vs international, with credit
magnitude as % of positive sales as a supporting check.

**Proposed vs proven.** The recommendation is a proposed business action. Any benefit (recovered accounts,
better forecasts) is expected but **not measured** here. It should be confirmed with the same KPI in later months.
""")

code(r"""
R["generated_with"] = {"python": platform.python_version(), "pandas": pd.__version__,
                       "numpy": np.__version__, "matplotlib": matplotlib.__version__}
with open(OUT / "results.json", "w") as f:
    json.dump(R, f, indent=2, default=str)
print("Saved:", OUT / "results.json", "and", len(list(FIG.glob('*.png'))), "figures in", FIG)
""")

md(r"""
### References
1. Chen, D. (2015). *Online Retail* [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5BW33
2. Chen, D., Sain, S. L., & Guo, K. (2012). Data mining for the online retail industry: A case study of RFM
   model-based customer segmentation using data mining. *Journal of Database Marketing & Customer Strategy
   Management, 19*(3), 197–208. https://doi.org/10.1057/dbm.2012.17
3. GOV.UK (2010). *PM welcomes announcement of date for Royal Wedding* (29 April 2011 declared a bank holiday).
   https://www.gov.uk/government/news/pm-welcomes-announcement-of-date-for-royal-wedding
""")

nb = nbf.v4.new_notebook()
nb["cells"] = cells
nb["metadata"] = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python"},
    "colab": {"provenance": [], "name": "CSE464_Case_Study_1.ipynb"},
}
OUT_NB.parent.mkdir(exist_ok=True)
nbf.write(nb, OUT_NB)
print("Wrote", OUT_NB)
