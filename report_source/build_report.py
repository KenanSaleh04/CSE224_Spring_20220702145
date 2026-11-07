"""Build the Case Study 1 report (editable DOCX + submission PDF) from the executed notebook outputs.

Usage:  python report_source/build_report.py
Inputs:  report_source/outputs/results.json and outputs/figures/*.png (written by the notebook)
Outputs: report_source/CSE464_Case_Study_1.docx (editable source)
         submission/CSE464_Case_Study_1.pdf (converted with LibreOffice)
Every number in the text is formatted from results.json, so the report always matches the notebook.
"""
import json
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
R = json.loads((HERE / "outputs" / "results.json").read_text())
FIG = HERE / "outputs" / "figures"
DOCX_OUT = HERE / "CSE464_Case_Study_1.docx"
PDF_OUT = ROOT / "submission" / "CSE464_Case_Study_1.pdf"
NOTEBOOK = "CSE464_Case_Study_1.ipynb"

BLUE = RGBColor(0x2F, 0x53, 0x96)   # template heading blue, measured from the template PDF
INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x55, 0x55, 0x55)
BODY_FONT, HEAD_FONT = "Times New Roman", "Arial"


# ----------------------------------------------------------------------------- number formatting
NEG = "−\u2060"  # minus sign + word joiner, so a line never breaks between "−" and "£"


def gbp(v, dp=0):
    return f"{NEG if v < 0 else ''}£{abs(v):,.{dp}f}"


def sgbp(v, dp=0):
    return ("+" if v > 0 else "") + gbp(v, dp)


def pct(v, dp=1, signed=True):
    s = f"{abs(v):.{dp}f}%"
    if not signed:
        return s
    return (NEG if v < 0 else "+") + s


def sint(v):
    return f"{NEG if v < 0 else ('+' if v > 0 else '')}{abs(int(v)):,}"


def kgbp(v, dp=0):
    return f"{NEG if v < 0 else ''}£{abs(v) / 1000:,.{dp}f}k"


M = R["monthly"]
mar, apr = M[R["prev_month"]], M[R["worst_month"]]
uk, intl = R["region"]["United Kingdom"], R["region"]["International"]
bm_mar, bm_apr = R["inv_by_month"][R["prev_month"]], R["inv_by_month"][R["worst_month"]]
C = R["contrib"]
acc = {a["customer"]: a for a in R["top_accounts"]}
sens = R["sensitivity"]
nonmerch_scope = [s for s in sens if s["scope"].startswith("Non-merchandise")][0]
dedup_scope = [s for s in sens if "repeated" in s["scope"]][0]
cust_scope = [s for s in sens if s["scope"].startswith("Only lines")][0]
ex = R["excluded_by_rule"]
r1 = [v for k, v in ex.items() if k.startswith("R1")][0]
r2 = sum(v for k, v in ex.items() if k.startswith("R2"))
r3 = [v for k, v in ex.items() if k.startswith("R3")][0]
days_share = R["days_effect"] / R["total_change"] * 100
large_drop = bm_apr["large_order_value"] - bm_mar["large_order_value"]
jar = [e for e in R["extreme_orders"] if e["date"].startswith("2011-01") and e["value"] > 0][0]
bird = [e for e in R["extreme_orders"] if e["date"].startswith("2011-12") and e["value"] > 0][0]
assert R["sensitivity_same"], "Conclusion changes across cleaning scopes – review the text before building"
assert R["worst_month"] == "2011-04" and R["prev_month"] == "2011-03", "Text below is written for March→April 2011"
assert R["days_prev"] - R["days_worst"] == 6, "Text below says 'six fewer' invoice days"

# ----------------------------------------------------------------------------- document helpers
# Layout measured from the original CSE464_Case_Study_Template.pdf: US Letter, 1-inch side margins,
# logo (258.7 pt wide) + "COMPUTER ENGINEERING DEPARTMENT" (Calibri Bold 12) as a header on every page,
# section headings Arial Bold 20 pt #2F5396, lettered subsections Arial 16 pt #2F5396, body Times New Roman 12,
# bullets Calibri 12 with a bold label.
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Pt(612), Pt(792)
sec.left_margin = sec.right_margin = Pt(72)
sec.top_margin, sec.bottom_margin = Pt(110), Pt(64)
sec.header_distance, sec.footer_distance = Pt(35), Pt(30)

normal = doc.styles["Normal"]
normal.font.name = BODY_FONT
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.0
BULLET_FONT = "Calibri"


def add_runs(p, text, size=None, color=None, font=None, italic=False, underline_bold=False):
    """Supports **bold** inline markup."""
    for i, part in enumerate(re.split(r"\*\*", text)):
        if not part:
            continue
        r = p.add_run(part)
        r.bold = (i % 2 == 1)
        r.italic = italic
        if underline_bold and r.bold:
            r.underline = True
        if size:
            r.font.size = Pt(size)
        if color is not None:
            r.font.color.rgb = color
        if font:
            r.font.name = font
            r._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    return p


def para(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None, after=4, italic=False, keep=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    if keep:
        p.paragraph_format.keep_with_next = True
    return add_runs(p, text, size=size, italic=italic)


def bullet(text, size=12, underline_bold=False, keep=False):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Pt(36)        # text at 108 pt, bullet at 90 pt, as in the template
    p.paragraph_format.first_line_indent = Pt(-18)
    if keep:
        p.paragraph_format.keep_with_next = True
    return add_runs(p, text, size=size, font=BULLET_FONT, underline_bold=underline_bold)


def section(num, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Pt(27)        # numeral at 99 pt
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.tab_stops.add_tab_stop(Pt(72))  # title at 144 pt
    add_runs(p, f"**{num}\t{title}**", size=20, color=BLUE, font=HEAD_FONT)


def sub(letter, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Pt(36)        # label at 108 pt
    p.paragraph_format.first_line_indent = Pt(-18) # letter at 90 pt
    p.paragraph_format.tab_stops.add_tab_stop(Pt(36))
    p.paragraph_format.keep_with_next = True
    add_runs(p, f"{letter}.\t**{title}**", size=16, color=BLUE, font=HEAD_FONT)


def shade(cell, hex_fill):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def fix_widths(t, widths):
    t.autofit = False
    grid = t._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j] / 2.54 * 1440)))
    for row in t.rows:
        for j, c in enumerate(row.cells):
            c.width = Cm(widths[j])


def table(rows, widths, caption, header_fill="DCE4F2", size=9, right_cols=()):
    cap = doc.add_paragraph()
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(2)
    cap.paragraph_format.keep_with_next = True
    add_runs(cap, caption, size=9.5, color=INK, font=HEAD_FONT)
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.width = Cm(widths[j])
            c.text = ""
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            if j in right_cols and i > 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            add_runs(p, f"**{val}**" if i == 0 else str(val), size=size, font=HEAD_FONT)
            if i == 0:
                shade(c, header_fill)
        # keep table rows together on one page
        for c in t.rows[i].cells:
            for p in c.paragraphs:
                p.paragraph_format.keep_with_next = i < len(rows) - 1
    fix_widths(t, widths)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = 0.8
    return t


def figure(name, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(FIG / f"{name}.png"), width=Cm(width_cm))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cap.paragraph_format.space_after = Pt(6)
    add_runs(cap, caption, size=9.5, color=GREY, font=HEAD_FONT)


def fill_header(header):
    """Template page header: university logo centred above the department name."""
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(17)
    p.add_run().add_picture(str(HERE / "assets" / "yeditepe_logo_template.png"), width=Pt(258.7))
    q = header.add_paragraph()
    q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    q.paragraph_format.space_after = Pt(0)
    add_runs(q, "**COMPUTER ENGINEERING DEPARTMENT**", size=12, font="Calibri", color=RGBColor(0, 0, 0))


def page_number_footer(footer):
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = txt
        run._r.append(el)
    run.font.size = Pt(9); run.font.color.rgb = GREY; run.font.name = "Calibri"


# ----------------------------------------------------------------------------- header, footer and cover page
sec.different_first_page_header_footer = True   # same header on the cover, but no page number there
fill_header(sec.header)
fill_header(sec.first_page_header)
page_number_footer(sec.footer)

BLACK = RGBColor(0, 0, 0)
spacer = doc.add_paragraph(); spacer.paragraph_format.space_before = Pt(118)  # template title sits at ~245 pt
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(14)
add_runs(p, "**CSE464 Case Study 1**", size=20, font="Calibri", color=BLACK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(0)
add_runs(p, "**INTRODUCTION TO DATA SCIENCE & BIG DATA ANALYTICS**", size=14, font="Calibri", color=BLACK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_runs(p, "**(FALL SEMESTER 2026)**", size=14, font="Calibri", color=BLACK)

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(60)
p.paragraph_format.space_after = Pt(2)
add_runs(p, "Where Did April 2011 Sales Go?", size=16, font="Calibri", color=BLACK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_runs(p, "A descriptive review of monthly net sales for a UK online gift retailer", size=12, font="Calibri", color=GREY)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
add_runs(p, "Selected issue 1(b)(ii): Sales figures / Company KPIs", size=12, font="Calibri", color=BLACK, italic=True)

spacer = doc.add_paragraph(); spacer.paragraph_format.space_before = Pt(70)
cover_rows = [
    ("Student", "Kenan Saleh"),
    ("Student number", "20220702115"),
    ("Semester", "Fall 2026"),
    ("Dataset", "UCI Online Retail (ID 352), DOI 10.24432/C5BW33"),
    ("Analysis notebook", f"{NOTEBOOK} (Python, NumPy, pandas, matplotlib)"),
]
t = doc.add_table(rows=len(cover_rows), cols=2)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(cover_rows):
    for j, val in enumerate((k, v)):
        c = t.cell(i, j); c.text = ""
        add_runs(c.paragraphs[0], f"**{val}**" if j == 0 else val, size=12, font="Calibri", color=BLACK)
        c.paragraphs[0].paragraph_format.space_after = Pt(3)
fix_widths(t, [4.0, 11.0])
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# ----------------------------------------------------------------------------- I. Executive summary
section("I.", "Executive Summary/Abstract")
para(f"**Challenge.** This study addresses issue 1(b)(ii) of the assignment: sales figures and company KPIs are far below "
     f"target, and management needs a decision within three weeks. The real targets are not available, so the preceding "
     f"complete month is used as a **historical benchmark**. The question is: “Across complete calendar months, when did "
     f"net invoiced sales fall most, which countries accounted for the change, and what should management investigate first?”")
para(f"**Solution.** As the company's data scientist, I analysed its own {R['raw_rows']:,} invoice lines (UCI Online Retail, "
     f"1 December 2010 – 9 December 2011) with descriptive analytics in Python. I cleaned the records, built a monthly net-sales "
     f"KPI and applied three methods: a month-to-month comparison, a distribution of order values and a country contribution "
     f"breakdown. The partial month December 2011 (1–9 December) was not compared with full months.")
para(f"**Outcome.** The largest fall was **March → April 2011: {gbp(R['worst_change'])} ({pct(R['worst_pct'])})**, from "
     f"{gbp(mar['net_sales'])} to {gbp(apr['net_sales'])}. April had {R['days_worst']} days with recorded invoices, compared "
     f"with {R['days_prev']} in March. At March's rate of {gbp(R['perday_prev'])} per invoice day, the six missing days "
     f"correspond arithmetically to {gbp(R['days_effect'])}, about {days_share:.0f}% of the decline; the records do not show "
     f"why no invoices were recorded on those days. Measured on the same two bases, UK net sales fell "
     f"{pct(uk['change_%'], signed=False)} for the whole month and {pct(uk['per_day_change_%'], signed=False)} per invoice "
     f"day, while international net sales fell {pct(intl['change_%'], signed=False)} for the whole month and "
     f"{pct(intl['per_day_change_%'], signed=False)} per invoice day (per-day values divide by {R['days_prev']} and "
     f"{R['days_worst']} days). Accounts in the Netherlands, Australia, EIRE and France ordered less in April "
     f"({gbp(R['top_intl_sum'])} together). The median order hardly changed ({gbp(bm_mar['median'])} → "
     f"{gbp(bm_apr['median'])}); there were {R['inv_prev'] - R['inv_worst']:,} fewer orders.")
para(f"**Proposed action.** Finance should first confirm the April 2011 trading calendar. The Head of Sales should then "
     f"review the international wholesale accounts that ordered less in April, following the proposed three-week "
     f"implementation plan (Table 4). The KPI to monitor is **net invoiced sales per invoice day, split UK vs "
     f"international**. All results come from the executed notebook {NOTEBOOK}.")

# ----------------------------------------------------------------------------- II. Background
section("II.", "Background/Introduction")
bullet(f"**About the Client/Company:** The client is an unnamed UK-based, registered, non-store online retailer that "
       f"mainly sells unique all-occasion gifts. Many of its customers are wholesalers [1]. Chen, Sain and Guo describe it "
       f"as a small business with some 80 staff [2]. Management gave me, as the company's data scientist, its transaction "
       f"log for 1 December 2010 – 9 December 2011: {R['raw_rows']:,} invoice lines on {R['raw_invoices']:,} invoice "
       f"numbers from customers in {R['raw_countries']} countries [1]. The log also contains Amazon fee lines, which shows "
       f"that some sales go through the Amazon marketplace.")
bullet("**Industry Context:** Online gift retail sells to both small shops (wholesale) and consumers. Demand is seasonal and "
       f"builds before Christmas: in this data, monthly net sales rose from {gbp(M['2011-08']['net_sales'])} in August to "
       f"{gbp(M['2011-11']['net_sales'])} in November 2011. Wholesale buying is lumpy: the top 5% of orders (above "
       f"{gbp(R['inv_p95'])}) bring {R['inv_large_share']:.1f}% of order value, so a change in one account's orders can "
       "move a whole month.")
bullet("**Goals and Objectives:** The company wants to bring sales back toward its targets. The objectives of this study are to "
       "(1) build a reliable monthly net-sales KPI from the raw records, (2) find the largest fall between complete months, "
       "(3) show where that fall came from, and (4) recommend what management should investigate first.")

# ----------------------------------------------------------------------------- III. Challenge
section("III.", "Challenge/Problem Statement")
sub("A", "The Problem")
para(f"Sales are reported to be far below target, but management does not know when the shortfall happened or where. "
     f"**KPI:** monthly net invoiced merchandise sales in GBP (positive sales minus credit notes). **Period:** the "
     f"{R['n_complete']} complete calendar months from December 2010 to November 2011; December 2011 stops on 9 December "
     f"and is treated as partial. **Benchmark:** the preceding complete month (a historical comparison, not a management "
     f"target). **Decision:** which issue management should investigate first within the three-week deadline.")
sub("B", "Why It Matters")
para("The assignment brief describes a company that is starting to compete with its existing products and services. "
     "Sales of those existing products are the most direct measure of whether that strategy is working, so a shortfall in "
     "the sales KPI is a strategic problem, not only an accounting one. "
     "If a calendar effect is read as a demand problem, or the other way round, the company will act on the wrong thing. "
     "It might redesign products or cut prices when the decline is concentrated in a few accounts that ordered less. Time "
     "and budget should go where sales were actually lower. The records can "
     "show **when** sales fell, **by how much**, **in which countries and accounts**, and whether the cause was fewer orders, "
     "smaller orders or more credits. They cannot show targets, costs, margins or profit, or **why** customers changed "
     "their behaviour. Those questions need input from the sales team and from the customers.")

# ----------------------------------------------------------------------------- IV. Solution
section("IV.", "Solution/Approach")
bullet("**Use of descriptive analytics, and why**", underline_bold=True, keep=True)
para("The question is about what already happened, and the data are one year of historical invoices with no outcome "
     "variable to predict. Descriptive analytics (cleaning, aggregation, summary statistics and charts) answers it directly "
     "and gives results that executives can check by hand. A predictive model would need several years of data to learn "
     "seasonality, and a two-month comparison would not support it. Descriptive results also fit the three-week deadline.")
sub("C", "Proposed Solution")
para(f"A reproducible notebook ({NOTEBOOK}) follows the nine steps of the course hand-out. Rows are cleaned with three "
     "explicit eligibility rules (Table 1), monthly net sales (Y) are built, and three descriptive methods are applied:")
bullet("**A. Time-series comparison** of adjacent complete months, in GBP (for ranking) and % (change ÷ previous month). "
       "It was chosen because the question is about change over time. Invoice days are counted to check coverage.")
bullet("**B. Distribution analysis** of order values (median, quartiles, 95th percentile). It was chosen because it separates "
       "'fewer orders' from 'smaller orders' and is not distorted by a few very large wholesale invoices.")
bullet("**C. Country contribution analysis:** each country's signed GBP change between the two months. It was chosen because "
       "the contributions add up exactly to the total, which shows where the money moved.")
table([
    ["Rule", "Invoice lines", "Reason"],
    ["R1 zero or negative unit price", f"{r1:,}", "Stock adjustments and 2 bad-debt entries; not sales"],
    ["R2 quantity sign ≠ invoice type", f"{r2:,}", "Guard rule; no line breaks it after R1"],
    ["R3 non-merchandise code", f"{r3:,}", "Postage, carriage, Amazon/bank fees, manual, discount, samples, vouchers"],
    ["Retained (eligible)", f"{R['retained_rows']:,}", f"Includes {R['retained_credit_lines']:,} credit lines and "
                                                       f"{R['retained_missing_customer']:,} lines without a CustomerID"],
    ["Total = raw rows", f"{R['raw_rows']:,}", f"{R['retained_rows']:,} retained + {R['excluded_rows']:,} excluded"],
], widths=[4.6, 2.3, 9.4], caption="Table 1. Row eligibility and reconciliation (notebook Step 3)", right_cols=(1,))
sub("D", "Key Features/Technologies")
para("**Tools:** Google Colab with Python 3; the official **ucimlrepo** import fetch_ucirepo(id=352) with UCI's Excel "
     "workbook as a documented fallback; **pandas** for parsing, filtering, groupby/pivot_table aggregation and calendar "
     "periods; **NumPy** for rule logic (np.select), quantiles and reconciliation checks (np.isclose); **matplotlib** "
     "for all charts. **Features kept:** InvoiceDate (month, day), Quantity × UnitPrice (signed line value), InvoiceNo and its "
     "'C' prefix (sales vs credits, order grouping), Country (UK vs international), CustomerID (drill-down only) and StockCode "
     "(scope filter). Description was not needed.")
sub("E", "Implementation Process")
para(f"Steps 1–2 load the data ({R['raw_rows']:,} rows, 8 columns) and profile types and missing values. UCI's metadata "
     f"says there are no missing values, but {R['missing_customer']:,} CustomerIDs and {R['missing_description']:,} "
     f"descriptions are blank. Step 3 parses dates and numbers, keeps InvoiceNo and StockCode as text, and applies rules "
     f"R1–R3. It keeps credit notes (so returns reduce sales once), lines without a CustomerID (the sale is still valid) and "
     f"{R['duplicate_rows']:,} identical repeated lines (no line number exists to prove they are errors). Step 4 builds Y and "
     f"checks that positive sales − credits = net sales for every month and that the monthly totals equal the sum of eligible "
     f"line values. Steps 5–6 plot X and Y and run methods A–C. Step 6A re-runs the ranking under three other cleaning "
     f"scopes. Steps 7–9 write the summary, problem and plan. The notebook runs top to bottom in a fresh runtime with "
     f"relative paths and saves every figure and number used here.")

# ----------------------------------------------------------------------------- V. Results
section("V.", "Results/Outcomes")
sub("F", "Quantifiable Results")
figure("fig1_monthly_net_sales", 13.6,
       f"Figure 1. Monthly net invoiced sales (GBP, merchandise lines), December 2010 – December 2011. December 2011 covers "
       f"1–9 December only (hatched) and is not ranked. Source: notebook Step 5.")
para(f"**Method A – time series.** Of {R['n_comparisons']} comparisons between adjacent complete months, "
     f"{R['n_declines']} were declines. The largest was **March → April 2011, {gbp(R['worst_change'])} "
     f"({pct(R['worst_pct'])})**, ahead of December 2010 → January 2011 ({gbp(R['second_change'])}, "
     f"{pct(R['second_pct'])}), the post-Christmas fall. April had {R['days_worst']} days with recorded invoices, compared "
     f"with {R['days_prev']} in March; no invoice is dated 22–25 April or 29–30 April. These dates coincide with the 2011 Easter weekend and "
     f"the 29 April royal-wedding bank holiday [3], but the records do not show whether the business was closed. Net sales "
     f"per invoice day fell {pct(R['perday_pct'], signed=False)} ({gbp(R['perday_prev'])} → {gbp(R['perday_worst'])}). "
     f"Valuing the six missing days at March's daily rate gives {gbp(R['days_effect'])} (about {days_share:.0f}% of the "
     f"decline, arithmetic only); lower sales per invoice day account for the other {gbp(R['rate_effect'])}. The ranking stays the same when repeated rows are removed or "
     f"only lines with a CustomerID are used. When postage and fee codes are added back, January comes within "
     f"{gbp(nonmerch_scope['gap_to_2nd_gbp'])} of April. For partial December 2011, the like-for-like 1–9 December "
     f"comparison is {pct(R['lfl_pct'])} against 2010.")
table([
    ["Measure (notebook Steps 4 and 6)", "March 2011", "April 2011", "Change"],
    ["Net invoiced sales (Y)", gbp(mar["net_sales"]), gbp(apr["net_sales"]), f"{gbp(R['worst_change'])} ({pct(R['worst_pct'])})"],
    ["Positive sales", gbp(mar["positive_sales"]), gbp(apr["positive_sales"]), gbp(apr["positive_sales"] - mar["positive_sales"])],
    ["Credit magnitude (returns)", gbp(mar["credit_magnitude"]), gbp(apr["credit_magnitude"]), sgbp(apr["credit_magnitude"] - mar["credit_magnitude"])],
    ["Invoice days", f"{R['days_prev']}", f"{R['days_worst']}", sint(R['days_worst'] - R['days_prev'])],
    ["Net sales per invoice day", gbp(R["perday_prev"]), gbp(R["perday_worst"]), pct(R["perday_pct"])],
    ["Sales invoices (orders)", f"{R['inv_prev']:,}", f"{R['inv_worst']:,}", sint(R['inv_worst'] - R['inv_prev'])],
    ["Median order value", gbp(bm_mar["median"]), gbp(bm_apr["median"]), sgbp(bm_apr["median"] - bm_mar["median"])],
    [f"Orders above 12-month P95 ({gbp(R['inv_p95'])})", f"{int(bm_mar['large_orders(>12m P95)'])}",
     f"{int(bm_apr['large_orders(>12m P95)'])}", sint(bm_apr['large_orders(>12m P95)'] - bm_mar['large_orders(>12m P95)'])],
    ["UK net sales, whole month", gbp(uk["net_prev"]), gbp(uk["net_worst"]), pct(uk["change_%"])],
    ["UK net sales per invoice day", gbp(uk["per_day_prev"]), gbp(uk["per_day_worst"]), pct(uk["per_day_change_%"])],
    ["International net sales, whole month", gbp(intl["net_prev"]), gbp(intl["net_worst"]), pct(intl["change_%"])],
    ["International net sales per invoice day", gbp(intl["per_day_prev"]), gbp(intl["per_day_worst"]), pct(intl["per_day_change_%"])],
], widths=[6.4, 2.9, 2.9, 4.1], caption=f"Table 2. March vs April 2011, all values in GBP unless stated. Per-invoice-day "
   f"values divide by all invoice days in the month ({R['days_prev']} and {R['days_worst']}).", right_cols=(1, 2, 3))
para(f"**Method B – order values.** Across {R['inv_n']:,} sales invoices in the complete months, the median order was "
     f"{gbp(R['inv_median'])} and half of all orders lay between {gbp(R['inv_q1'])} and {gbp(R['inv_q3'])}. The mean "
     f"({gbp(R['inv_mean'])}) sits above the median because of the long upper tail. Figure 2 shows that the typical "
     f"order in April was the same as in March. The fall came from {R['inv_prev'] - R['inv_worst']:,} fewer orders and "
     f"from large orders: their value fell by {kgbp(abs(large_drop))} ({gbp(bm_mar['large_order_value'])} → "
     f"{gbp(bm_apr['large_order_value'])}).")
figure("fig5_invoice_value_distribution", 13.2,
       "Figure 2. Value per sales invoice by complete month (GBP, log scale; box = Q1–Q3, line = median, whiskers = 5th–95th "
       "percentile; dashed line = 12-month P95). March and April 2011 are dark blue. Source: notebook Step 6B.")
figure("fig6_country_contributions", 12.8,
       f"Figure 3. Country contributions to the {gbp(R['total_change'])} change, March → April 2011 (GBP; red = decline, "
       f"blue = offset). The {R['n_countries_pair']} contributions sum exactly to the total. Source: notebook Step 6C.")
para(f"**Method C – countries (Figure 3).** The United Kingdom contributed {gbp(uk['change_gbp'])} "
     f"({uk['share_of_change_%']:.0f}% of the fall): its whole-month net sales fell {pct(uk['change_%'], signed=False)}, "
     f"but per invoice day they changed by only {pct(uk['per_day_change_%'])} ({gbp(uk['per_day_prev'])} → "
     f"{gbp(uk['per_day_worst'])}). International customers made up only {R['intl_share_mar']:.0f}% of March sales, yet "
     f"they contributed {gbp(intl['change_gbp'])} ({intl['share_of_change_%']:.0f}%): {pct(intl['change_%'], signed=False)} "
     f"for the whole month and {pct(intl['per_day_change_%'], signed=False)} per invoice day. The Netherlands ({gbp(C['Netherlands'])}), Australia ({gbp(C['Australia'])}), "
     f"EIRE ({gbp(C['EIRE'])}) and France ({gbp(C['France'])}) explain {gbp(R['top_intl_sum'])}. Most of these decreases come from a "
     f"few accounts that ordered less in April: Netherlands customer 14646 ({gbp(acc['14646']['prev'])} → {gbp(acc['14646']['worst'])}) and Australia "
     f"customer 12415 ({gbp(acc['12415']['prev'])} → {gbp(acc['12415']['worst'])}, net of a credit). Japan "
     f"({sgbp(C['Japan'])}), Singapore ({sgbp(C['Singapore'])}) and Hong Kong ({sgbp(C['Hong Kong'])}) were small offsets. "
     f"Countries present in only one month are included with zero for the other month.")
figure("fig2_region_monthly", 12.8,
       f"Figure 4. Net invoiced sales by region (GBP), complete months December 2010 – November 2011. Source: notebook Step 5.")
para(f"Figure 4 puts April in context. International net sales reached their lowest complete month in April 2011 "
     f"({gbp(intl['net_worst'])}, {R['intl_share_apr']:.1f}% of net sales against an average share of "
     f"{R['intl_share_avg']:.1f}%) and were higher again in May. The UK line dips in April; per invoice day it changed "
     f"by {pct(uk['per_day_change_%'])}. "
     f"Over the 12 complete months, the number of sales invoices and net sales move together closely (r = "
     f"{R['r_invoices_sales']:.2f}): the KPI tracks order volume more closely than order size.")
table([
    ["Method", "Why it was chosen", "Computed result", "One limitation"],
    ["A. Time-series comparison", "The question asks when sales fell between adjacent complete months",
     f"Mar → Apr 2011 {gbp(R['worst_change'])} ({pct(R['worst_pct'])}); per invoice day {pct(R['perday_pct'])}",
     "One year of data: cannot separate seasonality from a real problem"],
    ["B. Distribution analysis", "Separates fewer orders from smaller orders; median is robust to outliers",
     f"Median order {gbp(bm_mar['median'])} → {gbp(bm_apr['median'])}; orders {R['inv_prev']:,} → {R['inv_worst']:,}; "
     f"large orders {int(bm_mar['large_orders(>12m P95)'])} → {int(bm_apr['large_orders(>12m P95)'])}",
     "An invoice is not a customer; credits are not matched to the order they reverse"],
    ["C. Country contribution", "Shows where the change came from; contributions sum exactly to the total",
     f"UK {gbp(uk['change_gbp'])} (month {pct(uk['change_%'])}; per invoice day {pct(uk['per_day_change_%'])}); "
     f"international {gbp(intl['change_gbp'])} (month {pct(intl['change_%'])}; per invoice day {pct(intl['per_day_change_%'])})",
     "Shows where, not why; one large account can dominate a country"],
], widths=[3.2, 4.3, 5.0, 3.8], caption="Table 3. The three descriptive methods: reason, result and limitation (notebook Step 6)")
sub("G", "Qualitative Impact")
para("Management can now see that the worst month of the year coincides with six fewer days with recorded invoices and "
     "with lower orders from a short list of international wholesale accounts that can be named and contacted. UK net sales "
     f"per invoice day changed by {pct(uk['per_day_change_%'])} ({gbp(uk['per_day_prev'])} → {gbp(uk['per_day_worst'])}), "
     "although the records cannot show whether UK demand itself changed. "
     "This turns a general worry about low KPIs into a specific, checkable list of questions. The cleaned and reconciled KPI "
     "also gives a common, auditable definition of 'sales' for future reviews, and partial months and calendar gaps can be "
     "separated from performance changes before decisions are made. No customer feedback, satisfaction, savings or revenue gain was measured. These are "
     "possible future benefits of the proposed action, not results of this study.")

# ----------------------------------------------------------------------------- VI. Recommendations
section("VI.", "Recommendations")
para("**Priority recommendation.** The Head of Sales and the international account managers should review the international "
     "wholesale accounts that ordered less in April 2011, starting with the four largest country contributors in Figure 3 "
     "(Netherlands, Australia, EIRE and France). Before that, Finance should confirm the trading calendar, so that the review "
     "starts from a calendar-adjusted baseline. Table 4 gives a proposed three-week management timeline. This is a proposed "
     "business action; its benefits have not been measured and must be tracked with the KPI in Table 4.")
table([
    ["When", "Who", "What to investigate", "Output"],
    ["Week 1", "Finance and sales operations",
     f"Were 22–25 and 29–30 April 2011 non-trading days? Why did credits rise from {gbp(mar['credit_magnitude'])} to "
     f"{gbp(apr['credit_magnitude'])}?", "Calendar-adjusted April baseline"],
    ["Week 2", "Head of Sales, account managers",
     "Contact accounts NL 14646, AU 12415, EIRE 14156 and 14911, and the main French accounts. Why were their April "
     "orders lower (stock, price, timing, other)?", "Reasons, account by account"],
    ["Week 3", "Executive management", "Decide recovery actions per account; approve the KPI", "Decision memo"],
    ["Monthly", "Sales operations", "Net sales per invoice day, UK vs international; credits as % of positive sales",
     "KPI review on 1st working day"],
], widths=[1.6, 3.6, 7.6, 3.5], caption="Table 4. Proposed implementation timeline and metric to monitor")
sub("H", "Key Takeaways")
bullet("**Normalise before comparing.** Always compare complete months and also check sales per invoice day. Valued at "
       f"March's daily rate, April's six fewer days with recorded invoices correspond arithmetically to about "
       f"{days_share:.0f}% of the fall.")
bullet("**Large accounts carry concentrated risk.** The top 5% of orders bring "
       f"{R['inv_large_share']:.0f}% of value, and four international countries explain {abs(R['top_intl_sum'] / R['total_change'] * 100):.0f}% "
       "of the worst fall. Because the median order did not change, account-level order counts are the first thing to check.")
bullet("**Contributions must add up.** A signed country breakdown that reconciles to the total shows clearly where the money "
       "moved, and it is easy for executives to audit.")
sub("I", "Challenges Overcome")
bullet(f"**Credits and adjustments:** the raw data contained {R['credit_lines_raw']:,} credit-note lines. "
       f"{R['retained_credit_lines']:,} met the merchandise eligibility rules and were kept, each counted once as a negative "
       f"value (Table 1); the other {R['excluded_credit_lines']:,} were excluded "
       + (f"by R3 because they carry non-merchandise codes. " if list(R['credit_excluded_by_rule']) == ['R3'] else
          f"({'; '.join(f'{n:,} by {k}' for k, n in sorted(R['credit_excluded_by_rule'].items()))}). ")
       + f"{R['zero_price'] + R['negative_price']:,} zero- or negative-price adjustment lines were excluded by R1. Two huge "
       f"orders ({jar['qty']:,} and {bird['qty']:,} units) were each cancelled within minutes and net to zero in their month.")
bullet(f"**Hidden scope problems:** non-merchandise codes (e.g., Amazon fees up to £17,836 per line) were removed. The effect "
       f"of this choice and of the {R['duplicate_rows']:,} repeated rows was tested, and the conclusion did not change.")
table([["Cleaning scope (notebook Step 6A)", "Lines", "Largest decline", "Change", "Gap to 2nd (Dec → Jan)"]] + [
    [x["scope"], f"{x['lines']:,}", x["largest decline"].replace("→", " → "),
     f"{gbp(x['change_gbp'])} ({pct(x['change_pct'])})", gbp(x["gap_to_2nd_gbp"])] for x in sens],
    widths=[5.4, 1.9, 2.8, 3.6, 2.6], caption="Table 5. Sensitivity of the main result to cleaning decisions",
    right_cols=(1, 3, 4))
bullet(f"**Wrong metadata and partial periods:** UCI says there are no missing values, yet {R['missing_customer']:,} "
       "CustomerIDs are blank. These lines were kept for totals. December 2011 was compared only like-for-like.")
bullet("**Data access:** the notebook uses the official UCI import and automatically falls back to UCI's own workbook. "
       "Both routes were run and gave identical results.")

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(8)
p.paragraph_format.keep_with_next = True
add_runs(p, "**References**", size=11.5, color=BLUE, font=HEAD_FONT)
for ref in [
    "[1] Chen, D. (2015). Online Retail [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5BW33",
    "[2] Chen, D., Sain, S. L., & Guo, K. (2012). Data mining for the online retail industry: A case study of RFM model-based "
    "customer segmentation using data mining. Journal of Database Marketing & Customer Strategy Management, 19(3), 197–208. "
    "https://doi.org/10.1057/dbm.2012.17",
    "[3] GOV.UK (2010). PM welcomes announcement of date for Royal Wedding. "
    "https://www.gov.uk/government/news/pm-welcomes-announcement-of-date-for-royal-wedding",
    f"[4] Saleh, K. (2026). {NOTEBOOK} – executed analysis notebook submitted with this report.",
]:
    q = para(ref, align=WD_ALIGN_PARAGRAPH.LEFT, size=9.5, after=2)
    q.paragraph_format.left_indent = Cm(0.6)
    q.paragraph_format.first_line_indent = Cm(-0.6)

doc.core_properties.title = "CSE464 Case Study 1 – Kenan Saleh"
doc.core_properties.author = "Kenan Saleh"
doc.core_properties.comments = ""
doc.core_properties.subject = "CSE464 Case Study 1"
doc.core_properties.keywords = ""
doc.core_properties.last_modified_by = "Kenan Saleh"
import datetime as _dt
doc.core_properties.created = doc.core_properties.modified = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0, tzinfo=None)
doc.core_properties.revision = 1
doc.save(DOCX_OUT)
print("Saved", DOCX_OUT)

# ----------------------------------------------------------------------------- PDF via LibreOffice
with tempfile.TemporaryDirectory() as tmp:
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(DOCX_OUT)],
                   check=True, capture_output=True)
    PDF_OUT.parent.mkdir(exist_ok=True)
    shutil.move(str(Path(tmp) / (DOCX_OUT.stem + ".pdf")), PDF_OUT)
print("Saved", PDF_OUT)
