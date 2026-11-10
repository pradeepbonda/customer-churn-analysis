"""
Customer Churn Analysis
1. Loads the Telco customer data into a SQLite database.
2. Runs every query in sql/churn_analysis.sql and saves each result as a CSV.
3. Builds an Excel dashboard (outputs/churn_dashboard.xlsx) whose numbers are
   live COUNTIFS formulas over the customer data.

Run from the project folder:  python src/run_analysis.py
"""
import re
import sqlite3
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "telco_customer_churn.csv"
SQL_FILE = ROOT / "sql" / "churn_analysis.sql"
RESULTS = ROOT / "outputs" / "query_results"
DASHBOARD = ROOT / "outputs" / "churn_dashboard.xlsx"


# ---------- 1. Load and clean ----------
def load_data() -> pd.DataFrame:
    df = pd.read_csv(DATA)
    # TotalCharges is blank for 11 brand-new customers (tenure 0) -> treat as 0.
    df["TotalCharges"] = pd.to_numeric(df["TotalCharges"], errors="coerce").fillna(0)
    df["SeniorCitizen"] = df["SeniorCitizen"].map({1: "Yes", 0: "No"})
    # snake_case column names for SQL
    df = df.rename(columns={"customerID": "customer_id", "StreamingTV": "streaming_tv"})
    df.columns = [re.sub(r"(?<!^)(?=[A-Z])", "_", c).lower() for c in df.columns]
    return df


# ---------- 2. Run SQL ----------
def run_queries(df: pd.DataFrame) -> dict:
    con = sqlite3.connect(":memory:")
    df.to_sql("customers", con, index=False)
    text = SQL_FILE.read_text()
    blocks = re.split(r"^-- name:\s*", text, flags=re.M)[1:]
    results = {}
    RESULTS.mkdir(parents=True, exist_ok=True)
    for block in blocks:
        name, sql = block.split("\n", 1)
        name = name.strip()
        result = pd.read_sql_query(sql, con)
        result.to_csv(RESULTS / f"{name}.csv", index=False)
        results[name] = result
        print(f"\n=== {name} ===\n{result.to_string(index=False)}")
    con.close()
    return results


# ---------- 3. Excel dashboard ----------
FONT = "Arial"
NAVY = "1F3A5F"
LIGHT = "EAF0F6"
thin = Side(style="thin", color="C8D2DC")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def tenure_band(t: int) -> str:
    if t <= 12:
        return "0-12 months"
    if t <= 24:
        return "13-24 months"
    if t <= 48:
        return "25-48 months"
    return "49-72 months"


def build_dashboard(df: pd.DataFrame, results: dict):
    wb = Workbook()

    # --- Data sheet: the full customer table ---
    data = wb.active
    data.title = "Data"
    cols = ["customer_id", "gender", "senior_citizen", "tenure", "contract",
            "internet_service", "tech_support", "online_security",
            "payment_method", "monthly_charges", "total_charges", "churn"]
    out = df[cols].copy()
    out.insert(4, "tenure_band", out["tenure"].apply(tenure_band))
    headers = ["Customer ID", "Gender", "Senior", "Tenure (months)", "Tenure Band",
               "Contract", "Internet Service", "Tech Support", "Online Security",
               "Payment Method", "Monthly Charge ($)", "Total Charges ($)", "Churn"]
    data.append(headers)
    for row in out.itertuples(index=False):
        data.append(list(row))
    style_header(data, 1, len(headers))
    for i, w in enumerate([13, 9, 8, 10, 14, 15, 15, 13, 15, 24, 12, 13, 8], 1):
        data.column_dimensions[get_column_letter(i)].width = w
    data.freeze_panes = "A2"
    data.auto_filter.ref = data.dimensions
    n = len(out) + 1  # last data row
    for r in range(2, n + 1):
        data.cell(row=r, column=11).number_format = "$#,##0.00"
        data.cell(row=r, column=12).number_format = "$#,##0.00"
    for row in data.iter_rows():
        for c in row:
            c.font = Font(name=FONT, size=10, bold=c.row == 1, color="FFFFFF" if c.row == 1 else "000000")

    rng = lambda col: f"Data!${col}$2:${col}${n}"
    CHURN, CONTRACT, BAND, NET, PAY, TENURE, MONTHLY = (
        rng("M"), rng("F"), rng("E"), rng("G"), rng("J"), rng("D"), rng("K"))

    # --- Dashboard sheet ---
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Customer Churn Dashboard"
    ws["B2"].font = Font(name=FONT, size=18, bold=True, color=NAVY)
    ws["B3"] = f"Telco customer data, {len(df):,} customers. Every number below is a live formula over the Data sheet."
    ws["B3"].font = Font(name=FONT, size=10, italic=True, color="5A6B7D")

    # KPI tiles
    kpis = [
        ("Customers", f"=COUNTA({CHURN})", "#,##0"),
        ("Churned", f'=COUNTIF({CHURN},"Yes")', "#,##0"),
        ("Churn rate", "=C6/B6", "0.0%"),
        ("Revenue lost per month", f'=SUMIFS({MONTHLY},{CHURN},"Yes")', "$#,##0"),
    ]
    for i, (label, formula, fmt) in enumerate(kpis):
        col = 2 + i
        lab = ws.cell(row=5, column=col, value=label)
        lab.font = Font(name=FONT, size=10, bold=True, color="5A6B7D")
        lab.alignment = Alignment(horizontal="center")
        val = ws.cell(row=6, column=col, value=formula)
        val.font = Font(name=FONT, size=16, bold=True, color=NAVY)
        val.number_format = fmt
        val.alignment = Alignment(horizontal="center")
        for r in (5, 6):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=LIGHT)
            ws.cell(row=r, column=col).border = BOX

    def breakdown(start_row, title, categories, crit_range):
        ws.cell(row=start_row, column=2, value=title).font = Font(name=FONT, size=12, bold=True, color=NAVY)
        hr = start_row + 1
        for j, h in enumerate(["Group", "Customers", "Churned", "Churn rate"], 2):
            ws.cell(row=hr, column=j, value=h)
        for j in range(2, 6):
            c = ws.cell(row=hr, column=j)
            c.font = Font(name=FONT, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.border = BOX
            c.alignment = Alignment(horizontal="center")
        for k, cat in enumerate(categories):
            r = hr + 1 + k
            ws.cell(row=r, column=2, value=cat)
            ws.cell(row=r, column=3, value=f'=COUNTIF({crit_range},B{r})')
            ws.cell(row=r, column=4, value=f'=COUNTIFS({crit_range},B{r},{CHURN},"Yes")')
            ws.cell(row=r, column=5, value=f"=IF(C{r}=0,0,D{r}/C{r})")
            ws.cell(row=r, column=5).number_format = "0.0%"
            ws.cell(row=r, column=3).number_format = "#,##0"
            ws.cell(row=r, column=4).number_format = "#,##0"
            for j in range(2, 6):
                ws.cell(row=r, column=j).font = Font(name=FONT, size=10)
                ws.cell(row=r, column=j).border = BOX
        first, last = hr + 1, hr + len(categories)

        chart = BarChart()
        chart.type = "bar"
        chart.title = f"Churn rate by {title.split(' by ')[-1].lower()}"
        chart.style = 2
        chart.legend = None
        chart.y_axis.numFmt = "0%"
        chart.y_axis.majorGridlines = None
        chart.add_data(Reference(ws, min_col=5, min_row=hr, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=2, min_row=first, max_row=last))
        chart.series[0].graphicalProperties.solidFill = NAVY
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.height, chart.width = 5.2, 12
        ws.add_chart(chart, f"G{start_row}")
        return last + 3

    r = 9
    r = breakdown(r, "Churn by contract", ["Month-to-month", "One year", "Two year"], CONTRACT)
    r = breakdown(r, "Churn by tenure band",
                  ["0-12 months", "13-24 months", "25-48 months", "49-72 months"], BAND)
    r = breakdown(r, "Churn by internet service", ["Fiber optic", "DSL", "No"], NET)
    r = breakdown(r, "Churn by payment method",
                  ["Electronic check", "Mailed check", "Bank transfer (automatic)",
                   "Credit card (automatic)"], PAY)

    # High-risk segment block
    ws.cell(row=r, column=2, value="High-risk segment: month-to-month, first 12 months, fiber optic").font = \
        Font(name=FONT, size=12, bold=True, color="B03A2E")
    seg = f'{CONTRACT},"Month-to-month",{TENURE},"<=12",{NET},"Fiber optic"'
    count_cell = f"D{r+1}"  # segment customer count
    rows = [
        ("Customers in segment", f"=COUNTIFS({seg})", "#,##0"),
        ("Share of all customers", f"={count_cell}/$B$6", "0.0%"),
        ("Churn rate in segment", f'=IF({count_cell}=0,0,COUNTIFS({seg},{CHURN},"Yes")/{count_cell})', "0.0%"),
        ("Share of all churners", f'=COUNTIFS({seg},{CHURN},"Yes")/$C$6', "0.0%"),
    ]
    for k, (label, formula, fmt) in enumerate(rows):
        rr = r + 1 + k
        ws.cell(row=rr, column=2, value=label).font = Font(name=FONT, size=10, bold=True)
        cell = ws.cell(row=rr, column=4, value=formula)
        cell.number_format = fmt
        cell.font = Font(name=FONT, size=10, bold=True, color="B03A2E")
        for j in (2, 3, 4):
            ws.cell(row=rr, column=j).border = BOX
            ws.cell(row=rr, column=j).fill = PatternFill("solid", fgColor="FBEDEC")

    for col, w in zip("ABCDEF", [2, 30, 13, 13, 22, 2]):
        ws.column_dimensions[col].width = w

    # --- Recommendations sheet ---
    rec = wb.create_sheet("Recommendations", 1)
    rec.sheet_view.showGridLines = False
    rec["B2"] = "Retention recommendations"
    rec["B2"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    items = [
        ("Finding", "Recommendation"),
        ("Month-to-month customers churn at about 43%, versus 11% on one-year and 3% on two-year contracts.",
         "Offer a discount or bonus for switching to a one-year contract, targeted at month-to-month customers in their first year."),
        ("Nearly half of customers leave within their first 12 months (47% churn in that band).",
         "Build a first-90-days onboarding program: welcome call, setup check, and a 3-month satisfaction check-in."),
        ("The high-risk segment (month-to-month, first year, fiber optic) is about 13% of customers but about 34% of all churners, with roughly 70% churn.",
         "Make this segment the first target for retention offers; it gives the most churn reduction per dollar spent."),
        ("Fiber optic customers churn at 42% despite paying the most.",
         "Review fiber pricing and service quality; survey fiber customers who leave to find the reason."),
        ("Customers without tech support or online security churn far more than those with both.",
         "Bundle tech support and online security free for the first 6 months for new internet customers."),
        ("Electronic check payers churn at 45%, versus 15-17% for automatic payments.",
         "Encourage automatic payment with a small monthly credit."),
    ]
    for k, (a, b) in enumerate(items):
        rr = 4 + k
        rec.cell(row=rr, column=2, value=a)
        rec.cell(row=rr, column=3, value=b)
        for j in (2, 3):
            c = rec.cell(row=rr, column=j)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOX
            c.font = Font(name=FONT, size=10, bold=k == 0, color="FFFFFF" if k == 0 else "000000")
            if k == 0:
                c.fill = PatternFill("solid", fgColor=NAVY)
        rec.row_dimensions[rr].height = 20 if k == 0 else 48
    rec.column_dimensions["A"].width = 2
    rec.column_dimensions["B"].width = 60
    rec.column_dimensions["C"].width = 60
    rec.cell(row=12, column=2, value="Rates come from the SQL results in outputs/query_results and the Dashboard sheet.").font = \
        Font(name=FONT, size=9, italic=True, color="5A6B7D")

    DASHBOARD.parent.mkdir(parents=True, exist_ok=True)
    wb.save(DASHBOARD)
    print(f"\nSaved {DASHBOARD}")


if __name__ == "__main__":
    customers = load_data()
    print("Columns:", list(customers.columns))
    query_results = run_queries(customers)
    build_dashboard(customers, query_results)
