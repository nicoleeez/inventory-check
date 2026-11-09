"""
skz_sales_report.py

Turns the export of my photocard stock tracker (one HTML file) into a sales report.

Every time I export the tracker, the file keeps a copy of the inventory as it was at that
moment (member, category, quantity and list price of every card).
When I sell a card, I lower its quantity in the tracker. So by comparing an earlier and a
later copy, every quantity drop is a sale.

What the script does:
  1. Reads all inventory copies stored in the HTML file.
  2. Builds a sales log (one row per quantity drop) and the current inventory.
  3. Runs data checks and lists discrepancies that need review.
  4. Writes CSV files for further analysis and an Excel report with live formulas.

Usage:
  python skz_sales_report.py tracker_export.html --out output_folder
"""

import argparse
import datetime as dt
import json
import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
SALES_START = dt.datetime(2026, 9, 1)   # August = registering new stock, not selling yet
DEFAULT_DISCOUNT = 0.10                 # bundle rule: 3 cards 5% off, 5+ cards 10% off
UPLOAD_IN_PROGRESS = {"Lee Know", "Seungmin", "Changbin", "I.N"}
MEMBERS = ["Bang Chan", "Lee Know", "Changbin", "Hyunjin", "Han",
           "Felix", "Seungmin", "I.N", "Unit", "Album"]
CATEGORY_NAMES = {"album": "Album photocard", "special": "Special / POB",
                  "other": "Event & merch card", "skzoo": "SKZOO",
                  "sealed": "Sealed album"}
PRICE_OUTLIER = 50.0                    # list prices above this are flagged for review

# Purchase cost, from my cost workbook (sheets 专辑 and 小卡), in RMB
RMB_PER_EUR = 7.4957                    # rate used in my own cost workbook
CARD_COST_RMB = (30973.56, 4229)        # total paid, number of photocards bought
ALBUM_COST_RMB = (5556.20, 253)         # total paid, number of albums bought


# ----------------------------------------------------------------------------
# 1. Read snapshots
# ----------------------------------------------------------------------------
def load_snapshots(html_path):
    """Return a list of snapshots. Each snapshot is the dict saved by the tracker."""
    html = Path(html_path).read_text(encoding="utf-8")
    html = re.sub(r"data:image/[^\"')]+", "", html)        # drop embedded images
    decoder = json.JSONDecoder()
    snapshots = []
    for script in re.findall(r"<script[^>]*>(.*?)</script>", html, re.S):
        # data scripts start with the assignment; the app code also mentions it later
        match = re.match(r"\s*window\.__EXPORTED_CARDS__\s*=\s*", script)
        if match:
            cards, _ = decoder.raw_decode(script, match.end())
            snapshots.append(cards)
    return snapshots


def snapshot_time(cards):
    """The tracker does not store a save date, so use the newest timestamp inside
    the snapshot (card IDs and image uploads both carry one)."""
    text = json.dumps(cards)
    stamps = [int(x) / 1000 for x in re.findall(r"-(1[78]\d{11})\b", text)]
    stamps += [int(x) for x in re.findall(r"/v(1[78]\d{8})/", text)]
    return dt.datetime.fromtimestamp(max(stamps), dt.timezone.utc).replace(tzinfo=None)


def card_label(card):
    label = card.get("album") or card.get("name") or ""
    if card.get("members"):
        members = card["members"]
        label += " (" + (", ".join(members) if isinstance(members, list) else str(members)) + ")"
    return label.strip()


def flatten(cards):
    """Snapshot dict -> {key: row}. Key = member|category|card id (ids repeat across members)."""
    rows = {}
    for member, categories in cards.items():
        for category, items in categories.items():
            for card in items:
                price = card.get("price")
                rows[f"{member}|{category}|{card['id']}"] = {
                    "member": member,
                    "category": CATEGORY_NAMES.get(category, category),
                    "item": card_label(card),
                    "qty": card.get("qty", 0),
                    "list_price": price if isinstance(price, (int, float)) else None,
                }
    return rows


# ----------------------------------------------------------------------------
# 2. Sales log, inventory, stock trend
# ----------------------------------------------------------------------------
def build_tables(snapshots):
    times = [snapshot_time(s) for s in snapshots]
    flats = [flatten(s) for s in snapshots]

    # last known price per card, used when a card sold before prices were tracked
    latest_price = {}
    for flat in flats:
        for key, row in flat.items():
            if row["list_price"] is not None:
                latest_price[key] = row["list_price"]

    sales, trend = [], []
    for i, flat in enumerate(flats):
        sold = added = setup = 0
        if i > 0:
            prev = flats[i - 1]
            for key, row in flat.items():
                change = row["qty"] - prev[key]["qty"] if key in prev else row["qty"]
                if change > 0:
                    added += change
                elif change < 0 and times[i - 1] < SALES_START:
                    setup += -change            # August stock setup, not a sale
                elif change < 0:
                    sold += -change
                    price, basis = row["list_price"], "Listed price at the time"
                    if price is None and key in latest_price:
                        price, basis = latest_price[key], "Later listed price"
                    elif price is None:
                        basis = "No price recorded"
                    sales.append({
                        "period_start": times[i - 1].date(), "recorded_by": times[i].date(),
                        "month": times[i - 1].strftime("%Y-%m"),
                        "member": row["member"], "category": row["category"],
                        "item": row["item"], "key": key, "units_sold": -change,
                        "list_price": price, "price_basis": basis,
                    })
        trend.append({
            "version_date": times[i].date(),
            "phase": "Setup (stock registration)" if times[i] < SALES_START else "Selling",
            "listings": len(flat), "units_in_stock": sum(r["qty"] for r in flat.values()),
            "units_sold": sold, "units_added": added, "setup_adjustments": setup,
        })

    sales = pd.DataFrame(sales).sort_values(["recorded_by", "member", "item"]).reset_index(drop=True)
    inventory = pd.DataFrame([{"key": k, **v} for k, v in flats[-1].items()])
    inventory = inventory.sort_values(["member", "category", "item"]).reset_index(drop=True)
    return sales, inventory, pd.DataFrame(trend), times, flats


def read_costs(path):
    """Read total spend and quantity from my cost workbook (summary cells of the
    sheets 专辑 and 小卡). Falls back to the values in Settings if no file is given."""
    if not path:
        return CARD_COST_RMB, ALBUM_COST_RMB
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    cards = (wb["小卡"]["G2"].value, wb["小卡"]["F2"].value)
    albums = (wb["专辑"]["H2"].value, wb["专辑"]["G2"].value)
    return cards, albums


# ----------------------------------------------------------------------------
# 3. Data checks
# ----------------------------------------------------------------------------
def run_checks(sales, inventory, times, flats):
    """Return a list of (check, result, action) rows."""
    checks = []

    n = int(sales.loc[sales.price_basis == "No price recorded", "units_sold"].sum())
    checks.append(("Units sold with no list price on record", n,
                   "Left out of revenue; prices were only tracked from late September."))

    n = int(sales.loc[sales.price_basis == "Later listed price", "units_sold"].sum())
    checks.append(("Units sold before a price was set (later price used)", n,
                   "Valued at the first list price recorded afterwards."))

    n = int(inventory.list_price.isna().sum())
    checks.append(("Listings in stock without a list price", n,
                   "Stock value excludes them; prices to be added in the tracker."))

    removed = sum(1 for i in range(1, len(flats)) for k in flats[i - 1] if k not in flats[i])
    checks.append(("Listings deleted from the tracker", removed,
                   "Not counted as sales, as a deletion can also be a correction."))

    n = int(((inventory.member == "Album") | (inventory.category == "Sealed album")).sum())
    checks.append(("Whole-album listings entered after part of the stock had sold", n,
                   "Their sales are understated; treated as a lower bound."))

    big = sales[sales.units_sold >= 3]
    checks.append(("Single listings dropping by 3 or more units at once", len(big),
                   "Kept as sales (likely bundle orders); to verify against Vinted order history."))

    n = int((inventory.list_price > PRICE_OUTLIER).sum())
    checks.append((f"List prices above €{PRICE_OUTLIER:.0f} (possible typo)", n,
                   "Flagged for review before reporting stock value."))
    return checks


# ----------------------------------------------------------------------------
# 4a. CSV export
# ----------------------------------------------------------------------------
def write_csvs(sales, inventory, trend, out):
    s = sales.copy()
    s["est_revenue"] = s.units_sold * s.list_price * (1 - DEFAULT_DISCOUNT)
    s["upload_status"] = s.member.map(lambda m: "In progress" if m in UPLOAD_IN_PROGRESS else "Complete")
    s.drop(columns="key").to_csv(out / "sales.csv", index=False)
    inv = inventory.copy()
    inv["stock_value"] = inv.qty * inv.list_price
    inv.drop(columns="key").to_csv(out / "inventory.csv", index=False)
    trend.to_csv(out / "stock_trend.csv", index=False)


# ----------------------------------------------------------------------------
# 4b. Excel report
# ----------------------------------------------------------------------------
F = "Arial"
HEAD, HEAD_FILL = Font(name=F, bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3A5F")
BLUE, BLACK, GREEN = Font(name=F, color="0000FF"), Font(name=F), Font(name=F, color="008000")
BOLD, NOTE = Font(name=F, bold=True), Font(name=F, italic=True, size=9, color="555555")
INPUT = PatternFill("solid", fgColor="FFFF00")
EUR, INT, PCT = "€#,##0;(€#,##0);-", "#,##0;(#,##0);-", "0%;(0%);-"
EUR2 = "€#,##0.00;(€#,##0.00);-"
LINE = Border(bottom=Side(style="thin", color="BBBBBB"))


def header(ws, row, names, widths=None, col=1):
    for j, name in enumerate(names):
        c = ws.cell(row=row, column=col + j, value=name)
        c.font, c.fill = HEAD, HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for j, w in enumerate(widths or []):
        ws.column_dimensions[get_column_letter(col + j)].width = w


def cell(ws, row, col, value, font=BLACK, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fmt:
        c.number_format = fmt
    return c


def write_excel(sales, inventory, trend, checks, times, costs, path):
    wb = Workbook()
    report = wb.active
    report.title = "Report"

    # --- Inventory -----------------------------------------------------------
    inv = wb.create_sheet("Inventory")
    header(inv, 1, ["Key", "Member", "Category", "Item", "Qty in stock", "List price (€)",
                    "Stock value (€)", "Unit cost (€)", "Stock cost (€)"],
           [10, 12, 20, 34, 11, 12, 14, 12, 13])
    inv.column_dimensions["A"].hidden = True
    for r, row in enumerate(inventory.itertuples(), 2):
        for j, v in enumerate([row.key, row.member, row.category, row.item, row.qty,
                               None if pd.isna(row.list_price) else row.list_price], 1):
            cell(inv, r, j, v, BLUE)
        cell(inv, r, 5, row.qty, BLUE, INT)
        inv.cell(row=r, column=6).number_format = EUR2
        cell(inv, r, 7, f'=IF(F{r}="","",E{r}*F{r})', fmt=EUR2)
        # albums use the album cost, everything else the photocard cost (can be overwritten)
        cell(inv, r, 8, f'=IF(OR(B{r}="Album",C{r}="Sealed album"),Report!$E$7,Report!$E$6)', GREEN, EUR2)
        cell(inv, r, 9, f'=IF(H{r}="","",E{r}*H{r})', fmt=EUR2)
    inv_last = len(inventory) + 1
    inv.freeze_panes, inv.auto_filter.ref = "B2", f"A1:I{inv_last}"

    # --- Sales_Log -----------------------------------------------------------
    sl = wb.create_sheet("Sales_Log")
    header(sl, 1, ["Period start", "Recorded by", "Month", "Member", "Category", "Item", "Key",
                   "Units sold", "List price (€)", "Price basis", "Revenue at list (€)",
                   "Actual price per unit (€)", "Est. revenue (€)", "Unit cost (€)", "Est. profit (€)",
                   "Cost of cards sold (€)"],
           [12, 12, 9, 12, 20, 34, 10, 9, 11, 22, 13, 13, 13, 11, 12, 13])
    sl.column_dimensions["G"].hidden = True
    disc = "Report!$C$4"
    for r, row in enumerate(sales.itertuples(), 2):
        vals = [row.period_start, row.recorded_by, row.month, row.member, row.category,
                row.item, row.key, row.units_sold,
                None if pd.isna(row.list_price) else row.list_price, row.price_basis]
        for j, v in enumerate(vals, 1):
            cell(sl, r, j, v, BLUE)
        sl.cell(row=r, column=1).number_format = sl.cell(row=r, column=2).number_format = "yyyy-mm-dd"
        sl.cell(row=r, column=9).number_format = EUR2
        cell(sl, r, 11, f'=IF(I{r}="","",H{r}*I{r})', fmt=EUR2)
        sl.cell(row=r, column=12).fill = INPUT
        sl.cell(row=r, column=12).number_format = EUR2
        cell(sl, r, 13, f'=IF(L{r}<>"",H{r}*L{r},IF(I{r}="","",H{r}*I{r}*(1-{disc})))', fmt=EUR2)
        look = f"INDEX(Inventory!$H$2:$H${inv_last},MATCH(G{r},Inventory!$A$2:$A${inv_last},0))"
        cell(sl, r, 14, f'=IFERROR({look},Report!$E$6)', GREEN, EUR2)
        cell(sl, r, 15, f'=IF(M{r}="","",M{r}-H{r}*N{r})', fmt=EUR2)
        cell(sl, r, 16, f'=IF(M{r}="","",H{r}*N{r})', fmt=EUR2)
    sl_last = len(sales) + 1
    sl.freeze_panes, sl.auto_filter.ref = "A2", f"A1:P{sl_last}"

    S = lambda c: f"Sales_Log!${c}$2:${c}${sl_last}"
    I = lambda c: f"Inventory!${c}$2:${c}${inv_last}"

    # --- Stock_Trend ---------------------------------------------------------
    st = wb.create_sheet("Stock_Trend")
    header(st, 1, ["Version date", "Phase", "Listings", "Units in stock", "Units sold since previous",
                   "Units added since previous", "Setup adjustments (not sales)"],
           [13, 26, 10, 13, 15, 15, 16])
    for r, row in enumerate(trend.itertuples(), 2):
        for j, v in enumerate(row[1:], 1):
            cell(st, r, j, v, BLUE, "yyyy-mm-dd" if j == 1 else (INT if j > 2 else None))
    t_last = len(trend) + 1
    line = LineChart()
    line.title, line.legend, line.height, line.width = "Units in stock over time", None, 7, 16
    line.add_data(Reference(st, min_col=4, min_row=1, max_row=t_last), titles_from_data=True)
    line.set_categories(Reference(st, min_col=1, min_row=2, max_row=t_last))
    st.add_chart(line, "I2")

    # --- Data_Checks ---------------------------------------------------------
    dc = wb.create_sheet("Data_Checks")
    cell(dc, 1, 1, "Data checks: discrepancies found in the tracker data", Font(name=F, bold=True, size=13))
    header(dc, 3, ["Check", "Result", "How it was handled"], [52, 10, 70])
    for r, (name, result, action) in enumerate(checks, 4):
        cell(dc, r, 1, name); cell(dc, r, 2, result, BLUE, INT); cell(dc, r, 3, action)
        for j in (1, 2, 3):
            dc.cell(row=r, column=j).border = LINE

    # --- Report (front page) -------------------------------------------------
    ws = report
    ws.sheet_view.showGridLines = False
    cell(ws, 1, 1, "Photocard shop: sales & inventory report", Font(name=F, bold=True, size=15))
    cell(ws, 2, 1, f"Vinted shop nicoleeezz · sales period 1 Sep to {times[-1]:%d %b %Y} · "
                   f"source: {len(times)} stored copies of the inventory", NOTE)
    cell(ws, 4, 1, "Assumption: average discount off list price", BOLD)
    c = cell(ws, 4, 3, DEFAULT_DISCOUNT, BLUE, "0%"); c.fill = INPUT
    cell(ws, 4, 4, "Bundle rule 3 cards 5%, 5+ cards 10%, plus small extra reductions. Change to test.", NOTE)
    cell(ws, 5, 1, "Exchange rate (RMB per €)", BOLD)
    c = cell(ws, 5, 3, RMB_PER_EUR, BLUE, "0.0000"); c.fill = INPUT
    cell(ws, 5, 4, "Rate used in my cost workbook.", NOTE)
    (card_rmb, card_n), (alb_rmb, alb_n) = costs
    for r, label, total, n in ((6, "Avg. purchase cost per photocard", card_rmb, card_n),
                               (7, "Avg. purchase cost per album", alb_rmb, alb_n)):
        cell(ws, r, 1, label, BOLD)
        c = cell(ws, r, 3, f"={total}/{n}", BLUE, '"¥"0.00'); c.fill = INPUT
        cell(ws, r, 5, f"=C{r}/$C$5", fmt=EUR2)
        cell(ws, r, 6, f"¥{total:,.2f} for {n:,} bought (cost workbook)", NOTE)
    cell(ws, 6, 4, "in €:", NOTE); cell(ws, 7, 4, "in €:", NOTE)

    # KPI tiles (row 6 labels, row 7 values)
    kpis = [("Units sold", f"=SUM({S('H')})", INT),
            ("Est. revenue", f"=SUM({S('M')})", EUR),
            ("Cost of cards sold", f"=SUM({S('P')})", EUR),
            ("Gross profit", f"=SUM({S('O')})", EUR),
            ("Gross margin", "=IFERROR(E10/C10,0)", PCT),
            ("Units in stock", f"=SUM({I('E')})", INT),
            ("Stock value", f"=SUM({I('G')})", EUR)]
    for j, (label, formula, fmt) in enumerate(kpis):
        col = 2 + j
        lab = cell(ws, 9, col, label, Font(name=F, size=9, color="555555"))
        val = cell(ws, 10, col, formula, Font(name=F, bold=True, size=14), fmt)
        lab.alignment = val.alignment = Alignment(horizontal="center")
        val.border = Border(bottom=Side(style="medium", color="1F3A5F"))

    # Monthly sales
    r0 = 13
    cell(ws, r0, 1, "Sales by month", BOLD)
    header(ws, r0 + 1, ["Month", "Units sold", "Est. revenue (€)", "Units per day"])
    months = [("2026-09", "September", 30), ("2026-10", "October", 31)]
    for n, (key, name, _) in enumerate(months):
        r = r0 + 2 + n
        cell(ws, r, 1, key, BLUE)
        cell(ws, r, 2, f"=SUMIFS({S('H')},{S('C')},$A{r})", fmt=INT)
        cell(ws, r, 3, f"=SUMIFS({S('M')},{S('C')},$A{r})", fmt=EUR)
        start = dt.date(int(key[:4]), int(key[5:]), 1)
        end = dt.date(2026, 10, 1) if key == "2026-09" else times[-1].date()
        cell(ws, r, 4, f"=IFERROR(B{r}/{(end - start).days},0)", fmt="0.0")
    month_rows = (r0 + 2, r0 + 3)
    cell(ws, r0 + 4, 1, f"October covers 1 Oct to {times[-1]:%d %b} (latest stored copy).", NOTE)

    # By member and by category
    def breakdown(start, title, keys, s_col, i_col, status=False):
        cell(ws, start, 1, title, BOLD)
        cols = [title.split(" by ")[-1].capitalize(), "Units sold", "Est. revenue (€)",
                "Share of revenue", "Units in stock", "Sell-through"]
        header(ws, start + 1, cols + (["Upload status"] if status else []))
        for n, key in enumerate(keys):
            r = start + 2 + n
            cell(ws, r, 1, key)
            cell(ws, r, 2, f"=SUMIFS({S('H')},{S(s_col)},$A{r})", fmt=INT)
            cell(ws, r, 3, f"=SUMIFS({S('M')},{S(s_col)},$A{r})", fmt=EUR)
            cell(ws, r, 4, f"=IFERROR(C{r}/$C$10,0)", fmt=PCT)
            cell(ws, r, 5, f"=SUMIFS({I('E')},{I(i_col)},$A{r})", fmt=INT)
            cell(ws, r, 6, f"=IFERROR(B{r}/(B{r}+E{r}),0)", fmt=PCT)
            if status:
                cell(ws, r, 7, "In progress" if key in UPLOAD_IN_PROGRESS else "Complete", BLUE)
            for j in range(1, 8 if status else 7):
                ws.cell(row=r, column=j).border = LINE
        return start + 2, start + 1 + len(keys)

    m_first, m_last = breakdown(20, "Sales by member", MEMBERS, "D", "B", status=True)
    cell(ws, m_last + 1, 1, "Sell-through = sold / (sold + in stock). Members marked 'In progress' "
                            "are not fully uploaded yet. Album = whole albums, entered after part of the stock had sold, so their sales are understated.", NOTE)
    cats = list(CATEGORY_NAMES.values())
    c_first, c_last = breakdown(m_last + 3, "Sales by category", cats, "E", "C")
    cell(ws, c_last + 1, 1, "Sealed album sales are understated: the albums were entered after part of the stock had sold.", NOTE)

    # Findings: text built from formulas so it updates with the data
    f0 = c_last + 2
    cell(ws, f0, 1, "Findings", BOLD)
    hj, bc = m_first + MEMBERS.index("Hyunjin"), m_first + MEMBERS.index("Bang Chan")
    alb, pob = c_first, c_first + 1
    sep, octo = month_rows
    findings = [
        f'="Sales rose from "&TEXT(D{sep},"0.0")&" to "&TEXT(D{octo},"0.0")&" cards per day from September to October, while more stock was being uploaded."',
        f'="Two members drive most revenue: Hyunjin "&TEXT(D{hj},"0%")&" and Bang Chan "&TEXT(D{bc},"0%")&"."',
        f'="Special / POB cards earn about €"&TEXT(C{pob}/B{pob},"0.00")&" per card, against €"&TEXT(C{alb}/B{alb},"0.00")&" for album cards."',
        f'="After purchase cost, the gross margin is about "&TEXT(F10,"0%")&": roughly €"&TEXT(E10,"#,##0")&" gross profit on €"&TEXT(C10,"#,##0")&" estimated revenue."',
        f'="Stock covers roughly "&TEXT(IFERROR(G10/D{octo}/30,0),"0")&" months of sales at the October pace."',
    ]
    for n, f in enumerate(findings):
        cell(ws, f0 + 1 + n, 1, f)

    for col, w in zip("ABCDEFGH", [32, 13, 14, 16, 14, 13, 14, 14]):
        ws.column_dimensions[col].width = w

    bar = BarChart()
    bar.type, bar.title, bar.legend, bar.height, bar.width = "bar", "Est. revenue by member (€)", None, 8, 13
    bar.add_data(Reference(ws, min_col=3, min_row=m_first - 1, max_row=m_last), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=m_first, max_row=m_last))
    ws.add_chart(bar, "J13")
    bar2 = BarChart()
    bar2.title, bar2.legend, bar2.height, bar2.width = "Units sold per month", None, 6, 13
    bar2.add_data(Reference(ws, min_col=2, min_row=r0 + 1, max_row=r0 + 3), titles_from_data=True)
    bar2.set_categories(Reference(ws, min_col=1, min_row=r0 + 2, max_row=r0 + 3))
    ws.add_chart(bar2, "J30")
    ws.print_area = f"A1:S{f0 + len(findings)}"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1

    # --- Notes ---------------------------------------------------------------
    nt = wb.create_sheet("Notes")
    notes = [
        ("How the data is built", True),
        ("Source: export of my own stock tracker website. The file keeps a copy of the full inventory from each earlier export.", False),
        ("A sale = a drop in a card's quantity between an earlier and a later copy (I lower the quantity after each sale).", False),
        ("August 2026 = registering stock bought in China; changes then are setup, not sales. Selling started in September.", False),
        ("Revenue = actual price if entered in Sales_Log, otherwise list price x (1 - average discount on Report).", False),
        ("Costs: I sell second-hand on Vinted, so there is no seller fee, and buyers pay shipping.", False),
        ("Purchase cost = average cost per photocard and per album from my cost workbook, converted from RMB.", False),
        ("Packaging materials are not included. Unit costs in Inventory can be overwritten card by card.", False),
        ("Sales made before a listing was uploaded are not captured, so all sales figures are a lower bound. This affects whole albums most: they were entered after part of the stock had sold.", False),
        ("Timing: I reduce stock in batches after processing a week or so of orders, so a sale shows on the processing date, not the order date. My income records count it only after the buyer confirms receipt.", False),
        ("", False),
        ("Legend", True),
        ("Yellow = input cells (Report assumptions, actual sale prices). Blue = data. Black = formulas. Green = pulled from another sheet.", False),
    ]
    for r, (t, b) in enumerate(notes, 1):
        cell(nt, r, 1, t, BOLD if b else BLACK)
    nt.column_dimensions["A"].width = 110

    wb.save(path)


# ----------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("html", help="tracker export (.html)")
    parser.add_argument("--out", default="output", help="output folder")
    parser.add_argument("--costs", help="optional: my cost workbook (.xlsx) with sheets 专辑 and 小卡")
    args = parser.parse_args()

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    snapshots = load_snapshots(args.html)
    sales, inventory, trend, times, flats = build_tables(snapshots)
    checks = run_checks(sales, inventory, times, flats)
    write_csvs(sales, inventory, trend, out)
    write_excel(sales, inventory, trend, checks, times, read_costs(args.costs), out / "photocard_sales_report.xlsx")

    print(f"Stored copies read: {len(snapshots)} ({times[0]:%d %b} to {times[-1]:%d %b %Y})")
    print(f"Sales records:  {len(sales)} ({int(sales.units_sold.sum())} units)")
    print(f"Listings now:   {len(inventory)} ({int(inventory.qty.sum())} units)")
    for name, result, _ in checks:
        print(f"  check | {name}: {result}")


if __name__ == "__main__":
    main()
