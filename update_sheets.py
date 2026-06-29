"""
update_sheets.py  —  US+ Health Inventory Forecast Dashboard
-------------------------------------------------------------
Pulls live FBA inventory from SP-API, aggregates 6 months of sales
from the local history file (or SP-API for missing months), runs the
forecast model, and writes everything to a live Google Sheet.

Run manually:   python update_sheets.py
Schedule weekly via Task Scheduler (see README).

Sheets written:
  - Dashboard          : KPI summary
  - Forecast & Reorder : Full SKU table with status + reorder quantities
  - Action Items       : Filtered to Reorder Now / Monitor, ready to assign
  - Stock History      : Weekly available-stock snapshots (append-only)
"""

import os, sys, json, datetime, time, gzip, math
from pathlib import Path
import certifi, httpx

# SSL fix for Windows — must happen before sp_api imports
_orig = httpx.Client
class _SSL(httpx.Client):
    def __init__(self, *a, **kw):
        kw.setdefault("verify", certifi.where())
        super().__init__(*a, **kw)
httpx.Client = _SSL

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from sp_api.api import Inventories, Reports
from sp_api.base import Marketplaces
from dotenv import load_dotenv

load_dotenv(dotenv_path=Path(__file__).parent / ".env")

# ── Constants ──────────────────────────────────────────────────────────────────

LEAD_DAYS     = 60
Z_SCORE       = 1.65   # 95% service level
TARGET_MONTHS = 2
SHEET_NAME    = "US+ Health - Inventory Forecast"
HISTORY_FILE  = Path(__file__).parent / "data" / "sales_history.json"

SP_CREDS = {
    "lwa_app_id":       os.environ["LWA_APP_ID"],
    "lwa_client_secret": os.environ["LWA_CLIENT_SECRET"],
    "refresh_token":    os.environ["SP_API_REFRESH_TOKEN"],
}
MARKETPLACE_ID = os.environ.get("MARKETPLACE_ID", "ATVPDKIKX0DER")
SSL_VERIFY     = certifi.where()

STATUS_COLORS = {
    "Reorder Now":        {"red": 1.0,   "green": 0.780, "blue": 0.808},
    "Monitor":            {"red": 1.0,   "green": 0.922, "blue": 0.612},
    "OK":                 {"red": 0.776, "green": 0.937, "blue": 0.808},
    "Covered by Inbound": {"red": 0.741, "green": 0.843, "blue": 0.933},
}
HEADER_COLOR  = {"red": 0.122, "green": 0.306, "blue": 0.475}   # #1F4E79
HEADER2_COLOR = {"red": 0.173, "green": 0.243, "blue": 0.314}   # darker accent

ASIN_NAMES = {
    "B08Y7X8375": "Hydrogen Peroxide 3% 32oz",
    "B08Y83DNZ5": "Hydrogen Peroxide 3% 1 Gal",
    "B097HP7DQ6": "Castor Oil 1 Gal",
    "B097LTHS4S": "Vegetable Glycerin 32oz",
    "B097LVPKMP": "Vegetable Glycerin 8oz",
    "B0981HX5NG": "Mineral Oil 8oz",
    "B09CV925V4": "Castor Oil 10oz",
    "B09DZ2P2WJ": "Sweet Almond Oil 1 Gal",
    "B09DZDD71G": "Sweet Almond Oil 10oz",
    "B0BJH3RD1F": "Mineral Oil 32oz",
    "B0BR99MF15": "Vegetable Glycerin 1 Gal",
    "B0CCMHLX72": "Castor Oil 20oz",
    "B0DDWQ1515": "Organic Castor Oil 16oz",
    "B0DSCKXPQH": "Organic Jojoba Oil 16oz",
}


def _product_label(asin: str, item_name: str) -> str:
    """Return 'Short Name (ASIN)' using the known name map, or truncated API title."""
    short = ASIN_NAMES.get(asin, "")
    if not short:
        short = (item_name[:48] + "...") if len(item_name) > 48 else item_name
    return f"{short} ({asin})" if short else asin


# ── Google Sheets auth ─────────────────────────────────────────────────────────

def _gc() -> gspread.Client:
    creds_path = os.environ["GOOGLE_CREDENTIALS_PATH"]
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return gspread.authorize(creds)


def get_or_create_sheet(gc: gspread.Client) -> gspread.Spreadsheet:
    sheet_id = os.environ["FORECAST_SHEET_ID"].strip()
    return gc.open_by_key(sheet_id)


# ── SP-API helpers ─────────────────────────────────────────────────────────────

def _marketplace():
    for m in Marketplaces:
        if m.value[0] == MARKETPLACE_ID:
            return m
    return Marketplaces.US


def fetch_inventory() -> pd.DataFrame:
    api = Inventories(credentials=SP_CREDS, marketplace=_marketplace(), verify=SSL_VERIFY)
    rows, next_token = [], None
    while True:
        kw = dict(details=True)
        if next_token:
            kw["nextToken"] = next_token
        resp = api.get_inventory_summary_marketplace(**kw).payload
        for item in resp.get("inventorySummaries", []):
            d  = item.get("inventoryDetails", {})
            rv = d.get("reservedQuantity", {})
            rows.append({
                "sku":       item.get("sellerSku", ""),
                "asin":      item.get("asin", ""),
                "item_name": item.get("productName", ""),
                "available": d.get("fulfillableQuantity", 0) or 0,
                "inbound":   (d.get("inboundShippedQuantity", 0) or 0)
                           + (d.get("inboundReceivingQuantity", 0) or 0),
                "reserved":  rv.get("totalReservedQuantity", 0) or 0,
            })
        next_token = resp.get("pagination", {}).get("nextToken")
        if not next_token:
            break
    return pd.DataFrame(rows)


def _consolidate_by_asin(df: pd.DataFrame) -> pd.DataFrame:
    """Merge rows that share the same ASIN (multiple SKUs → one row per product).
    Sums available/inbound/reserved. Picks the SKU with most available stock,
    preferring non-auto-generated 'Amazon.Found.*' SKUs."""
    if df.empty:
        return df
    result = []
    for asin, group in df.groupby("asin"):
        clean = group[~group["sku"].str.startswith("Amazon.Found.")]
        primary = clean if not clean.empty else group
        best = primary.loc[primary["available"].idxmax()]
        result.append({
            "sku":       best["sku"],
            "asin":      asin,
            "item_name": best["item_name"],
            "available": int(group["available"].sum()),
            "inbound":   int(group["inbound"].sum()),
            "reserved":  int(group["reserved"].sum()),
        })
    consolidated = pd.DataFrame(result)
    dupes = df.groupby("asin").size()
    dupes = dupes[dupes > 1]
    if not dupes.empty:
        print(f"  Merged {len(dupes)} ASINs with multiple SKUs:")
        for asin in dupes.index:
            skus = df[df["asin"] == asin]["sku"].tolist()
            print(f"    {asin}: {' + '.join(skus)}")
    return consolidated


def _target_months() -> list[str]:
    last = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
    months = []
    for _ in range(6):
        months.append(last.strftime("%Y-%m"))
        last = last.replace(day=1) - datetime.timedelta(days=1)
    return list(reversed(months))


def _history_monthly(asins: list[str], months: list[str]) -> pd.DataFrame:
    if not HISTORY_FILE.exists():
        return pd.DataFrame()
    with open(HISTORY_FILE) as f:
        records = json.load(f)
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df["month"] = pd.to_datetime(df["date"]).dt.to_period("M").astype(str)
    df = df[df["asin"].isin(asins) & df["month"].isin(months)]
    if df.empty:
        return pd.DataFrame()
    pivot = df.groupby(["asin", "month"])["units"].sum().unstack(fill_value=0)
    for m in months:
        if m not in pivot.columns:
            pivot[m] = 0
    return pivot[months].reset_index()


def _fetch_sp_month(year: int, month: int) -> pd.DataFrame:
    start = datetime.date(year, month, 1)
    end   = (datetime.date(year, month + 1, 1) if month < 12
             else datetime.date(year + 1, 1, 1)) - datetime.timedelta(days=1)
    api = Reports(credentials=SP_CREDS, marketplace=_marketplace(), verify=SSL_VERIFY)
    resp = api.create_report(
        reportType="GET_SALES_AND_TRAFFIC_REPORT",
        dataStartTime=start.isoformat() + "T00:00:00Z",
        dataEndTime=end.isoformat() + "T23:59:59Z",
        reportOptions={"dateGranularity": "MONTH", "asinGranularity": "CHILD"},
    )
    rid = resp.payload["reportId"]
    for _ in range(30):
        time.sleep(10)
        s = api.get_report(rid).payload
        if s["processingStatus"] == "DONE":
            break
        if s["processingStatus"] in ("FATAL", "CANCELLED"):
            raise RuntimeError(f"Report {rid} {s['processingStatus']}")
    doc = api.get_report_document(s["reportDocumentId"]).payload
    raw = httpx.get(doc["url"], verify=certifi.where(), timeout=60).content
    if doc.get("compressionAlgorithm", "").upper() == "GZIP":
        raw = gzip.decompress(raw)
    rows = []
    for entry in json.loads(raw.decode("utf-8")).get("salesAndTrafficByAsin", []):
        asin = entry.get("childAsin") or entry.get("parentAsin", "")
        rows.append({"asin": asin, "units": entry.get("salesByAsin", {}).get("unitsOrdered", 0)})
    return pd.DataFrame(rows)


def _save_sp_month_to_history(month_str: str, df_sp: pd.DataFrame):
    """Cache a fetched SP-API monthly report into the local history file."""
    records_to_add = [
        {"date": f"{month_str}-01", "asin": str(r["asin"]), "units": int(r["units"])}
        for _, r in df_sp.iterrows() if int(r.get("units", 0)) > 0
    ]
    if not records_to_add:
        return
    existing = []
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE) as f:
            existing = json.load(f)
    # Remove any previously-saved records for this month to avoid duplicates
    existing = [r for r in existing if not r.get("date", "").startswith(month_str)]
    existing.extend(records_to_add)
    with open(HISTORY_FILE, "w") as f:
        json.dump(existing, f)
    print(f"  Cached {len(records_to_add)} records for {month_str} → history file")


def get_monthly_sales(asins: list[str]) -> tuple[pd.DataFrame, list[str]]:
    months = _target_months()
    pivot  = _history_monthly(asins, months)
    if pivot.empty:
        pivot = pd.DataFrame({"asin": asins})
        for m in months:
            pivot[m] = 0
    missing = [m for m in months if pivot[m].sum() == 0]
    for m in missing:
        y, mo = int(m[:4]), int(m[5:])
        print(f"  Fetching SP-API report for {m}...")
        try:
            df_sp = _fetch_sp_month(y, mo)
            for _, r in df_sp.iterrows():
                if r["asin"] in pivot["asin"].values:
                    pivot.loc[pivot["asin"] == r["asin"], m] = int(r["units"])
            _save_sp_month_to_history(m, df_sp)  # cache so we never re-fetch this month
        except Exception as e:
            print(f"  Warning: could not fetch {m}: {e}")
    return pivot, months


# ── Forecast model ─────────────────────────────────────────────────────────────

def run_forecast(inventory: pd.DataFrame, sales: pd.DataFrame, months: list[str]) -> pd.DataFrame:
    df = inventory.merge(sales, on="asin", how="left")
    for m in months:
        df[m] = df.get(m, pd.Series(0, index=df.index)).fillna(0)

    r3 = months[-3:]
    p3 = months[-6:-3] if len(months) >= 6 else months[:3]

    df["avg_6mo"]      = df[months].mean(axis=1)
    df["avg_3mo"]      = df[r3].mean(axis=1)
    df["avg_prev3mo"]  = df[p3].mean(axis=1)
    df["trend"]        = df.apply(
        lambda r: (r["avg_3mo"] - r["avg_prev3mo"]) / r["avg_prev3mo"]
        if r["avg_prev3mo"] > 0 else 0, axis=1)
    df["forecast"]     = (df["avg_3mo"] * (1 + df["trend"])).clip(lower=0).round().astype(int)
    df["std_dev"]      = df[months].std(axis=1, ddof=1).fillna(0)
    lt_mo              = LEAD_DAYS / 30
    df["safety_stock"] = (Z_SCORE * df["std_dev"] * math.sqrt(lt_mo)).round().astype(int)
    df["reorder_point"]= (df["safety_stock"] + df["forecast"] * lt_mo).round().astype(int)
    # Order qty mirrors Venus's formula:
    #   - Only order if available < reorder_point (reserved stock doesn't inflate the trigger)
    #   - Order enough to cover TARGET_MONTHS demand, netting out available + inbound only
    df["order_qty"] = df.apply(
        lambda r: max(0, round(r["forecast"] * TARGET_MONTHS - r["available"] - r["inbound"]))
        if r["available"] < r["reorder_point"] else 0, axis=1)
    df["days_of_stock"]= df.apply(
        lambda r: round(r["available"] / (r["forecast"] / 30), 1)
        if r["forecast"] > 0 else (None if r["available"] == 0 else 9999), axis=1)

    def _status(r):
        if r["order_qty"] > 0:
            return "Reorder Now"
        if r["available"] < r["reorder_point"] and r["inbound"] > 0:
            return "Covered by Inbound"
        if r["available"] < r["reorder_point"] * 1.2:  # 20% buffer = early warning zone
            return "Monitor"
        return "OK"

    df["status"] = df.apply(_status, axis=1)
    return df


# ── Sheets writing helpers ─────────────────────────────────────────────────────

def _rgb(d: dict) -> dict:
    return {"red": d["red"], "green": d["green"], "blue": d["blue"]}


def _fmt_cell(bg=None, bold=False, fg=None, size=10, halign="LEFT", valign="MIDDLE", wrap=False) -> dict:
    fmt: dict = {
        "textFormat": {
            "bold": bold,
            "fontSize": size,
            **({"foregroundColor": _rgb(fg)} if fg else {}),
        },
        "horizontalAlignment": halign,
        "verticalAlignment": valign,
        "wrapStrategy": "WRAP" if wrap else "OVERFLOW_CELL",
    }
    if bg:
        fmt["backgroundColor"] = _rgb(bg)
    return fmt


def _merge_req(sheet_id: int, r1: int, c1: int, r2: int, c2: int) -> dict:
    return {
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL",
        }
    }


def _col_width_req(sheet_id: int, col: int, width_px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": col, "endIndex": col + 1},
            "properties": {"pixelSize": width_px},
            "fields": "pixelSize",
        }
    }


def _row_height_req(sheet_id: int, row: int, height_px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": row, "endIndex": row + 1},
            "properties": {"pixelSize": height_px},
            "fields": "pixelSize",
        }
    }


def _freeze_req(sheet_id: int, rows: int = 1, cols: int = 0) -> dict:
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def _format_range_req(sheet_id: int, r1: int, c1: int, r2: int, c2: int, fmt: dict) -> dict:
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fmt.keys()) + ")",
        }
    }


def _get_or_add_tab(ss: gspread.Spreadsheet, title: str) -> gspread.Worksheet:
    try:
        ws = ss.worksheet(title)
        ws.clear()
        return ws
    except gspread.WorksheetNotFound:
        return ss.add_worksheet(title=title, rows=200, cols=30)


def _delete_default_sheet(ss: gspread.Spreadsheet):
    try:
        ws = ss.worksheet("Sheet1")
        if len(ss.worksheets()) > 1:
            ss.del_worksheet(ws)
    except gspread.WorksheetNotFound:
        pass


# ── Write: Forecast & Reorder ──────────────────────────────────────────────────

def write_forecast_tab(ss: gspread.Spreadsheet, df: pd.DataFrame, months: list[str]):
    ws = _get_or_add_tab(ss, "Forecast & Reorder")
    sid = ws.id

    month_labels = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %Y") for m in months]

    def _clean(v):
        import math
        if v is None:
            return ""
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return ""
        return v

    # Row 1: assumptions
    ws.update([[
        f"Last updated: {datetime.datetime.now():%Y-%m-%d %H:%M}  |  "
        f"Lead time: {LEAD_DAYS} days  |  Service level: 95%  |  Target coverage: {TARGET_MONTHS} months"
    ]], "A1")

    # Row 2: headers
    HEADERS = (
        ["SKU", "ASIN", "Product (Name + ASIN)", "Lead Time (Days)",
         "Available Stock", "Inbound Stock", "Reserved Stock"]
        + month_labels
        + ["Avg Monthly Demand (6-mo)", "3-Mo Moving Avg", "Trend (%)",
           "Forecasted Demand", "Std Dev", "Safety Stock", "Reorder Point",
           "Order Qty", "Status", "Days of Stock"]
    )
    ws.update([HEADERS], "A2")

    # Data rows
    rows = []
    for _, row in df.iterrows():
        dos = row["days_of_stock"]
        r = (
            [row["sku"], row["asin"], _product_label(row["asin"], row.get("item_name", "")), LEAD_DAYS,
             int(row["available"]), int(row["inbound"]), int(row["reserved"])]
            + [int(row.get(m, 0)) for m in months]
            + [_clean(round(row["avg_6mo"], 1)), _clean(round(row["avg_3mo"], 1)),
               _clean(round(row["trend"], 4)),
               int(row["forecast"]), _clean(round(row["std_dev"], 1)),
               int(row["safety_stock"]), int(row["reorder_point"]),
               int(row["order_qty"]), row["status"],
               _clean(dos)]
        )
        rows.append(r)

    if rows:
        ws.update(rows, "A3")

    # Sparkline column header + formulas (trend line per SKU)
    spark_col_letter = chr(ord("A") + len(HEADERS))   # column right after last header
    ws.update([["Trend"]], f"{spark_col_letter}2")
    spark_formulas = []
    month_start_col = chr(ord("A") + 7)               # column H = first month
    month_end_col   = chr(ord("A") + 6 + len(months)) # last month column
    for ri in range(3, 3 + len(rows)):
        spark_formulas.append([
            f'=SPARKLINE({month_start_col}{ri}:{month_end_col}{ri},'
            f'{{"charttype","column";"color1","#1F4E79";"color2","#C0392B";"negcolor","#C0392B"}})'
        ])
    if spark_formulas:
        ws.update(spark_formulas, f"{spark_col_letter}3", value_input_option="USER_ENTERED")

    # ── Batch formatting ──
    reqs = []

    # Header row background + bold
    reqs.append(_format_range_req(sid, 1, 0, 2, len(HEADERS),
        _fmt_cell(bg=HEADER_COLOR, bold=True,
                  fg={"red": 1, "green": 1, "blue": 1},
                  halign="CENTER", wrap=True)))

    # Info row styling
    reqs.append(_format_range_req(sid, 0, 0, 1, len(HEADERS),
        _fmt_cell(bg={"red": 0.9, "green": 0.9, "blue": 0.9},
                  fg={"red": 0.3, "green": 0.3, "blue": 0.3}, size=9)))

    # Freeze header rows + SKU column
    reqs.append(_freeze_req(sid, rows=2, cols=1))

    # Header row height
    reqs.append(_row_height_req(sid, 1, 48))

    # Column widths
    col_widths = (
        [160, 110, 280, 80, 80, 80, 80]   # SKU, ASIN, Name, lead, avail, inbound, reserved
        + [70] * len(months)               # month columns
        + [90, 90, 70, 90, 70, 80, 90, 80, 130, 80]  # calc columns
    )
    for ci, w in enumerate(col_widths):
        reqs.append(_col_width_req(sid, ci, w))

    # Trend column: percent format
    trend_col = 7 + len(months) + 2  # 0-based
    reqs.append({
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 2 + len(rows),
                      "startColumnIndex": trend_col, "endColumnIndex": trend_col + 1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    })

    # Status cell colors (per row)
    status_col_0 = HEADERS.index("Status")
    for ri, (_, row) in enumerate(df.iterrows()):
        color = STATUS_COLORS.get(row["status"])
        if color:
            reqs.append(_format_range_req(
                sid, 2 + ri, status_col_0, 3 + ri, status_col_0 + 1,
                _fmt_cell(bg=color, bold=True, halign="CENTER")))

    # Sparkline column header format + width
    spark_col_idx = len(HEADERS)
    reqs.append(_col_width_req(sid, spark_col_idx, 100))

    ss.batch_update({"requests": reqs})

    # Format sparkline header (must happen after batch_update since _hdr writes directly)
    spark_col_letter2 = chr(ord("A") + len(HEADERS))
    ws.update([["6-Mo Trend"]], f"{spark_col_letter2}2")

    print(f"  'Forecast & Reorder' updated — {len(rows)} SKUs")


# ── Write: Dashboard ───────────────────────────────────────────────────────────

_CARD = {
    "red":    ({"red": 0.753, "green": 0.224, "blue": 0.169}, {"red": 1, "green": 1, "blue": 1}),
    "orange": ({"red": 0.878, "green": 0.459, "blue": 0.098}, {"red": 1, "green": 1, "blue": 1}),
    "green":  ({"red": 0.133, "green": 0.545, "blue": 0.133}, {"red": 1, "green": 1, "blue": 1}),
    "blue":   ({"red": 0.118, "green": 0.416, "blue": 0.635}, {"red": 1, "green": 1, "blue": 1}),
    "dark":   (HEADER_COLOR,                                   {"red": 1, "green": 1, "blue": 1}),
    "navy":   (HEADER2_COLOR,                                  {"red": 0.78, "green": 0.87, "blue": 0.95}),
}

WHITE      = {"red": 1,     "green": 1,     "blue": 1}
SHEET_GRAY = {"red": 0.953, "green": 0.953, "blue": 0.953}


def write_dashboard_tab(ss: gspread.Spreadsheet, df: pd.DataFrame):
    # ── Compute KPI values ──
    counts    = df["status"].value_counts()
    valid_dos = df["days_of_stock"].replace([float("inf"), -float("inf")], pd.NA).dropna()
    valid_dos = valid_dos[valid_dos < 9999]
    med_days  = int(valid_dos.median()) if not valid_dos.empty else 0
    avg_trend = df["trend"].mean()
    reorder_n = int(counts.get("Reorder Now", 0))
    monitor_n = int(counts.get("Monitor", 0))
    ok_n      = int(counts.get("OK", 0))
    covered_n = int(counts.get("Covered by Inbound", 0))
    total_n   = len(df)
    forecast_n = int(df["forecast"].sum())

    reorder_skus   = df[df["order_qty"] > 0].sort_values("order_qty", ascending=False)
    n_reorder_skus = len(reorder_skus)

    # ── Dashboard tab — delete and recreate to remove all old merges/formatting ──
    try:
        ss.del_worksheet(ss.worksheet("Dashboard"))
    except gspread.WorksheetNotFound:
        pass
    ws  = ss.add_worksheet(title="Dashboard", rows=100, cols=25, index=0)
    sid = ws.id

    # Layout (0-indexed rows):
    #  0  = title bar
    #  1  = subtitle bar
    #  2  = spacer (14px)
    #  3  = KPI pair 1 labels   (Reorder Now | Monitor)
    #  4  = KPI pair 1 values
    #  5  = spacer (12px)
    #  6  = KPI pair 2 labels   (OK | Covered)
    #  7  = KPI pair 2 values
    #  8  = spacer (12px)
    #  9  = KPI pair 3 labels   (Total SKUs | Forecast Next Month)
    #  10 = KPI pair 3 values
    #  11 = spacer (12px)
    #  12 = KPI pair 4 labels   (Median Days of Stock | Avg Forecast Trend)
    #  13 = KPI pair 4 values
    #  14+= spacer (charts float to the right)
    #
    # Columns: A(0)=margin | B-E(1-4)=left card | F(5)=gap | G-J(6-9)=right card | K(10)=margin

    kpi_rows = [
        ["US+ Health — Inventory Forecast Dashboard"] + [""] * 10,
        [f"Updated: {datetime.datetime.now():%b %d, %Y  %H:%M}   |   "
         f"{total_n} Active SKUs   |   {LEAD_DAYS}-Day Lead Time   |   95% Service Level"] + [""] * 10,
        [""] * 11,
        ["", "REORDER NOW",         "", "", "", "", "MONITOR",              "", "", "", ""],
        ["", reorder_n,             "", "", "", "", monitor_n,              "", "", "", ""],
        [""] * 11,
        ["", "OK",                  "", "", "", "", "COVERED BY INBOUND",   "", "", "", ""],
        ["", ok_n,                  "", "", "", "", covered_n,              "", "", "", ""],
        [""] * 11,
        ["", "TOTAL SKUS TRACKED",  "", "", "", "", "FORECAST NEXT MONTH",  "", "", "", ""],
        ["", total_n,               "", "", "", "", forecast_n,             "", "", "", ""],
        [""] * 11,
        ["", "MEDIAN DAYS OF STOCK","", "", "", "", "AVG FORECAST TREND",   "", "", "", ""],
        ["", med_days,              "", "", "", "", round(avg_trend, 4),    "", "", "", ""],
    ]
    ws.update(kpi_rows, "A1")

    # Chart data: write to cols M-N (12-13) so charts on same sheet can reference them.
    # These columns will be set to 2px wide (invisible) and hidden under the chart overlays.
    # Status pie data: M3:N7  (row indices 2-6 in API)
    ws.update([
        ["Status",            "Count"],
        ["Reorder Now",        reorder_n],
        ["Monitor",            monitor_n],
        ["OK",                 ok_n],
        ["Covered by Inbound", covered_n],
    ], "M3")
    # Reorder bar data: M9:N9+n  (row indices 8+ in API)
    ws.update(
        [["Product", "Order Qty"]] + [
            [_product_label(r["asin"], r.get("item_name", "")), int(r["order_qty"])]
            for _, r in reorder_skus.iterrows()
        ],
        "M9"
    )

    # ── Batch format requests ──
    reqs = []

    # Hide gridlines
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"hideGridlines": True}},
            "fields": "gridProperties.hideGridlines",
        }
    })

    # Sheet background: light gray everywhere
    reqs.append(_format_range_req(sid, 0, 0, 50, 25,
        _fmt_cell(bg=SHEET_GRAY)))

    # Title
    reqs.append(_merge_req(sid, 0, 0, 1, 11))
    reqs.append(_format_range_req(sid, 0, 0, 1, 11,
        _fmt_cell(bg=HEADER_COLOR, bold=True, fg=WHITE, size=18,
                  halign="CENTER", valign="MIDDLE")))
    reqs.append(_row_height_req(sid, 0, 58))

    # Subtitle
    reqs.append(_merge_req(sid, 1, 0, 2, 11))
    reqs.append(_format_range_req(sid, 1, 0, 2, 11,
        _fmt_cell(bg={"red": 0.145, "green": 0.255, "blue": 0.38},
                  fg={"red": 0.74, "green": 0.84, "blue": 0.95}, size=9, halign="CENTER")))
    reqs.append(_row_height_req(sid, 1, 30))

    # Spacer row 2
    reqs.append(_row_height_req(sid, 2, 14))

    # KPI card pairs: each pair = (label_row, val_row, spacer_row, left_theme, right_theme)
    kpi_pairs = [
        (3,  4,  5,  "red",  "orange"),
        (6,  7,  8,  "green","blue"),
        (9,  10, 11, "dark", "dark"),
        (12, 13, 14, "navy", "navy"),
    ]
    for (lr, vr, sr, lt, rt) in kpi_pairs:
        bg_l, fg_l = _CARD[lt]
        bg_r, fg_r = _CARD[rt]

        # Merge label cells: cols 1-5 (B-E) and 6-10 (G-J)
        reqs.append(_merge_req(sid, lr, 1, lr + 1, 5))
        reqs.append(_merge_req(sid, lr, 6, lr + 1, 10))
        # Merge value cells
        reqs.append(_merge_req(sid, vr, 1, vr + 1, 5))
        reqs.append(_merge_req(sid, vr, 6, vr + 1, 10))

        # Label formatting
        reqs.append(_format_range_req(sid, lr, 1, lr + 1, 5,
            _fmt_cell(bg=bg_l, bold=True,
                      fg={"red": fg_l["red"], "green": fg_l["green"],
                          "blue": fg_l["blue"], "alpha": 0.80},
                      size=9, halign="CENTER", valign="MIDDLE")))
        reqs.append(_format_range_req(sid, lr, 6, lr + 1, 10,
            _fmt_cell(bg=bg_r, bold=True,
                      fg={"red": fg_r["red"], "green": fg_r["green"],
                          "blue": fg_r["blue"], "alpha": 0.80},
                      size=9, halign="CENTER", valign="MIDDLE")))

        # Value formatting (big number)
        reqs.append(_format_range_req(sid, vr, 1, vr + 1, 5,
            _fmt_cell(bg=bg_l, bold=True, fg=fg_l, size=36,
                      halign="CENTER", valign="MIDDLE")))
        reqs.append(_format_range_req(sid, vr, 6, vr + 1, 10,
            _fmt_cell(bg=bg_r, bold=True, fg=fg_r, size=36,
                      halign="CENTER", valign="MIDDLE")))

        # Row heights: label=26, value=72, spacer=12
        reqs.append(_row_height_req(sid, lr, 26))
        reqs.append(_row_height_req(sid, vr, 72))
        reqs.append(_row_height_req(sid, sr, 12))

    # Avg Trend: format as percentage
    reqs.append({
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 13, "endRowIndex": 14,
                      "startColumnIndex": 6, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    })

    # Column widths — KPI area
    col_widths = {0: 16, 1: 88, 2: 88, 3: 88, 4: 88,
                  5: 16, 6: 88, 7: 88, 8: 88, 9: 88, 10: 16,
                  11: 10,   # gap before chart data cols
                  12: 60,   # M — chart label data
                  13: 60,   # N — chart value data
                  }
    for ci, w in col_widths.items():
        reqs.append(_col_width_req(sid, ci, w))

    # Properly hide chart-data columns M-N (cols 12-13) so they don't show
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 12, "endIndex": 14},
            "properties": {"hiddenByUser": True},
            "fields": "hiddenByUser",
        }
    })

    ss.batch_update({"requests": reqs})

    # Charts reference the Dashboard tab itself (same-sheet references always work)
    _add_dashboard_charts(ss, sid, n_reorder_skus)
    print("  'Dashboard' updated")


# ── Dashboard charts ───────────────────────────────────────────────────────────

def _add_dashboard_charts(ss: gspread.Spreadsheet, sid: int, n_reorder_skus: int):
    """Delete any existing Dashboard charts then re-add donut + bar using same-sheet data."""
    # Delete existing charts on Dashboard
    try:
        resp = ss.client.request("get",
            f"https://sheets.googleapis.com/v4/spreadsheets/{ss.id}",
            params={"fields": "sheets(properties/sheetId,charts)"})
        for sheet in resp.json().get("sheets", []):
            if sheet["properties"]["sheetId"] == sid:
                existing = sheet.get("charts", [])
                if existing:
                    ss.batch_update({"requests": [
                        {"deleteEmbeddedObject": {"objectId": c["chartId"]}}
                        for c in existing
                    ]})
                break
    except Exception:
        pass

    ref = {"sheetId": sid}

    # Chart data layout on Dashboard (0-indexed rows, cols M=12, N=13):
    #   Row 2  (M3:N3) : "Status" | "Count"   ← header (skipped)
    #   Rows 3-6       : status data           ← pie chart source
    #   Row 8  (M9:N9) : "Product" | "Order Qty" ← header (skipped)
    #   Rows 9-9+n     : reorder data          ← bar chart source

    # Pie chart — status distribution
    donut = {
        "addChart": {"chart": {
            "spec": {
                "title": "Inventory Status Distribution",
                "titleTextFormat": {"bold": True, "fontSize": 11,
                    "foregroundColor": _rgb(HEADER_COLOR)},
                "backgroundColor": _rgb(WHITE),
                "hiddenDimensionStrategy": "SHOW_ALL",
                "pieChart": {
                    "legendPosition": "RIGHT_LEGEND",
                    "threeDimensional": False,
                    "domain": {"sourceRange": {"sources": [{
                        **ref,
                        "startRowIndex": 3, "endRowIndex": 7,
                        "startColumnIndex": 12, "endColumnIndex": 13}]}},
                    "series": {"sourceRange": {"sources": [{
                        **ref,
                        "startRowIndex": 3, "endRowIndex": 7,
                        "startColumnIndex": 13, "endColumnIndex": 14}]}},
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {**ref, "rowIndex": 2, "columnIndex": 14},
                "widthPixels": 400, "heightPixels": 255,
            }},
        }}
    }

    # Bar chart — units to reorder by product
    n = max(n_reorder_skus, 1)
    bar = {
        "addChart": {"chart": {
            "spec": {
                "title": "Units to Reorder by Product",
                "titleTextFormat": {"bold": True, "fontSize": 11,
                    "foregroundColor": _rgb(HEADER_COLOR)},
                "backgroundColor": _rgb(WHITE),
                "hiddenDimensionStrategy": "SHOW_ALL",
                "basicChart": {
                    "chartType": "BAR",
                    "legendPosition": "NO_LEGEND",
                    "axis": [{"position": "BOTTOM_AXIS", "title": "Units to Order"}],
                    "domains": [{"domain": {"sourceRange": {"sources": [{
                        **ref,
                        "startRowIndex": 9, "endRowIndex": 9 + n,
                        "startColumnIndex": 12, "endColumnIndex": 13}]}}}],
                    "series": [{"series": {"sourceRange": {"sources": [{
                        **ref,
                        "startRowIndex": 9, "endRowIndex": 9 + n,
                        "startColumnIndex": 13, "endColumnIndex": 14}]}},
                        "targetAxis": "BOTTOM_AXIS",
                        "color": _rgb({"red": 0.753, "green": 0.224, "blue": 0.169}),
                    }],
                },
            },
            "position": {"overlayPosition": {
                "anchorCell": {**ref, "rowIndex": 8, "columnIndex": 14},
                "widthPixels": 400, "heightPixels": 295,
            }},
        }}
    }

    try:
        ss.batch_update({"requests": [donut, bar]})
        print("  Charts added to Dashboard")
    except Exception as e:
        print(f"  Warning: could not add charts: {e}")


# ── Write: Action Items ────────────────────────────────────────────────────────

def write_action_items_tab(ss: gspread.Spreadsheet, df: pd.DataFrame):
    ws = _get_or_add_tab(ss, "Action Items")
    sid = ws.id

    HEADERS = ["SKU", "ASIN", "Product (Name + ASIN)", "Status", "Order Qty",
               "Owner", "Target Date", "Priority", "Notes"]

    priority_order = {"Reorder Now": 0, "Monitor": 1, "OK": 2, "Covered by Inbound": 3}
    sdf = df[df["status"].isin(["Reorder Now", "Monitor"])].sort_values(
        "status", key=lambda s: s.map(priority_order).fillna(4))

    rows = [[
        r["sku"], r["asin"], _product_label(r["asin"], r.get("item_name", "")),
        r["status"], int(r["order_qty"]), "", "", "", ""
    ] for _, r in sdf.iterrows()]

    ws.update([["Action Items — fill in Owner, Target Date, Priority below"]], "A1")
    ws.update([HEADERS], "A2")
    if rows:
        ws.update(rows, "A3")

    reqs = []

    # Title row
    reqs.append(_format_range_req(sid, 0, 0, 1, len(HEADERS),
        _fmt_cell(bg=HEADER_COLOR, bold=True,
                  fg={"red": 1, "green": 1, "blue": 1}, size=11)))
    reqs.append(_row_height_req(sid, 0, 36))

    # Header row
    reqs.append(_format_range_req(sid, 1, 0, 2, len(HEADERS),
        _fmt_cell(bg=HEADER2_COLOR, bold=True,
                  fg={"red": 1, "green": 1, "blue": 1}, halign="CENTER")))

    # Freeze
    reqs.append(_freeze_req(sid, rows=2))

    # Status colors per row
    for ri, (_, row) in enumerate(sdf.iterrows()):
        color = STATUS_COLORS.get(row["status"])
        if color:
            reqs.append(_format_range_req(
                sid, 2 + ri, 3, 3 + ri, 4,
                _fmt_cell(bg=color, bold=True, halign="CENTER")))

    # Column widths
    for ci, w in enumerate([160, 110, 300, 140, 80, 120, 110, 90, 200]):
        reqs.append(_col_width_req(sid, ci, w))

    ss.batch_update({"requests": reqs})
    print(f"  'Action Items' updated — {len(rows)} flagged SKUs")


# ── Write: Stock Snapshots (daily available-stock log, append-only) ───────────

def write_stock_history_tab(ss: gspread.Spreadsheet, df: pd.DataFrame):
    TAB_NAME = "Stock Snapshots"

    # Migrate: delete old "Stock History" tab if it exists
    try:
        ss.del_worksheet(ss.worksheet("Stock History"))
    except gspread.WorksheetNotFound:
        pass

    # Get existing tab WITHOUT clearing — this tab is strictly append-only
    try:
        ws = ss.worksheet(TAB_NAME)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=TAB_NAME, rows=500, cols=50)
    sid = ws.id

    today    = datetime.date.today().isoformat()
    existing = ws.get_all_values()

    # First run: write header row with product name + ASIN
    if not existing or existing[0] == []:
        header = ["Date"] + [
            f"{ASIN_NAMES.get(r['asin'], r.get('item_name', r['asin'])[:28])}\n({r['asin']})"
            for _, r in df.iterrows()
        ]
        ws.update([header], "A1")
        reqs = [
            _format_range_req(sid, 0, 0, 1, len(header),
                _fmt_cell(bg=HEADER_COLOR, bold=True, fg=WHITE,
                          halign="CENTER", valign="MIDDLE", wrap=True)),
            _row_height_req(sid, 0, 52),
            _freeze_req(sid, rows=1, cols=1),
            _col_width_req(sid, 0, 95),   # Date column
        ]
        for ci in range(1, len(header)):
            reqs.append(_col_width_req(sid, ci, 115))
        ss.batch_update({"requests": reqs})
        existing = [header]

    # Don't double-log the same day
    dates_logged = [r[0] for r in existing[1:] if r]
    if today in dates_logged:
        print(f"  '{TAB_NAME}' already has today's snapshot — skipped")
        return

    snapshot = [today] + [int(row["available"]) for _, row in df.iterrows()]
    next_row  = len(existing) + 1
    ws.update([snapshot], f"A{next_row}")
    print(f"  '{TAB_NAME}' snapshot added for {today} ({len(df)} SKUs)")


# ── Write: Instructions ───────────────────────────────────────────────────────

def write_instructions_tab(ss: gspread.Spreadsheet):
    ws = _get_or_add_tab(ss, "Instructions")
    sid = ws.id

    # (section_label, content, is_header)
    sections = [
        ("US+ Health — Inventory Forecast Dashboard", "", True),
        ("", "", False),
        ("WHAT THIS SHEET DOES", "", "subheader"),
        ("This Google Sheet connects live to Amazon FBA via SP-API and automatically calculates "
         "how much inventory to reorder for every active SKU. It updates each time the script runs "
         "(schedule it daily for best results).", "", False),
        ("", "", False),
        ("HOW THE FORECAST WORKS", "", "subheader"),
        ("6-Month Average Demand", "Average units sold across the last 6 months (Dec 2025 – present).", False),
        ("3-Month Moving Average", "More recent 3-month average — weights recent trends more heavily.", False),
        ("Trend (%)", "Month-over-month sales trend. Positive = growing, negative = declining.", False),
        ("Forecasted Demand", "Projected units needed next month, adjusted for trend.", False),
        ("Safety Stock", "Buffer stock to cover demand variability (95% service level, Z = 1.65).", False),
        ("Reorder Point", "Stock level that triggers a reorder = Safety Stock + (Daily Demand × Lead Days).", False),
        ("Order Qty", "Units to order = max(0, Forecast × 2 − Available − Inbound), but only if Available < Reorder Point.", False),
        ("Days of Stock", "How many days current available stock will last at current demand.", False),
        ("", "", False),
        ("STATUS MEANINGS", "", "subheader"),
        ("🔴  Reorder Now", "Available < Reorder Point AND Order Qty > 0. Stock is below the 2-month safety threshold — place an order.", False),
        ("🟠  Monitor", "Available is within 20% of Reorder Point — watch closely, may need to order soon.", False),
        ("🟢  OK", "Available is more than 20% above Reorder Point. No action needed.", False),
        ("🔵  Covered by Inbound", "Available < Reorder Point but inbound shipment is on its way to cover the gap.", False),
        ("", "", False),
        ("TABS EXPLAINED", "", "subheader"),
        ("Dashboard", "KPI summary cards + charts. Refreshes on every run.", False),
        ("Forecast & Reorder", "Full table — all 45 SKUs with forecast model, order quantities, and sparklines.", False),
        ("Action Items", "Filtered to Reorder Now + Monitor only. Fill in Owner, Target Date, Priority.", False),
        ("Sales History", "Monthly units sold per SKU going back to Dec 2025. Red cells = SP-API quota gap (auto-fixes).", False),
        ("Stock Snapshots", "Daily available FBA stock log. Appends one row per day — builds into history over time.", False),
        ("", "", False),
        ("DATA SOURCES", "", "subheader"),
        ("FBA Inventory", "Live from Amazon SP-API — fetched fresh on every run.", False),
        ("Sales Dec 2025 – Mar 2026", "SP-API monthly Sales & Traffic reports.", False),
        ("Sales Apr 2026 – present", "Local daily sales history file (sales_history.json).", False),
        ("", "", False),
        ("KEY SETTINGS", "", "subheader"),
        (f"Lead Time", f"{LEAD_DAYS} days (time from order to FBA receipt).", False),
        ("Service Level", "95% — safety stock covers 95% of demand variability scenarios.", False),
        ("Target Coverage", f"{TARGET_MONTHS} months of forward stock.", False),
        ("", "", False),
        ("KNOWN LIMITATIONS", "", "subheader"),
        ("March 2026 data", "May show 0 due to SP-API quota limits from running the script multiple times. Auto-fixes on next run.", False),
        ("Sales history depth", "Daily data available from Apr 3, 2026. Earlier months use monthly SP-API reports.", False),
        ("Seasonal spikes", "Forecast uses 6-month average — does not account for seasonal demand peaks.", False),
        ("", "", False),
        ("HOW TO RUN / SCHEDULE", "", "subheader"),
        ("Manual run", "Open a terminal in the amazon-inventory-dashboard folder and run:  python update_sheets.py", False),
        ("Auto-schedule", "Use Windows Task Scheduler to run daily at 7 AM. Ask the team to set this up.", False),
        ("", "", False),
        ("CONTACT", "", "subheader"),
        ("Script location", r"C:\Users\Admin\OneDrive\Julian\amazon-inventory-dashboard\update_sheets.py", False),
        ("Google Sheet ID", "1DXk4CKgLGfyBP-SwMKg97f-5EpxVdg0_rpG2SEwQQ8s", False),
    ]

    # Build flat row list
    rows = []
    for item in sections:
        label, value, kind = item
        if kind is True:          # main title
            rows.append([label])
        elif kind == "subheader":
            rows.append([label])
        elif value:               # two-column label + explanation
            rows.append([label, value])
        else:
            rows.append([label])

    ws.update(rows, "A1")

    reqs = []

    # Column widths
    reqs.append(_col_width_req(sid, 0, 260))
    reqs.append(_col_width_req(sid, 1, 600))

    row_idx = 0
    for item in sections:
        label, value, kind = item

        if kind is True:  # Main title
            reqs.append(_merge_req(sid, row_idx, 0, row_idx + 1, 2))
            reqs.append(_format_range_req(sid, row_idx, 0, row_idx + 1, 2,
                _fmt_cell(bg=HEADER_COLOR, bold=True, fg=WHITE, size=16, halign="CENTER", valign="MIDDLE")))
            reqs.append(_row_height_req(sid, row_idx, 52))

        elif kind == "subheader":
            reqs.append(_merge_req(sid, row_idx, 0, row_idx + 1, 2))
            reqs.append(_format_range_req(sid, row_idx, 0, row_idx + 1, 2,
                _fmt_cell(bg=HEADER2_COLOR, bold=True, fg=WHITE, size=10, valign="MIDDLE")))
            reqs.append(_row_height_req(sid, row_idx, 28))

        elif label == "":  # Spacer
            reqs.append(_row_height_req(sid, row_idx, 8))

        elif value:  # Label + explanation (2-col)
            reqs.append(_format_range_req(sid, row_idx, 0, row_idx + 1, 1,
                _fmt_cell(bold=True, size=10, valign="MIDDLE")))
            reqs.append(_format_range_req(sid, row_idx, 1, row_idx + 1, 2,
                _fmt_cell(fg={"red": 0.2, "green": 0.2, "blue": 0.2}, size=10, valign="MIDDLE", wrap=True)))
            reqs.append(_row_height_req(sid, row_idx, 22))

        else:  # Single description line
            reqs.append(_merge_req(sid, row_idx, 0, row_idx + 1, 2))
            reqs.append(_format_range_req(sid, row_idx, 0, row_idx + 1, 2,
                _fmt_cell(fg={"red": 0.2, "green": 0.2, "blue": 0.2}, size=10, valign="MIDDLE", wrap=True)))
            reqs.append(_row_height_req(sid, row_idx, 22))

        row_idx += 1

    ss.batch_update({"requests": reqs})
    print("  'Instructions' tab written")


# ── Write: Sales History (monthly units sold per SKU, going back 6 months) ────

def write_sales_history_tab(ss: gspread.Spreadsheet, df: pd.DataFrame, months: list[str]):
    ws  = _get_or_add_tab(ss, "Sales History")
    sid = ws.id

    month_labels = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %Y") for m in months]
    HEADERS = ["SKU", "ASIN", "Product", "Category"] + month_labels + ["6-Mo Total"]

    # Category lookup from the known ASIN → category mapping
    ASIN_CATEGORY = {
        "B08Y7X8375": "Hydrogen Peroxide", "B08Y83DNZ5": "Hydrogen Peroxide",
        "B097HP7DQ6": "Carrier Oils",      "B097LTHS4S": "Carrier Oils",
        "B097LVPKMP": "Carrier Oils",      "B0981HX5NG": "Carrier Oils",
        "B09CV925V4": "Carrier Oils",      "B09DZ2P2WJ": "Carrier Oils",
        "B09DZDD71G": "Carrier Oils",      "B0BJH3RD1F": "Carrier Oils",
        "B0BR99MF15": "Carrier Oils",      "B0CCMHLX72": "Carrier Oils",
        "B0DDWQ1515": "Organic Oils",      "B0DSCKXPQH": "Organic Oils",
    }

    rows = []
    for _, row in df.iterrows():
        monthly = [int(row.get(m, 0)) for m in months]
        total   = sum(monthly)
        rows.append([
            row["sku"],
            row["asin"],
            _product_label(row["asin"], row.get("item_name", "")),
            ASIN_CATEGORY.get(row["asin"], "Other"),
        ] + monthly + [total])

    # Sort by 6-month total descending (top sellers first)
    rows.sort(key=lambda r: r[-1], reverse=True)

    ws.update([[
        "Monthly units sold per SKU  |  Source: Amazon SP-API  |  "
        "Months showing 0 may reflect SP-API quota limits — will update on next run"
    ]], "A1")
    ws.update([HEADERS], "A2")
    if rows:
        ws.update(rows, "A3")

    reqs = []

    # Info row
    reqs.append(_format_range_req(sid, 0, 0, 1, len(HEADERS),
        _fmt_cell(bg={"red": 0.9, "green": 0.9, "blue": 0.9},
                  fg={"red": 0.35, "green": 0.35, "blue": 0.35}, size=9)))

    # Header row
    reqs.append(_format_range_req(sid, 1, 0, 2, len(HEADERS),
        _fmt_cell(bg=HEADER_COLOR, bold=True, fg=WHITE, halign="CENTER", wrap=True)))
    reqs.append(_row_height_req(sid, 1, 48))

    # Freeze header + SKU col
    reqs.append(_freeze_req(sid, rows=2, cols=1))

    # Month value cells: center-aligned
    for ci in range(4, 4 + len(months)):
        reqs.append(_format_range_req(sid, 2, ci, 2 + len(rows), ci + 1,
            _fmt_cell(halign="CENTER")))

    # 6-Mo Total column: highlighted
    total_ci = 4 + len(months)
    reqs.append(_format_range_req(sid, 1, total_ci, 1 + 1, total_ci + 1,
        _fmt_cell(bg=HEADER_COLOR, bold=True, fg=WHITE, halign="CENTER")))
    reqs.append(_format_range_req(sid, 2, total_ci, 2 + len(rows), total_ci + 1,
        _fmt_cell(bg={"red": 0.855, "green": 0.918, "blue": 0.996},
                  bold=True, halign="CENTER")))

    # Zero cells: light red to flag missing data
    for ri, row_vals in enumerate(rows):
        for ci, val in enumerate(row_vals[4:4 + len(months)], start=4):
            if val == 0:
                reqs.append(_format_range_req(
                    sid, 2 + ri, ci, 3 + ri, ci + 1,
                    _fmt_cell(bg={"red": 1.0, "green": 0.92, "blue": 0.92},
                              fg={"red": 0.6, "green": 0.0, "blue": 0.0},
                              halign="CENTER")))

    # Column widths
    col_widths = [150, 110, 250, 120] + [75] * len(months) + [90]
    for ci, w in enumerate(col_widths):
        reqs.append(_col_width_req(sid, ci, w))

    ss.batch_update({"requests": reqs})
    print(f"  'Sales History' updated — {len(rows)} SKUs, {len(months)} months")


# ── Main ───────────────────────────────────────────────────────────────────────

def run():
    print("-" * 55)
    print("  US+ Health - Inventory Forecast (Google Sheets)")
    print(f"  {datetime.date.today():%B %d, %Y}")
    print("-" * 55)

    print("\n[1/5] Fetching FBA inventory from SP-API...")
    inventory = fetch_inventory()
    print(f"      {len(inventory)} SKUs found")
    inventory = _consolidate_by_asin(inventory)
    print(f"      {len(inventory)} unique ASINs after consolidation")

    asins = inventory["asin"].dropna().unique().tolist()

    print("\n[2/5] Loading monthly sales data...")
    sales, months = get_monthly_sales(asins)
    print(f"      Months: {', '.join(months)}")

    print("\n[3/5] Running forecast calculations...")
    forecast_df = run_forecast(inventory, sales, months)
    reorder = (forecast_df["status"] == "Reorder Now").sum()
    print(f"      {reorder} SKUs flagged for reorder")

    print("\n[4/5] Connecting to Google Sheets...")
    gc = _gc()
    ss = get_or_create_sheet(gc)
    _delete_default_sheet(ss)
    print(f"      Sheet: {ss.title}")
    print(f"      URL: https://docs.google.com/spreadsheets/d/{ss.id}")

    print("\n[5/5] Writing tabs...")
    write_instructions_tab(ss)
    write_forecast_tab(ss, forecast_df, months)
    write_dashboard_tab(ss, forecast_df)
    write_action_items_tab(ss, forecast_df)
    write_sales_history_tab(ss, forecast_df, months)
    write_stock_history_tab(ss, forecast_df)

    print(f"\nDone. Open your sheet:")
    print(f"  https://docs.google.com/spreadsheets/d/{ss.id}\n")


if __name__ == "__main__":
    run()
