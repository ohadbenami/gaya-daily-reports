#!/usr/bin/env python3
"""
Pallets Distribution Report - Gaya Foods
Source: Supabase table `הפצה` (snapshot of the old Monday distribution board).
        Distribution data of record = Priority DOCUMENTS_D + Supabase `הפצה`.
Modes: daily | weekly | monthly
Usage: python pallets_report.py [daily|weekly|monthly]
"""
import sys
import os
import requests
from datetime import datetime, timezone, timedelta
from collections import defaultdict

# ── Config ────────────────────────────────────────────────────────────────────
# Supabase (data project that hosts the `הפצה` table).
# Env vars take precedence (set in GitHub Actions); fall back to the data-project defaults.
SUPABASE_URL    = os.environ.get("SUPABASE_URL", "https://uwfbirjpzzberwrhkson.supabase.co").rstrip("/")
SUPABASE_KEY    = os.environ.get(
    "SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InV3ZmJpcmpwenpiZXJ3cmhrc29uIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjIwNzAzNTAsImV4cCI6MjA3NzY0NjM1MH0.ar3kfCjkVCqsyqx9zBsSbfn2AORxL9Ph7KLkQUjM6-I",
)
SUPABASE_TABLE  = "הפצה"
TIMELINES_TOKEN = "f40ecfc9-31e8-4905-a920-b27e5559fabc"
WHATSAPP_PHONE  = "972528012869"
ISRAEL_TZ       = timezone(timedelta(hours=2))

# Supabase `הפצה` column names (Hebrew, snapshot of old Monday board).
COL_DATE     = "ת. הפצה"      # distribution date, format YYYY-MM-DD
COL_DRIVER   = "נהג"          # driver
COL_CUSTOMER = "שם לקוח"      # customer name
COL_PALLETS  = "משטחים"       # pallets

SUPABASE_HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
}

# סדר תצוגת נהגים
DRIVER_ORDER = ["שי", "אורי", "אורי נגלה 2", "שי נגלה 2", "BL", "שפע תובלה", "לא שויך"]

HEBREW_MONTHS = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני",
                 "יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"]

# ── Supabase: שליפת רשומות ההפצה ───────────────────────────────────────────────
def fetch_items(date_from=None, date_to=None):
    """שלוף רשומות הפצה מטבלת `הפצה` ב-Supabase, מסונן לפי טווח תאריכי הפצה.

    date_from / date_to הם מחרוזות 'YYYY-MM-DD' (כולל). אם לא נמסרו — מושך הכל.
    שמות עמודות עם רווח/נקודה מצוטטים ב-PostgREST (לדוגמה "ת. הפצה").
    """
    all_rows = []
    offset = 0
    page_size = 1000
    col_date_q = f'"{COL_DATE}"'

    while True:
        params = {"select": "*", "limit": str(page_size), "offset": str(offset)}
        if date_from:
            params[col_date_q] = f"gte.{date_from}"
        if date_to:
            # PostgREST: שני תנאים על אותה עמודה דורשים תחביר and=()
            if date_from:
                params.pop(col_date_q, None)
                params["and"] = f'({col_date_q}.gte.{date_from},{col_date_q}.lte.{date_to})'
            else:
                params[col_date_q] = f"lte.{date_to}"

        resp = requests.get(
            f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}",
            headers=SUPABASE_HEADERS,
            params=params,
            timeout=60,
        )
        resp.raise_for_status()
        batch = resp.json()
        all_rows.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size

    print(f"  → נשלפו {len(all_rows)} רשומות הפצה מ-Supabase")
    return all_rows


# ── עיבוד נתונים ─────────────────────────────────────────────────────────────
def parse_item(item):
    return {
        "date":     item.get(COL_DATE) or "",
        "driver":   item.get(COL_DRIVER) or "לא שויך",
        "customer": item.get(COL_CUSTOMER) or "",
        "pallets":  float(item.get(COL_PALLETS) or 0),
    }


def group_by_driver(items):
    """קבץ לפי נהג → לקוח → סכום משטחים."""
    drivers = defaultdict(lambda: defaultdict(float))
    for item in items:
        p = parse_item(item)
        drivers[p["driver"]][p["customer"]] += p["pallets"]
    return drivers


def shorten(name):
    for suffix in [" ובניו שיווק בע\"מ", " שיווק והפצה בע\"מ", " שיווק בע\"מ",
                   " בע\"מ", " (1999)", " (2002)", " (1996)", " (1985)",
                   " - פניני השף", " - אגודה שיתופית חקלאית"]:
        name = name.replace(suffix, "")
    return name.strip()[:22]


def ordered_drivers(drivers):
    """החזר נהגים בסדר מוגדר — קודם ידועים, אחר כך שאר."""
    known = [d for d in DRIVER_ORDER if d in drivers]
    others = [d for d in drivers if d not in DRIVER_ORDER]
    return known + others


def driver_block(driver, customers, bold=True):
    total = int(sum(customers.values()))
    name = f"*{driver}*" if bold else driver
    lines = [f"🚛 {name} — {total} משטחים"]
    for cust, p in sorted(customers.items(), key=lambda x: -x[1]):
        lines.append(f"• {shorten(cust)} — {int(p)}")
    return "\n".join(lines)


# ── שליחה לוואטסאפ ───────────────────────────────────────────────────────────
def send_whatsapp(text):
    resp = requests.post(
        "https://app.timelines.ai/integrations/api/messages",
        headers={"Authorization": f"Bearer {TIMELINES_TOKEN}", "Content-Type": "application/json"},
        json={"phone": WHATSAPP_PHONE, "text": text},
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()


# ── DAILY ─────────────────────────────────────────────────────────────────────
def daily_report(items, now):
    today = now.strftime("%Y-%m-%d")
    yesterday = (now - timedelta(days=1)).strftime("%Y-%m-%d")

    filtered = [i for i in items if parse_item(i)["date"] == today]
    note = ""
    date_str = now.strftime("%d.%m.%Y")

    if not filtered:
        filtered = [i for i in items if parse_item(i)["date"] == yesterday]
        note = "\n_(נתוני אתמול — הנתונים להיום טרם עודכנו)_"
        date_str = (now - timedelta(days=1)).strftime("%d.%m.%Y")

    if not filtered:
        return f"📦 *חלוקה {date_str}*\n\nאין נתונים זמינים."

    drivers = group_by_driver(filtered)
    total_p = int(sum(sum(c.values()) for c in drivers.values()))
    total_d = len(filtered)

    summary = " | ".join(
        f"{drv} {int(sum(drivers[drv].values()))}"
        for drv in ordered_drivers(drivers)
    )

    lines = [
        f"📦 *חלוקה {date_str}*{note}",
        "",
        f"*{total_p} משטחים* | {total_d} שורות",
        summary,
        "",
        "━━━━━━━━━━",
    ]
    for drv in ordered_drivers(drivers):
        lines.append(driver_block(drv, drivers[drv]))
        lines.append("")

    return "\n".join(lines).strip()


# ── WEEKLY ────────────────────────────────────────────────────────────────────
def weekly_report(items, now):
    today = now.date()
    monday = today - timedelta(days=today.weekday())   # ראשון לשבוע (ב')
    week_dates = set()
    d = monday
    while d <= today:
        week_dates.add(d.strftime("%Y-%m-%d"))
        d += timedelta(days=1)

    filtered = [i for i in items if parse_item(i)["date"] in week_dates]
    week_str = f"{monday.strftime('%d.%m')} — {today.strftime('%d.%m.%Y')}"

    if not filtered:
        return f"📦 *סיכום שבועי | {week_str}*\n\nאין נתונים לשבוע זה."

    drivers = group_by_driver(filtered)
    total_p = int(sum(sum(c.values()) for c in drivers.values()))
    total_d = len(filtered)

    lines = [
        f"📦 *סיכום שבועי | {week_str}*",
        "",
        f"*{total_p} משטחים* | {total_d} שורות",
        "",
        "━━━━━━━━━━",
    ]
    for drv in ordered_drivers(drivers):
        total = int(sum(drivers[drv].values()))
        if total > 0:
            lines.append(f"🚛 *{drv}* — {total}")

    return "\n".join(lines).strip()


# ── MONTHLY ───────────────────────────────────────────────────────────────────
def monthly_report(items, now):
    first_this_month = now.date().replace(day=1)
    last_month_end   = first_this_month - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)

    month_dates = set()
    d = last_month_start
    while d <= last_month_end:
        month_dates.add(d.strftime("%Y-%m-%d"))
        d += timedelta(days=1)

    filtered = [i for i in items if parse_item(i)["date"] in month_dates]
    month_name = HEBREW_MONTHS[last_month_end.month - 1]
    year = last_month_end.year

    if not filtered:
        return f"📦 *סיכום חודשי | {month_name} {year}*\n\nאין נתונים לחודש זה."

    drivers = group_by_driver(filtered)
    total_p = int(sum(sum(c.values()) for c in drivers.values()))
    total_d = len(filtered)

    lines = [
        f"📦 *סיכום חודשי | {month_name} {year}*",
        "",
        f"*{total_p} משטחים* חולקו | {total_d} שורות",
        "",
        "━━━━━━━━━━",
    ]
    for drv in ordered_drivers(drivers):
        total = int(sum(drivers[drv].values()))
        if total > 0:
            lines.append(f"🚛 *{drv}* — {total}")

    return "\n".join(lines).strip()


# ── Main ──────────────────────────────────────────────────────────────────────
def date_window(mode, now):
    """החזר (date_from, date_to) לטעינה מ-Supabase לפי המוד (כולל מרווח ביטחון)."""
    if mode == "daily":
        # היום + אתמול (fallback) — מושכים יומיים אחרונים
        return (now - timedelta(days=1)).strftime("%Y-%m-%d"), now.strftime("%Y-%m-%d")
    if mode == "weekly":
        today = now.date()
        monday = today - timedelta(days=today.weekday())
        return monday.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")
    if mode == "monthly":
        first_this_month = now.date().replace(day=1)
        last_month_end = first_this_month - timedelta(days=1)
        last_month_start = last_month_end.replace(day=1)
        return last_month_start.strftime("%Y-%m-%d"), last_month_end.strftime("%Y-%m-%d")
    raise ValueError(f"Unknown mode: {mode}. Use daily/weekly/monthly")


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("PALLETS_MODE", "daily")
    now  = datetime.now(ISRAEL_TZ)
    print(f"[pallets_report] mode={mode} | {now.strftime('%d.%m.%Y %H:%M')} Israel")

    print("שולף נתונים מ-Supabase (טבלת הפצה)...")
    date_from, date_to = date_window(mode, now)
    items = fetch_items(date_from, date_to)

    if mode == "daily":
        msg = daily_report(items, now)
    elif mode == "weekly":
        msg = weekly_report(items, now)
    elif mode == "monthly":
        msg = monthly_report(items, now)
    else:
        raise ValueError(f"Unknown mode: {mode}. Use daily/weekly/monthly")

    print("הודעה:\n" + msg)
    result = send_whatsapp(msg)
    print(f"✅ נשלח! uid: {result.get('data', {}).get('message_uid', '-')}")


if __name__ == "__main__":
    main()
