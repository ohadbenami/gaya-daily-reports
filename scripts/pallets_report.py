#!/usr/bin/env python3
"""
Pallets Distribution Report - Gaya Foods
Source: Monday.com board 5089475109 (הפצה)
Modes: daily | weekly | monthly
Usage: python pallets_report.py [daily|weekly|monthly]
"""
import sys
import os
import requests
from datetime import datetime, timezone, timedelta
from collections import defaultdict

# ── Config ────────────────────────────────────────────────────────────────────
MONDAY_TOKEN   = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjU1MTE4OTg5MiwiYWFpIjoxMSwidWlkIjo3MDYwMzkyMSwiaWFkIjoiMjAyNS0wOC0xN1QxNjo0MDo0OC4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjczNTY5NjgsInJnbiI6ImV1YzEifQ.DX4YtZ2uq-E4WcTK0n0AN-CW7lzdEp075QesM-CdITE"
BOARD_ID       = 5089475109
TIMELINES_TOKEN = "f40ecfc9-31e8-4905-a920-b27e5559fabc"
WHATSAPP_PHONE  = "972528012869"
MONDAY_API      = "https://api.monday.com/v2"
ISRAEL_TZ       = timezone(timedelta(hours=2))

MONDAY_HEADERS = {
    "Authorization": f"Bearer {MONDAY_TOKEN}",
    "Content-Type": "application/json",
    "API-Version": "2024-01"
}

# סדר תצוגת נהגים
DRIVER_ORDER = ["שי", "אורי", "אורי נגלה 2", "שי נגלה 2", "BL", "שפע תובלה", "לא שויך"]

HEBREW_MONTHS = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני",
                 "יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"]

# ── Monday.com: שליפת כל הדפים ────────────────────────────────────────────────
def fetch_all_items():
    """שלוף את כל הפריטים מהלוח — עובר דפים עד cursor=null."""
    all_items = []

    # דף ראשון
    query = """{
      boards(ids: [%d]) {
        items_page(limit: 500) {
          cursor
          items {
            name
            column_values(ids: ["date4", "color_mkz4z0q4", "numeric_mkz4s8sc"]) { id text }
          }
        }
      }
    }""" % BOARD_ID
    resp = requests.post(MONDAY_API, headers=MONDAY_HEADERS, json={"query": query}, timeout=60)
    resp.raise_for_status()
    page = resp.json()["data"]["boards"][0]["items_page"]
    all_items.extend(page["items"])
    cursor = page.get("cursor")

    # דפים נוספים
    while cursor:
        query = """{
          next_items_page(limit: 500, cursor: "%s") {
            cursor
            items {
              name
              column_values(ids: ["date4", "color_mkz4z0q4", "numeric_mkz4s8sc"]) { id text }
            }
          }
        }""" % cursor
        resp = requests.post(MONDAY_API, headers=MONDAY_HEADERS, json={"query": query}, timeout=60)
        resp.raise_for_status()
        page = resp.json()["data"]["next_items_page"]
        all_items.extend(page["items"])
        cursor = page.get("cursor")

    print(f"  → נשלפו {len(all_items)} פריטים מהלוח")
    return all_items


# ── עיבוד נתונים ─────────────────────────────────────────────────────────────
def parse_item(item):
    cols = {cv["id"]: cv["text"] for cv in item["column_values"]}
    return {
        "date":     cols.get("date4") or "",
        "driver":   cols.get("color_mkz4z0q4") or "לא שויך",
        "customer": item["name"],
        "pallets":  float(cols.get("numeric_mkz4s8sc") or 0),
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
def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("PALLETS_MODE", "daily")
    now  = datetime.now(ISRAEL_TZ)
    print(f"[pallets_report] mode={mode} | {now.strftime('%d.%m.%Y %H:%M')} Israel")

    print("שולף נתונים מ-Monday.com...")
    items = fetch_all_items()

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
