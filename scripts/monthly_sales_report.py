#!/usr/bin/env python3
"""
Monthly Sales Report - Gaya Foods
Uses Supabase REST API (anon key) → no DB password needed
"""

import requests
from datetime import datetime, timezone, timedelta

# ── Config ────────────────────────────────────────────────────────────────────
SUPABASE_URL    = "https://uwfbirjpzzberwrhkson.supabase.co"
SUPABASE_KEY    = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InV3ZmJpcmpwenpiZXJ3cmhrc29uIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjIwNzAzNTAsImV4cCI6MjA3NzY0NjM1MH0.ar3kfCjkVCqsyqx9zBsSbfn2AORxL9Ph7KLkQUjM6-I"
TIMELINES_TOKEN = "f40ecfc9-31e8-4905-a920-b27e5559fabc"
WHATSAPP_PHONE  = "972528012869"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json"
}

# ── SQL Queries ───────────────────────────────────────────────────────────────
SQL_SALES = """
SELECT c."סוכן", SUM(p."כמות" * p."מחיר ליחידה") AS sum
FROM "פירוט חשבוניות מרכזות" p
JOIN "חשבוניות מרכזות" c ON c."חשבונית מרכזת" = p."חשבונית מרכזת"
WHERE c."תאריך" >= DATE_TRUNC('month', CURRENT_DATE)
  AND c."חיוב/זיכוי" = 'חיוב'
GROUP BY c."סוכן"
ORDER BY sum DESC
"""

SQL_ORDERS = """
SELECT "סוכן", SUM("סכום הזמנה") AS sum, COUNT(*) AS cnt
FROM "הזמנות"
WHERE "סטטוס הזמנה מפריוריטי" NOT IN ('סגורה', 'מבוטלת')
GROUP BY "סוכן"
ORDER BY sum DESC
"""

SQL_UNINVOICED = """
SELECT COUNT(DISTINCT m."תעודת משלוח") AS cnt, SUM(p."סה\u05f4כ מחיר") AS sum
FROM "משלוחים" m
JOIN "פירוט תעודות משלוח" p ON p."תעודת משלוח מקשרת" = m."תעודת משלוח"
WHERE m."חויבה" = 'N'
  AND m."תאריך" >= (CURRENT_DATE - INTERVAL '90 days')
"""


# ── Helpers ───────────────────────────────────────────────────────────────────
def run_sql(query: str) -> list:
    resp = requests.post(
        f"{SUPABASE_URL}/rest/v1/rpc/execute_sql",
        headers=HEADERS,
        json={"query": query.strip()},
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()


def fmt(n) -> str:
    return f"₪{int(round(float(n))):,}"


def hebrew_month(dt: datetime) -> str:
    months = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני",
              "יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"]
    return months[dt.month - 1]


def send_whatsapp(text: str):
    resp = requests.post(
        "https://app.timelines.ai/integrations/api/messages",
        headers={"Authorization": f"Bearer {TIMELINES_TOKEN}", "Content-Type": "application/json"},
        json={"phone": WHATSAPP_PHONE, "text": text},
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    israel = timezone(timedelta(hours=2))
    now = datetime.now(israel)
    timestamp = now.strftime("%d.%m.%Y בשעה %H:%M")
    month_name = hebrew_month(now)
    year = now.year

    sales    = run_sql(SQL_SALES)
    orders   = run_sql(SQL_ORDERS)
    uninv    = run_sql(SQL_UNINVOICED)

    sales_total  = sum(float(r.get("sum", 0)) for r in sales)
    orders_total = sum(float(r.get("sum", 0)) for r in orders)
    orders_count = sum(int(r.get("cnt", 0)) for r in orders)
    uni_sum      = float(uninv[0].get("sum", 0)) if uninv else 0
    uni_count    = int(uninv[0].get("cnt", 0)) if uninv else 0

    AGENTS = {"חיים שחרור": "חיים", "אוראל כהן": "אוראל", "פאר מגיד": "פאר", "אוהד": "אוהד"}
    short = lambda n: AGENTS.get(n, n)

    lines = [
        f"📊 *{month_name} {year}*",
        f"_נכון לתאריך {timestamp}_",
        "",
        "━━━━━━━━━━━━━━",
        f'💰 *סה"כ חשבוניות שיצאו החודש*',
        f'*{fmt(sales_total)}* סה"כ',
        *[f'{short(r["סוכן"])} {fmt(r["sum"])}' for r in sales],
        "",
        "━━━━━━━━━━━━━━",
        "📋 *הזמנות פתוחות*",
        f'*{fmt(orders_total)}* סה"כ | {orders_count} הזמנות',
        *[f'{short(r["סוכן"])} {fmt(r["sum"])}' for r in orders],
        "",
        "━━━━━━━━━━━━━━",
        "⚠️ *ת.מ שטרם חויבו*",
        f"*{fmt(uni_sum)}* | {uni_count} תעודות",
    ]

    message = "\n".join(lines)
    result = send_whatsapp(message)
    print(f"✅ נשלח! uid: {result.get('data', {}).get('message_uid', '-')}")


if __name__ == "__main__":
    main()
