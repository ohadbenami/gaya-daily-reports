#!/usr/bin/env python3
"""
Monthly Sales Report - Gaya Foods
- מכירות החודש לפי סוכן (חשבוניות)
- הזמנות פתוחות לפי סוכן
- ת.מ שטרם חויבו (90 יום)
- שליחה לוואטסאפ דרך TimelineAI
"""

import os
import requests
import psycopg2
import psycopg2.extras
from datetime import datetime

# ── Config ────────────────────────────────────────────────────────────────────
SUPABASE_DB_URL  = os.environ.get('SUPABASE_DB_URL')   # Transaction pooler URL (full)
TIMELINES_TOKEN  = os.environ.get('TIMELINES_TOKEN', 'f40ecfc9-31e8-4905-a920-b27e5559fabc')
WHATSAPP_PHONE   = os.environ.get('WHATSAPP_PHONE', '972528012869')

if not SUPABASE_DB_URL:
    raise ValueError("Missing SUPABASE_DB_URL secret. Add it in GitHub Secrets.")

# ── Queries ───────────────────────────────────────────────────────────────────
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
SELECT "סוכן", SUM("סכום הזמנה") AS sum, COUNT(*) AS count
FROM "הזמנות"
WHERE "סטטוס הזמנה מפריוריטי" NOT IN ('סגורה', 'מבוטלת')
GROUP BY "סוכן"
ORDER BY sum DESC
"""

SQL_UNINVOICED = """
SELECT COUNT(DISTINCT m."תעודת משלוח") AS count, SUM(p."סה\u05bcכ מחיר") AS sum
FROM "משלוחים" m
JOIN "פירוט תעודות משלוח" p ON p."תעודת משלוח מקשרת" = m."תעודת משלוח"
WHERE m."חויבה" = 'N'
  AND m."תאריך" >= (CURRENT_DATE - INTERVAL '90 days')
"""


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(n):
    """Format number as ₪X,XXX,XXX"""
    return f"₪{int(round(n)):,}"


def hebrew_month(dt):
    months = {
        1: 'ינואר', 2: 'פברואר', 3: 'מרץ', 4: 'אפריל',
        5: 'מאי', 6: 'יוני', 7: 'יולי', 8: 'אוגוסט',
        9: 'ספטמבר', 10: 'אוקטובר', 11: 'נובמבר', 12: 'דצמבר'
    }
    return months[dt.month]


def send_whatsapp(text):
    resp = requests.post(
        "https://app.timelines.ai/integrations/api/messages",
        headers={
            "Authorization": f"Bearer {TIMELINES_TOKEN}",
            "Content-Type": "application/json"
        },
        json={"phone": WHATSAPP_PHONE, "text": text},
        timeout=30
    )
    resp.raise_for_status()
    return resp.json()


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    now = datetime.now()
    timestamp = now.strftime("%d.%m.%Y בשעה %H:%M")
    month_name = hebrew_month(now)
    year = now.year

    # Connect via Transaction Pooler (IPv4, works from GitHub Actions)
    conn = psycopg2.connect(
        SUPABASE_DB_URL,
        sslmode="require",
        options="-c client_encoding=UTF8"
    )
    conn.set_client_encoding('UTF8')
    cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Query 1 — Sales
    cur.execute(SQL_SALES)
    sales_rows = cur.fetchall()
    sales_total = sum(r['sum'] for r in sales_rows)

    # Query 2 — Open orders
    cur.execute(SQL_ORDERS)
    orders_rows = cur.fetchall()
    orders_total = sum(r['sum'] for r in orders_rows)
    orders_count = sum(r['count'] for r in orders_rows)

    # Query 3 — Uninvoiced delivery notes
    cur.execute(SQL_UNINVOICED)
    uni = cur.fetchone()
    uni_count = int(uni['count']) if uni['count'] else 0
    uni_sum = float(uni['sum']) if uni['sum'] else 0

    cur.close()
    conn.close()

    # ── Build known agents lookup ─────────────────────────────────────────────
    AGENTS = {'חיים שחרור': 'חיים', 'אוראל כהן': 'אוראל', 'פאר מגיד': 'פאר', 'אוהד': 'אוהד'}

    def agent_short(name):
        return AGENTS.get(name, name)

    # ── Build WhatsApp message ────────────────────────────────────────────────
    lines = [
        f"📊 *{month_name} {year}*",
        f"_נכון לתאריך {timestamp}_",
        "",
        "━━━━━━━━━━━━━━",
        "💰 *סה\"כ חשבוניות שיצאו החודש*",
        f"*{fmt(sales_total)}* סה\"כ",
    ]
    for r in sales_rows:
        lines.append(f"{agent_short(r['סוכן'])} {fmt(r['sum'])}")

    lines += [
        "",
        "━━━━━━━━━━━━━━",
        "📋 *הזמנות פתוחות*",
        f"*{fmt(orders_total)}* סה\"כ | {orders_count} הזמנות",
    ]
    for r in orders_rows:
        lines.append(f"{agent_short(r['סוכן'])} {fmt(r['sum'])}")

    lines += [
        "",
        "━━━━━━━━━━━━━━",
        "⚠️ *ת.מ שטרם חויבו*",
        f"*{fmt(uni_sum)}* | {uni_count} תעודות",
    ]

    message = "\n".join(lines)

    # Send
    result = send_whatsapp(message)
    print(f"✅ נשלח! uid: {result.get('data', {}).get('message_uid', '-')}")
    print(f"\n{message}")


if __name__ == "__main__":
    main()
