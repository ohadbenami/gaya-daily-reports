#!/usr/bin/env python3
"""
Daily Reorder Alerts - Gaya Foods
Queries Priority ERP for stock levels, expiry, containers.
Sends WhatsApp alerts to Ohad and Kiril every morning.
"""

import os
import json
import requests
from datetime import datetime, timedelta
from pathlib import Path

# --- Configuration ---
PRIORITY_API_USER = os.environ["PRIORITY_API_USER"]
PRIORITY_API_PASS = os.environ["PRIORITY_API_PASS"]
TIMELINEAI_TOKEN = os.environ["TIMELINEAI_TOKEN"]

PRIORITY_BASE_URL = "https://p.priority-connect.online/odata/Priority/tabzfdbb.ini/a230521"

# Load config
CONFIG_PATH = Path(__file__).parent / "reorder_config.json"
with open(CONFIG_PATH) as f:
    CONFIG = json.load(f)

REORDER_POINTS = CONFIG["reorder_points"]
DEFAULT_REORDER_POINT = CONFIG["default_reorder_point"]
EXPIRY_WARNING_DAYS = CONFIG["expiry_warning_days"]
RECIPIENTS = CONFIG["recipients"]


def query_priority(table, params, timeout=60):
    """Query Priority ERP OData API."""
    url = f"{PRIORITY_BASE_URL}/{table}"
    try:
        response = requests.get(
            url,
            params=params,
            auth=(PRIORITY_API_USER, PRIORITY_API_PASS),
            headers={"Accept": "application/json"},
            timeout=timeout
        )
        response.raise_for_status()
        return response.json().get("value", [])
    except requests.exceptions.Timeout:
        print(f"  TIMEOUT querying {table}")
        return []
    except requests.exceptions.HTTPError as e:
        print(f"  HTTP ERROR querying {table}: {e}")
        return []


def get_low_stock():
    """Get products below reorder point - only products with defined reorder points."""
    print("  Querying stock levels...")

    # Only query products that have reorder points defined
    # Build filter for specific SKUs
    sku_filters = " or ".join(f"PARTNAME eq '{sku}'" for sku in REORDER_POINTS.keys())
    if not sku_filters:
        return []

    raw = query_priority("LOGPART", {
        "$filter": sku_filters,
        "$select": "PARTNAME,PARTDES,FAMILYNAME",
        "$expand": "PARTBALANCE_SUBFORM($select=BALANCE,WARHSNAME)",
    }, timeout=90)

    low_stock = []
    for item in raw:
        partname = item.get("PARTNAME", "")
        partdes = item.get("PARTDES", "")
        balances = item.get("PARTBALANCE_SUBFORM", [])
        total_balance = sum(b.get("BALANCE", 0) for b in balances)

        reorder_point = REORDER_POINTS.get(partname, DEFAULT_REORDER_POINT)

        if total_balance < reorder_point:
            low_stock.append({
                "partname": partname,
                "partdes": partdes,
                "balance": int(total_balance),
                "reorder_point": reorder_point,
                "deficit_pct": round((1 - total_balance / reorder_point) * 100) if reorder_point > 0 else 0
            })

    # Sort by deficit percentage (most critical first)
    low_stock.sort(key=lambda x: x["deficit_pct"], reverse=True)
    return low_stock


def get_expiring_products():
    """Get products expiring within warning period."""
    print("  Querying expiring products...")
    cutoff = (datetime.now() + timedelta(days=EXPIRY_WARNING_DAYS)).strftime("%Y-%m-%dT00:00:00+02:00")

    raw = query_priority("LOGPART", {
        "$select": "PARTNAME,PARTDES",
        "$expand": f"PARTBALANCE_SUBFORM($filter=EXPIRYDATE lt {cutoff} and BALANCE gt 0;$select=BALANCE,EXPIRYDATE,WARHSNAME)",
        "$top": 500
    }, timeout=90)

    expiring = []
    for item in raw:
        balances = item.get("PARTBALANCE_SUBFORM", [])
        if not balances:
            continue
        for b in balances:
            exp_date = b.get("EXPIRYDATE", "")
            if not exp_date:
                continue
            try:
                exp_dt = datetime.fromisoformat(exp_date.replace("Z", "+00:00"))
                days_left = (exp_dt.date() - datetime.now().date()).days
            except (ValueError, TypeError):
                days_left = 999

            if 0 < days_left <= EXPIRY_WARNING_DAYS:
                expiring.append({
                    "partname": item.get("PARTNAME", ""),
                    "partdes": item.get("PARTDES", ""),
                    "balance": int(b.get("BALANCE", 0)),
                    "expiry_date": exp_date[:10],
                    "days_left": days_left,
                    "warehouse": b.get("WARHSNAME", "")
                })

    # Sort by days_left (most urgent first)
    expiring.sort(key=lambda x: x["days_left"])
    return expiring


def get_containers_in_transit():
    """Get purchase orders in transit (containers at sea/port)."""
    print("  Querying containers in transit...")
    raw = query_priority("PORDERS", {
        "$filter": "STATDES eq 'בדרך' or STATDES eq 'כנ\"מ ללא BL'",
        "$select": "ORDNAME,CDES,CURDATE,STATDES,NOA_ETA,NOA_KONTAINER",
        "$expand": "PORDERITEMS_SUBFORM($select=PARTNAME,TQUANT,QPRICE)",
        "$orderby": "NOA_ETA asc"
    })

    containers = []
    for po in raw:
        eta = po.get("NOA_ETA", "")
        if eta:
            try:
                eta_str = datetime.fromisoformat(eta.replace("Z", "+00:00")).strftime("%d.%m.%Y")
            except (ValueError, TypeError):
                eta_str = eta[:10]
        else:
            eta_str = "לא ידוע"
        items = po.get("PORDERITEMS_SUBFORM", [])
        item_names = ", ".join(set(i.get("PARTNAME", "") for i in items[:3]))
        if len(items) > 3:
            item_names += f" +{len(items)-3}"

        containers.append({
            "ordname": po.get("ORDNAME", ""),
            "supplier": po.get("CDES", ""),
            "status": po.get("STATDES", ""),
            "eta": eta_str,
            "container": po.get("NOA_KONTAINER", ""),
            "items": item_names,
            "total_value": sum(i.get("QPRICE", 0) for i in items)
        })

    return containers


def get_pending_purchase_orders():
    """Get approved purchase orders not yet shipped."""
    print("  Querying pending purchase orders...")
    raw = query_priority("PORDERS", {
        "$filter": "STATDES eq 'מאושרת'",
        "$select": "ORDNAME,CDES,CURDATE,NOA_ETA",
        "$expand": "PORDERITEMS_SUBFORM($select=PARTNAME,TQUANT)",
        "$orderby": "CURDATE desc",
        "$top": 20
    })

    pending = []
    for po in raw:
        items = po.get("PORDERITEMS_SUBFORM", [])
        item_names = ", ".join(set(i.get("PARTNAME", "") for i in items[:3]))
        if len(items) > 3:
            item_names += f" +{len(items)-3}"

        pending.append({
            "ordname": po.get("ORDNAME", ""),
            "supplier": po.get("CDES", ""),
            "date": po.get("CURDATE", "")[:10],
            "items": item_names
        })

    return pending


def build_message(low_stock, expiring, containers, pending):
    """Build the WhatsApp alert message."""
    today = datetime.now().strftime("%d.%m.%Y")
    lines = [
        f"🌅 *התראות בוקר - גאיה פודס*",
        f"📅 {today}",
        ""
    ]

    # Low stock alerts
    if low_stock:
        lines.append(f"🔴 *צריך להזמין ({len(low_stock)} פריטים):*")
        for item in low_stock[:15]:
            emoji = "🔴" if item["deficit_pct"] > 50 else "🟠"
            lines.append(
                f"{emoji} {item['partname']} {item['partdes'][:20]} - "
                f"{item['balance']:,} קרטונים (מתחת ל-{item['reorder_point']:,})"
            )
        if len(low_stock) > 15:
            lines.append(f"   ... ועוד {len(low_stock)-15} פריטים")
        lines.append("")
    else:
        lines.append("✅ *מלאי תקין* - אין פריטים מתחת לנקודת הזמנה")
        lines.append("")

    # Expiring products
    if expiring:
        lines.append(f"🟠 *תפוגות קרובות ({len(expiring)} פריטים, <{EXPIRY_WARNING_DAYS} יום):*")
        for item in expiring[:10]:
            if item["days_left"] < 30:
                emoji = "🔴"
            elif item["days_left"] < 45:
                emoji = "🟠"
            else:
                emoji = "🟡"
            try:
                exp_formatted = datetime.strptime(item["expiry_date"], "%Y-%m-%d").strftime("%d.%m")
            except ValueError:
                exp_formatted = item["expiry_date"]
            lines.append(
                f"{emoji} {item['partname']} {item['partdes'][:20]} - "
                f"{item['balance']:,} קרטונים, תפוגה {exp_formatted} ({item['days_left']} יום)"
            )
        if len(expiring) > 10:
            lines.append(f"   ... ועוד {len(expiring)-10} פריטים")
        lines.append("")

    # Containers in transit
    if containers:
        lines.append(f"🚢 *מכולות בדרך ({len(containers)}):*")
        for c in containers:
            container_info = f", מכולה {c['container']}" if c['container'] else ""
            lines.append(
                f"📦 {c['ordname']} {c['supplier'][:15]} - "
                f"ETA {c['eta']}{container_info}"
            )
            if c['items']:
                lines.append(f"   └ {c['items']}")
        lines.append("")

    # Pending purchase orders
    if pending:
        lines.append(f"📋 *הזמנות רכש ממתינות ({len(pending)}):*")
        for p in pending:
            lines.append(f"• {p['ordname']} {p['supplier'][:15]} - {p['items']}")
        lines.append("")

    # Summary line
    total_issues = len(low_stock) + len(expiring)
    if total_issues == 0:
        lines.append("✅ הכל תקין! יום טוב 🎉")
    else:
        lines.append(f"⚠️ סה\"כ {total_issues} פריטים דורשים תשומת לב")

    return "\n".join(lines)


def send_whatsapp(text, phone):
    """Send a WhatsApp message via TimelineAI."""
    url = "https://app.timelines.ai/integrations/api/messages"
    headers = {
        "Authorization": f"Bearer {TIMELINEAI_TOKEN}",
        "Content-Type": "application/json"
    }
    payload = {"phone": phone, "text": text}

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()


def main():
    print(f"[{datetime.now()}] Starting daily reorder alerts...")

    # 1. Query all data
    print("Querying Priority ERP...")
    low_stock = get_low_stock()
    print(f"  Found {len(low_stock)} products below reorder point")

    expiring = get_expiring_products()
    print(f"  Found {len(expiring)} expiring products")

    containers = get_containers_in_transit()
    print(f"  Found {len(containers)} containers in transit")

    pending = get_pending_purchase_orders()
    print(f"  Found {len(pending)} pending purchase orders")

    # 2. Build message
    message = build_message(low_stock, expiring, containers, pending)
    print(f"\n--- Message Preview ---\n{message}\n--- End Preview ---\n")

    # 3. Send to all recipients
    for recipient in RECIPIENTS:
        print(f"Sending to {recipient['name']} ({recipient['phone']})...")
        try:
            result = send_whatsapp(message, recipient["phone"])
            msg_uid = result.get("data", {}).get("message_uid", "N/A")
            print(f"  Sent! Message UID: {msg_uid}")
        except Exception as e:
            print(f"  ERROR sending to {recipient['name']}: {e}")

    print(f"\n[{datetime.now()}] Done!")


if __name__ == "__main__":
    main()
