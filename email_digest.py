#!/usr/bin/env python3
"""
Morning Email Digest - Gaya Foods
Fetches emails from MS365, summarizes with Claude, sends via WhatsApp.
Runs daily at 08:05 Israel time.
"""

import os
import json
import requests
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any

# --- Configuration from environment variables ---
MS365_CLIENT_ID = os.environ["MS365_CLIENT_ID"]
MS365_CLIENT_SECRET = os.environ["MS365_CLIENT_SECRET"]
MS365_TENANT_ID = os.environ["MS365_TENANT_ID"]
MS365_USER_EMAIL = os.environ["MS365_USER_EMAIL"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
TIMELINEAI_TOKEN = os.environ["TIMELINEAI_TOKEN"]
WHATSAPP_PHONE = os.environ.get("WHATSAPP_PHONE", "972528012869")

# Microsoft Graph API endpoints
GRAPH_API_URL = "https://graph.microsoft.com/v1.0"
TOKEN_URL = f"https://login.microsoftonline.com/{MS365_TENANT_ID}/oauth2/v2.0/token"

# Claude API
CLAUDE_API_URL = "https://api.anthropic.com/v1/messages"


def get_ms365_token() -> str:
    """Get access token for Microsoft Graph API using client credentials."""
    data = {
        "client_id": MS365_CLIENT_ID,
        "client_secret": MS365_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(TOKEN_URL, data=data, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]


def fetch_emails(token: str) -> List[Dict[str, Any]]:
    """Fetch emails from yesterday 19:00 to now."""
    # Calculate time range: yesterday 19:00 Israel time to now
    israel_tz = timezone(timedelta(hours=2))  # IST (winter)
    now = datetime.now(israel_tz)
    yesterday_7pm = (now - timedelta(days=1)).replace(hour=19, minute=0, second=0, microsecond=0)

    # Convert to ISO format for Graph API
    from_time = yesterday_7pm.strftime("%Y-%m-%dT%H:%M:%SZ")

    url = f"{GRAPH_API_URL}/users/{MS365_USER_EMAIL}/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    params = {
        "$filter": f"receivedDateTime ge {from_time}",
        "$orderby": "receivedDateTime desc",
        "$top": 50,
        "$select": "id,subject,from,receivedDateTime,importance,isRead,bodyPreview,hasAttachments"
    }

    response = requests.get(url, headers=headers, params=params, timeout=60)
    response.raise_for_status()
    return response.json().get("value", [])


def categorize_emails(emails: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """Categorize emails by topic."""
    categories = {
        "finance": [],      # כספים
        "orders": [],       # הזמנות
        "internal": [],     # פנימי
        "urgent": [],       # דחוף
        "other": []         # אחר
    }

    finance_keywords = ["חשבונית", "תשלום", "העברה", "invoice", "payment", "bank", "בנק"]
    order_keywords = ["הזמנה", "משלוח", "po", "order", "shipment", "delivery", "מכולה"]
    urgent_keywords = ["דחוף", "urgent", "asap", "חשוב", "important", "מיידי"]

    for email in emails:
        subject = email.get("subject", "").lower()
        body_preview = email.get("bodyPreview", "").lower()
        from_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        importance = email.get("importance", "normal")

        combined_text = f"{subject} {body_preview}"

        # Check urgency first
        if importance == "high" or any(kw in combined_text for kw in urgent_keywords):
            categories["urgent"].append(email)
        # Check categories
        elif any(kw in combined_text for kw in finance_keywords):
            categories["finance"].append(email)
        elif any(kw in combined_text for kw in order_keywords):
            categories["orders"].append(email)
        elif "gaya-foods" in from_email or "gaya" in from_email:
            categories["internal"].append(email)
        else:
            categories["other"].append(email)

    return categories


def summarize_with_claude(emails: List[Dict[str, Any]], categories: Dict[str, List]) -> str:
    """Use Claude to create a smart summary of the emails."""
    # Build email data for Claude
    email_data = []
    for email in emails[:20]:  # Limit to 20 most recent
        email_data.append({
            "subject": email.get("subject", ""),
            "from": email.get("from", {}).get("emailAddress", {}).get("name", "Unknown"),
            "time": email.get("receivedDateTime", ""),
            "preview": email.get("bodyPreview", "")[:200],
            "importance": email.get("importance", "normal"),
            "hasAttachments": email.get("hasAttachments", False)
        })

    prompt = f"""אתה עוזר למנכ"ל לסכם מיילים בבוקר.

הנה המיילים שהתקבלו מאתמול 19:00:
{json.dumps(email_data, ensure_ascii=False, indent=2)}

סטטיסטיקה:
- סה"כ: {len(emails)} מיילים
- דחופים: {len(categories['urgent'])}
- כספים: {len(categories['finance'])}
- הזמנות: {len(categories['orders'])}
- פנימי: {len(categories['internal'])}
- אחר: {len(categories['other'])}

צור סיכום קצר לוואטסאפ (מקסימום 450 תווים) בפורמט:
1. שורת פתיחה עם תאריך וסטטיסטיקה
2. אם יש דחופים - רשום אותם קודם
3. קטגוריות לפי נושא (רק אם יש)
4. שורת action items (מה דורש תשומת לב)

השתמש באימוג'ים: ☀️📬🔴💰📦⚡
כתוב בעברית. תמציתי. ישיר."""

    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    payload = {
        "model": "claude-3-5-haiku-20241022",
        "max_tokens": 500,
        "messages": [
            {"role": "user", "content": prompt}
        ]
    }

    response = requests.post(CLAUDE_API_URL, headers=headers, json=payload, timeout=60)
    response.raise_for_status()

    result = response.json()
    return result["content"][0]["text"]


def create_fallback_summary(emails: List[Dict], categories: Dict) -> str:
    """Create a simple summary if Claude API fails."""
    today = datetime.now().strftime("%d.%m")

    lines = [f"☀️ בוקר טוב! | {today}"]
    lines.append(f"📬 {len(emails)} מיילים | 🔴 {len(categories['urgent'])} דחופים")

    if categories['urgent']:
        lines.append("\n⚡ דחוף:")
        for email in categories['urgent'][:2]:
            sender = email.get("from", {}).get("emailAddress", {}).get("name", "")[:15]
            subject = email.get("subject", "")[:25]
            lines.append(f"• {sender}: {subject}")

    if categories['finance']:
        lines.append(f"\n💰 כספים ({len(categories['finance'])})")

    if categories['orders']:
        lines.append(f"📦 הזמנות ({len(categories['orders'])})")

    return "\n".join(lines)[:500]


def send_whatsapp(message: str) -> Dict:
    """Send the digest message via WhatsApp."""
    url = "https://app.timelines.ai/integrations/api/messages"
    headers = {
        "Authorization": f"Bearer {TIMELINEAI_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "phone": WHATSAPP_PHONE,
        "text": message
    }

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    return response.json()


def main():
    print(f"[{datetime.now()}] Starting morning email digest...")

    # 1. Get MS365 token
    print("Getting MS365 access token...")
    token = get_ms365_token()
    print("  Token acquired")

    # 2. Fetch emails
    print("Fetching emails from MS365...")
    emails = fetch_emails(token)
    print(f"  Found {len(emails)} emails since yesterday 19:00")

    if not emails:
        # Still send a message if no emails
        message = f"☀️ בוקר טוב! | {datetime.now().strftime('%d.%m')}\n\n📬 אין מיילים חדשים מאתמול 19:00\n\n🎯 יום פרודוקטיבי!"
        send_whatsapp(message)
        print("No emails - sent empty inbox message")
        return

    # 3. Categorize emails
    print("Categorizing emails...")
    categories = categorize_emails(emails)
    print(f"  Urgent: {len(categories['urgent'])}, Finance: {len(categories['finance'])}, "
          f"Orders: {len(categories['orders'])}, Internal: {len(categories['internal'])}")

    # 4. Summarize with Claude
    print("Generating summary with Claude...")
    try:
        summary = summarize_with_claude(emails, categories)
    except Exception as e:
        print(f"  Claude API failed: {e}")
        print("  Using fallback summary...")
        summary = create_fallback_summary(emails, categories)

    print(f"  Summary length: {len(summary)} chars")

    # 5. Send via WhatsApp
    print("Sending via WhatsApp...")
    result = send_whatsapp(summary)
    print(f"  Sent! Message UID: {result.get('data', {}).get('message_uid', 'N/A')}")

    print(f"[{datetime.now()}] Done!")


if __name__ == "__main__":
    main()
