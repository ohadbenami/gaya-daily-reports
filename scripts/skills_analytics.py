#!/usr/bin/env python3
"""
Skills Analytics - Cloud Version
מקור נתונים: Supabase skills_registry + conversations
ללא תלות בפיילסיסטם מקומי.
"""

import os
import requests
from datetime import datetime, timedelta, timezone

# ─── Config (from env / GitHub Secrets) ─────────────────────────────
SUPABASE_URL    = os.environ["SUPABASE_URL"]
SUPABASE_KEY    = os.environ["SUPABASE_KEY"]
TIMELINES_TOKEN = os.environ["TIMELINEAI_TOKEN"]
WHATSAPP_PHONE  = os.environ["WHATSAPP_PHONE"]

SUPA_HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
}

NOW = datetime.now(timezone.utc)


# ─── Supabase helpers ────────────────────────────────────────────────

def supa_get(endpoint, params=""):
    url = f"{SUPABASE_URL}/rest/v1/{endpoint}?{params}"
    r = requests.get(url, headers=SUPA_HEADERS, timeout=20)
    r.raise_for_status()
    return r.json()


def supa_patch(endpoint, filters, data):
    url = f"{SUPABASE_URL}/rest/v1/{endpoint}?{filters}"
    r = requests.patch(url, headers={**SUPA_HEADERS, "Prefer": "return=minimal"},
                       json=data, timeout=20)
    return r


# ─── Data fetching ───────────────────────────────────────────────────

def get_skills():
    """שולף את כל הסקילים מ-skills_registry."""
    rows = supa_get("skills_registry",
                    "select=name,description,usage_count,last_used_at,created_at"
                    "&order=name.asc&limit=200")
    return {r["name"]: r for r in rows}


def get_conversations():
    """שולף שיחות אחרונות לניתוח מילות מפתח."""
    rows = supa_get("conversations",
                    "select=title,summary,topic_tags,created_at"
                    "&order=created_at.desc&limit=500")
    return rows


# ─── Analytics logic ─────────────────────────────────────────────────

SKILL_KEYWORDS = {
    "morning-briefing":          ["בוקר טוב", "בוקר", "daily summary", "דשבורד", "morning"],
    "monthly-sales":             ["מכירות חודש", "sales", "דשבורד מכירות", "monthly sales"],
    "daily-pallets":             ["משטחים", "חלוקה", "pallets", "נהגים"],
    "smart-reorder":             ["מה להזמין", "רשימת רכש", "חסרים", "reorder", "הזמן"],
    "containers-report":         ["מכולות", "דוח מכולות", "ETA", "container"],
    "customer-card":             ["כרטיס לקוח", "לקוח", "customer card", "obligo"],
    "product-card":              ["מוצר", "כרטיס מוצר", "product card", "מקט"],
    "product-card-supabase":     ["מוצר supa", "product supabase", "כרטיס מוצר מהיר"],
    "smart-collections":         ["גביה", "חשבוניות פתוחות", "מי חייב", "collections"],
    "inventory-intake":          ["קליטות", "דוח קליטות", "intake"],
    "gaya-cfo":                  ["cfo", "כספים", "תזרים", "רווחיות", "finance"],
    "gaya-sales":                ["מכירות סוכנים", "sales agent", "demand"],
    "gaya-ops":                  ["מלאי ops", "reorder ops", "ספקים"],
    "gaya-orchestrator":         ["דשבורד מנכל", "executive", "orchestrator"],
    "speech-generator":          ["שלח קול", "voice", "generate speech", "tts"],
    "music-generator":           ["צור מוזיקה", "music"],
    "fal-video":                 ["סרטון ai", "fal", "veo3", "kling"],
    "kinetic-video-creator":     ["סרטון קינטי", "kinetic"],
    "gaya-cinematic-video":      ["סרטון מוצר", "cinematic", "4k video"],
    "napkin-ai":                 ["napkin", "diagram"],
    "social-post":               ["פוסט", "social post"],
    "frontend-design":           ["frontend", "בנה ממשק", "design component"],
    "fireflies":                 ["פגישות", "תמלולים", "fireflies", "meeting summary"],
    "transcribe":                ["תמלל", "transcribe", "הקלטה"],
    "daily-email-digest":        ["מיילים", "email digest"],
    "gemini-research":           ["gemini", "מחקר ספק", "deep research"],
    "perplexity":                ["pplx", "perplexity", "מחקר"],
    "openai-gpt":                ["gpt", "dalle", "whisper", "openai"],
    "recall":                    ["מי קיבל", "מנה", "recall", "עקיבות"],
    "churn-analysis":            ["ניתוח נטישה", "churn"],
    "priority-create-part":      ["מקט חדש", "פתח מקט", "create part"],
    "priority-agent-transfer":   ["העבר לקוחות", "agent transfer"],
    "priority-mastery":          ["priority query", "raw_query"],
    "procurement-manager":       ["רכש", "procurement"],
    "brainstorming":             ["brainstorm", "סיעור מוחות"],
    "git-worktrees":             ["git worktree", "branch isolation"],
    "systematic-debugging":      ["debug", "באג", "שגיאה"],
    "writing-skills":            ["כתוב סקיל", "write skill"],
    "react-best-practices":      ["react", "next.js", "component"],
    "security-audit":            ["security audit", "semgrep"],
    "agent-browser":             ["browser automation", "playwright"],
    "nvr-cameras":               ["מצלמות", "nvr", "cameras", "cctv"],
    "task-manager":              ["משימות", "standup", "focus", "inbox"],
}


def days_ago(dt_str):
    if not dt_str:
        return 9999
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        return (NOW - dt).days
    except Exception:
        return 9999


def count_mentions(skill_name, conversations):
    keywords = SKILL_KEYWORDS.get(skill_name, [skill_name.replace("-", " ")])
    count = 0
    for conv in conversations:
        text = " ".join([
            str(conv.get("title", "")),
            str(conv.get("summary", "")),
            " ".join(conv.get("topic_tags") or []),
        ]).lower()
        if any(kw.lower() in text for kw in keywords):
            count += 1
    return count


def categorize(skill_row, conv_count):
    last_used = skill_row.get("last_used_at")
    created   = skill_row.get("created_at", "")
    usage     = (skill_row.get("usage_count") or 0) + conv_count

    # best freshness signal
    d = min(days_ago(last_used), days_ago(created))

    if d <= 7:
        return "hot", usage, d
    elif d <= 30:
        return "active", usage, d
    else:
        return "cold", usage, d


# ─── WhatsApp ────────────────────────────────────────────────────────

def send_whatsapp(text):
    resp = requests.post(
        "https://app.timelines.ai/integrations/api/messages",
        headers={
            "Authorization": f"Bearer {TIMELINES_TOKEN}",
            "Content-Type": "application/json",
        },
        json={"phone": WHATSAPP_PHONE, "text": text},
        timeout=30,
    )
    if resp.status_code not in (200, 201):
        raise RuntimeError(f"WhatsApp error {resp.status_code}: {resp.text[:200]}")
    return resp


# ─── Build message ───────────────────────────────────────────────────

def build_message(skills, conversations):
    today = datetime.now().strftime("%d.%m.%Y")
    hot, active, cold = [], [], []

    for name, row in skills.items():
        conv_count = count_mentions(name, conversations)
        cat, usage, d = categorize(row, conv_count)
        item = (name, usage, d)
        if cat == "hot":
            hot.append(item)
        elif cat == "active":
            active.append(item)
        else:
            cold.append(item)

    hot.sort(key=lambda x: (-x[1], x[2]))
    active.sort(key=lambda x: (-x[1], x[2]))
    cold.sort(key=lambda x: x[2])

    lines = []
    lines.append(f"📊 *Skills Analytics | {today}*")
    lines.append("━━━━━━━━━━━━━━━━━━━━━")
    lines.append(f"🧠 *{len(skills)} סקילים* | {len(hot)} חמים | {len(active)} פעילים | {len(cold)} קרים")
    lines.append(f"💬 שיחות מתועדות: {len(conversations)}")
    lines.append("")

    if hot:
        lines.append(f"🔥 *חמים - שבוע אחרון ({len(hot)}):*")
        for name, usage, d in hot[:10]:
            ago = "היום" if d == 0 else f"{d}d ago"
            u = f"×{usage}" if usage > 0 else "—"
            lines.append(f"  • `{name}` {u} | {ago}")
        if len(hot) > 10:
            lines.append(f"  _...ועוד {len(hot)-10}_")
        lines.append("")

    if active:
        lines.append(f"✅ *פעילים - חודש אחרון ({len(active)}):*")
        for name, usage, _ in active[:8]:
            u = f"×{usage}" if usage > 0 else "—"
            lines.append(f"  • `{name}` {u}")
        if len(active) > 8:
            lines.append(f"  _...ועוד {len(active)-8}_")
        lines.append("")

    if cold:
        lines.append(f"🧊 *לא בשימוש ({len(cold)}):*")
        cold_names = [n for n, _, _ in cold[:12]]
        lines.append("  " + ", ".join(f"`{n}`" for n in cold_names))
        if len(cold) > 12:
            lines.append(f"  _...ועוד {len(cold)-12}_")
        lines.append("")

    # Top 5
    all_items = hot + active + cold
    top = sorted(all_items, key=lambda x: -x[1])[:5]
    if any(u > 0 for _, u, _ in top):
        lines.append("🏆 *הכי בשימוש:*")
        for i, (name, usage, _) in enumerate(top, 1):
            if usage > 0:
                lines.append(f"  {i}. `{name}` ×{usage}")
        lines.append("")

    lines.append("_Gaya AI System • Skills Analytics_")
    return "\n".join(lines)


# ─── Main ────────────────────────────────────────────────────────────

def main():
    print(f"📊 Skills Analytics | {NOW.strftime('%Y-%m-%d %H:%M')} UTC")

    print("1. שולף skills_registry...")
    skills = get_skills()
    print(f"   {len(skills)} סקילים")

    print("2. שולף conversations...")
    conversations = get_conversations()
    print(f"   {len(conversations)} שיחות")

    print("3. בונה הודעה...")
    message = build_message(skills, conversations)
    print(message)

    print("4. שולח WhatsApp...")
    send_whatsapp(message)
    print("✅ נשלח!")


if __name__ == "__main__":
    main()
