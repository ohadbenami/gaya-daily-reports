# Gaya Daily Reports

Automated daily reports for Gaya Foods via GitHub Actions.

## Reports

### 🚢 Containers Report (מכולות בכנ"מ)
- **Schedule:** Daily at 8:30 AM Israel time (except Saturday)
- **Recipients:** Yuval, Ohad (WhatsApp)
- **Content:** Excel report with all containers in customs

## Setup

### GitHub Secrets Required

Add these secrets in repository Settings → Secrets → Actions:

| Secret | Value |
|--------|-------|
| `PRIORITY_API_HOST` | `https://p.priority-connect.online/odata/Priority/tabzfdbb.ini/a230521` |
| `PRIORITY_API_TOKEN` | `EC18DD66CFA644519E2979AC0CC007ED` |
| `PRIORITY_API_PASSWORD` | `PAT` |
| `TIMELINES_API_KEY` | `f40ecfc9-31e8-4905-a920-b27e5559fabc` |

### Manual Trigger

You can manually trigger the workflow from the Actions tab.

## Local Testing

```bash
# Set environment variables
export PRIORITY_API_TOKEN="your_token"
export TIMELINES_API_KEY="your_key"

# Run
python scripts/send_containers_report.py
```
