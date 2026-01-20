#!/usr/bin/env python3
"""
Daily Containers Report - Gaya Foods
Fetches container data from Priority, creates Excel report, sends via WhatsApp
"""

import os
import requests
import json
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Configuration
PRIORITY_API_HOST = os.environ.get('PRIORITY_API_HOST', 'https://p.priority-connect.online/odata/Priority/tabzfdbb.ini/a230521')
PRIORITY_API_TOKEN = os.environ.get('PRIORITY_API_TOKEN')
PRIORITY_API_PASSWORD = os.environ.get('PRIORITY_API_PASSWORD', 'PAT')
TIMELINES_API_KEY = os.environ.get('TIMELINES_API_KEY')

# WhatsApp recipients
RECIPIENTS = [
    {'name': 'יובל', 'phone': '972505267110'},
    {'name': 'אוהד', 'phone': '972528012869'},
]

def fetch_containers():
    """Fetch container data from Priority ERP"""
    url = f"{PRIORITY_API_HOST}/PORDERS"
    params = {
        '$filter': "STATDES eq 'כנ\"מ ללא BL' or STATDES eq 'בדרך'",
        '$select': 'ORDNAME,SUPNAME,CDES,CURDATE,QPRICE,STATDES,IMPFNUM,NOA_ETA,NOA_ETD,NOA_KONTAINER',
        '$orderby': 'NOA_ETA desc',
        '$top': '50'
    }

    response = requests.get(
        url,
        params=params,
        auth=(PRIORITY_API_TOKEN, PRIORITY_API_PASSWORD),
        headers={'Content-Type': 'application/json'}
    )

    if response.status_code == 200:
        data = response.json()
        return data.get('value', [])
    else:
        print(f"Error fetching data: {response.status_code}")
        print(response.text)
        return []

def calculate_days_in_port(eta_str):
    """Calculate days since ETA (days in port)"""
    if not eta_str:
        return 0
    try:
        eta = datetime.fromisoformat(eta_str.replace('Z', '+00:00'))
        today = datetime.now(eta.tzinfo) if eta.tzinfo else datetime.now()
        delta = today - eta
        return max(0, delta.days)
    except:
        return 0

def format_date(date_str):
    """Format date to DD.MM.YY"""
    if not date_str:
        return '-'
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime('%d.%m.%y')
    except:
        return date_str

def create_excel_report(containers):
    """Create formatted Excel report"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "דשבורד מכולות"

    # Colors
    RED_FILL = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    ORANGE_FILL = PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid")
    GREEN_FILL = PatternFill(start_color="77DD77", end_color="77DD77", fill_type="solid")
    HEADER_FILL = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
    LIGHT_BLUE = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    DARK_RED = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")

    # Fonts
    TITLE_FONT = Font(name='Arial', size=24, bold=True, color="1A5276")
    HEADER_FONT = Font(name='Arial', size=12, bold=True, color="FFFFFF")
    BOLD_FONT = Font(name='Arial', size=11, bold=True)
    BIG_NUMBER = Font(name='Arial', size=28, bold=True, color="2C3E50")
    ALERT_FONT = Font(name='Arial', size=14, bold=True, color="C0392B")

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Process containers
    processed = []
    for c in containers:
        days = calculate_days_in_port(c.get('NOA_ETA'))
        processed.append({
            'po': c.get('ORDNAME', ''),
            'container': c.get('NOA_KONTAINER', c.get('IMPFNUM', '')),
            'eta': format_date(c.get('NOA_ETA')),
            'fob': c.get('QPRICE', 0),
            'days': days,
            'supplier': c.get('CDES', '')
        })

    # Sort by days descending
    processed.sort(key=lambda x: x['days'], reverse=True)

    # Count critical
    critical_count = sum(1 for c in processed if c['days'] > 30)
    warning_count = sum(1 for c in processed if 14 < c['days'] <= 30)
    total_fob = sum(c['fob'] for c in processed)

    # Set column widths
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 18

    ws.sheet_view.rightToLeft = True

    # Title
    today_str = datetime.now().strftime('%d.%m.%Y')
    ws.merge_cells('B2:G2')
    ws['B2'] = f"🚢 דוח מכולות בכנ\"מ - Gaya Foods"
    ws['B2'].font = TITLE_FONT
    ws['B2'].alignment = Alignment(horizontal='center')

    ws.merge_cells('B3:G3')
    ws['B3'] = f"📅 תאריך: {today_str} | עדכון אוטומטי יומי"
    ws['B3'].font = Font(size=12, color="7F8C8D")
    ws['B3'].alignment = Alignment(horizontal='center')

    # Summary boxes
    ws.merge_cells('B5:C6')
    ws['B5'] = str(len(processed))
    ws['B5'].font = BIG_NUMBER
    ws['B5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B5'].fill = LIGHT_BLUE
    ws.merge_cells('B7:C7')
    ws['B7'] = "סה\"כ מכולות"
    ws['B7'].font = BOLD_FONT
    ws['B7'].alignment = Alignment(horizontal='center')

    ws.merge_cells('D5:E6')
    ws['D5'] = f"${total_fob/1000000:.2f}M"
    ws['D5'].font = BIG_NUMBER
    ws['D5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D5'].fill = LIGHT_BLUE
    ws.merge_cells('D7:E7')
    ws['D7'] = "שווי כולל"
    ws['D7'].font = BOLD_FONT
    ws['D7'].alignment = Alignment(horizontal='center')

    ws.merge_cells('F5:G6')
    ws['F5'] = f"{critical_count} 🔴"
    ws['F5'].font = Font(name='Arial', size=28, bold=True, color="C0392B")
    ws['F5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['F5'].fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
    ws.merge_cells('F7:G7')
    ws['F7'] = "קריטי (>30 יום)"
    ws['F7'].font = BOLD_FONT
    ws['F7'].alignment = Alignment(horizontal='center')

    # Alert
    if critical_count > 0:
        ws.merge_cells('B9:G9')
        ws['B9'] = f"⚠️ התראה: {critical_count} מכולות מעל 30 יום בנמל!"
        ws['B9'].font = ALERT_FONT
        ws['B9'].alignment = Alignment(horizontal='center')
        ws['B9'].fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")

    # Table header
    headers = ["#", "הזמנה", "מכולה", "ETA", "FOB $", "ימים", "סטטוס"]
    row = 11
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col+1, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border

    # Table data
    for i, cont in enumerate(processed, 1):
        row = 11 + i

        if cont['days'] > 100:
            status = "🔴🔴 קריטי!"
            fill = DARK_RED
            font_color = "FFFFFF"
        elif cont['days'] > 30:
            status = "🔴 קריטי"
            fill = RED_FILL
            font_color = "000000"
        elif cont['days'] > 14:
            status = "🟠 אזהרה"
            fill = ORANGE_FILL
            font_color = "000000"
        else:
            status = "🟢 תקין"
            fill = GREEN_FILL
            font_color = "000000"

        data = [i, cont['po'], cont['container'], cont['eta'], f"${cont['fob']:,.0f}", f"{cont['days']}", status]

        for col, value in enumerate(data, 1):
            cell = ws.cell(row=row, column=col+1, value=value)
            cell.font = Font(name='Arial', size=11, color=font_color if col == 7 else "000000")
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            if col in [6, 7]:
                cell.fill = fill
                if cont['days'] > 100:
                    cell.font = Font(name='Arial', size=11, bold=True, color="FFFFFF")

    # Total row
    total_row = 11 + len(processed) + 1
    ws.merge_cells(f'B{total_row}:D{total_row}')
    ws[f'B{total_row}'] = "סה\"כ"
    ws[f'B{total_row}'].font = HEADER_FONT
    ws[f'B{total_row}'].alignment = Alignment(horizontal='center')
    ws[f'B{total_row}'].fill = HEADER_FILL

    ws[f'F{total_row}'] = f"${total_fob:,.0f}"
    ws[f'F{total_row}'].font = Font(name='Arial', size=12, bold=True, color="FFFFFF")
    ws[f'F{total_row}'].fill = HEADER_FILL
    ws[f'F{total_row}'].alignment = Alignment(horizontal='center')

    # Save
    filename = f"דוח_מכולות_כנמ_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    wb.save(filename)
    return filename

def upload_file(filepath):
    """Upload file to TimelineAI"""
    url = "https://app.timelines.ai/integrations/api/files_upload"
    headers = {"Authorization": f"Bearer {TIMELINES_API_KEY}"}

    with open(filepath, 'rb') as f:
        response = requests.post(url, headers=headers, files={'file': f})

    if response.status_code == 200:
        data = response.json()
        return data.get('data', {}).get('uid')
    else:
        print(f"Error uploading file: {response.status_code}")
        return None

def send_whatsapp(phone, file_uid, text):
    """Send WhatsApp message with file"""
    url = "https://app.timelines.ai/integrations/api/messages"
    headers = {
        "Authorization": f"Bearer {TIMELINES_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "phone": phone,
        "file_uid": file_uid,
        "text": text
    }

    response = requests.post(url, headers=headers, json=payload)
    return response.status_code == 200

def main():
    print("🚢 Fetching container data from Priority...")
    containers = fetch_containers()

    if not containers:
        print("No containers found or error fetching data")
        return

    print(f"Found {len(containers)} containers")

    print("📊 Creating Excel report...")
    filename = create_excel_report(containers)
    print(f"Created: {filename}")

    print("📤 Uploading file...")
    file_uid = upload_file(filename)

    if not file_uid:
        print("Failed to upload file")
        return

    print("📱 Sending to WhatsApp...")
    today = datetime.now().strftime('%d.%m.%Y')
    text = f"📊 דוח מכולות בכנ\"מ - Gaya Foods\n📅 {today}\n\n🤖 עדכון אוטומטי יומי"

    for recipient in RECIPIENTS:
        success = send_whatsapp(recipient['phone'], file_uid, text)
        status = "✅" if success else "❌"
        print(f"{status} {recipient['name']}: {recipient['phone']}")

    print("\n✅ Done!")

if __name__ == "__main__":
    main()
