#!/usr/bin/env python3
"""
Daily Containers Report - Gaya Foods (Enhanced Version)
- Container list from Monday.com (status, ETA)
- Items from Priority ERP (PORDERITEMS_SUBFORM)
- Landing cost calculations with shipping input cell
- Sends via WhatsApp to Ohad and Kiril
"""

import os
import requests
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Configuration
MONDAY_API_TOKEN = os.environ.get('MONDAY_API_TOKEN')
MONDAY_API_URL = "https://api.monday.com/v2"
PRIORITY_API_HOST = os.environ.get('PRIORITY_API_HOST', 'https://p.priority-connect.online/odata/Priority/tabzfdbb.ini/a230521')
PRIORITY_API_TOKEN = os.environ.get('PRIORITY_API_TOKEN')
PRIORITY_API_PASSWORD = os.environ.get('PRIORITY_API_PASSWORD', 'PAT')
TIMELINES_API_KEY = os.environ.get('TIMELINES_API_KEY')

# Board IDs
ORDERS_BOARD_ID = 1900622333
CURRENCIES_BOARD_ID = 1958318760

# Fixed costs
PORT_COST_ILS = 1107

# WhatsApp recipients
RECIPIENTS = [
    {'name': 'אוהד', 'phone': '972528012869'},
    {'name': 'קיריל', 'phone': '972538470070'},
]

# Styles
HEADER_FILL = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
SECTION_FILL = PatternFill(start_color="5B768A", end_color="5B768A", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
CALC_FILL = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")
LIGHT_BLUE = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")
RED_FILL = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")
GREEN_FILL = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")

HEADER_FONT = Font(name='Arial', size=14, bold=True, color="FFFFFF")
TITLE_FONT = Font(name='Arial', size=20, bold=True, color="2C3E50")
LABEL_FONT = Font(name='Arial', size=14, bold=True, color="333333")
DATA_FONT = Font(name='Arial', size=14, color="333333")
INPUT_FONT = Font(name='Arial', size=16, bold=True, color="C0392B")
CALC_FONT = Font(name='Arial', size=14, bold=True, color="27AE60")
BIG_NUMBER = Font(name='Arial', size=28, bold=True, color="2C3E50")

thin_border = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)

medium_border = Border(
    left=Side(style='medium', color='333333'),
    right=Side(style='medium', color='333333'),
    top=Side(style='medium', color='333333'),
    bottom=Side(style='medium', color='333333')
)


def monday_query(query):
    """Execute Monday.com GraphQL query"""
    headers = {
        "Authorization": MONDAY_API_TOKEN,
        "Content-Type": "application/json"
    }
    response = requests.post(MONDAY_API_URL, json={"query": query}, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Monday API error: {response.status_code}")
        return None


def priority_query(table, params):
    """Execute Priority OData query"""
    url = f"{PRIORITY_API_HOST}/{table}"
    response = requests.get(
        url,
        params=params,
        auth=(PRIORITY_API_TOKEN, PRIORITY_API_PASSWORD),
        headers={'Content-Type': 'application/json'}
    )
    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print(f"Priority API error: {response.status_code} - {response.text}")
        return []


def fetch_usd_rate():
    """Fetch USD exchange rate from Monday.com"""
    query = f'''
    {{
        boards(ids: [{CURRENCIES_BOARD_ID}]) {{
            items_page(limit: 10) {{
                items {{
                    name
                    column_values {{ id text }}
                }}
            }}
        }}
    }}
    '''
    result = monday_query(query)
    if result and 'data' in result:
        items = result['data']['boards'][0]['items_page']['items']
        for item in items:
            if item['name'] == 'USD':
                for col in item['column_values']:
                    if col['id'] == 'numeric_mkqyfw35':
                        return float(col['text']) if col['text'] else 3.5
    return 3.5


def fetch_containers_from_monday():
    """Fetch containers with active shipping statuses from Monday (בדרך, באוניה, כנמ ללא BL, etc.)"""
    query = f'''
    {{
        boards(ids: [{ORDERS_BOARD_ID}]) {{
            items_page(limit: 100) {{
                items {{
                    id
                    name
                    column_values {{ id text }}
                }}
            }}
        }}
    }}
    '''
    result = monday_query(query)
    if not result or 'data' not in result:
        return []
    
    containers = []
    items = result['data']['boards'][0]['items_page']['items']
    
    for item in items:
        cols = {c['id']: c['text'] for c in item['column_values']}
        status = cols.get('color_mkpn4sz9', '')
        
        # Filter by status - include all active shipping statuses
        valid_statuses = [
            'בדרך',           # On the way
            'באוניה',         # On ship
            'כנ"מ ללא BL',    # At port without BL
            'כנמ ללא BL',     # At port without BL (variant)
            'בנמל',           # At port
            'ממתין לשחרור',   # Waiting for release
        ]
        if status not in valid_statuses:
            continue
        
        po_number = cols.get('text_mkpnmg1y', '')
        if not po_number:
            continue
        
        containers.append({
            'po': po_number,
            'container': cols.get('text_mkpnqkdj', ''),
            'supplier': cols.get('text_mkpnsenz', '') or 'לא ידוע',
            'eta': cols.get('date_mkpnbh0z', ''),
            'fob_total': float(cols.get('numeric_mkpnhbgt', '0') or '0'),
            'currency': cols.get('text_mkpnkq', '$'),
            'status': status,
        })
    
    return containers


def fetch_items_from_priority(po_number):
    """Fetch items for a PO from Priority PORDERITEMS_SUBFORM"""
    params = {
        '$filter': f"ORDNAME eq '{po_number}'",
        '$select': 'ORDNAME,CDES',
        '$expand': 'PORDERITEMS_SUBFORM($select=PARTNAME,PDES,TQUANT,PRICE,QPRICE)'
    }
    data = priority_query('PORDERS', params)
    
    if not data:
        return []
    
    items = []
    for order in data:
        for item in order.get('PORDERITEMS_SUBFORM', []):
            qty = item.get('TQUANT', 0)
            if qty <= 0:
                continue
            items.append({
                'sku': item.get('PARTNAME', ''),
                'description': item.get('PDES', ''),
                'quantity': int(qty),
                'unit': 'קרט',
                'unit_price': item.get('PRICE', 0)
            })
    
    return items


def calculate_days_in_port(eta_str):
    """Calculate days since ETA"""
    if not eta_str:
        return 0
    try:
        eta = datetime.strptime(eta_str, '%Y-%m-%d')
        delta = datetime.now() - eta
        return max(0, delta.days)
    except:
        return 0


def create_summary_sheet(ws, containers, usd_rate):
    """Create summary dashboard sheet"""
    ws.sheet_view.rightToLeft = True
    
    widths = [3, 15, 18, 25, 12, 15, 12, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Title
    ws.merge_cells('B2:H2')
    ws['B2'] = f"🚢 דוח מכולות בכנ\"מ - Gaya Foods"
    ws['B2'].font = TITLE_FONT
    ws['B2'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('B3:H3')
    ws['B3'] = f"📅 {datetime.now().strftime('%d.%m.%Y')} | 💵 שער: {usd_rate}"
    ws['B3'].font = Font(size=12, color="7F8C8D")
    ws['B3'].alignment = Alignment(horizontal='center')
    
    # KPIs
    total_fob = sum(c['fob_total'] for c in containers)
    critical_count = sum(1 for c in containers if calculate_days_in_port(c['eta']) > 30)
    
    row = 5
    ws.merge_cells(f'B{row}:C{row+1}')
    ws[f'B{row}'] = str(len(containers))
    ws[f'B{row}'].font = BIG_NUMBER
    ws[f'B{row}'].fill = LIGHT_BLUE
    ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'B{row+2}:C{row+2}')
    ws[f'B{row+2}'] = "סה\"כ מכולות"
    ws[f'B{row+2}'].font = LABEL_FONT
    ws[f'B{row+2}'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells(f'D{row}:E{row+1}')
    ws[f'D{row}'] = f"${total_fob/1000:.0f}K"
    ws[f'D{row}'].font = BIG_NUMBER
    ws[f'D{row}'].fill = LIGHT_BLUE
    ws[f'D{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'D{row+2}:E{row+2}')
    ws[f'D{row+2}'] = "FOB כולל"
    ws[f'D{row+2}'].font = LABEL_FONT
    ws[f'D{row+2}'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells(f'F{row}:G{row+1}')
    ws[f'F{row}'] = f"{critical_count} 🔴"
    ws[f'F{row}'].font = Font(name='Arial', size=28, bold=True, color="C0392B")
    ws[f'F{row}'].fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
    ws[f'F{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'F{row+2}:G{row+2}')
    ws[f'F{row+2}'] = "קריטי (>30 יום)"
    ws[f'F{row+2}'].font = LABEL_FONT
    ws[f'F{row+2}'].alignment = Alignment(horizontal='center')
    
    # Table
    row = 10
    headers = ["#", "הזמנה", "מכולה", "ספק", "ETA", "FOB $", "ימים", "גיליון"]
    for col, header in enumerate(headers, 2):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    containers_sorted = sorted(containers, key=lambda x: calculate_days_in_port(x['eta']), reverse=True)
    
    for i, cont in enumerate(containers_sorted, 1):
        row += 1
        days = calculate_days_in_port(cont['eta'])
        
        if days > 30:
            row_fill = RED_FILL
        elif days > 14:
            row_fill = ORANGE_FILL
        else:
            row_fill = GREEN_FILL
        
        eta_fmt = ''
        if cont['eta']:
            try:
                eta_fmt = datetime.strptime(cont['eta'], '%Y-%m-%d').strftime('%d.%m.%y')
            except:
                eta_fmt = cont['eta']
        
        values = [i, cont['po'], cont['container'] or '-', cont['supplier'][:20], 
                  eta_fmt or '-', f"${cont['fob_total']:,.0f}", str(days), f"→ {cont['po']}"]
        
        for col, value in enumerate(values, 2):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = DATA_FONT
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            if col == 8:
                cell.fill = row_fill


def create_container_sheet(wb, container, items, usd_rate):
    """Create detailed sheet for a container with items from Priority"""
    sheet_name = container['po'][:31]
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_view.rightToLeft = True
    
    widths = [4, 20, 12, 40, 10, 14, 14, 14, 14, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Title
    row = 2
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = f"📦 עלות נחיתה - מכולה {container['po']}"
    ws[f'B{row}'].font = TITLE_FONT
    ws[f'B{row}'].alignment = Alignment(horizontal='center')
    
    # Container details
    row = 4
    eta_fmt = ''
    if container['eta']:
        try:
            eta_fmt = datetime.strptime(container['eta'], '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            eta_fmt = container['eta']
    
    ws.cell(row=row, column=2, value="מספר מכולה:").font = LABEL_FONT
    ws.cell(row=row, column=3, value=container['container'] or '-').font = DATA_FONT
    ws.cell(row=row, column=5, value="ספק:").font = LABEL_FONT
    ws.cell(row=row, column=6, value=container['supplier']).font = DATA_FONT
    row += 1
    ws.cell(row=row, column=2, value="ETA:").font = LABEL_FONT
    ws.cell(row=row, column=3, value=eta_fmt or '-').font = DATA_FONT
    ws.cell(row=row, column=5, value="סטטוס:").font = LABEL_FONT
    ws.cell(row=row, column=6, value=container['status']).font = DATA_FONT
    
    # Parameters section
    row += 2
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = "⚙️ פרמטרים לחישוב"
    ws[f'B{row}'].font = HEADER_FONT
    ws[f'B{row}'].fill = SECTION_FILL
    ws[f'B{row}'].alignment = Alignment(horizontal='center')
    for col in range(2, 11):
        ws.cell(row=row, column=col).fill = SECTION_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    rate_row = row
    ws.cell(row=row, column=2, value="שער דולר:").font = LABEL_FONT
    rate_cell = ws.cell(row=row, column=3, value=usd_rate)
    rate_cell.font = DATA_FONT
    rate_cell.number_format = '0.000'
    rate_cell.fill = LIGHT_BLUE
    rate_cell.border = thin_border
    
    ws.cell(row=row, column=5, value="עלות נמל (₪):").font = LABEL_FONT
    ws.cell(row=row, column=6, value=PORT_COST_ILS).font = DATA_FONT
    ws.cell(row=row, column=6).fill = LIGHT_BLUE
    ws.cell(row=row, column=6).border = thin_border
    
    ws.cell(row=row, column=8, value="מכס:").font = LABEL_FONT
    ws.cell(row=row, column=9, value="0%").font = DATA_FONT
    
    # Shipping input
    row += 1
    shipping_row = row
    ws.cell(row=row, column=2, value="👇 עלות הובלה סה\"כ ($):").font = Font(name='Arial', size=14, bold=True, color="C0392B")
    shipping_cell = ws.cell(row=row, column=3, value="")
    shipping_cell.fill = INPUT_FILL
    shipping_cell.border = medium_border
    shipping_cell.font = INPUT_FONT
    shipping_cell.alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=5, value="← מלא כאן").font = Font(size=12, italic=True, color="888888")
    
    # Items table
    row += 2
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = "📋 פירוט מק\"טים ועלות נחיתה"
    ws[f'B{row}'].font = HEADER_FONT
    ws[f'B{row}'].fill = SECTION_FILL
    ws[f'B{row}'].alignment = Alignment(horizontal='center')
    for col in range(2, 11):
        ws.cell(row=row, column=col).fill = SECTION_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    row += 1
    headers = ["#", "מק\"ט", "כמות", "תיאור", "FOB/יח' $", "FOB סה\"כ $", "הובלה/יח' $", "נמל/יח' ₪", "עלות נחיתה/יח' ₪"]
    for col, header in enumerate(headers, 2):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[row].height = 35
    
    if not items:
        row += 1
        ws.merge_cells(f'B{row}:J{row}')
        ws[f'B{row}'] = "אין פריטים להציג"
        ws[f'B{row}'].font = Font(size=12, italic=True, color="888888")
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        return
    
    total_units = sum(item['quantity'] for item in items if item['quantity'] > 0)
    if total_units == 0:
        total_units = 1
    
    row += 1
    first_data_row = row
    shipping_ref = f"$C${shipping_row}"
    rate_ref = f"$C${rate_row}"
    
    for i, item in enumerate(items, 1):
        if item['quantity'] <= 0:
            continue
        
        fob_total = item['quantity'] * item['unit_price']
        
        ws.cell(row=row, column=2, value=i).font = DATA_FONT
        ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
        
        ws.cell(row=row, column=3, value=item['sku']).font = DATA_FONT
        ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')
        
        ws.cell(row=row, column=4, value=item['quantity']).font = DATA_FONT
        ws.cell(row=row, column=4).number_format = '#,##0'
        ws.cell(row=row, column=4).alignment = Alignment(horizontal='center')
        
        desc = item['description'][:35] if item['description'] else ''
        ws.cell(row=row, column=5, value=desc).font = DATA_FONT
        
        ws.cell(row=row, column=6, value=item['unit_price']).font = DATA_FONT
        ws.cell(row=row, column=6).number_format = '$#,##0.00'
        ws.cell(row=row, column=6).alignment = Alignment(horizontal='center')
        
        ws.cell(row=row, column=7, value=fob_total).font = DATA_FONT
        ws.cell(row=row, column=7).number_format = '$#,##0'
        ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')
        
        # Shipping per unit formula
        ship_formula = f'=IF({shipping_ref}="","",{shipping_ref}/{total_units})'
        ws.cell(row=row, column=8, value=ship_formula).font = CALC_FONT
        ws.cell(row=row, column=8).fill = CALC_FILL
        ws.cell(row=row, column=8).number_format = '$#,##0.00'
        ws.cell(row=row, column=8).alignment = Alignment(horizontal='center')
        
        # Port per unit
        port_per_unit = PORT_COST_ILS / total_units
        ws.cell(row=row, column=9, value=port_per_unit).font = DATA_FONT
        ws.cell(row=row, column=9).fill = CALC_FILL
        ws.cell(row=row, column=9).number_format = '₪#,##0.00'
        ws.cell(row=row, column=9).alignment = Alignment(horizontal='center')
        
        # Landing cost formula
        landing_formula = f'=IF(H{row}="",F{row}*{rate_ref}+I{row},(F{row}+H{row})*{rate_ref}+I{row})'
        ws.cell(row=row, column=10, value=landing_formula)
        ws.cell(row=row, column=10).font = Font(name='Arial', size=14, bold=True, color="27AE60")
        ws.cell(row=row, column=10).fill = CALC_FILL
        ws.cell(row=row, column=10).number_format = '₪#,##0.00'
        ws.cell(row=row, column=10).alignment = Alignment(horizontal='center')
        
        for col in range(2, 11):
            ws.cell(row=row, column=col).border = thin_border
        ws.row_dimensions[row].height = 25
        row += 1
    
    last_data_row = row - 1
    
    # Totals
    ws.merge_cells(f'B{row}:E{row}')
    ws.cell(row=row, column=2, value="סה\"כ").font = HEADER_FONT
    ws.cell(row=row, column=2).fill = HEADER_FILL
    ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
    for col in range(2, 6):
        ws.cell(row=row, column=col).fill = HEADER_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    ws.cell(row=row, column=6).fill = HEADER_FILL
    ws.cell(row=row, column=6).border = thin_border
    
    ws.cell(row=row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})").font = HEADER_FONT
    ws.cell(row=row, column=7).fill = HEADER_FILL
    ws.cell(row=row, column=7).number_format = '$#,##0'
    ws.cell(row=row, column=7).alignment = Alignment(horizontal='center')
    ws.cell(row=row, column=7).border = thin_border
    
    for col in [8, 9, 10]:
        ws.cell(row=row, column=col).fill = HEADER_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    # Legend
    row += 2
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = f"📝 מלא הובלה בתא הצהוב ← החישובים יתעדכנו | סה\"כ יחידות: {total_units:,}"
    ws[f'B{row}'].font = Font(size=11, italic=True, color="666666")
    ws[f'B{row}'].alignment = Alignment(horizontal='center')


def create_excel_report(containers, usd_rate):
    """Create full Excel with summary + per-container sheets"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "סיכום מכולות"
    
    create_summary_sheet(ws, containers, usd_rate)
    
    for container in containers:
        print(f"  Fetching items for {container['po']}...")
        items = fetch_items_from_priority(container['po'])
        print(f"    Found {len(items)} items")
        create_container_sheet(wb, container, items, usd_rate)
    
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
        return response.json().get('data', {}).get('uid')
    print(f"Upload error: {response.status_code}")
    return None


def send_whatsapp(phone, file_uid, text):
    """Send WhatsApp message with file"""
    url = "https://app.timelines.ai/integrations/api/messages"
    headers = {"Authorization": f"Bearer {TIMELINES_API_KEY}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, json={"phone": phone, "file_uid": file_uid, "text": text})
    return response.status_code == 200


def main():
    print("💵 Fetching USD rate...")
    usd_rate = fetch_usd_rate()
    print(f"USD Rate: {usd_rate}")
    
    print("🚢 Fetching containers from Monday...")
    containers = fetch_containers_from_monday()
    
    if not containers:
        print("No containers found")
        return
    
    print(f"Found {len(containers)} containers")
    
    print("📊 Creating Excel with items from Priority...")
    filename = create_excel_report(containers, usd_rate)
    print(f"Created: {filename}")
    
    print("📤 Uploading...")
    file_uid = upload_file(filename)
    if not file_uid:
        print("Upload failed")
        return
    
    print("📱 Sending WhatsApp...")
    today = datetime.now().strftime('%d.%m.%Y')
    critical = sum(1 for c in containers if calculate_days_in_port(c['eta']) > 30)
    text = f"🚢 דוח מכולות בכנ\"מ - Gaya Foods\n📅 {today}\n📦 {len(containers)} מכולות"
    if critical > 0:
        text += f"\n🔴 {critical} קריטי (>30 יום)"
    text += "\n\n🤖 עדכון יומי"
    
    for r in RECIPIENTS:
        ok = send_whatsapp(r['phone'], file_uid, text)
        print(f"{'✅' if ok else '❌'} {r['name']}: {r['phone']}")
    
    print("\n✅ Done!")


if __name__ == "__main__":
    main()
