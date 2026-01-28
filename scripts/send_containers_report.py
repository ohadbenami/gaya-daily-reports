#!/usr/bin/env python3
"""
Daily Containers Report - Gaya Foods (Enhanced Version)
Fetches container data from Monday.com, creates Excel with landing cost calculations,
sends via WhatsApp to Ohad and Kiril
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
        print(response.text)
        return None


def fetch_usd_rate():
    """Fetch USD exchange rate from Monday.com Currencies board"""
    query = f'''
    {{
        boards(ids: [{CURRENCIES_BOARD_ID}]) {{
            items_page(limit: 10) {{
                items {{
                    name
                    column_values {{
                        id
                        text
                    }}
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
    return 3.5  # Default


def fetch_containers():
    """Fetch containers with status 'בדרך' or 'כנמ ללא BL' from Monday.com"""
    query = f'''
    {{
        boards(ids: [{ORDERS_BOARD_ID}]) {{
            items_page(limit: 100) {{
                items {{
                    id
                    name
                    column_values {{
                        id
                        text
                        value
                    }}
                    subitems {{
                        id
                        name
                        column_values {{
                            id
                            text
                        }}
                    }}
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
        
        # Filter: only "בדרך" or "כנ"מ ללא BL"
        if status not in ['בדרך', 'כנ"מ ללא BL', 'כנמ ללא BL']:
            continue
        
        # Skip items without PO number
        po_number = cols.get('text_mkpnmg1y', '')
        if not po_number:
            continue
        
        # Parse subitems
        subitems = []
        for si in item.get('subitems', []):
            si_cols = {c['id']: c['text'] for c in si['column_values']}
            qty_str = si_cols.get('numeric_mkqd6mgd', '0')
            price_str = si_cols.get('numeric_mkqd8cs2', '0')
            
            # Skip discount lines (negative quantities)
            try:
                qty = float(qty_str) if qty_str else 0
                if qty <= 0:
                    continue
            except:
                qty = 0
            
            try:
                price = float(price_str) if price_str else 0
            except:
                price = 0
            
            subitems.append({
                'sku': si['name'],
                'description': si_cols.get('long_text_mkqdfhcn', ''),
                'quantity': int(qty),
                'unit': si_cols.get('text_mkqdkx16', 'קרט'),
                'unit_price': price
            })
        
        # Get FOB total
        fob_str = cols.get('numeric_mkpnhbgt', '0')
        try:
            fob_total = float(fob_str) if fob_str else 0
        except:
            fob_total = 0
        
        # Get supplier name from relation
        supplier = cols.get('board_relation_mkr56mp8', '')
        if not supplier:
            supplier = cols.get('text_mkpnsenz', 'לא ידוע')
        
        containers.append({
            'po': po_number,
            'container': cols.get('text_mkpnqkdj', ''),
            'supplier': supplier,
            'eta': cols.get('date_mkpnbh0z', ''),
            'etd': cols.get('date_mkpnywp8', ''),
            'fob_total': fob_total,
            'currency': cols.get('text_mkpnkq', '$'),
            'status': status,
            'items': subitems
        })
    
    return containers


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
    
    # Column widths
    widths = [3, 15, 18, 20, 12, 15, 12, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Title
    ws.merge_cells('B2:H2')
    ws['B2'] = f"🚢 דוח מכולות בכנ\"מ - Gaya Foods"
    ws['B2'].font = TITLE_FONT
    ws['B2'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('B3:H3')
    ws['B3'] = f"📅 תאריך: {datetime.now().strftime('%d.%m.%Y')} | 💵 שער דולר: {usd_rate}"
    ws['B3'].font = Font(size=12, color="7F8C8D")
    ws['B3'].alignment = Alignment(horizontal='center')
    
    # KPIs
    total_fob = sum(c['fob_total'] for c in containers)
    critical_count = sum(1 for c in containers if calculate_days_in_port(c['eta']) > 30)
    
    row = 5
    # Total containers
    ws.merge_cells(f'B{row}:C{row+1}')
    ws[f'B{row}'] = str(len(containers))
    ws[f'B{row}'].font = BIG_NUMBER
    ws[f'B{row}'].fill = LIGHT_BLUE
    ws[f'B{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'B{row+2}:C{row+2}')
    ws[f'B{row+2}'] = "סה\"כ מכולות"
    ws[f'B{row+2}'].font = LABEL_FONT
    ws[f'B{row+2}'].alignment = Alignment(horizontal='center')
    
    # Total FOB
    ws.merge_cells(f'D{row}:E{row+1}')
    ws[f'D{row}'] = f"${total_fob/1000:.0f}K"
    ws[f'D{row}'].font = BIG_NUMBER
    ws[f'D{row}'].fill = LIGHT_BLUE
    ws[f'D{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'D{row+2}:E{row+2}')
    ws[f'D{row+2}'] = "FOB כולל"
    ws[f'D{row+2}'].font = LABEL_FONT
    ws[f'D{row+2}'].alignment = Alignment(horizontal='center')
    
    # Critical
    ws.merge_cells(f'F{row}:G{row+1}')
    ws[f'F{row}'] = f"{critical_count} 🔴"
    ws[f'F{row}'].font = Font(name='Arial', size=28, bold=True, color="C0392B")
    ws[f'F{row}'].fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
    ws[f'F{row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'F{row+2}:G{row+2}')
    ws[f'F{row+2}'] = "קריטי (>30 יום)"
    ws[f'F{row+2}'].font = LABEL_FONT
    ws[f'F{row+2}'].alignment = Alignment(horizontal='center')
    
    # Table header
    row = 10
    headers = ["#", "הזמנה", "מכולה", "ספק", "ETA", "FOB $", "ימים", "גיליון"]
    for col, header in enumerate(headers, 2):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    # Sort by days descending
    containers_sorted = sorted(containers, key=lambda x: calculate_days_in_port(x['eta']), reverse=True)
    
    # Table data
    for i, cont in enumerate(containers_sorted, 1):
        row += 1
        days = calculate_days_in_port(cont['eta'])
        
        # Color based on days
        if days > 30:
            row_fill = RED_FILL
        elif days > 14:
            row_fill = ORANGE_FILL
        else:
            row_fill = GREEN_FILL
        
        eta_formatted = ''
        if cont['eta']:
            try:
                eta_dt = datetime.strptime(cont['eta'], '%Y-%m-%d')
                eta_formatted = eta_dt.strftime('%d.%m.%y')
            except:
                eta_formatted = cont['eta']
        
        values = [
            i,
            cont['po'],
            cont['container'] or '-',
            cont['supplier'][:15] if cont['supplier'] else '-',
            eta_formatted or '-',
            f"${cont['fob_total']:,.0f}",
            str(days),
            f"→ {cont['po']}"
        ]
        
        for col, value in enumerate(values, 2):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = DATA_FONT
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            if col in [8]:  # Days column
                cell.fill = row_fill
    
    # Navigation note
    row += 2
    ws.merge_cells(f'B{row}:H{row}')
    ws[f'B{row}'] = "💡 לחץ על שם הגיליון למטה לפירוט מק\"טים ועלות נחיתה"
    ws[f'B{row}'].font = Font(size=11, italic=True, color="7F8C8D")
    ws[f'B{row}'].alignment = Alignment(horizontal='center')


def create_container_sheet(wb, container, usd_rate):
    """Create detailed sheet for a container"""
    # Sanitize sheet name
    sheet_name = container['po'][:31] if container['po'] else 'Unknown'
    ws = wb.create_sheet(title=sheet_name)
    ws.sheet_view.rightToLeft = True
    
    # Column widths
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
    eta_formatted = ''
    if container['eta']:
        try:
            eta_dt = datetime.strptime(container['eta'], '%Y-%m-%d')
            eta_formatted = eta_dt.strftime('%d.%m.%Y')
        except:
            eta_formatted = container['eta']
    
    details = [
        ("מספר מכולה:", container['container'] or '-', "ספק:", container['supplier'] or '-'),
        ("ETA:", eta_formatted or '-', "סטטוס:", container['status']),
    ]
    for d in details:
        ws.cell(row=row, column=2, value=d[0]).font = LABEL_FONT
        ws.cell(row=row, column=3, value=d[1]).font = DATA_FONT
        ws.cell(row=row, column=5, value=d[2]).font = LABEL_FONT
        ws.cell(row=row, column=6, value=d[3]).font = DATA_FONT
        row += 1
    
    # Parameters section
    row += 1
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = "⚙️ פרמטרים לחישוב"
    ws[f'B{row}'].font = HEADER_FONT
    ws[f'B{row}'].fill = SECTION_FILL
    ws[f'B{row}'].alignment = Alignment(horizontal='center')
    for col in range(2, 11):
        ws.cell(row=row, column=col).fill = SECTION_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    # Rate + Port
    row += 1
    rate_row = row
    ws.cell(row=row, column=2, value="שער דולר:").font = LABEL_FONT
    rate_cell = ws.cell(row=row, column=3, value=usd_rate)
    rate_cell.font = DATA_FONT
    rate_cell.number_format = '0.000'
    rate_cell.fill = LIGHT_BLUE
    rate_cell.border = thin_border
    
    ws.cell(row=row, column=5, value="עלות נמל (₪):").font = LABEL_FONT
    port_cell = ws.cell(row=row, column=6, value=PORT_COST_ILS)
    port_cell.font = DATA_FONT
    port_cell.number_format = '#,##0'
    port_cell.fill = LIGHT_BLUE
    port_cell.border = thin_border
    
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
    
    ws.cell(row=row, column=5, value="← מלא כאן").font = Font(name='Arial', size=12, italic=True, color="888888")
    
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
    
    # Table headers
    row += 1
    headers = ["#", "מק\"ט", "כמות", "תיאור", "FOB/יח' $", "FOB סה\"כ $", "הובלה/יח' $", "נמל/יח' ₪", "עלות נחיתה/יח' ₪"]
    for col, header in enumerate(headers, 2):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[row].height = 35
    
    # Calculate total units
    items = container.get('items', [])
    if not items:
        row += 1
        ws.merge_cells(f'B{row}:J{row}')
        ws[f'B{row}'] = "אין פריטים להציג"
        ws[f'B{row}'].font = Font(size=12, italic=True, color="888888")
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        return
    
    total_units = sum(item['quantity'] for item in items if item['quantity'] > 0)
    if total_units == 0:
        total_units = 1  # Avoid division by zero
    
    # Data rows
    row += 1
    first_data_row = row
    shipping_cell_ref = f"$C${shipping_row}"
    rate_cell_ref = f"$C${rate_row}"
    
    for i, item in enumerate(items, 1):
        if item['quantity'] <= 0:
            continue
        
        fob_total = item['quantity'] * item['unit_price']
        
        ws.cell(row=row, column=2, value=i).font = DATA_FONT
        ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
        
        ws.cell(row=row, column=3, value=item['sku']).font = DATA_FONT
        ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')
        
        qty_cell = ws.cell(row=row, column=4, value=item['quantity'])
        qty_cell.font = DATA_FONT
        qty_cell.number_format = '#,##0'
        qty_cell.alignment = Alignment(horizontal='center')
        
        desc = item['description'][:35] if item['description'] else ''
        ws.cell(row=row, column=5, value=desc).font = DATA_FONT
        
        fob_unit_cell = ws.cell(row=row, column=6, value=item['unit_price'])
        fob_unit_cell.font = DATA_FONT
        fob_unit_cell.number_format = '$#,##0.00'
        fob_unit_cell.alignment = Alignment(horizontal='center')
        
        fob_total_cell = ws.cell(row=row, column=7, value=fob_total)
        fob_total_cell.font = DATA_FONT
        fob_total_cell.number_format = '$#,##0'
        fob_total_cell.alignment = Alignment(horizontal='center')
        
        # Shipping per unit formula
        shipping_formula = f'=IF({shipping_cell_ref}="","",{shipping_cell_ref}/{total_units})'
        ship_cell = ws.cell(row=row, column=8, value=shipping_formula)
        ship_cell.font = CALC_FONT
        ship_cell.fill = CALC_FILL
        ship_cell.number_format = '$#,##0.00'
        ship_cell.alignment = Alignment(horizontal='center')
        
        # Port per unit
        port_per_unit = PORT_COST_ILS / total_units
        port_cell = ws.cell(row=row, column=9, value=port_per_unit)
        port_cell.font = DATA_FONT
        port_cell.fill = CALC_FILL
        port_cell.number_format = '₪#,##0.00'
        port_cell.alignment = Alignment(horizontal='center')
        
        # Landing cost formula
        landing_formula = f'=IF(H{row}="",F{row}*{rate_cell_ref}+I{row},(F{row}+H{row})*{rate_cell_ref}+I{row})'
        landing_cell = ws.cell(row=row, column=10, value=landing_formula)
        landing_cell.font = Font(name='Arial', size=14, bold=True, color="27AE60")
        landing_cell.fill = CALC_FILL
        landing_cell.number_format = '₪#,##0.00'
        landing_cell.alignment = Alignment(horizontal='center')
        
        for col in range(2, 11):
            ws.cell(row=row, column=col).border = thin_border
        
        ws.row_dimensions[row].height = 25
        row += 1
    
    last_data_row = row - 1
    
    # Totals row
    ws.merge_cells(f'B{row}:E{row}')
    ws.cell(row=row, column=2, value="סה\"כ").font = HEADER_FONT
    ws.cell(row=row, column=2).fill = HEADER_FILL
    ws.cell(row=row, column=2).alignment = Alignment(horizontal='center')
    for col in range(2, 6):
        ws.cell(row=row, column=col).fill = HEADER_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    ws.cell(row=row, column=6).fill = HEADER_FILL
    ws.cell(row=row, column=6).border = thin_border
    
    fob_sum = ws.cell(row=row, column=7, value=f"=SUM(G{first_data_row}:G{last_data_row})")
    fob_sum.font = HEADER_FONT
    fob_sum.fill = HEADER_FILL
    fob_sum.number_format = '$#,##0'
    fob_sum.alignment = Alignment(horizontal='center')
    fob_sum.border = thin_border
    
    for col in [8, 9, 10]:
        ws.cell(row=row, column=col).fill = HEADER_FILL
        ws.cell(row=row, column=col).border = thin_border
    
    # Legend
    row += 2
    ws.merge_cells(f'B{row}:J{row}')
    ws[f'B{row}'] = f"📝 מלא את עלות ההובלה בתא הצהוב ← כל החישובים יתעדכנו | סה\"כ יחידות: {total_units:,}"
    ws[f'B{row}'].font = Font(size=11, italic=True, color="666666")
    ws[f'B{row}'].alignment = Alignment(horizontal='center')


def create_excel_report(containers, usd_rate):
    """Create full Excel report with summary + per-container sheets"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "סיכום מכולות"
    
    # Create summary sheet
    create_summary_sheet(ws, containers, usd_rate)
    
    # Create per-container sheets
    for container in containers:
        create_container_sheet(wb, container, usd_rate)
    
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
    print("💵 Fetching USD rate from Monday.com...")
    usd_rate = fetch_usd_rate()
    print(f"USD Rate: {usd_rate}")
    
    print("🚢 Fetching containers from Monday.com...")
    containers = fetch_containers()
    
    if not containers:
        print("No containers found with status 'בדרך' or 'כנמ ללא BL'")
        return
    
    print(f"Found {len(containers)} containers")
    
    print("📊 Creating Excel report...")
    filename = create_excel_report(containers, usd_rate)
    print(f"Created: {filename}")
    
    print("📤 Uploading file...")
    file_uid = upload_file(filename)
    
    if not file_uid:
        print("Failed to upload file")
        return
    
    print("📱 Sending to WhatsApp...")
    today = datetime.now().strftime('%d.%m.%Y')
    critical_count = sum(1 for c in containers if calculate_days_in_port(c['eta']) > 30)
    text = f"🚢 דוח מכולות בכנ\"מ - Gaya Foods\n📅 {today}\n📦 {len(containers)} מכולות"
    if critical_count > 0:
        text += f"\n🔴 {critical_count} קריטי (>30 יום)"
    text += "\n\n🤖 עדכון אוטומטי יומי"
    
    for recipient in RECIPIENTS:
        success = send_whatsapp(recipient['phone'], file_uid, text)
        status = "✅" if success else "❌"
        print(f"{status} {recipient['name']}: {recipient['phone']}")
    
    print("\n✅ Done!")


if __name__ == "__main__":
    main()
