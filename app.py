
import os
import io
import json
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from flask import Flask, request, jsonify
from docx_generator import generate_quote_docx
 
# Google Sheets
import gspread
from google.oauth2.service_account import Credentials
 
app = Flask(__name__)
 
 
@app.after_request
def add_cors_headers(response):
    """Allow the Netlify-hosted quote form to call this API directly."""
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response
 
 
def get_sheets_client():
    """Build gspread client from environment variable JSON credentials."""
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json:
        print("No GOOGLE_CREDENTIALS_JSON set — skipping Sheets logging")
        return None
    creds_dict = json.loads(creds_json)
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)
 
 
def log_to_sheets(quote_data):
    """Append a row to the estimates log Google Sheet."""
    try:
        gc = get_sheets_client()
        if not gc:
            return
 
        sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        if not sheet_id:
            print("No GOOGLE_SHEET_ID set — skipping Sheets logging")
            return
 
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1
 
        # Add header row if sheet is empty
        if ws.row_count == 0 or ws.cell(1, 1).value != 'Reference':
            ws.append_row([
                'Reference', 'Date', 'Name', 'Phone', 'Email', 'Address',
                'Gate Type', 'Width', 'Motor', 'Driveway', 'Estimate (inc GST)',
                'Preferred Date', 'Notes', 'Status'
            ], value_input_option='RAW')
 
        ws.append_row([
            quote_data['ref'],
            quote_data['date'],
            quote_data['name'],
            quote_data['phone'],
            quote_data['email'],
            quote_data['address'],
            quote_data['gate_type'],
            quote_data['gate_width'],
            quote_data['motor'],
            quote_data['driveway'],
            quote_data['total'],
            f"{quote_data['preferred_date']} — {quote_data['preferred_time']}",
            quote_data['notes'],
            'Pending',   # default status — update manually in Sheets
        ], value_input_option='RAW')
 
        print(f"Logged to Sheets: {quote_data['ref']}")
 
    except Exception as e:
        print(f"Sheets logging failed (non-fatal): {e}")
 
@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'Auto Gates Vic webhook running'}), 200
 
@app.route('/webhook', methods=['POST', 'OPTIONS'])
def webhook():
    if request.method == 'OPTIONS':
        return '', 204
 
    try:
        # Web3Forms sends JSON or form data
        if request.is_json:
            data = request.get_json()
        else:
            data = request.form.to_dict()
 
        print(f"Received submission: {json.dumps(data, indent=2)}")
 
        # Build quote data from form fields
        quote_data = build_quote_data(data)
 
        # Generate Word doc in memory
        docx_buffer = generate_quote_docx(quote_data)
 
        # Send email with attachment
        send_email(quote_data, docx_buffer)
 
        # Log to Google Sheets
        log_to_sheets(quote_data)
 
        return jsonify({'success': True, 'ref': quote_data['ref']}), 200
 
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500
 
 
def build_quote_data(data):
    """Map Web3Forms fields to quote template fields."""
    import random
 
    ref = data.get('reference') or f"AGV-{random.randint(10000, 99999)}"
    today = datetime.date.today().strftime('%d %B %Y')
 
    # Parse items/groups from submission
    # Electrical
    elec_m = int(data.get('electrical', '10m').replace('m', '') or 10)
    elec_cost = elec_m * 44
 
    # Driveway surface → civils cost
    driveway = data.get('driveway', 'Other')
    is_hard = driveway in ['Concrete', 'Asphalt']
    civils_labour = 1210 if is_hard else 1980
    concrete_cost = 500 if is_hard else 700
 
    # Gate width
    width_str = data.get('gate_width', '3.0m')
    width = float(width_str.replace('m', '') or 3.0)
 
    # Infill rate per metre
    infill = data.get('infill', 'Custom / Unsure')
    infill_rates = {
        'Vertical Hardwood Battens': 153,
        'Horizontal Timber Decking': 175,
        'Colorbond Steel': 80,
        'Aluminium Battens': 220,
        'Tubular Steel': 90,
        'Custom / Unsure': 140,
    }
    infill_rate = infill_rates.get(infill, 140)
    timber_cost = round(width * infill_rate, 2)
 
    rack_cost  = round(width * 14.95, 2)
    track_cost = round(width * 2 * 9.47, 2)
    frame_cost = 1200
    motor_cost = 2889
    dress_labour   = 880
    install_labour = 990
 
    mats_cost = motor_cost + frame_cost + rack_cost + track_cost + timber_cost + concrete_cost
    markup = round(mats_cost * 0.10, 2)
 
    # Add-ons
    addon_prices = {
        'Safety Beam Sensors': 300,
        'External Antenna': 150,
        'Keypad Entry': 350,
        'Battery Backup': 220,
        'Video Intercom': 480,
        'Smart App Control': 650,
        'Loop Detector': 180,
        'Warning Light / Siren': 120,
    }
    addons_str = data.get('addons', 'Safety Beam Sensors')
    addon_names = [a.strip() for a in addons_str.split(',') if a.strip()]
    addon_items = []
    addons_total = 0
    for name in addon_names:
        price = addon_prices.get(name, 0)
        addons_total += price
        label = f"{name} (included)" if name == 'Safety Beam Sensors' else name
        addon_items.append({'desc': label, 'amt': f'${price:,.2f}'})
 
    labour_total   = elec_cost + civils_labour + dress_labour + install_labour
    materials_total = motor_cost + frame_cost + concrete_cost + timber_cost + rack_cost + track_cost + markup
 
    subtotal = labour_total + materials_total + addons_total
    gst      = round(subtotal * 0.10, 2)
    total    = subtotal + gst
 
    return {
        'ref':            ref,
        'date':           today,
        'valid_days':     30,
        'name':           f"{data.get('name', '')}".strip(),
        'phone':          data.get('phone', ''),
        'email':          data.get('email', ''),
        'address':        data.get('address', ''),
        'gate_type':      data.get('gate_type', ''),
        'gate_width':     data.get('gate_width', ''),
        'gate_height':    data.get('gate_height', ''),
        'infill':         infill,
        'motor':          data.get('motor', ''),
        'driveway':       driveway,
        'slope':          data.get('slope', 'Flat').capitalize(),
        'electrical':     data.get('electrical', '10m'),
        'remotes':        data.get('remotes', '2 remotes (included)'),
        'access':         data.get('access', ''),
        'preferred_date': data.get('preferred_date', '—'),
        'preferred_time': data.get('preferred_time', '—'),
        'notes':          data.get('notes', 'None provided.'),
        'groups': [
            {
                'label': 'Labour',
                'total': f'${labour_total:,.2f}',
                'items': [
                    {'desc': f'Electrical — conduit & trenching ({elec_m}m)', 'amt': f'${elec_cost:,.2f}'},
                    {'desc': f'Civil works — {driveway.lower()} driveway',    'amt': f'${civils_labour:,.2f}'},
                    {'desc': 'Gate dress & preparation',                       'amt': f'${dress_labour:,.2f}'},
                    {'desc': 'Gate install & commissioning',                   'amt': f'${install_labour:,.2f}'},
                ]
            },
            {
                'label': 'Materials & Hardware',
                'total': f'${materials_total:,.2f}',
                'items': [
                    {'desc': data.get('motor', 'DC Motor') + ' — controller & hardware', 'amt': f'${motor_cost:,.2f}'},
                    {'desc': f'Custom gate frame ({width:.1f}m)',                          'amt': f'${frame_cost:,.2f}'},
                    {'desc': 'Concrete supply',                                            'amt': f'${concrete_cost:,.2f}'},
                    {'desc': f'{infill} ({width:.1f}m @ ${infill_rate}/m)',               'amt': f'${timber_cost:,.2f}'},
                    {'desc': f'Rack ({width:.1f}m @ $14.95/m)',                            'amt': f'${rack_cost:,.2f}'},
                    {'desc': f'Track ({width*2:.1f}m @ $9.47/m)',                          'amt': f'${track_cost:,.2f}'},
                    {'desc': 'Materials markup (10%)',                                     'amt': f'${markup:,.2f}'},
                ]
            },
            {
                'label': 'Add-Ons',
                'total': f'${addons_total:,.2f}',
                'items': addon_items or [{'desc': 'Safety Beam Sensors (included)', 'amt': '$300.00'}]
            },
        ],
        'subtotal': f'${subtotal:,.2f}',
        'gst':      f'${gst:,.2f}',
        'total':    f'${total:,.2f}',
    }
 
 
def send_email(quote_data, docx_buffer):
    """Send email with Word doc attachment."""
    smtp_host = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
    smtp_port = int(os.environ.get('SMTP_PORT', 587))
    smtp_user = os.environ.get('SMTP_USER')
    smtp_pass = os.environ.get('SMTP_PASS')
    to_email  = os.environ.get('TO_EMAIL', smtp_user)
 
    msg = MIMEMultipart()
    msg['From']    = smtp_user
    msg['To']      = to_email
    msg['Subject'] = f"New Quote Request — {quote_data['ref']} — {quote_data['name']}"
 
    body = f"""Hi James,
 
A new quote request has come in via the website.
 
Reference:  {quote_data['ref']}
Customer:   {quote_data['name']}
Phone:      {quote_data['phone']}
Email:      {quote_data['email']}
Address:    {quote_data['address']}
 
Gate Type:  {quote_data['gate_type']}
Width:      {quote_data['gate_width']}
Motor:      {quote_data['motor']}
Driveway:   {quote_data['driveway']}
 
Estimate:   {quote_data['total']} (inc GST)
 
Preferred Site Visit:  {quote_data['preferred_date']} — {quote_data['preferred_time']}
 
Notes:  {quote_data['notes']}
 
The populated Word quote is attached. Review, adjust pricing if needed, then save as PDF and send to the customer.
 
— Auto Gates Vic Website
"""
 
    msg.attach(MIMEText(body, 'plain'))
 
    # Attach Word doc
    filename = f"AutoGatesVic_Quote_{quote_data['ref']}.docx"
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(docx_buffer.getvalue())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    msg.attach(part)
 
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
 
    print(f"Email sent: {filename} → {to_email}")
 
 
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
