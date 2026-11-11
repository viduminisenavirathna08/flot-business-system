import os
import re
import hmac
import threading
import calendar
from collections import defaultdict
from datetime import date
from functools import wraps

import pandas as pd
import openpyxl
import pymysql
import pymysql.cursors
from dotenv import load_dotenv
from flask import (Flask, render_template, request, jsonify,
                   session, redirect, url_for)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'change-this-secret-key')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ITEMS_FILE = os.path.join(BASE_DIR, 'items.xlsx')   # project.py thiyena folder eke items.xlsx
ITEM_COST_FILE = os.path.join(BASE_DIR, 'itemcost.xlsx')   # item cost / packing cost thiyena excel eka
cost_lock = threading.Lock()
ORDERS_LOG_FILE = os.path.join(BASE_DIR, 'orders_log.xlsx')   # okkoma orders save wena excel eka
orders_log_lock = threading.Lock()
excel_lock = threading.Lock()   # items.xlsx ekata ekama welawata godak write wenna epa

ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', 'admin123')

# Budget eka dawas gana update karala nathnam alert ekak yanawa
BUDGET_REMINDER_DAYS = int(os.getenv('BUDGET_REMINDER_DAYS', '3'))
# Postpone karapu order ekak, postpone date ekata dawas gana kalin notification enna one da
POSTPONE_NOTICE_DAYS = int(os.getenv('POSTPONE_NOTICE_DAYS', '3'))


# ------------------------------------------------------------------
# Setup helpers
# ------------------------------------------------------------------
def init_excel_file():
    if not os.path.exists(ITEMS_FILE):
        data = {
            'Code': ['BL002', 'BLF01', 'BLP04', 'BLN03', 'SN', 'SGB'],
            'Name': ['Crystal', 'Full Stone', 'Plain', 'Plain with design',
                     'Shiva Necklace', 'Shiva Bangle'],
            'Category': ['Bangles', 'Bangles', 'Bangles', 'Bangles',
                         'Shiva & Couple', 'Shiva & Couple'],
            'Stock': [50, 30, 40, 25, 50, 30],
            'MinStock': [10, 10, 10, 10, 10, 10]
        }
        pd.DataFrame(data).to_excel(ITEMS_FILE, index=False)


init_excel_file()


_order_items_ready = False


def get_db_connection():
    global _order_items_ready
    connection = pymysql.connect(
        host=os.getenv('DB_HOST', 'localhost'),
        user=os.getenv('DB_USER', 'root'),
        password=os.getenv('DB_PASSWORD', ''),
        database=os.getenv('DB_NAME', 'business_management_system_db'),
        cursorclass=pymysql.cursors.DictCursor
    )
    if not _order_items_ready:
        # Order ekakata items godak thiyenna puluwan: order_items table eka automatically hadanawa
        with connection.cursor() as cursor:
            cursor.execute(
                "CREATE TABLE IF NOT EXISTS order_items ("
                "id INT AUTO_INCREMENT PRIMARY KEY, order_id INT NOT NULL, "
                "item_code VARCHAR(50) DEFAULT '', product_name VARCHAR(150) DEFAULT '', "
                "quantity INT DEFAULT 1, price DECIMAL(10,2) DEFAULT 0, INDEX idx_order (order_id))")
            try:
                cursor.execute("ALTER TABLE orders MODIFY tracking_number VARCHAR(40) NULL DEFAULT NULL")
            except Exception:
                pass   # tracking_number column eka nathnam awulak naha
            cursor.execute(
                "CREATE TABLE IF NOT EXISTS barcode_pool ("
                "barcode VARCHAR(40) PRIMARY KEY, type VARCHAR(20) NOT NULL, "
                "order_id INT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, "
                "INDEX idx_pool (type, order_id))")
            # Parcel status (Shipped / Delivered / Returned) saha e wenas wuna welawa
            for column, ddl in (('delivery_status', "VARCHAR(20) NOT NULL DEFAULT 'Shipped'"),
                                ('status_updated_at', "DATETIME NULL")):
                cursor.execute("SHOW COLUMNS FROM orders LIKE %s", (column,))
                if not cursor.fetchone():
                    cursor.execute("ALTER TABLE orders ADD COLUMN " + column + " " + ddl)
        connection.commit()
        _order_items_ready = True
    return connection


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get('logged_in'):
            if request.path.startswith('/api/'):
                return jsonify({'status': 'error', 'message': 'Not logged in'}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return wrapper


# ------------------------------------------------------------------
# Excel stock helpers
# ------------------------------------------------------------------
ITEM_COLUMNS = ['Code', 'Name', 'Category', 'Stock', 'MinStock']

# Excel eke column names wenas widihata thibunath (Item Code, Qty, ...) wada karanna
COLUMN_ALIASES = {
    'code': 'Code', 'itemcode': 'Code', 'sku': 'Code',
    'name': 'Name', 'itemname': 'Name', 'productname': 'Name', 'item': 'Name',
    'category': 'Category',
    'stock': 'Stock', 'qty': 'Stock', 'quantity': 'Stock',
    'remainingstock': 'Stock', 'stockqty': 'Stock',
    'minstock': 'MinStock', 'minstockalert': 'MinStock',
    'minimumstock': 'MinStock', 'minqty': 'MinStock',
}


def read_items_df():
    if not os.path.exists(ITEMS_FILE):
        return pd.DataFrame(columns=ITEM_COLUMNS)

    df = pd.read_excel(ITEMS_FILE)

    rename_map = {}
    for col in df.columns:
        key = str(col).strip().lower().replace(' ', '').replace('_', '')
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename_map)

    for col in ITEM_COLUMNS:
        if col not in df.columns:
            df[col] = 0 if col in ('Stock', 'MinStock') else ''

    return df[ITEM_COLUMNS]


def adjust_stock(item_code, product_name, delta):
    """delta negative = stock adu karanna, positive = stock ayeth danna."""
    with excel_lock:
        if not os.path.exists(ITEMS_FILE):
            return False
        df = read_items_df()
        code = str(item_code or '').strip().lower()
        name = str(product_name or '').strip().lower()

        for idx, row in df.iterrows():
            row_code = str(row.get('Code', '')).strip().lower()
            row_name = str(row.get('Name', '')).strip().lower()
            if (code and row_code == code) or (not code and name and row_name == name):
                current = int(row.get('Stock', 0) or 0)
                df.at[idx, 'Stock'] = max(0, current + delta)
                df.to_excel(ITEMS_FILE, index=False)
                return True
    return False


# ------------------------------------------------------------------
# Auth routes
# ------------------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        if (hmac.compare_digest(username, ADMIN_USERNAME) and
                hmac.compare_digest(password, ADMIN_PASSWORD)):
            session['logged_in'] = True
            return redirect(url_for('index'))
        error = 'Username or password is incorrect.'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ------------------------------------------------------------------
# Page routes
# ------------------------------------------------------------------
@app.route('/')
@login_required
def index():
    return render_template('order.html')


@app.route('/product')
@login_required
def product_page():
    return render_template('product.html')


@app.route('/customers')
@login_required
def customers_page():
    return render_template('customer.html')


# ------------------------------------------------------------------
# Orders API
# ------------------------------------------------------------------
STATUS_FILL_COLORS = {'Confirmed': 'E2F0D9', 'Cancelled': 'F8D7DA', 'Postponed': 'DDEBF7'}
DEFAULT_FILL_COLOR = 'FFF2CC'   # New Lead / Callback / No Answer


def export_orders_excel():
    """DB eke okkoma orders ekama excel file ekakata (orders_log.xlsx) liyanawa. Light colors status walata."""
    try:
        from openpyxl.styles import PatternFill, Font
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT * FROM orders ORDER BY order_date ASC, id ASC")
            orders = cursor.fetchall()
            lines_map = load_order_lines(cursor, orders)
        connection.close()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Orders'
        ws.append(['Order ID', 'Date', 'Customer', 'Phone', 'Address', 'Item Codes', 'Items',
                   'Quantity', 'Amount', 'Delivery', 'Status', 'Postpone Date', 'Tracking No', 'Parcel Status'])
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for o in orders:
            lines = lines_map[o['id']]
            ws.append([
                o['id'], o['order_date'].isoformat() if o.get('order_date') else '',
                o['customer_name'], o['phone_number'], o['address'],
                ', '.join(l['code'] for l in lines if l['code']),
                ' + '.join(('%s x%s' % (l['name'] or l['code'], l['qty'])) for l in lines),
                sum(l['qty'] for l in lines), float(o['amount'] or 0),
                o['delivery_method'], o['status'],
                o['postpone_date'].isoformat() if o.get('postpone_date') else '',
                o.get('tracking_number') or '', o.get('delivery_status') or 'Shipped'])
            fill = PatternFill('solid', start_color=STATUS_FILL_COLORS.get(o['status'], DEFAULT_FILL_COLOR))
            for cell in ws[ws.max_row]:
                cell.fill = fill
        for col, width in zip('ABCDEFGHIJKLMN', (9, 12, 22, 14, 36, 18, 40, 9, 12, 14, 12, 14, 18, 13)):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = 'A2'

        with orders_log_lock:
            wb.save(ORDERS_LOG_FILE)
        return True
    except Exception as e:
        print('orders_log.xlsx update failed (Excel eke open wela nam close karanna):', e)
        return False


def clean_lines(raw_items):
    """Frontend eke items list eka clean karanawa: [{'code','name','qty','price'}]"""
    lines = []
    for it in raw_items or []:
        code = str(it.get('code', '') or '').strip()
        name = str(it.get('name', '') or '').strip()
        if not code and not name:
            continue
        try:
            qty = max(1, int(float(it.get('qty') or 1)))
        except (TypeError, ValueError):
            qty = 1
        price = to_money(it.get('price'), 0.0) or 0.0
        lines.append({'code': code, 'name': name, 'qty': qty, 'price': max(price, 0.0)})
    return lines


def order_aggregate(lines):
    return {'item_code': ', '.join(l['code'] for l in lines if l['code']),
            'product_name': ' + '.join(l['name'] or l['code'] for l in lines),
            'quantity': sum(l['qty'] for l in lines),
            'amount': round(sum(l['price'] for l in lines), 2)}


def lines_from_payload(d):
    lines = clean_lines(d.get('items'))
    if not lines:   # purana format (item_code / item / quantity / price)
        lines = clean_lines([{'code': d.get('item_code'), 'name': d.get('item'),
                              'qty': d.get('quantity'), 'price': d.get('price')}])
    return lines


def write_order_items(cursor, order_id, lines):
    cursor.execute("DELETE FROM order_items WHERE order_id = %s", (order_id,))
    for l in lines:
        cursor.execute(
            "INSERT INTO order_items (order_id, item_code, product_name, quantity, price) "
            "VALUES (%s, %s, %s, %s, %s)", (order_id, l['code'], l['name'], l['qty'], l['price']))


def load_order_lines(cursor, orders):
    """order id -> [{'code','name','qty','price'}] (order_items nathnam order eke ma columns walin)"""
    ids = [o['id'] for o in orders]
    by_order = defaultdict(list)
    if ids:
        cursor.execute("SELECT order_id, item_code, product_name, quantity, price FROM order_items "
                       "ORDER BY id ASC")
        for r in cursor.fetchall():
            by_order[r['order_id']].append({
                'code': r['item_code'] or '', 'name': r['product_name'] or '',
                'qty': int(r['quantity'] or 1), 'price': float(r['price'] or 0)})
    result = {}
    for o in orders:
        lines = by_order.get(o['id'])
        if not lines:
            lines = [{'code': o.get('item_code') or '', 'name': o.get('product_name') or '',
                      'qty': int(o.get('quantity') or 1), 'price': float(o.get('amount') or 0)}]
        result[o['id']] = lines
    return result


def adjust_order_stock(lines, sign):
    """sign = -1 stock adu karanna, +1 ayeth danna"""
    for l in lines:
        adjust_stock(l['code'], l['name'], sign * l['qty'])


def insert_order(cursor, d):
    """Order ekak save karanawa. Duplicate nam False return karanawa."""
    phone = str(d.get('phone', '')).strip()
    lines = lines_from_payload(d)
    agg = order_aggregate(lines)

    # Duplicate check: phone + items (code nathnam name) ekama thiyenawada
    if agg['item_code']:
        cursor.execute(
            "SELECT id FROM orders WHERE phone_number = %s AND item_code = %s LIMIT 1",
            (phone, agg['item_code']))
    else:
        cursor.execute(
            "SELECT id FROM orders WHERE phone_number = %s AND product_name = %s LIMIT 1",
            (phone, agg['product_name']))
    if cursor.fetchone():
        return False

    cursor.execute(
        """INSERT INTO orders
           (customer_name, phone_number, address, item_code, product_name,
            quantity, amount, delivery_method, status, order_date)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
        (d.get('name'), phone, d.get('address') or '', agg['item_code'], agg['product_name'],
         agg['quantity'] or 1, agg['amount'],
         d.get('delivery', 'Standard') or 'Standard', 'New Lead',
         d.get('order_date')))
    write_order_items(cursor, cursor.lastrowid, lines)
    return True


@app.route('/api/get_orders', methods=['GET'])
@login_required
def get_orders():
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT * FROM orders ORDER BY id DESC")
            orders = cursor.fetchall()
            lines_map = load_order_lines(cursor, orders)
        connection.close()
        for o in orders:
            o['items'] = lines_map[o['id']]
            o['order_date'] = o['order_date'].isoformat() if o.get('order_date') else ''
            o['postpone_date'] = o['postpone_date'].isoformat() if o.get('postpone_date') else ''
            o['created_at'] = str(o['created_at']) if o.get('created_at') else ''
            o['status_updated_at'] = str(o['status_updated_at'])[:16] if o.get('status_updated_at') else ''
            o['delivery_status'] = o.get('delivery_status') or 'Shipped'
        return jsonify(orders)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/save_order', methods=['POST'])
@login_required
def save_order():
    data = request.get_json() or {}
    if not data.get('name') or not data.get('phone'):
        return jsonify({'status': 'error', 'message': 'Name and phone are required'})
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            saved = insert_order(cursor, data)
            connection.commit()
        connection.close()
        if saved:
            export_orders_excel()
        return jsonify({'status': 'success' if saved else 'duplicate'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/bulk_save_orders', methods=['POST'])
@login_required
def bulk_save_orders():
    """Excel upload eka: orders okkoma ekama request ekakin."""
    payload = request.get_json() or {}
    orders = payload.get('orders', [])
    added = duplicates = skipped = 0
    duplicate_list = []
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            for d in orders:
                if not d.get('name') or not d.get('phone'):
                    skipped += 1
                    continue
                if insert_order(cursor, d):
                    added += 1
                else:
                    duplicates += 1
                    agg = order_aggregate(lines_from_payload(d))
                    duplicate_list.append({
                        'name': d.get('name'), 'phone': d.get('phone'),
                        'address': d.get('address') or '',
                        'item_code': agg['item_code'], 'product_name': agg['product_name'],
                        'amount': agg['amount'], 'order_date': d.get('order_date')})
            connection.commit()
        connection.close()
        if added:
            export_orders_excel()
        return jsonify({'status': 'success', 'added': added, 'duplicates': duplicates,
                        'skipped': skipped, 'duplicate_list': duplicate_list})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/update_order', methods=['POST'])
@login_required
def update_order():
    """Order eka edit: customer details, delivery method, items (code / qty / price)."""
    d = request.get_json() or {}
    order_id = d.get('id')
    if not order_id:
        return jsonify({'status': 'error', 'message': 'Order id is required'})
    if not str(d.get('name', '')).strip() or not str(d.get('phone', '')).strip():
        return jsonify({'status': 'error', 'message': 'Name and phone are required'})
    lines = lines_from_payload(d)
    if not lines:
        return jsonify({'status': 'error', 'message': 'Add at least one item'})
    agg = order_aggregate(lines)

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT * FROM orders WHERE id = %s", (order_id,))
            order = cursor.fetchone()
            if not order:
                connection.close()
                return jsonify({'status': 'error', 'message': 'Order not found'})
            old_lines = load_order_lines(cursor, [order])[order['id']]

            cursor.execute(
                "UPDATE orders SET customer_name = %s, phone_number = %s, address = %s, "
                "item_code = %s, product_name = %s, quantity = %s, amount = %s, "
                "delivery_method = %s WHERE id = %s",
                (str(d.get('name')).strip(), str(d.get('phone')).strip(), d.get('address') or '',
                 agg['item_code'], agg['product_name'], agg['quantity'], agg['amount'],
                 d.get('delivery') or 'Standard', order_id))
            write_order_items(cursor, order_id, lines)
            sync_order_barcode(cursor, order_id)
            connection.commit()
        connection.close()

        # Confirmed order ekak nam stock eka aluth items walata samana karanawa
        if order['status'] == 'Confirmed':
            adjust_order_stock(old_lines, +1)
            adjust_order_stock(lines, -1)
        export_orders_excel()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


PARCEL_STATUSES = ('Shipped', 'Delivered', 'Returned')


@app.route('/api/parcels/scan', methods=['POST'])
@login_required
def scan_parcels():
    """Barcode (tracking number) list ekak Delivered / Returned kirima. Barcode reader eken ekin eka pass karanna puluwan."""
    d = request.get_json() or {}
    status = d.get('status')
    if status not in PARCEL_STATUSES:
        return jsonify({'status': 'error', 'message': 'Invalid parcel status'})
    codes = []
    for c in d.get('barcodes', []):
        c = str(c or '').strip().upper()
        if c and c not in codes:
            codes.append(c)
    if not codes:
        return jsonify({'status': 'error', 'message': 'No barcodes'})

    try:
        connection = get_db_connection()
        updated, already, found_codes = [], [], set()
        with connection.cursor() as cursor:
            marks = ', '.join(['%s'] * len(codes))
            cursor.execute(
                "SELECT id, customer_name, phone_number, tracking_number, status, delivery_status "
                "FROM orders WHERE UPPER(TRIM(tracking_number)) IN (" + marks + ")", codes)
            rows = cursor.fetchall()
            for r in rows:
                code = str(r['tracking_number']).strip().upper()
                found_codes.add(code)
                info = {'barcode': code, 'customer': r['customer_name'], 'phone': r['phone_number'],
                        'previous': r['delivery_status'] or 'Shipped'}
                if (r['delivery_status'] or 'Shipped') == status:
                    already.append(info)
                else:
                    cursor.execute("UPDATE orders SET delivery_status = %s, status_updated_at = NOW() "
                                   "WHERE id = %s", (status, r['id']))
                    updated.append(info)
            connection.commit()
        connection.close()
        if updated:
            export_orders_excel()
        return jsonify({'status': 'success', 'updated': updated, 'already': already,
                        'not_found': [c for c in codes if c not in found_codes]})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


def is_courier_method(method):
    t = str(method or '').lower()
    return 'courier' in t or 'domex' in t or 'express' in t


def pool_type_for(method):
    return 'Courier' if is_courier_method(method) else 'Post Office'


def assign_pool_barcode(cursor, order_id, kind):
    """Pool eken ilaga available barcode eka order ekata denawa. Barcode eka return karanawa (nathnam None)."""
    cursor.execute("SELECT barcode FROM barcode_pool WHERE type = %s AND order_id IS NULL "
                   "ORDER BY barcode ASC LIMIT 1", (kind,))
    row = cursor.fetchone()
    if not row:
        return None
    cursor.execute("UPDATE barcode_pool SET order_id = %s WHERE barcode = %s AND order_id IS NULL",
                   (order_id, row['barcode']))
    if cursor.rowcount != 1:
        return None
    cursor.execute("UPDATE orders SET tracking_number = %s WHERE id = %s", (row['barcode'], order_id))
    return row['barcode']


def release_pool_barcode(cursor, order_id):
    """Order ekata dunna pool barcode eka ayeth pool ekata (available) damanawa."""
    cursor.execute("SELECT barcode FROM barcode_pool WHERE order_id = %s", (order_id,))
    for r in cursor.fetchall():
        cursor.execute("UPDATE barcode_pool SET order_id = NULL WHERE barcode = %s", (r['barcode'],))
        cursor.execute("UPDATE orders SET tracking_number = NULL WHERE id = %s AND tracking_number = %s",
                       (order_id, r['barcode']))


def sync_order_barcode(cursor, order_id):
    """Confirmed order ekakata delivery type ekata adala pool barcode ekak thiyenna hadanawa.
    Confirmed nathnam / delivery type wenas unoth pool barcode eka ayeth pool ekata yanawa."""
    cursor.execute("SELECT status, delivery_method, tracking_number FROM orders WHERE id = %s", (order_id,))
    o = cursor.fetchone()
    if not o:
        return
    cursor.execute("SELECT barcode, type FROM barcode_pool WHERE order_id = %s", (order_id,))
    held = cursor.fetchone()
    kind = pool_type_for(o['delivery_method'])
    if o['status'] != 'Confirmed':
        if held:
            release_pool_barcode(cursor, order_id)
        return
    if held and held['type'] != kind:
        release_pool_barcode(cursor, order_id)
        held = None
        o['tracking_number'] = None
    if not held and not str(o['tracking_number'] or '').strip():
        assign_pool_barcode(cursor, order_id, kind)


def expand_barcode_token(token):
    """'FL000001-FL000100' nathnam 'FL000001-100' wage range ekak barcodes list ekak karanawa."""
    token = token.strip().upper()
    if '-' not in token:
        return [token]
    left, right = token.split('-', 1)
    ml = re.match(r'^([A-Z]*)(\d+)([A-Z]*)$', left)
    mr = re.match(r'^([A-Z]*)(\d+)([A-Z]*)$', right)
    if not ml or not mr:
        return [token]
    prefix, start, suffix = ml.group(1), ml.group(2), ml.group(3)
    if (mr.group(1) and mr.group(1) != prefix) or (mr.group(3) and mr.group(3) != suffix):
        return [token]
    a, b = int(start), int(mr.group(2))
    if b < a or b - a >= 5000:
        return [token]
    width = len(start)
    return [prefix + str(n).zfill(width) + suffix for n in range(a, b + 1)]


@app.route('/api/barcode_pool/add', methods=['POST'])
@login_required
def barcode_pool_add():
    """Post Office / Courier barcodes pool ekata add karanawa (range ekath: FL000001-FL000100).
    Line eke 'BARCODE 07xxxxxxxx' nam e phone eke parcel ekatama denawa. Ithuru barcodes pool ekata."""
    d = request.get_json() or {}
    kind = d.get('type')
    if kind not in ('Post Office', 'Courier'):
        return jsonify({'status': 'error', 'message': 'Choose Post Office or Courier'})

    plain, by_phone = [], []
    for line in str(d.get('text', '')).replace(',', ' ').replace(';', ' ').splitlines():
        tokens = line.split()
        if not tokens:
            continue
        digits = tokens[1].replace('-', '') if len(tokens) > 1 else ''
        if len(tokens) == 2 and digits.isdigit() and 9 <= len(digits) <= 12:
            by_phone.append((tokens[0].strip().upper(), digits if len(digits) != 9 else '0' + digits))
        else:
            for t in tokens:
                plain.extend(expand_barcode_token(t))
    if not plain and not by_phone:
        return jsonify({'status': 'error', 'message': 'No barcodes'})

    try:
        connection = get_db_connection()
        added = duplicates = assigned = 0
        notes = []
        with connection.cursor() as cursor:
            cursor.execute("SELECT UPPER(TRIM(tracking_number)) AS t FROM orders "
                           "WHERE tracking_number IS NOT NULL AND TRIM(tracking_number) <> ''")
            used = {r['t'] for r in cursor.fetchall()}
            cursor.execute("SELECT barcode FROM barcode_pool")
            in_pool = {r['barcode'] for r in cursor.fetchall()}

            for code in plain:
                if len(code) > 40:
                    notes.append(code + ' - too long')
                elif code in used or code in in_pool:
                    duplicates += 1
                else:
                    cursor.execute("INSERT INTO barcode_pool (barcode, type) VALUES (%s, %s)", (code, kind))
                    in_pool.add(code)
                    added += 1

            cursor.execute("SELECT id, phone_number, delivery_method FROM orders "
                           "WHERE status = 'Confirmed' AND (tracking_number IS NULL OR TRIM(tracking_number) = '') "
                           "ORDER BY order_date ASC, id ASC")
            waiting = [r for r in cursor.fetchall() if pool_type_for(r['delivery_method']) == kind]

            for code, phone in by_phone:
                if code in used or code in in_pool:
                    duplicates += 1
                    continue
                target = next((r for r in waiting
                               if str(r['phone_number']).replace('-', '').lstrip('0') == phone.lstrip('0')), None)
                if not target:
                    notes.append(code + ' - no waiting parcel for ' + phone)
                    continue
                waiting.remove(target)
                cursor.execute("INSERT INTO barcode_pool (barcode, type, order_id) VALUES (%s, %s, %s)",
                               (code, kind, target['id']))
                cursor.execute("UPDATE orders SET tracking_number = %s WHERE id = %s", (code, target['id']))
                in_pool.add(code)
                added += 1
                assigned += 1

            # Barcode nathi Confirmed parcels walata pool eken auto assign
            for r in waiting:
                if assign_pool_barcode(cursor, r['id'], kind):
                    assigned += 1
            connection.commit()
        connection.close()
        if assigned:
            export_orders_excel()
        return jsonify({'status': 'success', 'added': added, 'duplicates': duplicates,
                        'assigned': assigned, 'notes': notes})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/barcode_pool/summary', methods=['GET'])
@login_required
def barcode_pool_summary():
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT type, COUNT(*) AS n FROM barcode_pool WHERE order_id IS NULL GROUP BY type")
            available = {r['type']: int(r['n']) for r in cursor.fetchall()}
            cursor.execute("SELECT delivery_method FROM orders WHERE status = 'Confirmed' "
                           "AND (tracking_number IS NULL OR TRIM(tracking_number) = '')")
            needed = {'Post Office': 0, 'Courier': 0}
            for r in cursor.fetchall():
                needed[pool_type_for(r['delivery_method'])] += 1
            cursor.execute("SELECT COUNT(*) AS n FROM orders WHERE status = 'Postponed'")
            hold = int(cursor.fetchone()['n'])
        connection.close()
        return jsonify({'post_office': {'available': available.get('Post Office', 0), 'needed': needed['Post Office']},
                        'courier': {'available': available.get('Courier', 0), 'needed': needed['Courier']},
                        'hold': hold})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/parcels/status', methods=['POST'])
@login_required
def set_parcel_status():
    d = request.get_json() or {}
    order_id = d.get('id')
    status = d.get('delivery_status')
    if not order_id or status not in PARCEL_STATUSES:
        return jsonify({'status': 'error', 'message': 'Order id and valid status are required'})
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("UPDATE orders SET delivery_status = %s, status_updated_at = NOW() "
                           "WHERE id = %s", (status, order_id))
            connection.commit()
        connection.close()
        export_orders_excel()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/update_delivery', methods=['POST'])
@login_required
def update_delivery():
    d = request.get_json() or {}
    order_id = d.get('id')
    method = str(d.get('delivery', '') or '').strip()
    if not order_id or not method or len(method) > 50:
        return jsonify({'status': 'error', 'message': 'Order id and delivery method are required'})
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("UPDATE orders SET delivery_method = %s WHERE id = %s", (method, order_id))
            sync_order_barcode(cursor, order_id)
            connection.commit()
        connection.close()
        export_orders_excel()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/save_tracking', methods=['POST'])
@login_required
def save_tracking():
    d = request.get_json() or {}
    order_id = d.get('id')
    tracking = str(d.get('tracking_number', '') or '').strip()
    if not order_id:
        return jsonify({'status': 'error', 'message': 'Order id is required'})
    if len(tracking) > 40:
        return jsonify({'status': 'error', 'message': 'Tracking number is too long (max 40 characters)'})
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            if tracking:
                cursor.execute("SELECT id, customer_name FROM orders WHERE UPPER(TRIM(tracking_number)) = %s "
                               "AND id <> %s LIMIT 1", (tracking.upper(), order_id))
                other = cursor.fetchone()
                if other:
                    connection.close()
                    return jsonify({'status': 'error',
                                    'message': 'This barcode is already used for ' + str(other['customer_name'])})
            cursor.execute("UPDATE orders SET tracking_number = %s WHERE id = %s",
                           (tracking or None, order_id))
            # Pool barcode ekak wenas karoth pala eka ayeth pool ekata, aluth eka pool eke nam e order ekata
            cursor.execute("UPDATE barcode_pool SET order_id = NULL WHERE order_id = %s AND barcode <> %s",
                           (order_id, tracking.upper()))
            if tracking:
                cursor.execute("UPDATE barcode_pool SET order_id = %s WHERE barcode = %s AND order_id IS NULL",
                               (order_id, tracking.upper()))
            connection.commit()
        connection.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


@app.route('/api/update_order_status', methods=['POST'])
@login_required
def update_order_status():
    data = request.get_json() or {}
    order_id = data.get('id')
    phone = data.get('phone')       # purana frontend ekata support
    new_status = data.get('status')

    postpone_date = data.get('postpone_date') or None

    if new_status == 'Postponed':
        if not postpone_date:
            return jsonify({'status': 'error', 'message': 'Please select a postpone date'})
        try:
            date.fromisoformat(postpone_date)
        except ValueError:
            return jsonify({'status': 'error', 'message': 'Invalid postpone date'})
    else:
        postpone_date = None   # wenath status ekakata gaththama postpone date eka ain wenawa

    if new_status not in ('New Lead', 'Confirmed', 'Cancelled', 'Callback', 'No Answer', 'Postponed'):
        return jsonify({'status': 'error', 'message': 'Invalid status'})

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            if order_id:
                cursor.execute("SELECT * FROM orders WHERE id = %s", (order_id,))
            else:
                cursor.execute("SELECT * FROM orders WHERE phone_number = %s "
                               "ORDER BY id DESC LIMIT 1", (phone,))
            order = cursor.fetchone()
            if not order:
                connection.close()
                return jsonify({'status': 'error', 'message': 'Order not found'})

            old_status = order['status']
            cursor.execute("UPDATE orders SET status = %s, postpone_date = %s WHERE id = %s",
                           (new_status, postpone_date, order['id']))
            sync_order_barcode(cursor, order['id'])
            connection.commit()

            order_lines = load_order_lines(cursor, [order])[order['id']]
            # Confirmed wenakota stock adu karanawa (eka parak witharai)
            if new_status == 'Confirmed' and old_status != 'Confirmed':
                adjust_order_stock(order_lines, -1)
            # Confirmed idan wenath ekakata gaththama stock ayeth danawa
            elif old_status == 'Confirmed' and new_status != 'Confirmed':
                adjust_order_stock(order_lines, +1)
        connection.close()
        export_orders_excel()
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ------------------------------------------------------------------
# Items API (items.xlsx)
# ------------------------------------------------------------------
@app.route('/api/item_catalog', methods=['GET'])
@login_required
def item_catalog():
    """Order form / Excel import walata: item code, name, category, quantity, selling price."""
    try:
        costs = load_item_costs()
        catalog, seen = [], set()
        for key in sorted(costs.keys(), key=lambda k: (k[0], k[1])):
            info = costs[key]
            seen.add(key[0])
            catalog.append({'code': info['code'], 'name': info['name'], 'category': info['category'],
                            'qty': info['qty'], 'price': info['price']})
        with excel_lock:
            df = read_items_df().fillna('')
        for _, row in df.iterrows():
            code = str(row['Code']).strip()
            if code and code.lower() not in seen:
                seen.add(code.lower())
                catalog.append({'code': code, 'name': str(row['Name']),
                                'category': str(row['Category']), 'qty': 1, 'price': 0.0})
        return jsonify(catalog)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/get_items', methods=['GET'])
@login_required
def get_items():
    try:
        with excel_lock:
            df = read_items_df()
        df = df.fillna('')
        return jsonify(df.to_dict(orient='records'))
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/api/save_item', methods=['POST'])
@login_required
def save_item():
    data = request.get_json() or {}
    code = str(data.get('code', '')).strip()
    if not code:
        return jsonify({'status': 'error', 'message': 'Item code is required'})

    try:
        with excel_lock:
            df = read_items_df()
            mask = df['Code'].astype(str).str.strip() == code

            if mask.any():
                idx = df[mask].index[0]
                df.at[idx, 'Name'] = data.get('name')
                df.at[idx, 'Category'] = data.get('category')
                df.at[idx, 'Stock'] = data.get('stock')
                df.at[idx, 'MinStock'] = data.get('min_stock')
            else:
                new_row = pd.DataFrame([{
                    'Code': code, 'Name': data.get('name'),
                    'Category': data.get('category'),
                    'Stock': data.get('stock'),
                    'MinStock': data.get('min_stock')
                }])
                df = pd.concat([df, new_row], ignore_index=True)

            df.to_excel(ITEMS_FILE, index=False)
        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})


# ------------------------------------------------------------------
# Budget API (parcel budget, daily ad cost, monthly expenses, settings)
# ------------------------------------------------------------------
def r2(value):
    return round(float(value or 0), 2)


def to_money(value, default=0.0):
    # Number ekak float ekak karanawa. Hisa / waradi nam default return karanawa.
    try:
        if value is None or str(value).strip() == '':
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def parse_date(value):
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        return None


def parse_month(value):
    # '2026-10' -> date(2026, 10, 1)
    try:
        year, month = str(value).split('-')[:2]
        return date(int(year), int(month), 1)
    except (TypeError, ValueError):
        return None


def month_bounds(first_day):
    last = calendar.monthrange(first_day.year, first_day.month)[1]
    return first_day, date(first_day.year, first_day.month, last)


def api_error(e):
    return jsonify({'status': 'error', 'message': str(e)}), 500


def pick_cost(override_value, default_value):
    return float(override_value) if override_value is not None else default_value


def load_budget_defaults(cursor):
    cursor.execute("SELECT setting_key, setting_value FROM budget_settings")
    values = {row['setting_key']: float(row['setting_value'] or 0) for row in cursor.fetchall()}
    return {'packing': values.get('default_packing_cost', 0.0),
            'delivery': values.get('default_delivery_cost', 0.0)}


def _norm_header(value):
    return str(value or '').strip().lower().replace(' ', '').replace('_', '')


def _cost_columns(ws):
    """Header row eken Code / Name / Category / Quantity / Item cost / Packing cost columns (1-based)."""
    cols = {}
    names = {'code': ('code', 'itemcode', 'sku'), 'name': ('name', 'itemname'),
             'category': ('category',), 'qty': ('quantity', 'qty'),
             'cost': ('itemcost', 'cost', 'costprice'), 'packing': ('packingcost', 'packing'),
             'price': ('sellingprice', 'price', 'saleprice')}
    for idx, cell in enumerate(ws[1], start=1):
        key = _norm_header(cell.value)
        for field, aliases in names.items():
            if key in aliases:
                cols.setdefault(field, idx)
    return cols


def _num_or_none(value):
    if value is None or str(value).strip() == '':
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _qty_of(value):
    n = _num_or_none(value)
    return int(n) if n and n >= 1 else 1


def load_item_costs(cursor=None):
    # itemcost.xlsx eken: (item code lowercase, quantity) -> {'code','name','category','qty','cost','packing'}
    # Excel eke code ekakata quantity 1, 2 ... wenama rows thiyenawa; cost = e quantity ekata full cost.
    result = {}
    if not os.path.exists(ITEM_COST_FILE):
        return result
    with cost_lock:
        wb = openpyxl.load_workbook(ITEM_COST_FILE, data_only=True)
        ws = wb.active
        cols = _cost_columns(ws)
        if 'code' in cols:
            def cell(row, field):
                return row[cols[field] - 1] if field in cols else None
            for row in ws.iter_rows(min_row=2, values_only=True):
                code = str(cell(row, 'code') or '').strip()
                if not code:
                    continue
                qty = _qty_of(cell(row, 'qty'))
                result[(code.lower(), qty)] = {
                    'code': code, 'qty': qty,
                    'name': str(cell(row, 'name') or '').strip(),
                    'category': str(cell(row, 'category') or '').strip(),
                    'cost': _num_or_none(cell(row, 'cost')) or 0.0,
                    'packing': _num_or_none(cell(row, 'packing')),
                    'price': _num_or_none(cell(row, 'price')) or 0.0}
        wb.close()
    return result


def lookup_item_cost(item_costs, code, qty):
    """Exact quantity row thiyenawanam eka; nathnam quantity 1 row eka x qty. Return (info, scaled)."""
    code = str(code or '').strip().lower()
    exact = item_costs.get((code, qty))
    if exact:
        return exact, False
    base = item_costs.get((code, 1))
    if base:
        return base, True
    return {}, False


def save_item_costs(entries):
    """entries: [{'code','qty','name','category','cost','packing'}] -> itemcost.xlsx (Code + Quantity eken update / aluth row)."""
    with cost_lock:
        if os.path.exists(ITEM_COST_FILE):
            wb = openpyxl.load_workbook(ITEM_COST_FILE)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(['Code', 'Name', 'Category', 'Quantity', 'Item cost', 'Packing cost'])
        cols = _cost_columns(ws)
        for field, header in (('code', 'Code'), ('qty', 'Quantity'),
                              ('cost', 'Item cost'), ('packing', 'Packing cost')):
            if field not in cols:
                cols[field] = ws.max_column + 1
                ws.cell(row=1, column=cols[field], value=header)

        rows = {}
        for r in range(2, ws.max_row + 1):
            code = str(ws.cell(row=r, column=cols['code']).value or '').strip().lower()
            if code:
                rows[(code, _qty_of(ws.cell(row=r, column=cols['qty']).value))] = r

        for e in entries:
            key = (e['code'].lower(), e['qty'])
            r = rows.get(key)
            if r is None:
                r = ws.max_row + 1
                rows[key] = r
                ws.cell(row=r, column=cols['code'], value=e['code'])
                ws.cell(row=r, column=cols['qty'], value=e['qty'])
                if 'name' in cols:
                    ws.cell(row=r, column=cols['name'], value=e.get('name', ''))
                if 'category' in cols:
                    ws.cell(row=r, column=cols['category'], value=e.get('category', ''))
            ws.cell(row=r, column=cols['cost']).value = e['cost']
            ws.cell(row=r, column=cols['packing']).value = e['packing']
        wb.save(ITEM_COST_FILE)


def load_delivery_costs(cursor):
    # delivery method (lowercase) -> cost
    cursor.execute("SELECT method, cost FROM delivery_costs")
    return {str(row['method']).strip().lower(): float(row['cost'] or 0)
            for row in cursor.fetchall()}


def month_expense_totals(cursor, first_day):
    # Me masayata apply wena monthly Expenses saha Savings total.
    cursor.execute(
        "SELECT category, SUM(amount) AS total FROM monthly_expenses "
        "WHERE start_month <= %s GROUP BY category", (first_day,))
    totals = {'Expense': 0.0, 'Savings': 0.0}
    for row in cursor.fetchall():
        totals[row['category']] = float(row['total'] or 0)
    return totals['Expense'], totals['Savings']


def compute_parcels(cursor, start, end):
    # start..end athara Confirmed orders (parcels), cost saha profit ekka.
    # Dawasaka ad cost eka, e dawase parcels gana walin bedanawa.
    defaults = load_budget_defaults(cursor)
    item_costs = load_item_costs(cursor)
    delivery_costs = load_delivery_costs(cursor)

    cursor.execute("SELECT order_id, product_cost, packing_cost, delivery_cost, extra_cost "
                   "FROM parcel_costs")
    overrides = {row['order_id']: row for row in cursor.fetchall()}

    cursor.execute(
        "SELECT id, customer_name, phone_number, product_name, item_code, quantity, "
        "amount, delivery_method, order_date, delivery_status FROM orders "
        "WHERE status = 'Confirmed' AND order_date BETWEEN %s AND %s "
        "ORDER BY order_date ASC, id ASC", (start, end))
    orders = cursor.fetchall()

    cursor.execute(
        "SELECT ad_date, amount FROM daily_ad_costs WHERE ad_date BETWEEN %s AND %s",
        (start, end))
    ad_costs = {row['ad_date']: float(row['amount'] or 0) for row in cursor.fetchall()}

    # Item category (items.xlsx eken), item code eken
    categories = {}
    try:
        with excel_lock:
            items_df = read_items_df()
        for _, row in items_df.iterrows():
            categories[str(row['Code']).strip().lower()] = str(row['Category'] or '').strip()
    except Exception:
        categories = {}

    lines_map = load_order_lines(cursor, orders)

    parcels_per_day = defaultdict(int)
    for o in orders:
        parcels_per_day[o['order_date']] += 1

    parcels = []
    for o in orders:
        day = o['order_date']
        order_lines = lines_map[o['id']]
        qty = sum(l['qty'] for l in order_lines) or 1
        returned = (o.get('delivery_status') == 'Returned')
        # Return una parcel ekata income naha (sale = 0), item eka ayeth enawa nisa item cost naha. Packing, delivery, ad saha extra cost loss wenawa.
        sale = 0.0 if returned else float(o['amount'] or 0)
        ov = overrides.get(o['id']) or {}

        # Order eke item line ekin ekata cost (Code + Quantity), packing = line walin wadi eka
        product_default = 0.0
        packings = []
        line_categories = []
        for l in order_lines:
            info, scaled = lookup_item_cost(item_costs, l['code'], l['qty'])
            if not returned:
                product_default += info.get('cost', 0.0) * (l['qty'] if scaled else 1)
            if info.get('packing') is not None:
                packings.append(info['packing'])
            cat = categories.get(str(l['code']).strip().lower()) or info.get('category', '')
            if cat and cat not in line_categories:
                line_categories.append(cat)
        item_packing = max(packings) if packings else None
        method = str(o['delivery_method'] or '').strip().lower()

        # Cost priority: parcel ekata wenama set karapu eka > item / method cost > default
        product = pick_cost(ov.get('product_cost'), product_default)
        packing = pick_cost(ov.get('packing_cost'),
                            item_packing if item_packing is not None else defaults['packing'])
        delivery = pick_cost(ov.get('delivery_cost'),
                             delivery_costs.get(method, defaults['delivery']))
        extra = pick_cost(ov.get('extra_cost'), 0.0)
        ad_share = ad_costs.get(day, 0.0) / parcels_per_day[day]
        total_cost = product + packing + delivery + extra + ad_share

        parcels.append({
            'id': o['id'], 'date': day.isoformat(),
            'customer_name': o['customer_name'], 'phone_number': o['phone_number'],
            'product_name': o['product_name'], 'item_code': o['item_code'],
            'lines': order_lines,
            'category': ', '.join(line_categories),
            'delivery_method': o['delivery_method'] or '',
            'quantity': qty, 'sale': r2(sale), 'returned': returned,
            'delivery_status': o.get('delivery_status') or 'Shipped',
            'product_cost': r2(product), 'packing_cost': r2(packing),
            'delivery_cost': r2(delivery), 'extra_cost': r2(extra),
            'ad_share': r2(ad_share),
            'total_cost': r2(total_cost), 'profit': r2(sale - total_cost),
            'custom': {
                'product': ov.get('product_cost') is not None,
                'packing': ov.get('packing_cost') is not None,
                'delivery': ov.get('delivery_cost') is not None,
                'extra': ov.get('extra_cost') is not None
            }
        })
    return parcels, ad_costs


def summarize(parcels, ad_cost):
    sales = sum(p['sale'] for p in parcels)
    product = sum(p['product_cost'] for p in parcels)
    packing = sum(p['packing_cost'] for p in parcels)
    delivery = sum(p['delivery_cost'] for p in parcels)
    extra = sum(p['extra_cost'] for p in parcels)
    gross = sales - product - packing - delivery - extra - ad_cost
    returned = [p for p in parcels if p.get('returned')]
    return {
        'parcels': len(parcels), 'sales': r2(sales),
        'returned': len(returned),
        'returned_loss': r2(sum(p['product_cost'] + p['packing_cost'] + p['delivery_cost'] + p['extra_cost'] for p in returned)),
        'product_cost': r2(product), 'packing_cost': r2(packing),
        'delivery_cost': r2(delivery), 'extra_cost': r2(extra),
        'parcel_costs': r2(product + packing + delivery + extra),
        'ad_cost': r2(ad_cost), 'gross_profit': r2(gross)
    }


@app.route('/api/budget/day', methods=['GET'])
@login_required
def budget_day():
    day = parse_date(request.args.get('date')) or date.today()
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            parcels, ad_costs = compute_parcels(cursor, day, day)
            first_day, last_day = month_bounds(day.replace(day=1))
            exp_total, sav_total = month_expense_totals(cursor, first_day)
        connection.close()

        days_in_month = last_day.day
        expense_share = exp_total / days_in_month
        savings_share = sav_total / days_in_month

        summary = summarize(parcels, ad_costs.get(day, 0.0))
        summary['expense_share'] = r2(expense_share)
        summary['savings_share'] = r2(savings_share)
        summary['net_profit'] = r2(summary['gross_profit'] - expense_share - savings_share)
        return jsonify({'date': day.isoformat(), 'parcels': parcels, 'summary': summary})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/month', methods=['GET'])
@login_required
def budget_month():
    first_day = parse_month(request.args.get('month')) or date.today().replace(day=1)
    first_day, last_day = month_bounds(first_day)
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            parcels, ad_costs = compute_parcels(cursor, first_day, last_day)
            exp_total, sav_total = month_expense_totals(cursor, first_day)
        connection.close()

        days_in_month = last_day.day
        fixed_share = (exp_total + sav_total) / days_in_month

        by_day = defaultdict(list)
        for p in parcels:
            by_day[p['date']].append(p)
        active_days = sorted(set(by_day.keys()) | {d.isoformat() for d in ad_costs.keys()})

        rows = []
        for d in active_days:
            s = summarize(by_day.get(d, []), ad_costs.get(date.fromisoformat(d), 0.0))
            s['date'] = d
            s['fixed_share'] = r2(fixed_share)
            s['net_profit'] = r2(s['gross_profit'] - fixed_share)
            rows.append(s)

        total = summarize(parcels, sum(ad_costs.values()))
        total['expenses'] = r2(exp_total)
        total['savings'] = r2(sav_total)
        total['net_profit'] = r2(total['gross_profit'] - exp_total - sav_total)
        return jsonify({'month': first_day.strftime('%Y-%m'), 'days': rows, 'summary': total})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/parcel_cost', methods=['POST'])
@login_required
def save_parcel_cost():
    data = request.get_json() or {}
    order_id = data.get('order_id')
    if not order_id:
        return jsonify({'status': 'error', 'message': 'Order id is required'})

    # Hisa / null nam e cost eka default (item / method) eken ganna
    keys = ('product_cost', 'packing_cost', 'delivery_cost', 'extra_cost')
    values = [to_money(data.get(k), None) for k in keys]
    if any(v is not None and v < 0 for v in values):
        return jsonify({'status': 'error', 'message': 'Costs cannot be negative'})

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute(
                "INSERT INTO parcel_costs "
                "(order_id, product_cost, packing_cost, delivery_cost, extra_cost) "
                "VALUES (%s, %s, %s, %s, %s) "
                "ON DUPLICATE KEY UPDATE product_cost = VALUES(product_cost), "
                "packing_cost = VALUES(packing_cost), delivery_cost = VALUES(delivery_cost), "
                "extra_cost = VALUES(extra_cost)",
                (order_id, values[0], values[1], values[2], values[3]))
            connection.commit()
        connection.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/ad_costs', methods=['GET'])
@login_required
def get_ad_costs():
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT ad_date, amount FROM daily_ad_costs "
                           "ORDER BY ad_date DESC LIMIT 60")
            rows = cursor.fetchall()
        connection.close()
        return jsonify([{'ad_date': r['ad_date'].isoformat(), 'amount': float(r['amount'] or 0)}
                        for r in rows])
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/ad_cost', methods=['POST'])
@login_required
def save_ad_cost():
    data = request.get_json() or {}
    ad_date = parse_date(data.get('ad_date'))
    amount = to_money(data.get('amount'), None)
    if ad_date is None:
        return jsonify({'status': 'error', 'message': 'Valid date is required'})
    if amount is None or amount < 0:
        return jsonify({'status': 'error', 'message': 'Amount must be 0 or more'})

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute(
                "INSERT INTO daily_ad_costs (ad_date, amount) VALUES (%s, %s) "
                "ON DUPLICATE KEY UPDATE amount = VALUES(amount), updated_at = NOW()",
                (ad_date, amount))
            connection.commit()
        connection.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/monthly_expenses', methods=['GET'])
@login_required
def get_monthly_expenses():
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT id, name, category, amount, start_month "
                           "FROM monthly_expenses ORDER BY category ASC, id ASC")
            rows = cursor.fetchall()
            this_month, last_day = month_bounds(date.today().replace(day=1))
            exp_total, sav_total = month_expense_totals(cursor, this_month)
        connection.close()

        items = [{'id': r['id'], 'name': r['name'], 'category': r['category'],
                  'amount': float(r['amount'] or 0),
                  'start_month': r['start_month'].strftime('%Y-%m') if r['start_month'] else ''}
                 for r in rows]
        return jsonify({'items': items,
                        'totals': {'expenses': r2(exp_total), 'savings': r2(sav_total),
                                   'days_in_month': last_day.day}})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/monthly_expenses', methods=['POST'])
@login_required
def add_monthly_expense():
    data = request.get_json() or {}
    name = str(data.get('name', '')).strip()
    category = data.get('category')
    amount = to_money(data.get('amount'), None)
    start_month = parse_month(data.get('start_month')) or date.today().replace(day=1)

    if not name or len(name) > 100:
        return jsonify({'status': 'error', 'message': 'Name is required (max 100 characters)'})
    if category not in ('Expense', 'Savings'):
        return jsonify({'status': 'error', 'message': 'Invalid category'})
    if amount is None or amount <= 0:
        return jsonify({'status': 'error', 'message': 'Amount must be greater than 0'})

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute(
                "INSERT INTO monthly_expenses (name, category, amount, start_month) "
                "VALUES (%s, %s, %s, %s)", (name, category, amount, start_month))
            connection.commit()
        connection.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/monthly_expenses/delete', methods=['POST'])
@login_required
def delete_monthly_expense():
    data = request.get_json() or {}
    expense_id = data.get('id')
    if not expense_id:
        return jsonify({'status': 'error', 'message': 'Id is required'})
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("DELETE FROM monthly_expenses WHERE id = %s", (expense_id,))
            connection.commit()
        connection.close()
        return jsonify({'status': 'success'})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/settings', methods=['GET'])
@login_required
def get_budget_settings():
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            defaults = load_budget_defaults(cursor)
            item_costs = load_item_costs(cursor)
            delivery_costs = load_delivery_costs(cursor)
            cursor.execute("SELECT method FROM delivery_costs")
            saved_methods = [str(r['method']).strip() for r in cursor.fetchall()]
            cursor.execute("SELECT DISTINCT delivery_method FROM orders "
                           "WHERE delivery_method IS NOT NULL AND delivery_method <> ''")
            used_methods = [str(r['delivery_method']).strip() for r in cursor.fetchall()]
        connection.close()

        # Items: itemcost.xlsx rows (code + quantity) + items.xlsx eke cost nathi codes (quantity 1)
        items = []
        seen_codes = set()
        for key in sorted(item_costs.keys(), key=lambda k: (item_costs[k]['code'].lower(), k[1])):
            info = item_costs[key]
            seen_codes.add(key[0])
            items.append({'code': info['code'], 'qty': info['qty'], 'name': info['name'],
                          'category': info['category'],
                          'cost_price': info['cost'], 'packing_cost': info['packing']})
        with excel_lock:
            df = read_items_df().fillna('')
        for _, row in df.iterrows():
            code = str(row['Code']).strip()
            if not code or code.lower() in seen_codes:
                continue
            seen_codes.add(code.lower())
            items.append({'code': code, 'qty': 1, 'name': str(row['Name']),
                          'category': str(row['Category']),
                          'cost_price': 0.0, 'packing_cost': None})

        # Delivery methods: saved + orders walin + standard 3 (case-insensitive dupes ain karanawa)
        methods, seen = [], set()
        for name in saved_methods + used_methods + ['Post Office', 'Speed Post', 'Courier']:
            key = name.strip().lower()
            if key and key not in seen:
                seen.add(key)
                methods.append(name.strip())
        method_rows = [{'method': m, 'cost': delivery_costs.get(m.lower())} for m in methods]

        return jsonify({'default_packing_cost': defaults['packing'],
                        'default_delivery_cost': defaults['delivery'],
                        'items': items, 'delivery_methods': method_rows})
    except Exception as e:
        return api_error(e)


@app.route('/api/budget/settings', methods=['POST'])
@login_required
def save_budget_settings():
    data = request.get_json() or {}
    packing = to_money(data.get('default_packing_cost'), 0.0)
    delivery = to_money(data.get('default_delivery_cost'), 0.0)
    if packing < 0 or delivery < 0:
        return jsonify({'status': 'error', 'message': 'Costs cannot be negative'})

    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            for key, value in (('default_packing_cost', packing),
                               ('default_delivery_cost', delivery)):
                cursor.execute(
                    "INSERT INTO budget_settings (setting_key, setting_value) VALUES (%s, %s) "
                    "ON DUPLICATE KEY UPDATE setting_value = VALUES(setting_value)",
                    (key, value))

            # Delivery method cost (hisa nam default delivery cost eka use wenawa)
            for entry in data.get('delivery_methods', []):
                method = str(entry.get('method', '')).strip()
                method_cost = to_money(entry.get('cost'), None)
                if not method or len(method) > 50:
                    continue
                if method_cost is None:
                    cursor.execute("DELETE FROM delivery_costs WHERE method = %s", (method,))
                elif method_cost >= 0:
                    cursor.execute(
                        "INSERT INTO delivery_costs (method, cost) VALUES (%s, %s) "
                        "ON DUPLICATE KEY UPDATE cost = VALUES(cost)", (method, method_cost))
            connection.commit()
        connection.close()

        # Item cost (per unit) saha packing cost (per parcel; hisa nam default) -> itemcost.xlsx
        entries = []
        for item in data.get('item_costs', []):
            code = str(item.get('code', '')).strip()
            cost = to_money(item.get('cost_price'), 0.0)
            item_packing = to_money(item.get('packing_cost'), None)
            if not code or cost < 0 or (item_packing is not None and item_packing < 0):
                continue
            entries.append({'code': code, 'qty': _qty_of(item.get('qty')),
                            'name': str(item.get('name', '')).strip(),
                            'category': str(item.get('category', '')).strip(),
                            'cost': cost, 'packing': item_packing})
        if entries:
            try:
                save_item_costs(entries)
            except PermissionError:
                return jsonify({'status': 'error', 'message': 'itemcost.xlsx is open in Excel. Close it and save again.'})
        return jsonify({'status': 'success'})
    except Exception as e:
        return api_error(e)


# ------------------------------------------------------------------
# Notifications API
# ------------------------------------------------------------------
@app.route('/api/notifications', methods=['GET'])
@login_required
def get_notifications():
    notifications = []

    # 1. Low stock (stock eka min stock ekata samanai nathnam adu)
    try:
        with excel_lock:
            df = read_items_df()
        df['Stock'] = pd.to_numeric(df['Stock'], errors='coerce').fillna(0).astype(int)
        df['MinStock'] = pd.to_numeric(df['MinStock'], errors='coerce').fillna(10).astype(int)
        low = df[df['Stock'] <= df['MinStock']]
        for _, r in low.iterrows():
            out_of_stock = r['Stock'] == 0
            notifications.append({
                'type': 'low_stock',
                'level': 'danger' if out_of_stock else 'warning',
                'title': f"{r['Code']} - {r['Name']}",
                'message': 'Out of stock!' if out_of_stock
                           else f"Only {r['Stock']} left (min stock: {r['MinStock']})"
            })
    except Exception as e:
        notifications.append({'type': 'low_stock', 'level': 'warning',
                              'title': 'Could not check stock', 'message': str(e)})

    # 2. Daily ad cost / budget update reminder
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute("SELECT MAX(updated_at) AS last_update FROM daily_ad_costs")
            last_update = cursor.fetchone()['last_update']
        connection.close()

        if last_update is None:
            notifications.append({
                'type': 'budget', 'level': 'warning',
                'title': 'Budget not updated yet',
                'message': 'No daily ad cost entered yet. Please update your budget.'})
        else:
            days = (date.today() - last_update.date()).days
            if days >= BUDGET_REMINDER_DAYS:
                notifications.append({
                    'type': 'budget', 'level': 'danger',
                    'title': 'Budget needs an update',
                    'message': f'Daily ad cost was last updated {days} day(s) ago. Please update it.'})
    except Exception as e:
        notifications.append({
            'type': 'budget', 'level': 'warning',
            'title': 'Could not check budget',
            'message': 'Make sure the budget tables exist (run budget_tables.sql). ' + str(e)})

    # 3. Postponed orders: postpone date ekata dawas 3k (POSTPONE_NOTICE_DAYS) kalin idan alert enawa
    try:
        connection = get_db_connection()
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT customer_name, phone_number, product_name, status, postpone_date "
                "FROM orders WHERE status = 'Postponed' AND postpone_date IS NOT NULL "
                "ORDER BY postpone_date ASC, id ASC")
            rows = cursor.fetchall()
        connection.close()
        for o in rows:
            days_left = (o['postpone_date'] - date.today()).days
            if days_left > POSTPONE_NOTICE_DAYS:
                continue   # thama kalin, alert ekak dena eka wenne naa
            if days_left < 0:
                when = f"Overdue by {-days_left} day(s)"
            elif days_left == 0:
                when = "Due TODAY"
            else:
                when = f"Due in {days_left} day(s)"
            notifications.append({
                'type': 'postponed',
                'level': 'danger' if days_left <= 0 else 'warning',
                'title': f"{o['customer_name']} ({o['phone_number']})",
                'message': f"Postponed until {o['postpone_date'].isoformat()} - {when} - {o['product_name'] or ''}"
            })
    except Exception as e:
        notifications.append({'type': 'postponed', 'level': 'warning',
                              'title': 'Could not check postponed orders', 'message': str(e)})

    return jsonify({'notifications': notifications})


if __name__ == '__main__':
    app.run(debug=True)