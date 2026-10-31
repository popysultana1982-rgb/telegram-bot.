# ==============================================================================
# CyberNest Digital Mart — Telegram Auto-Shop & Business Engine
# Complete Architecture: aiogram/python-telegram-bot, SQLite, OpenPyXL, ReportLab, PyOTP
#
# Configured:
# - Super Admin ID: 8919985167
# - Token: Configured via env var BOT_TOKEN or fallback to new token
# - Render Web Service Health Check Server included
# ==============================================================================

import os
import sys
import io
import time
import json
import logging
import sqlite3
import datetime
import secrets
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from typing import Optional, List, Dict, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pyotp

# ReportLab for Official PDF Receipt Generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.barcode import qr

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    KeyboardButton,
    InputFile,
)
from telegram.constants import ParseMode
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# ------------------------------------------------------------------------------
# 1. LOGGING & CORE CONFIGURATION
# ------------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("CyberNestBot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "8656707191:AAEqK_Dn6M_OhFP12ATl6g5WyNCY02Db_sk")
SUPER_ADMIN_ID = int(os.getenv("SUPER_ADMIN_ID", "8919985167"))
DB_FILE = "store_database.db"

# ------------------------------------------------------------------------------
# 2. RENDER WEB SERVICE HEALTH CHECK SERVER (Keeps Render Free Instance Live)
# ------------------------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"CyberNest Bot is running and healthy!")

    def log_message(self, format, *args):
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    server.serve_forever()

# ------------------------------------------------------------------------------
# 3. HIGH-SPEED PERSISTENT DUAL-ENGINE DATABASE LAYER (PostgreSQL + SQLite)
# ------------------------------------------------------------------------------
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
IS_POSTGRES = bool(DATABASE_URL and ("postgres" in DATABASE_URL))

_GLOBAL_CONN = None

def get_raw_connection():
    global _GLOBAL_CONN
    if IS_POSTGRES:
        import psycopg2
        import psycopg2.extras
        if _GLOBAL_CONN is None or _GLOBAL_CONN.closed:
            url = DATABASE_URL
            if url.startswith("postgres://"):
                url = url.replace("postgres://", "postgresql://", 1)
            _GLOBAL_CONN = psycopg2.connect(url, cursor_factory=psycopg2.extras.RealDictCursor)
            _GLOBAL_CONN.autocommit = True
        return _GLOBAL_CONN
    else:
        if _GLOBAL_CONN is None:
            _GLOBAL_CONN = sqlite3.connect(DB_FILE, check_same_thread=False)
            _GLOBAL_CONN.row_factory = sqlite3.Row
        return _GLOBAL_CONN

class SmartCursor:
    def __init__(self, raw_cursor, is_pg):
        self.raw = raw_cursor
        self.is_pg = is_pg
        self._lastrowid = None

    def execute(self, sql, params=None):
        if self.is_pg:
            # 1. Translate parameter placeholder
            sql = sql.replace("?", "%s")
            # 2. DDL translations
            sql = sql.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY")
            sql = sql.replace("AUTOINCREMENT", "")
            sql = sql.replace("INSERT OR IGNORE INTO", "INSERT INTO")
            if "INSERT INTO settings" in sql and "ON CONFLICT" not in sql:
                sql += " ON CONFLICT (key) DO NOTHING"
            if "INSERT INTO users" in sql and "ON CONFLICT" not in sql:
                sql += " ON CONFLICT (user_id) DO NOTHING"

            # 3. Handle RETURNING id for lastrowid
            if sql.strip().upper().startswith("INSERT INTO") and ("RETURNING" not in sql.upper()) and ("SETTINGS" not in sql.upper()) and ("USERS" not in sql.upper()):
                sql += " RETURNING id"
                try:
                    res = self.raw.execute(sql, params) if params is not None else self.raw.execute(sql)
                    row = self.raw.fetchone()
                    if row:
                        self._lastrowid = row["id"] if isinstance(row, dict) else row[0]
                    return res
                except Exception as e:
                    logger.warning(f"Insert returning error: {e}")
                    pass
            # Normal execution
            try:
                if params is not None:
                    return self.raw.execute(sql, params)
                return self.raw.execute(sql)
            except Exception as e:
                try:
                    self.raw.connection.rollback()
                except Exception:
                    pass
                logger.error(f"SQL execution error: {e} in {sql}")
                raise
        else:
            if params is not None:
                res = self.raw.execute(sql, params)
            else:
                res = self.raw.execute(sql)
            self._lastrowid = getattr(self.raw, "lastrowid", None)
            return res

    @property
    def lastrowid(self):
        return self._lastrowid

    def fetchone(self):
        return self.raw.fetchone()

    def fetchall(self):
        return self.raw.fetchall()

class DBWrapper:
    def __init__(self):
        self.is_pg = IS_POSTGRES
        self.conn = get_raw_connection()

    def cursor(self):
        try:
            return SmartCursor(self.conn.cursor(), self.is_pg)
        except Exception:
            # Reconnect if connection was dropped by server
            global _GLOBAL_CONN
            _GLOBAL_CONN = None
            self.conn = get_raw_connection()
            return SmartCursor(self.conn.cursor(), self.is_pg)

    def commit(self):
        if not self.is_pg:
            self.conn.commit()

    def close(self):
        pass  # Keep persistent connection alive for blazing fast speed

def get_db():
    return DBWrapper()

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    # Users Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id BIGINT PRIMARY KEY,
        name TEXT,
        username TEXT,
        balance_bdt REAL DEFAULT 0.0,
        joined_date TEXT,
        is_admin INTEGER DEFAULT 0,
        current_state TEXT,
        state_data TEXT
    )
    """)

    # Categories Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        icon TEXT DEFAULT '🛍️'
    )
    """)

    # Subcategories Table (Brand / Category Subdivision)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS subcategories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id INTEGER,
        name TEXT NOT NULL,
        FOREIGN KEY (category_id) REFERENCES categories (id) ON DELETE CASCADE
    )
    """)

    # Products Table (Specific duration / package / variant)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subcategory_id INTEGER,
        name TEXT NOT NULL,
        price_bdt REAL NOT NULL,
        price_usdt REAL NOT NULL,
        image_file_id TEXT DEFAULT NULL,
        is_on_demand INTEGER DEFAULT 0,
        description TEXT,
        FOREIGN KEY (subcategory_id) REFERENCES subcategories (id) ON DELETE CASCADE
    )
    """)

    # Stock Items Table (For Excel Auto-Delivery like Mail, Proxy)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        content TEXT NOT NULL,
        is_sold INTEGER DEFAULT 0,
        sold_to BIGINT,
        sold_at TEXT,
        FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
    )
    """)

    # Orders Table (Stores lifetime purchases, receipt numbers, dates & credentials)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        receipt_no TEXT UNIQUE,
        user_id BIGINT,
        product_id INTEGER,
        product_name TEXT,
        category_name TEXT,
        quantity INTEGER,
        unit_price REAL,
        total_bdt REAL,
        total_usdt REAL,
        status TEXT DEFAULT 'completed',
        delivered_data TEXT,
        order_date TEXT,
        day_of_week TEXT,
        FOREIGN KEY (user_id) REFERENCES users (user_id)
    )
    """)

    # Deposits Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS deposits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id BIGINT,
        method TEXT,
        amount REAL,
        trx_id TEXT,
        status TEXT DEFAULT 'pending',
        created_at TEXT,
        FOREIGN KEY (user_id) REFERENCES users (user_id)
    )
    """)

    # Key-Value Settings Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    conn.commit()

    # Seed Default Settings
    default_settings = {
        "shop_name": "Will be Earn Shop",
        "shop_tagline": "Digital Products & Services",
        "usdt_rate": "125.0",
        "support_username": "@Talha_juba098",
        "bkash_number": "01XXXXXXXXX (Personal - Send Money)",
        "nagad_number": "01XXXXXXXXX (Personal - Send Money)",
        "binance_id": "Pay ID: 12345678 / USDT (TRC20): TQxxxxxxxxxxxxxxxxxx",
        "mail_code_url_1": "https://code.yamin.bd/",
        "mail_code_url_2": "https://outlook-code.example.com",
    }
    for k, v in default_settings.items():
        if conn.is_pg:
            cursor.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO NOTHING", (k, v))
        else:
            cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
    conn.commit()

    # Auto-healing: Ensure categories, subcategories, and products are fully populated
    cursor.execute("SELECT COUNT(*) as cnt FROM categories")
    row = cursor.fetchone()
    cnt = row["cnt"] if isinstance(row, dict) else row[0]
    cursor.execute("SELECT COUNT(*) as cnt FROM subcategories")
    s_row = cursor.fetchone()
    s_cnt = s_row["cnt"] if isinstance(s_row, dict) else s_row[0]
    if cnt == 0 or s_cnt == 0:
        logger.info("Database categories or subcategories missing, auto-healing and re-seeding...")
        try:
            cursor.execute("DELETE FROM products")
            cursor.execute("DELETE FROM subcategories")
            cursor.execute("DELETE FROM categories")
            conn.commit()
        except Exception:
            pass
        seed_clean_store(conn)
        logger.info("Auto-healing and seeding completed successfully.")

    # Force update verified settings in database
    try:
        cursor.execute("UPDATE settings SET value = 'https://code.yamin.bd/' WHERE key = 'mail_code_url_1'")
        cursor.execute("UPDATE settings SET value = '@Talha_juba098' WHERE key = 'support_username'")
        conn.commit()
    except Exception:
        pass
    conn.close()

def seed_clean_store(conn):
    cursor = conn.cursor()

    # 1. Categories (Clean without 'Other Domain Mail')
    cats = [
        ("🛡️ VPN", "🛡️"),
        ("🌐 Proxy", "🌐"),
        ("✉️ Mail", "✉️"),
        ("✉️ Recovery Added Mail", "✉️"),
    ]
    for cname, cicon in cats:
        cursor.execute("INSERT INTO categories (name, icon) VALUES (?, ?)", (cname, cicon))
    conn.commit()

    cursor.execute("SELECT id, name FROM categories")
    cat_map = {row["name"]: row["id"] for row in cursor.fetchall()}

    # 2. VPN Hierarchy: VPN -> Brands (Subcategories) -> Durations/Variants (Products)
    vpn_id = cat_map.get("🛡️ VPN")
    vpn_brands = [
        "Proton VPN", "Express VPN", "Nord VPN", "Surfshark VPN",
        "Cyberghost VPN", "IPVanish VPN", "Bitdefender VPN"
    ]
    durations = [
        ("3 Days", 15.0, 0.12),
        ("7 Days", 20.0, 0.16),
        ("10 Days", 25.0, 0.20),
        ("15 Days", 35.0, 0.28),
        ("30 Days", 60.0, 0.48),
    ]

    for brand in vpn_brands:
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (vpn_id, brand))
        sub_id = cursor.lastrowid
        for dur_name, bdt, usdt in durations:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 1)",
                (sub_id, dur_name, bdt, usdt),
            )

    # 3. Proxy Hierarchy: Proxy -> Brands (Subcategories) -> MB Packages (Products)
    proxy_id = cat_map.get("🌐 Proxy")
    proxy_brands = ["9 Proxy", "Rapid Proxy", "711 Proxy"]
    proxy_packages = [
        ("100 MB", 12.0, 0.096),
        ("200 MB", 23.0, 0.184),
        ("500 MB", 60.0, 0.480),
        ("1 GB", 110.0, 0.880),
    ]

    for pbrand in proxy_brands:
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (proxy_id, pbrand))
        psub_id = cursor.lastrowid
        for pkg_name, bdt, usdt in proxy_packages:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (psub_id, pkg_name, bdt, usdt),
            )

    # 4. Mail Hierarchy: Mail -> Subcategory -> Accounts
    mail_id = cat_map.get("✉️ Mail")
    cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (mail_id, "Accounts"))
    mid = cursor.lastrowid
    for mname, bdt, usdt in [("FR Outlook", 12.0, 0.096), ("Hotmail", 10.0, 0.080), ("Outlook", 10.0, 0.080)]:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
            (mid, mname, bdt, usdt),
        )

    # 5. Recovery Added Mail
    rec_mail_id = cat_map.get("✉️ Recovery Added Mail")
    cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (rec_mail_id, "Recovery Accounts"))
    rmid = cursor.lastrowid
    for rname, bdt, usdt in [("Recovery Added FR Outlook", 15.0, 0.120), ("Recovery Added Hotmail", 12.0, 0.096), ("Recovery Added Outlook", 12.0, 0.096)]:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
            (rmid, rname, bdt, usdt),
        )

    conn.commit()

# DB Query Helpers
def get_setting(key: str, default: str = "") -> str:
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key = ?", (key,))
    row = c.fetchone()
    conn.close()
    return row["value"] if row else default

def set_setting(key: str, value: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def get_or_create_user(user_id: int, name: str, username: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    if not row:
        joined = datetime.datetime.now().strftime("%Y-%m-%d")
        is_adm = 1 if user_id == SUPER_ADMIN_ID else 0
        c.execute(
            "INSERT INTO users (user_id, name, username, balance_bdt, joined_date, is_admin) VALUES (?, ?, ?, ?, ?, ?)",
            (user_id, name, username or "", 0.0, joined, is_adm),
        )
        conn.commit()
        c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
        row = c.fetchone()
    else:
        if user_id == SUPER_ADMIN_ID and row["is_admin"] != 1:
            c.execute("UPDATE users SET is_admin = 1 WHERE user_id = ?", (user_id,))
            conn.commit()
            c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
            row = c.fetchone()
    conn.close()
    return dict(row)

def set_user_state(user_id: int, state: Optional[str], data: Optional[Dict] = None):
    conn = get_db()
    c = conn.cursor()
    data_str = json.dumps(data) if data else "{}"
    c.execute("UPDATE users SET current_state = ?, state_data = ? WHERE user_id = ?", (state, data_str, user_id))
    conn.commit()
    conn.close()

def get_user_state(user_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT current_state, state_data FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        data = json.loads(row["state_data"]) if row["state_data"] else {}
        return row["current_state"], data
    return None, {}

def is_admin(user_id: int) -> bool:
    if user_id == SUPER_ADMIN_ID:
        return True
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT is_admin FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    return bool(row and row["is_admin"] == 1)

# ------------------------------------------------------------------------------
# 4. KEYBOARD LAYOUTS
# ------------------------------------------------------------------------------
def get_main_reply_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    """Returns the persistent main menu keyboard.
    👑 Admin Panel is displayed ONLY for Super Admin / Admin."""
    buttons = [
        [KeyboardButton("🛍️ Buy Products")],
        [KeyboardButton("👤 My Profile"), KeyboardButton("💳 Deposit")],
        [KeyboardButton("🔐 Get Code"), KeyboardButton("💬 Support")],
    ]
    if is_admin(user_id):
        buttons.append([KeyboardButton("👑 Admin Panel")])
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True)

# ------------------------------------------------------------------------------
# 5. EXCEL & PDF RECEIPT GENERATORS
# ------------------------------------------------------------------------------
def generate_excel_delivery(items: List[str], product_name: str, receipt_no: str) -> io.BytesIO:
    """Generates an aesthetic .xlsx file containing the delivered credentials."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Credentials"

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = ["SL", "Product / Variant", "Credential / Session / Key Data", "Delivery Date"]
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
    for idx, item in enumerate(items, start=1):
        row_data = [idx, product_name, item, now_str]
        ws.append(row_data)
        row_idx = idx + 1
        for c in range(1, len(row_data) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = data_font
            cell.border = thin_border
            if c in (1, 4):
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 75
    ws.column_dimensions['D'].width = 22

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def read_items_from_excel(file_bytes: bytes) -> List[str]:
    """Reads non-empty accounts from uploaded .xlsx spreadsheet."""
    items = []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell is not None:
                val = str(cell).strip()
                if val.lower() in ("sl", "no", "mail", "email", "password", "account", "credential", "product / variant", "delivery date"):
                    continue
                if len(val) > 2:
                    items.append(val)
    return items

def generate_pdf_receipt(
    receipt_no: str,
    date_str: str,
    day_name: str,
    customer_name: str,
    telegram_id: int,
    username: str,
    item_desc: str,
    category_name: str,
    qty: int,
    unit_price: float,
    total_bdt: float,
    bot_username: str,
) -> io.BytesIO:
    """Generates an official payment receipt PDF matching Receipt_RCP-287D10DA5A.pdf."""
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )

    styles = getSampleStyleSheet()
    shop_title_style = ParagraphStyle(
        'ShopTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=24,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=1,
    )
    shop_tag_style = ParagraphStyle(
        'ShopTag',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        textColor=colors.HexColor('#E0E0E0'),
        alignment=1,
    )
    receipt_title_style = ParagraphStyle(
        'ReceiptTitle',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=colors.HexColor('#FFFFFF'),
        alignment=1,
    )

    story = []

    # 1. Header Banner Table with Diagonal 'PAID IN FULL' badge effect
    shop_name = get_setting("shop_name", "Will be Earn Shop")
    shop_tagline = get_setting("shop_tagline", "Digital Products & Services")

    header_html = f"""
    <b>{shop_name}</b><br/>
    <font size=10 color='#E0E0E0'>{shop_tagline}</font><br/><br/>
    <font size=12 color='#FFD700'><b>★ OFFICIAL PAYMENT RECEIPT ★</b></font>
    """
    p_header = Paragraph(header_html, shop_title_style)

    badge_html = "<br/><br/><font color='#000000' size=11><b>PAID IN FULL</b></font>"
    p_badge = Paragraph(badge_html, ParagraphStyle('Badge', alignment=1, textColor=colors.black))

    header_table = Table([[p_header, p_badge]], colWidths=[400, 120])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#4A154B')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#FFCC00')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 16),
        ('TOPPADDING', (0, 0), (-1, -1), 16),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 20))

    # 2. Receipt Metadata Grid
    meta_data = [
        [
            Paragraph(f"<b>RECEIPT NO.</b><br/><font color='#C2185B' size=12><b>{receipt_no}</b></font>", styles['Normal']),
            Paragraph(f"<b>DATE ISSUED</b><br/>{date_str} ({day_name})", styles['Normal']),
        ],
        [
            Paragraph(f"<b>CUSTOMER</b><br/>{customer_name}", styles['Normal']),
            Paragraph(f"<b>TELEGRAM ID</b><br/><code>{telegram_id}</code>", styles['Normal']),
        ],
        [
            Paragraph(f"<b>USERNAME</b><br/>@{username if username else 'N/A'}", styles['Normal']),
            Paragraph("<b>PAYMENT METHOD</b><br/><font color='#2E7D32'><b>Wallet Balance</b></font>", styles['Normal']),
        ]
    ]
    meta_table = Table(meta_data, colWidths=[260, 260])
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('LINEBELOW', (0, -1), (-1, -1), 1, colors.HexColor('#E0E0E0')),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    # 3. Category Badge & Items Table
    badge_data = [[Paragraph(f"<font color='#4A154B'><b>CATEGORIES: {category_name.upper()}</b></font>", styles['Normal'])]]
    t_badge = Table(badge_data, colWidths=[520])
    t_badge.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F3E5F5')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_badge)
    story.append(Spacer(1, 8))

    item_table_data = [
        ["ITEM DESCRIPTION", "QTY", "PRICE", "AMOUNT"],
        [Paragraph(f"<b>{item_desc}</b><br/><font size=8 color='#666'>Instant Account / Service Delivery</font>", styles['Normal']),
         str(qty), f"{unit_price:.2f}", f"{total_bdt:.2f}"],
    ]
    item_table = Table(item_table_data, colWidths=[280, 60, 90, 90])
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A237E')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
    ]))
    story.append(item_table)

    # 4. Total Paid Highlight Bar
    total_data = [["TOTAL PAID (BDT)", f"{total_bdt:.2f} BDT"]]
    total_table = Table(total_data, colWidths=[370, 150])
    total_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#7B1FA2')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 13),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(total_table)
    story.append(Spacer(1, 25))

    # 5. QR Code Deep-Link (Scannable link that opens bot order directly)
    clean_bot = bot_username.lstrip('@')
    verify_url = f"https://t.me/{clean_bot}?start=order_{receipt_no}"

    qr_widget = qr.QrCodeWidget(verify_url)
    qr_widget.barWidth = 85
    qr_widget.barHeight = 85
    qr_widget.qrVersion = 2
    d_qr = Drawing(85, 85)
    d_qr.add(qr_widget)

    ref_text = (
        f"<b>Order Reference:</b><br/>"
        f"<b>{receipt_no}</b><br/>"
        f"<font color='#2E7D32'>✔ Payment Verified</font><br/><br/>"
        f"<font size=8 color='#555'>Scan the QR code to verify this order directly in Telegram.</font>"
    )
    p_ref = Paragraph(ref_text, styles['Normal'])

    footer_grid = Table([[p_ref, d_qr]], colWidths=[400, 120])
    footer_grid.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
    ]))
    story.append(footer_grid)
    story.append(Spacer(1, 20))

    # 6. Bottom System Disclaimer
    disc = (
        "<i>This is a system-generated receipt and does not require a signature.</i><br/>"
        f"<b>Thank you for shopping with {shop_name}!</b>"
    )
    story.append(Paragraph(disc, ParagraphStyle('Disc', alignment=1, fontSize=9, textColor=colors.HexColor('#555555'))))

    doc.build(story)
    bio.seek(0)
    return bio

# ------------------------------------------------------------------------------
# 6. USER HANDLERS (Shop, Profile, Deposit, 2FA, Orders)
# ------------------------------------------------------------------------------
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    shop_name = get_setting("shop_name", "Will be Earn Shop")

    # Check for Deep-Linking verification argument: /start order_<receipt_no>
    if context.args and len(context.args) > 0:
        arg = context.args[0]
        if arg.startswith("order_"):
            receipt_code = arg.replace("order_", "")
            await show_order_verification(update, receipt_code)
            return

    welcome_text = (
        f"🎉 <b>স্বাগতম, {user.full_name}!</b> ✨\n\n"
        f"💎 <b>{shop_name}</b>\n"
        f"─────────────────────────────\n"
        f"🛍️ প্রিমিয়াম VPN, Proxy আর Mail সার্ভিস — সবচেয়ে দ্রুত ডেলিভারিতে, ২৪/৭।\n\n"
        f"📌 নিচের মেনু থেকে বেছে নিন 👇"
    )
    reply_kb = get_main_reply_keyboard(user.id)
    await update.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)

async def show_order_verification(update: Update, receipt_no: str):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM orders WHERE receipt_no = ?", (receipt_no,))
    order = c.fetchone()
    conn.close()

    if not order:
        await update.message.reply_text("❌ দুঃখিত! এই রিসিট নম্বরের কোনো অর্ডার ডেটাবেজে পাওয়া যায়নি।")
        return

    txt = (
        f"🧾 <b>অর্ডার ভেরিফিকেশন কার্ড (VERIFIED)</b>\n"
        f"─────────────────────────────\n"
        f"🔹 <b>Receipt No:</b> <code>{order['receipt_no']}</code>\n"
        f"📅 <b>তারিখ ও বার:</b> {order['order_date']} ({order['day_of_week']})\n"
        f"👤 <b>কাস্টমার আইডি:</b> <code>{order['user_id']}</code>\n"
        f"📦 <b>প্রোডাক্ট:</b> {order['product_name']}\n"
        f"🔢 <b>পরিমাণ:</b> {order['quantity']} টি\n"
        f"💰 <b>পরিশোধিত মূল্য:</b> {order['total_bdt']:.2f} BDT\n"
        f"─────────────────────────────\n"
        f"✔ <i>এই ট্রানজ্যাকশনটি সফল ও শতভাগ অথেনটিক।</i>"
    )
    await update.message.reply_text(txt, parse_mode=ParseMode.HTML)

async def profile_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    rate = float(get_setting("usdt_rate", "125.0"))
    balance_bdt = db_user["balance_bdt"]
    balance_usdt = balance_bdt / rate if rate > 0 else 0.0

    profile_text = (
        f"👤 <b>My Profile</b>\n\n"
        f"👤 <b>Name</b> : {db_user['name']}\n"
        f"💰 <b>Balance</b> : {balance_bdt:.2f} BDT / {balance_usdt:.4f} USDT\n"
        f"🎉 <b>Joined</b> : {db_user['joined_date']}\n"
        f"─────────────────────────────"
    )
    inline_kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("💰 Balance Details", callback_data="prof_balance_details"),
            InlineKeyboardButton("🔄 Deposit History", callback_data="prof_deposit_history"),
        ],
        [
            InlineKeyboardButton("👥 Referral", callback_data="prof_referral"),
            InlineKeyboardButton("🛒 Purchases", callback_data="prof_purchases"),
        ]
    ])
    await update.message.reply_text(profile_text, parse_mode=ParseMode.HTML, reply_markup=inline_kb)

async def buy_products_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, name FROM categories")
    cats = c.fetchall()
    conn.close()

    text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
    keyboard = []
    row = []
    for cat in cats:
        row.append(InlineKeyboardButton(cat["name"], callback_data=f"cat_{cat['id']}"))
        if len(row) == 2:
            keyboard.append(row)
            row = []
    if row:
        keyboard.append(row)

    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(keyboard))

async def get_code_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "🔐 <b>Get Code & Authentication Tools</b>\n\n"
        "হটমেইল/আউটলুক ওটিপি কোড এবং লাইভ 2FA কোড পাওয়ার জন্য নিচের অপশন থেকে নির্বাচন করুন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 2FA Authenticator (Live Code)", callback_data="tool_2fa_prompt")],
        [InlineKeyboardButton("🌐 Hotmail & Outlook Code Server", url="https://code.yamin.bd/")],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "💬 <b>কাস্টমার সাপোর্ট</b>\n\n"
        "যেকোনো সমস্যায় নিচে এডমিনকে মেসেজ দিন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📨 Message Admin", url="tg://user?id=8919985167")]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

# Deposit Menu with 3 Separate Buttons: bKash, Nagad, Binance
async def deposit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    rate = get_setting("usdt_rate", "125.0")

    text = (
        f"💳 <b>অ্যাড ব্যালেন্স / ডিপোজিট</b>\n"
        f"─────────────────────────────\n"
        f"💰 বর্তমান ব্যালেন্স: <b>{db_user['balance_bdt']:.2f} BDT</b>\n"
        f"💵 ডলার রেট: <b>1 USDT = {rate} BDT</b>\n\n"
        f"নিচের যে মাধ্যমে টাকা পাঠাতে চান, সেই বাটনটিতে চাপ দিন:"
    )
    kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📱 bKash", callback_data="dep_view_bKash"),
            InlineKeyboardButton("📱 Nagad", callback_data="dep_view_Nagad"),
        ],
        [
            InlineKeyboardButton("💎 Binance (USDT)", callback_data="dep_view_Binance"),
        ]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

# ------------------------------------------------------------------------------
# 7. INLINE CALLBACK QUERY DISPATCHER (All Navigation & Purchase Routing)
# ------------------------------------------------------------------------------
async def callback_query_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    db_user = get_or_create_user(user_id, query.from_user.full_name, query.from_user.username)

    # 1. Category Click -> Show Subcategories / Brands (Level 2)
    if data.startswith("cat_"):
        cat_id = int(data.split("_")[1])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ?", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        text = f"<b>{cat_name} — Select Brand / Section:</b>"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(s["name"], callback_data=f"sub_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="buy_menu_root")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Root Categories
    if data == "buy_menu_root":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories")
        cats = c.fetchall()
        conn.close()
        text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
        keyboard = []
        row = []
        for cat in cats:
            row.append(InlineKeyboardButton(cat["name"], callback_data=f"cat_{cat['id']}"))
            if len(row) == 2:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(keyboard))
        return

    # 2. Subcategory Click -> Show Variants / Durations (Level 3)
    if data.startswith("sub_"):
        sub_id = int(data.split("_")[1])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt, price_usdt FROM products WHERE subcategory_id = ?", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} — {meta['sub_name']}" if meta else "Select Plan"
        text = f"<b>{title}</b>\n\nনিচের প্ল্যান বা মেয়াদ নির্বাচন করুন:"
        kb = []
        for p in prods:
            btn_text = f"{p['name']} • {p['price_bdt']:.2f} BDT / {p['price_usdt']:.4f} USDT"
            kb.append([InlineKeyboardButton(btn_text, callback_data=f"prod_{p['id']}")])
        back_cat_id = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"cat_{back_cat_id}")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # 3. Product Click -> Show Clean Card / Specific Image & Quantity Selector
    if data.startswith("prod_"):
        prod_id = int(data.split("_")[1])
        await show_product_card(query, prod_id)
        return

    # 4. Quantity Selection
    if data.startswith("qty_"):
        parts = data.split("_")
        prod_id = int(parts[1])
        qty = int(parts[2])
        await process_order(query, prod_id, qty)
        return

    # 5. Custom Quantity Prompt
    if data.startswith("custqty_"):
        prod_id = int(data.split("_")[1])
        set_user_state(user_id, "WAITING_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text(
            "✏️ <b>Custom Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):",
            parse_mode=ParseMode.HTML,
        )
        return

    # 6. Deposit Details Views for 3 Methods
    if data == "dep_view_bKash":
        num = get_setting("bkash_number")
        text = (
            f"📱 <b>bKash Personal Payment:</b>\n\n"
            f"নাম্বার: <code>{num}</code> (Send Money)\n\n"
            f"টাকা পাঠানোর পর নিচের বাটনে চাপ দিয়ে TrxID এবং টাকার পরিমাণ সাবমিট করুন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Submit bKash TrxID", callback_data="dep_form_bKash")],
            [InlineKeyboardButton("< Back", callback_data="dep_menu_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "dep_view_Nagad":
        num = get_setting("nagad_number")
        text = (
            f"📱 <b>Nagad Personal Payment:</b>\n\n"
            f"নাম্বার: <code>{num}</code> (Send Money)\n\n"
            f"টাকা পাঠানোর পর নিচের বাটনে চাপ দিয়ে TrxID এবং টাকার পরিমাণ সাবমিট করুন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Submit Nagad TrxID", callback_data="dep_form_Nagad")],
            [InlineKeyboardButton("< Back", callback_data="dep_menu_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "dep_view_Binance":
        num = get_setting("binance_id")
        text = (
            f"💎 <b>Binance / Crypto Payment:</b>\n\n"
            f"ঠিকানা / আইডি: <code>{num}</code>\n\n"
            f"টাকা/ডলার পাঠানোর পর নিচের বাটনে চাপ দিয়ে ট্রানজ্যাকশন আইডি ও টাকার পরিমাণ সাবমিট করুন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Submit Binance TrxID", callback_data="dep_form_Binance")],
            [InlineKeyboardButton("< Back", callback_data="dep_menu_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "dep_menu_root":
        rate = get_setting("usdt_rate", "125.0")
        text = (
            f"💳 <b>অ্যাড ব্যালেন্স / ডিপোজিট</b>\n"
            f"─────────────────────────────\n"
            f"💰 বর্তমান ব্যালেন্স: <b>{db_user['balance_bdt']:.2f} BDT</b>\n"
            f"💵 ডলার রেট: <b>1 USDT = {rate} BDT</b>\n\n"
            f"নিচের যে মাধ্যমে টাকা পাঠাতে চান, সেই বাটনটিতে চাপ দিন:"
        )
        kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("📱 bKash", callback_data="dep_view_bKash"),
                InlineKeyboardButton("📱 Nagad", callback_data="dep_view_Nagad"),
            ],
            [
                InlineKeyboardButton("💎 Binance (USDT)", callback_data="dep_view_Binance"),
            ]
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data.startswith("dep_form_"):
        method = data.split("_")[2]
        set_user_state(user_id, "WAITING_DEP_AMOUNT", {"method": method})
        await query.message.reply_text(f"✅ আপনি <b>{method}</b> নির্বাচন করেছেন।\nকত টাকা পাঠিয়েছেন তা লিখুন (যেমন: 500):", parse_mode=ParseMode.HTML)
        return

    # 7. 2FA Tool
    if data == "tool_2fa_prompt":
        set_user_state(user_id, "WAITING_2FA_KEY", {})
        await query.message.reply_text(
            "🔑 <b>2FA Authenticator:</b>\nআপনার 2FA Secret Key টি এখানে পেস্ট করে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    # 8. Profile Purchases (Lifetime History with Day of Week, Date, Time, Total)
    if data == "prof_purchases":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 15", (user_id,))
        orders = c.fetchall()
        conn.close()
        if not orders:
            await query.message.reply_text("🛒 আপনার পূর্বের কোনো কেনাকাটার হিস্টোরি পাওয়া যায়নি।")
            return
        msg = "🛒 <b>আপনার সর্বমোট অর্ডার ও কেনাকাটা হিস্টোরি:</b>\n─────────────────────────────\n"
        for o in orders:
            msg += (
                f"🧾 <b>Receipt:</b> <code>{o['receipt_no']}</code>\n"
                f"📦 <b>আইটেম:</b> {o['product_name']}\n"
                f"🔢 <b>পরিমাণ:</b> {o['quantity']} টি | মোট: <b>{o['total_bdt']:.2f} BDT</b>\n"
                f"📅 <b>তারিখ ও বার:</b> {o['order_date']} ({o['day_of_week']})\n"
                f"─────────────────────────────\n"
            )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data == "prof_deposit_history":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("🔄 আপনার কোনো ডিপোজিট হিস্টোরি পাওয়া যায়নি।")
            return
        msg = "🔄 <b>আপনার ডিপোজিট হিস্টোরি:</b>\n─────────────────────────────\n"
        for d in deps:
            status_icon = "⏳ Pending" if d['status'] == "pending" else ("✅ Approved" if d['status'] == "approved" else "❌ Rejected")
            msg += (
                f"💳 মেথড: {d['method']} | TrxID: <code>{d['trx_id']}</code>\n"
                f"💰 পরিমাণ: {d['amount']:.2f} BDT | অবস্থা: {status_icon}\n"
                f"📅 তারিখ: {d['created_at']}\n"
                f"─────────────────────────────\n"
            )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data == "prof_balance_details":
        rate = float(get_setting("usdt_rate", "125.0"))
        bdt = db_user['balance_bdt']
        usdt = bdt / rate if rate > 0 else 0.0
        msg = (
            f"💰 <b>ব্যালেন্স বিবরণী:</b>\n\n"
            f"💵 মূল ব্যালেন্স: <b>{bdt:.2f} BDT</b>\n"
            f"💎 ডলার সমমূল্য: <b>{usdt:.4f} USDT</b>\n"
            f"📈 এক্সচেঞ্জ রেট: 1 USDT = {rate} BDT"
        )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    if data == "prof_referral":
        bot_info = await context.bot.get_me()
        ref_link = f"https://t.me/{bot_info.username}?start=ref_{user_id}"
        msg = (
            f"👥 <b>রেফারাল প্রোগ্রাম</b>\n\n"
            f"আপনার বন্ধুদের আমন্ত্রণ জানান এবং স্পেশাল অফার উপভোগ করুন!\n\n"
            f"🔗 আপনার রেফার লিংক:\n<code>{ref_link}</code>"
        )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    # 9. Admin Direct Action Routing (Fixed bug: Handles all adm_ and admin_ routes)
    if data.startswith("adm_") or data.startswith("admin_"):
        if not is_admin(user_id):
            return
        await handle_admin_callbacks(query, data, context)
        return

# ------------------------------------------------------------------------------
# 8. PRODUCT CARD & PURCHASE ENGINE
# ------------------------------------------------------------------------------
async def show_product_card(query, prod_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    SELECT p.*, s.name as sub_name, c.name as cat_name
    FROM products p
    JOIN subcategories s ON p.subcategory_id = s.id
    JOIN categories c ON s.category_id = c.id
    WHERE p.id = ?
    """, (prod_id,))
    p = c.fetchone()

    c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
    avail_stock = c.fetchone()["cnt"]
    conn.close()

    if not p:
        await query.message.reply_text("প্রোডাক্টটি পাওয়া যায়নি।")
        return

    stock_badge = "⚡ Instant Request Delivery (২৪/৭ ডেলিভারি)" if p["is_on_demand"] else f"{avail_stock} টি স্টকে আছে"

    desc_part = f"\n📝 <i>{p['description']}</i>\n" if p['description'] else ""
    caption = (
        f"<b>{p['sub_name']} — {p['name']}</b>\n"
        f"─────────────────────────────{desc_part}\n"
        f"💰 <b>Price:</b> {p['price_bdt']:.2f} BDT / {p['price_usdt']:.4f} USDT per unit\n"
        f"📊 <b>Stock:</b> {stock_badge}\n\n"
        f"Select quantity:"
    )

    kb = [
        [
            InlineKeyboardButton(f"1× {p['price_bdt']:.0f} BDT", callback_data=f"qty_{p['id']}_1"),
            InlineKeyboardButton(f"2× {p['price_bdt']*2:.0f} BDT", callback_data=f"qty_{p['id']}_2"),
        ],
        [
            InlineKeyboardButton(f"3× {p['price_bdt']*3:.0f} BDT", callback_data=f"qty_{p['id']}_3"),
            InlineKeyboardButton(f"4× {p['price_bdt']*4:.0f} BDT", callback_data=f"qty_{p['id']}_4"),
        ],
        [
            InlineKeyboardButton(f"5× {p['price_bdt']*5:.0f} BDT", callback_data=f"qty_{p['id']}_5"),
            InlineKeyboardButton(f"6× {p['price_bdt']*6:.0f} BDT", callback_data=f"qty_{p['id']}_6"),
        ],
        [
            InlineKeyboardButton(f"7× {p['price_bdt']*7:.0f} BDT", callback_data=f"qty_{p['id']}_7"),
            InlineKeyboardButton(f"8× {p['price_bdt']*8:.0f} BDT", callback_data=f"qty_{p['id']}_8"),
        ],
        [
            InlineKeyboardButton(f"9× {p['price_bdt']*9:.0f} BDT", callback_data=f"qty_{p['id']}_9"),
            InlineKeyboardButton(f"10× {p['price_bdt']*10:.0f} BDT", callback_data=f"qty_{p['id']}_10"),
        ],
        [InlineKeyboardButton("✏️ Custom Quantity", callback_data=f"custqty_{p['id']}")],
        [InlineKeyboardButton("< Back", callback_data=f"sub_{p['subcategory_id']}")],
    ]

    # If Admin uploaded an individual photo for this variant, show it! Otherwise clean text only!
    if p["image_file_id"]:
        try:
            await query.message.reply_photo(
                photo=p["image_file_id"],
                caption=caption,
                parse_mode=ParseMode.HTML,
                reply_markup=InlineKeyboardMarkup(kb),
            )
            return
        except Exception as e:
            logger.warning(f"Error sending custom photo: {e}")

    # Clean text layout with zero unwanted images
    await query.message.reply_text(caption, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))

async def process_order(query, prod_id: int, quantity: int):
    user_id = query.from_user.id
    user_name = query.from_user.full_name
    username = query.from_user.username or ""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    SELECT p.*, s.name as sub_name, c.name as cat_name
    FROM products p
    JOIN subcategories s ON p.subcategory_id = s.id
    JOIN categories c ON s.category_id = c.id
    WHERE p.id = ?
    """, (prod_id,))
    prod = c.fetchone()
    c.execute("SELECT balance_bdt FROM users WHERE user_id = ?", (user_id,))
    user_balance = c.fetchone()["balance_bdt"]

    if not prod:
        conn.close()
        await query.message.reply_text("প্রোডাক্ট পাওয়া যায়নি।")
        return

    total_bdt = prod["price_bdt"] * quantity
    total_usdt = prod["price_usdt"] * quantity

    # Balance Verification
    if user_balance < total_bdt:
        conn.close()
        shortage = total_bdt - user_balance
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("💳 Deposit Now", callback_data="dep_menu_root")]])
        await query.message.reply_text(
            f"❌ <b>অপর্যাপ্ত ব্যালেন্স!</b>\n\n"
            f"প্রয়োজনীয় টাকা: {total_bdt:.2f} BDT\n"
            f"আপনার ব্যালেন্স: {user_balance:.2f} BDT\n"
            f"ঘাটতি: {shortage:.2f} BDT\n\n"
            f"দয়া করে ব্যালেন্স রিচার্জ করুন।",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d %I:%M:%S %p")
    days_bn = {"Monday": "সোমবার", "Tuesday": "মঙ্গলবার", "Wednesday": "বুধবার", "Thursday": "বৃহস্পতিবার", "Friday": "শুক্রবার", "Saturday": "শনিবার", "Sunday": "রবিবার"}
    day_name = days_bn.get(now.strftime("%A"), now.strftime("%A"))
    receipt_no = f"RCP-{secrets.token_hex(5).upper()}"
    full_product_name = f"{prod['sub_name']} ({prod['name']})"

    # 1. On-Demand Products (VPN)
    if prod["is_on_demand"]:
        c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
        c.execute("""
        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, order_date, day_of_week)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending_delivery', ?, ?)
        """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, now_str, day_name))
        order_id = c.lastrowid
        conn.commit()
        conn.close()

        # Confirmation to user
        await query.message.reply_text(
            f"✅ <b>অর্ডার সফল হয়েছে!</b>\n"
            f"🧾 <b>Receipt:</b> <code>{receipt_no}</code>\n"
            f"🛡️ <b>প্রোডাক্ট:</b> {full_product_name}\n"
            f"🔢 <b>পরিমাণ:</b> {quantity} টি | মোট: {total_bdt:.2f} BDT\n"
            f"📅 <b>তারিখ:</b> {now_str} ({day_name})\n\n"
            f"⏳ এডমিন খুব দ্রুত ইনবক্সে ভিপিএন অ্যাক্সেস পাঠিয়ে দেবে।",
            parse_mode=ParseMode.HTML,
        )

        # Alert Super Admin
        adm_kb = InlineKeyboardMarkup([[InlineKeyboardButton(f"🚀 Deliver Order #{order_id}", callback_data=f"adm_fulfill_{order_id}")]])
        try:
            await query.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>নতুন VPN অর্ডার এসেছে!</b>\n"
                    f"🧾 <b>Receipt:</b> <code>{receipt_no}</code> (Order #{order_id})\n"
                    f"👤 <b>কাস্টমার:</b> {user_name} (ID: <code>{user_id}</code>)\n"
                    f"🛡️ <b>প্রোডাক্ট:</b> {full_product_name}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি | পেইড: {total_bdt:.2f} BDT"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
        except Exception as e:
            logger.error(f"Failed to alert admin: {e}")
        return

    # 2. Instant Stock Products (Mail, Proxy)
    c.execute("SELECT id, content FROM stock WHERE product_id = ? AND is_sold = 0 LIMIT ?", (prod_id, quantity))
    stock_rows = c.fetchall()
    if len(stock_rows) < quantity:
        conn.close()
        await query.message.reply_text(
            f"⚠️ দুঃখিত! পর্যাপ্ত স্টক নেই।\nবর্তমানে স্টকে আছে: <b>{len(stock_rows)}</b> টি।",
            parse_mode=ParseMode.HTML,
        )
        return

    # Deduct Balance
    c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))

    # Mark Stock as Sold
    stock_ids = [r["id"] for r in stock_rows]
    items_content = [r["content"] for r in stock_rows]
    c.execute(
        f"UPDATE stock SET is_sold = 1, sold_to = ?, sold_at = ? WHERE id IN ({','.join(['?']*len(stock_ids))})",
        [user_id, now_str] + stock_ids,
    )

    # Insert Order Record
    delivered_json = json.dumps(items_content)
    c.execute("""
    INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, delivered_data, order_date, day_of_week)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', ?, ?, ?)
    """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, delivered_json, now_str, day_name))
    order_id = c.lastrowid
    conn.commit()
    conn.close()

    bot_info = await query.bot.get_me()

    # 1. Generate & Send Excel File
    excel_bio = generate_excel_delivery(items_content, full_product_name, receipt_no)
    excel_caption = (
        f"🎉 <b>আপনার অর্ডার সফলভাবে ডেলিভারি হয়েছে!</b>\n"
        f"🧾 <b>Receipt:</b> <code>{receipt_no}</code>\n"
        f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
        f"🔢 <b>পরিমাণ:</b> {quantity} টি | মোট: {total_bdt:.2f} BDT\n\n"
        f"📂 আপনার ডেলিভারি ডেটার এক্সেল ফাইলটি উপরে দেওয়া হলো।"
    )
    filename = f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"
    await query.message.reply_document(
        document=InputFile(excel_bio, filename=filename),
        caption=excel_caption,
        parse_mode=ParseMode.HTML,
    )

    # 2. Generate & Send Official Payment Receipt PDF with QR Code
    try:
        pdf_bio = generate_pdf_receipt(
            receipt_no=receipt_no,
            date_str=now_str,
            day_name=day_name,
            customer_name=user_name,
            telegram_id=user_id,
            username=username,
            item_desc=full_product_name,
            category_name=prod["cat_name"],
            qty=quantity,
            unit_price=prod["price_bdt"],
            total_bdt=total_bdt,
            bot_username=bot_info.username,
        )
        receipt_caption = f"🧾 <b>Official Payment Receipt:</b> {receipt_no}\nকিউআর কোড স্ক্যান করে ভেরিফাই করতে পারবেন।"
        await query.message.reply_document(
            document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
            caption=receipt_caption,
            parse_mode=ParseMode.HTML,
        )
    except Exception as e:
        logger.error(f"Failed to generate receipt PDF: {e}")

# ------------------------------------------------------------------------------
# 9. ADMIN PANEL ENGINE (All Button-Driven with Photo & CRUD Management)
# ------------------------------------------------------------------------------
async def admin_panel_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        return

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) as cnt FROM deposits WHERE status = 'pending'")
    pend_dep = c.fetchone()["cnt"]
    c.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = 'pending_delivery'")
    pend_ord = c.fetchone()["cnt"]
    conn.close()

    text = (
        f"👑 <b>সুপার এডমিন কন্ট্রোল ড্যাশবোর্ড</b>\n"
        f"─────────────────────────────\n"
        f"সব কাজ বাটন দিয়ে পরিচালনা করুন। কোনো কমান্ড লিখতে হবে না।\n\n"
        f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
        f"🛡️ পেন্ডিং VPN অর্ডার: <b>{pend_ord}</b> টি"
    )
    kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📦 Manage Stock", callback_data="adm_stock_menu"),
            InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_buttons_menu"),
        ],
        [
            InlineKeyboardButton("💳 Payment Methods (bKash/Nagad/Binance)", callback_data="adm_payment_methods"),
            InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_photos_menu"),
        ],
        [
            InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
            InlineKeyboardButton(f"🛡️ VPN Orders ({pend_ord})", callback_data="adm_pending_orders"),
        ],
        [
            InlineKeyboardButton("👥 User Search & Balance", callback_data="adm_user_mgmt"),
            InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
        ],
        [
            InlineKeyboardButton("⚙️ General Settings", callback_data="adm_settings"),
        ]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def handle_admin_callbacks(query, data: str, context: ContextTypes.DEFAULT_TYPE):
    user_id = query.from_user.id
    try:
        await _dispatch_admin_callback(query, data, context, user_id)
    except Exception as e:
        logger.error(f"Error in admin callback {data}: {e}", exc_info=True)
        try:
            await query.message.reply_text(f"⚠️ এডমিন রিকোয়েস্টে সমস্যা হয়েছে: {e}\nদয়া করে আবার চেষ্টা করুন।")
        except Exception:
            pass

async def _dispatch_admin_callback(query, data: str, context: ContextTypes.DEFAULT_TYPE, user_id: int):

    # Root Admin Dashboard
    if data in ("adm_root", "admin_root"):
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM deposits WHERE status = 'pending'")
        pend_dep = c.fetchone()["cnt"]
        c.execute("SELECT COUNT(*) as cnt FROM orders WHERE status = 'pending_delivery'")
        pend_ord = c.fetchone()["cnt"]
        conn.close()

        text = (
            f"👑 <b>সুপার এডমিন কন্ট্রোল ড্যাশবোর্ড</b>\n"
            f"─────────────────────────────\n"
            f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
            f"🛡️ পেন্ডিং VPN অর্ডার: <b>{pend_ord}</b> টি"
        )
        kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("📦 Manage Stock", callback_data="adm_stock_menu"),
                InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_buttons_menu"),
            ],
            [
                InlineKeyboardButton("💳 Payment Methods (bKash/Nagad/Binance)", callback_data="adm_payment_methods"),
                InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_photos_menu"),
            ],
            [
                InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
                InlineKeyboardButton(f"🛡️ VPN Orders ({pend_ord})", callback_data="adm_pending_orders"),
            ],
            [
                InlineKeyboardButton("👥 User Search & Balance", callback_data="adm_user_mgmt"),
                InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
            ],
            [
                InlineKeyboardButton("⚙️ General Settings", callback_data="adm_settings"),
            ]
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    # Stock Management Menu
    if data == "adm_stock_menu":
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.id, s.name as sub_name, p.name as prod_name, COUNT(st.id) as in_stock
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        LEFT JOIN stock st ON p.id = st.product_id AND st.is_sold = 0
        WHERE p.is_on_demand = 0
        GROUP BY p.id, s.name, p.name
        """)
        rows = c.fetchall()
        conn.close()

        stock_summary = "📦 <b>বর্তমান অটো-স্টক পরিসংখ্যান:</b>\n─────────────────────────────\n"
        for r in rows:
            stock_summary += f"• <b>{r['sub_name']} ({r['prod_name']})</b> : {r['in_stock']} টি\n"
        stock_summary += "\nস্টক যোগ করতে নিচের বাটনে চাপুন:"

        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Add Stock (Excel / Text)", callback_data="adm_addstock_select")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ])
        await query.edit_message_text(stock_summary, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    # Select Product for Stock Upload (Bug Fixed)
    if data == "adm_addstock_select":
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.id, s.name as sub_name, p.name as prod_name
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        WHERE p.is_on_demand = 0
        """)
        prods = c.fetchall()
        conn.close()
        kb = []
        for p in prods:
            kb.append([InlineKeyboardButton(f"{p['sub_name']} - {p['prod_name']}", callback_data=f"adm_uploadstock_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_stock_menu")])
        await query.edit_message_text("কোন প্রোডাক্টের স্টক যোগ করবেন সিলেক্ট করুন:", reply_markup=InlineKeyboardMarkup(kb))
        return

    # Prompt Stock File / Text
    if data.startswith("adm_uploadstock_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_STOCK_FILE", {"prod_id": prod_id})
        await query.message.reply_text(
            f"📥 <b>স্টক আপলোড মোড চালু হয়েছে!</b>\n\n"
            f"• আপনি একটি এক্সেল (<code>.xlsx</code>) ফাইল পাঠাতে পারেন।\n"
            f"• অথবা চ্যাটে সরাসরি এক লাইনে একটি করে অ্যাকাউন্ট/প্রক্সি পেস্ট করে পাঠিয়ে দিন।\n\n"
            f"বট নিজে থেকে সবগুলো গুনে স্টকে ঢুকিয়ে নেবে।",
            parse_mode=ParseMode.HTML,
        )
        return

    # Photo Manager Menu
    if data == "adm_photos_menu":
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.id, s.name as sub_name, p.name as prod_name, p.image_file_id
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        """)
        prods = c.fetchall()
        conn.close()
        kb = []
        for p in prods:
            has_photo = "🖼️" if p["image_file_id"] else "⚪"
            kb.append([InlineKeyboardButton(f"{has_photo} {p['sub_name']} - {p['prod_name']}", callback_data=f"adm_photo_item_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_root")])
        await query.edit_message_text(
            "🖼️ <b>প্রোডাক্ট ফটো ম্যানেজার:</b>\nযে প্রোডাক্টের ছবি যোগ বা পরিবর্তন করতে চান, সেটি বেছে নিন (⚪ = ছবি নেই, 🖼️ = নিজস্ব ছবি আছে):",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    if data.startswith("adm_photo_item_"):
        prod_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.*, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        status = "✅ নিজস্ব ছবি আপলোড করা আছে" if p["image_file_id"] else "❌ কোনো ছবি নেই (শুধু ক্লিন টেক্সট শো করবে)"
        txt = (
            f"🖼️ <b>{p['sub_name']} — {p['name']}</b>\n\n"
            f"বর্তমান ফটো অবস্থা: {status}\n\n"
            f"নতুন ছবি দিতে চাইলে নিচের বাটনে চাপ দিন:"
        )
        kb = [
            [InlineKeyboardButton("📸 Upload / Change Photo", callback_data=f"adm_setphoto_{p['id']}")],
        ]
        if p["image_file_id"]:
            kb.append([InlineKeyboardButton("🗑️ Remove Photo", callback_data=f"adm_delphoto_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_photos_menu")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_setphoto_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PHOTO_UPLOAD", {"prod_id": prod_id})
        await query.message.reply_text("📸 অনুগ্রহ করে এই প্রোডাক্টের জন্য যে ছবিটি দিতে চান, সেটি টেলিগ্রামের চ্যাটে সেন্ড করুন:")
        return

    if data.startswith("adm_delphoto_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET image_file_id = NULL WHERE id = ?", (prod_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ ছবি সফলভাবে মুছে ফেলা হয়েছে। এখন এটি ক্লিন টেক্সটে শো করবে।")
        return

    # Structured Product & Variant Management Menu (Clean Category-wise Hierarchy)
    if data == "adm_buttons_menu":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories")
        cats = c.fetchall()
        conn.close()
        kb = []
        for cat in cats:
            kb.append([InlineKeyboardButton(f"{cat['name']}", callback_data=f"admm_cat_{cat['id']}")])
        kb.append([InlineKeyboardButton("➕ Add New Category", callback_data="admm_add_cat")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("🏷️ <b>প্রোডাক্ট ও ভ্যারিয়েন্ট ম্যানেজমেন্ট:</b>\nযে ক্যাটাগরির প্রোডাক্ট পরিচালনা করতে চান, সেটি নির্বাচন করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Category Selected -> Show Brands / Subcategories
    if data.startswith("admm_cat_"):
        cat_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ?", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        kb = []
        for s in subs:
            kb.append([InlineKeyboardButton(f"📁 {s['name']}", callback_data=f"admm_sub_{s['id']}")])
        kb.append([InlineKeyboardButton(f"➕ Add New Brand in {cat_name}", callback_data=f"admm_add_sub_{cat_id}")])
        kb.append([InlineKeyboardButton(f"✏️ Rename Category", callback_data=f"admm_ren_cat_{cat_id}"), InlineKeyboardButton(f"🗑️ Delete Category", callback_data=f"admm_del_cat_{cat_id}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_buttons_menu")])
        await query.edit_message_text(f"📁 <b>{cat_name}</b>\nনিচের ব্র্যান্ড বা সেকশন নির্বাচন করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Brand Selected -> Show Variants / Products under this Brand
    if data.startswith("admm_sub_"):
        sub_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT s.name as sub_name, s.category_id, c.name as cat_name FROM subcategories s JOIN categories c ON s.category_id = c.id WHERE s.id = ?", (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt FROM products WHERE subcategory_id = ?", (sub_id,))
        prods = c.fetchall()
        conn.close()

        sub_name = meta["sub_name"] if meta else "Section"
        kb = []
        for p in prods:
            kb.append([InlineKeyboardButton(f"📦 {p['name']} ({p['price_bdt']:.0f} Tk)", callback_data=f"admm_prod_{p['id']}")])
        kb.append([InlineKeyboardButton(f"➕ Add New Variant / Plan", callback_data=f"admm_add_prod_{sub_id}")])
        kb.append([InlineKeyboardButton(f"✏️ Rename Brand", callback_data=f"admm_ren_sub_{sub_id}"), InlineKeyboardButton(f"🗑️ Delete Brand", callback_data=f"admm_del_sub_{sub_id}")])
        back_cat = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"admm_cat_{back_cat}")])
        await query.edit_message_text(f"📦 <b>{meta['cat_name']} ➔ {sub_name}</b>\nভ্যারিয়েন্টের বিবরণ বা দাম পরিবর্তন করতে নিচের অপশন বেছে নিন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Specific Variant Selected -> View Full Details & Edit Options
    if data.startswith("admm_prod_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.*, s.name as sub_name, s.category_id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        photo_status = "🖼️ নিজস্ব ছবি সেট করা আছে" if p["image_file_id"] else "⚪ কোনো ছবি নেই (ক্লিন টেক্সট)"
        desc_txt = p["description"] if p["description"] else "কোনো কাস্টম বিবরণ দেওয়া নেই"
        txt = (
            f"📦 <b>{p['sub_name']} — {p['name']}</b>\n"
            f"─────────────────────────────\n"
            f"💰 <b>বর্তমান দাম:</b> {p['price_bdt']:.2f} BDT ({p['price_usdt']:.4f} USDT)\n"
            f"📸 <b>ফটো স্ট্যাটাস:</b> {photo_status}\n"
            f"📝 <b>কাস্টম টেক্সট/বিবরণ:</b>\n<i>{desc_txt}</i>\n\n"
            f"কী পরিবর্তন করতে চান?"
        )
        kb = [
            [InlineKeyboardButton("💰 Edit Price (BDT)", callback_data=f"adm_chgprice_{p['id']}")],
            [InlineKeyboardButton("✏️ Edit Variant Name", callback_data=f"adm_rename_{p['id']}")],
            [InlineKeyboardButton("📝 Edit Custom Text / Description", callback_data=f"admm_edit_desc_{p['id']}")],
            [InlineKeyboardButton("📸 Change / Set Photo", callback_data=f"adm_setphoto_{p['id']}")],
            [InlineKeyboardButton("🗑️ Delete Variant", callback_data=f"adm_delprod_{p['id']}")],
            [InlineKeyboardButton("< Back", callback_data=f"admm_sub_{p['subcategory_id']}")],
        ]
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Variant Price, Rename, Description, Delete actions
    if data.startswith("adm_chgprice_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PRICE", {"prod_id": prod_id})
        await query.message.reply_text("নতুন BDT দাম লিখুন (যেমন: 25):")
        return

    if data.startswith("adm_rename_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_NAME", {"prod_id": prod_id})
        await query.message.reply_text("নতুন নাম বা মেয়াদ লিখুন (যেমন: 15 Days বা 1 GB):")
        return

    if data.startswith("admm_edit_desc_"):
        prod_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_DESC", {"prod_id": prod_id})
        await query.message.reply_text("📝 এই প্রোডাক্ট কার্ডে প্রদর্শনের জন্য যে বিবরণ বা নির্দেশিকা টেক্সট দিতে চান, তা লিখে পাঠান:")
        return

    if data.startswith("adm_delprod_"):
        prod_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM products WHERE id = ?", (prod_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ ভ্যারিয়েন্টটি সফলভাবে মুছে ফেলা হয়েছে।")
        return

    # Add New Variant Prompt
    if data.startswith("admm_add_prod_"):
        sub_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_NEW_PROD_NAME", {"sub_id": sub_id})
        await query.message.reply_text("➕ <b>নতুন ভ্যারিয়েন্ট যোগ:</b>\nভ্যারিয়েন্টের নাম বা মেয়াদ লিখুন (যেমন: 10 Days বা 500 MB):", parse_mode=ParseMode.HTML)
        return

    # Add New Brand / Subcategory Prompt
    if data.startswith("admm_add_sub_"):
        cat_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_NEW_SUB_NAME", {"cat_id": cat_id})
        await query.message.reply_text("➕ <b>নতুন ব্র্যান্ড / সেকশন যোগ:</b>\nব্র্যান্ডের নাম লিখুন (যেমন: Privado VPN বা Star Proxy):", parse_mode=ParseMode.HTML)
        return

    # Add New Category Prompt
    if data == "admm_add_cat":
        set_user_state(user_id, "ADMIN_WAITING_NEW_CAT_NAME", {})
        await query.message.reply_text("➕ <b>নতুন ক্যাটাগরি তৈরি:</b>\nক্যাটাগরির নাম লিখুন (যেমন: 🎮 Gaming Services):", parse_mode=ParseMode.HTML)
        return

    # Rename & Delete Brand
    if data.startswith("admm_ren_sub_"):
        sub_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_SUB", {"sub_id": sub_id})
        await query.message.reply_text("ব্র্যান্ডের নতুন নাম লিখুন:")
        return

    if data.startswith("admm_del_sub_"):
        sub_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ ব্র্যান্ডটি সফলভাবে মুছে ফেলা হয়েছে।")
        return

    # Rename & Delete Category
    if data.startswith("admm_ren_cat_"):
        cat_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_CAT", {"cat_id": cat_id})
        await query.message.reply_text("ক্যাটাগরির নতুন নাম লিখুন:")
        return

    if data.startswith("admm_del_cat_"):
        cat_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ ক্যাটাগরিটি মুছে ফেলা হয়েছে।")
        return

    # Pending Deposits View
    if data == "adm_pending_deposits":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT d.*, u.name as uname FROM deposits d JOIN users u ON d.user_id = u.user_id WHERE d.status = 'pending'")
        deps = c.fetchall()
        conn.close()
        if not deps:
            await query.message.reply_text("✅ বর্তমানে কোনো পেন্ডিং ডিপোজিট নেই।")
            return
        for d in deps:
            kb = InlineKeyboardMarkup([
                [
                    InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{d['id']}"),
                    InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{d['id']}"),
                ]
            ])
            await query.message.reply_text(
                f"💳 <b>Deposit Request #{d['id']}</b>\n"
                f"👤 ইউজার: {d['uname']} (ID: <code>{d['user_id']}</code>)\n"
                f"💰 পরিমাণ: {d['amount']:.2f} BDT\n"
                f"📱 মেথড: {d['method']}\n"
                f"🔢 TrxID: <code>{d['trx_id']}</code>",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        return

    # Pending VPN Orders
    if data == "adm_pending_orders":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, u.name as uname FROM orders o JOIN users u ON o.user_id = u.user_id WHERE o.status = 'pending_delivery'")
        ords = c.fetchall()
        conn.close()
        if not ords:
            await query.message.reply_text("✅ বর্তমানে কোনো পেন্ডিং VPN অর্ডার নেই।")
            return
        for o in ords:
            kb = InlineKeyboardMarkup([[InlineKeyboardButton(f"🚀 Deliver Order #{o['id']}", callback_data=f"adm_fulfill_{o['id']}")]])
            await query.message.reply_text(
                f"🛡️ <b>Pending Order #{o['id']}</b>\n"
                f"🧾 Receipt: <code>{o['receipt_no']}</code>\n"
                f"👤 ইউজার: {o['uname']} (ID: <code>{o['user_id']}</code>)\n"
                f"📦 প্রোডাক্ট: {o['product_name']}\n"
                f"🔢 পরিমাণ: {o['quantity']} টি\n"
                f"💰 পরিশোধিত: {o['total_bdt']:.2f} BDT",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        return

    # Interactive User Management (List of all users without typing IDs)
    if data.startswith("adm_user_mgmt"):
        page = 0
        if "_" in data:
            parts = data.split("_")
            if len(parts) >= 4 and parts[3].isdigit():
                page = int(parts[3])

        limit = 8
        offset = page * limit
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as total FROM users")
        total_u = c.fetchone()["total"]

        c.execute("SELECT user_id, name, balance_bdt FROM users ORDER BY user_id DESC LIMIT ? OFFSET ?", (limit, offset))
        u_list = c.fetchall()
        conn.close()

        txt = (
            f"👥 <b>ইউজার ম্যানেজমেন্ট ও প্রোফাইল লিস্ট</b>\n"
            f"─────────────────────────────\n"
            f"মোট রেজিস্টার্ড ইউজার: <b>{total_u}</b> জন\n"
            f"নিচের যেকোনো ইউজারের উপর ক্লিক করে তার প্রোফাইল, কেনাকাটার হিসাব ও ব্যালেন্স দেখতে পারবেন:"
        )

        kb = []
        for u in u_list:
            u_name = u["name"][:16] if u["name"] else "User"
            btn_title = f"👤 {u_name} (ID: {u['user_id']}) - {u['balance_bdt']:.0f} Tk"
            kb.append([InlineKeyboardButton(btn_title, callback_data=f"adm_uview_{u['user_id']}")])

        # Pagination buttons
        nav_row = []
        if page > 0:
            nav_row.append(InlineKeyboardButton("◀️ Previous", callback_data=f"adm_user_mgmt_page_{page - 1}"))
        if (offset + limit) < total_u:
            nav_row.append(InlineKeyboardButton("Next ▶️", callback_data=f"adm_user_mgmt_page_{page + 1}"))
        if nav_row:
            kb.append(nav_row)

        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # View Specific User Card
    if data.startswith("adm_uview_"):
        target_uid = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE user_id = ?", (target_uid,))
        u = c.fetchone()
        c.execute("SELECT COUNT(*) as cnt, SUM(total_bdt) as spent FROM orders WHERE user_id = ? AND status = 'completed'", (target_uid,))
        stats = c.fetchone()
        c.execute("SELECT product_name, quantity, total_bdt, order_date FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 5", (target_uid,))
        recent_orders = c.fetchall()
        conn.close()

        if not u:
            await query.message.reply_text("❌ এই ইউজারকে পাওয়া যায়নি।")
            return

        spent = stats["spent"] or 0.0
        contact_url = f"tg://user?id={u['user_id']}"

        orders_txt = ""
        if recent_orders:
            orders_txt = "\n<b>সাম্প্রতিক কেনাকাটা:</b>\n"
            for ro in recent_orders:
                orders_txt += f"• {ro['product_name']} ({ro['quantity']} টি) - {ro['total_bdt']:.0f} Tk\n"

        info = (
            f"👤 <b>ইউজার প্রোফাইল:</b>\n"
            f"─────────────────────────────\n"
            f"🆔 <b>আইডি:</b> <code>{u['user_id']}</code>\n"
            f"👤 <b>নাম:</b> {u['name']}\n"
            f"💬 <b>ইউজারনেম:</b> @{u['username'] if u['username'] else 'None'}\n"
            f"💰 <b>বর্তমান ব্যালেন্স:</b> <b>{u['balance_bdt']:.2f} BDT</b>\n"
            f"🛍️ <b>মোট সফল অর্ডার:</b> {stats['cnt']} টি (মোট খরচ: {spent:.2f} BDT)\n"
            f"📅 <b>জয়েনিং তারিখ:</b> {u['joined_date']}\n"
            f"{orders_txt}"
        )
        ukb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("➕ Add Balance", callback_data=f"adm_user_addbal_{u['user_id']}"),
                InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm_user_dedbal_{u['user_id']}"),
            ],
            [
                InlineKeyboardButton("💬 Contact User (Direct Chat)", url=contact_url),
            ],
            [
                InlineKeyboardButton("< Back to Users", callback_data="adm_user_mgmt"),
            ]
        ])
        await query.edit_message_text(info, parse_mode=ParseMode.HTML, reply_markup=ukb)
        return

    # Statistics
    if data == "adm_statistics":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as total_users, SUM(balance_bdt) as total_bal FROM users")
        u_stat = c.fetchone()
        c.execute("SELECT COUNT(*) as total_orders, SUM(total_bdt) as total_rev FROM orders WHERE status = 'completed'")
        o_stat = c.fetchone()
        conn.close()

        rev = o_stat["total_rev"] or 0.0
        bal = u_stat["total_bal"] or 0.0
        stat_text = (
            f"📊 <b>শপ সার্বিক পরিসংখ্যান</b>\n"
            f"─────────────────────────────\n"
            f"👥 মোট রেজিস্টার্ড ইউজার: <b>{u_stat['total_users']}</b> জন\n"
            f"🛍️ মোট সফল অর্ডার: <b>{o_stat['total_orders']}</b> টি\n"
            f"💰 মোট বিক্রয় আয়: <b>{rev:.2f} BDT</b>\n"
            f"💳 ইউজারদের কাছে বর্তমান মোট ব্যালেন্স: <b>{bal:.2f} BDT</b>"
        )
        await query.message.reply_text(stat_text, parse_mode=ParseMode.HTML)
        return

    # Dedicated Payment Methods Editor (bKash, Nagad, Binance, USDT Rate)
    if data == "adm_payment_methods":
        bkash = get_setting("bkash_number")
        nagad = get_setting("nagad_number")
        binance = get_setting("binance_id")
        rate = get_setting("usdt_rate", "125.0")
        text = (
            f"💳 <b>পেমেন্ট মেথড কন্ট্রোল ও এডিটর:</b>\n"
            f"─────────────────────────────\n"
            f"📱 <b>bKash Personal:</b> <code>{bkash}</code>\n"
            f"📱 <b>Nagad Personal:</b> <code>{nagad}</code>\n"
            f"💎 <b>Binance Info:</b> <code>{binance}</code>\n"
            f"💵 <b>USDT Rate:</b> 1 USDT = {rate} BDT\n\n"
            f"যেটি পরিবর্তন করতে চান, নিচের বাটনটিতে চাপ দিন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✏️ Edit bKash Number", callback_data="adm_set_bkash")],
            [InlineKeyboardButton("✏️ Edit Nagad Number", callback_data="adm_set_nagad")],
            [InlineKeyboardButton("✏️ Edit Binance Pay ID / USDT", callback_data="adm_set_binance")],
            [InlineKeyboardButton("💵 Edit USDT Exchange Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    # Settings
    if data == "adm_settings":
        rate = get_setting("usdt_rate", "125.0")
        supp = get_setting("support_username")
        bkash = get_setting("bkash_number")
        nagad = get_setting("nagad_number")
        binance = get_setting("binance_id")
        text = (
            f"⚙️ <b>বট কনফিগারেশন সেটিংস:</b>\n"
            f"─────────────────────────────\n"
            f"💵 ডলার রেট: 1 USDT = {rate} BDT\n"
            f"💬 সাপোর্ট হ্যান্ডেল: {supp}\n"
            f"📱 bKash: {bkash}\n"
            f"📱 Nagad: {nagad}\n"
            f"💎 Binance: {binance}\n\n"
            f"যেকোনো ভ্যালু আপডেট করতে বাটন সিলেক্ট করুন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("💵 Change USDT Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("💬 Change Support Handle", callback_data="adm_set_support")],
            [InlineKeyboardButton("📱 Change bKash Number", callback_data="adm_set_bkash")],
            [InlineKeyboardButton("📱 Change Nagad Number", callback_data="adm_set_nagad")],
            [InlineKeyboardButton("💎 Change Binance Info", callback_data="adm_set_binance")],
            [InlineKeyboardButton("< Back", callback_data="adm_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_set_rate":
        set_user_state(user_id, "ADMIN_WAITING_RATE", {})
        await query.message.reply_text("নতুন USDT এক্সচেঞ্জ রেট লিখুন (যেমন: 128.5):")
        return

    if data == "adm_set_support":
        set_user_state(user_id, "ADMIN_WAITING_SUPPORT", {})
        await query.message.reply_text("নতুন সাপোর্ট ইউজারনেম লিখুন (যেমন: @FastSupport):")
        return

    if data == "adm_set_bkash":
        set_user_state(user_id, "ADMIN_WAITING_BKASH", {})
        await query.message.reply_text("নতুন বিকাশ পার্সোনাল নাম্বার লিখুন:")
        return

    if data == "adm_set_nagad":
        set_user_state(user_id, "ADMIN_WAITING_NAGAD", {})
        await query.message.reply_text("নতুন নগদ পার্সোনাল নাম্বার লিখুন:")
        return

    if data == "adm_set_binance":
        set_user_state(user_id, "ADMIN_WAITING_BINANCE", {})
        await query.message.reply_text("নতুন বাইনান্স পে আইডি বা USDT অ্যাড্রেস লিখুন:")
        return

    # Interactive Balance Buttons Handlers
    if data.startswith("adm_user_addbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_ADDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💰 ইউজার <code>{target_uid}</code>-এর একাউন্টে কত টাকা যোগ করতে চান লিখুন (যেমন: 100):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_user_dedbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_DEDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💸 ইউজার <code>{target_uid}</code>-এর একাউন্ট থেকে কত টাকা কাটতে চান লিখুন (যেমন: 50):", parse_mode=ParseMode.HTML)
        return

    # Deposit Approval / Rejection
    if data.startswith("adm_dep_"):
        parts = data.split("_")
        action = parts[2]
        dep_id = int(parts[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE id = ?", (dep_id,))
        dep = c.fetchone()
        if not dep:
            conn.close()
            await query.message.reply_text("❌ ডিপোজিট রিকোয়েস্ট পাওয়া যায়নি।")
            return

        if action == "approve":
            c.execute("UPDATE deposits SET status = 'approved' WHERE id = ?", (dep_id,))
            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (dep["amount"], dep["user_id"]))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"✅ ডিপোজিট #{dep_id} অনুমোদিত হয়েছে ({dep['amount']} BDT যুক্ত হয়েছে)।")
            try:
                await context.bot.send_message(
                    chat_id=dep["user_id"],
                    text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে {dep['amount']:.2f} BDT যুক্ত হয়েছে।",
                    parse_mode=ParseMode.HTML,
                )
            except Exception:
                pass
        elif action == "reject":
            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (dep_id,))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"❌ ডিপোজিট #{dep_id} বাতিল করা হয়েছে।")
            try:
                await context.bot.send_message(
                    chat_id=dep["user_id"],
                    text=f"❌ আপনার {dep['amount']:.2f} BDT ডিপোজিট রিকোয়েস্ট বাতিল করা হয়েছে।",
                    parse_mode=ParseMode.HTML,
                )
            except Exception:
                pass
        return

    # VPN Order Delivery
    if data.startswith("adm_fulfill_"):
        order_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_VPN_TEXT", {"order_id": order_id})
        await query.message.reply_text(
            f"🚀 <b>Order #{order_id} ডেলিভারি করুন:</b>\n"
            f"কাস্টমারকে পাঠানোর জন্য ভিপিএন লগইন ও কি লিখে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

# ------------------------------------------------------------------------------
# 10. TEXT, PHOTO & DOCUMENT ROUTER
# ------------------------------------------------------------------------------
async def text_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text.strip()
    state, sdata = get_user_state(user.id)

    # Main Reply Keyboard
    if text == "🛍️ Buy Products":
        await buy_products_menu_handler(update, context)
        return
    if text == "👤 My Profile":
        await profile_handler(update, context)
        return
    if text == "💳 Deposit":
        await deposit_menu_handler(update, context)
        return
    if text == "🔐 Get Code":
        await get_code_menu_handler(update, context)
        return
    if text == "💬 Support":
        await support_handler(update, context)
        return
    if text == "👑 Admin Panel":
        if is_admin(user.id):
            await admin_panel_menu_handler(update, context)
        return

    # 2FA Key Generator
    if state == "WAITING_2FA_KEY":
        secret = text.replace(" ", "").strip()
        try:
            totp = pyotp.TOTP(secret)
            code = totp.now()
            time_remaining = 30 - (int(time.time()) % 30)
            set_user_state(user.id, None)
            res_text = (
                f"🔑 <b>2FA Live Code</b>\n\n"
                f"🔢 কোড: <code>{code}</code> (ক্লিক করে কপি করুন)\n"
                f"⏳ মেয়াদ বাকি: <b>{time_remaining} সেকেন্ড</b>\n\n"
                f"📌 <i>প্রতি ৩০ সেকেন্ড পর পর কোড পরিবর্তিত হবে।</i>"
            )
            await update.message.reply_text(res_text, parse_mode=ParseMode.HTML)
        except Exception:
            await update.message.reply_text("❌ অবৈধ 2FA Secret Key! দয়া করে সঠিক কি দিন।")
        return

    # Custom Quantity
    if state == "WAITING_CUSTOM_QTY":
        try:
            qty = int(text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 15):")
            return
        prod_id = sdata.get("prod_id")
        set_user_state(user.id, None)
        class MockQuery:
            def __init__(self, msg, u):
                self.message = msg
                self.from_user = u
                self.bot = context.bot
        await process_order(MockQuery(update.message, user), prod_id, qty)
        return

    # Deposit submission
    if state == "WAITING_DEP_AMOUNT":
        try:
            amt = float(text)
            if amt <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("দয়া করে টাকার সঠিক পরিমাণ লিখুন (যেমন: 500):")
            return
        sdata["amount"] = amt
        set_user_state(user.id, "WAITING_DEP_TRX", sdata)
        await update.message.reply_text("এবার আপনার পেমেন্টের <b>Transaction ID (TrxID)</b> টি লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if state == "WAITING_DEP_TRX":
        trx = text
        method = sdata.get("method", "Manual")
        amount = sdata.get("amount", 0.0)
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")

        conn = get_db()
        c = conn.cursor()
        c.execute(
            "INSERT INTO deposits (user_id, method, amount, trx_id, status, created_at) VALUES (?, ?, ?, ?, 'pending', ?)",
            (user.id, method, amount, trx, now_str),
        )
        dep_id = c.lastrowid
        conn.commit()
        conn.close()

        set_user_state(user.id, None)
        await update.message.reply_text(
            f"✅ <b>ডিপোজিট রিকোয়েস্ট সফলভাবে জমা হয়েছে!</b>\n\n"
            f"আইডি: #{dep_id}\n"
            f"মেথড: {method}\n"
            f"পরিমাণ: {amount:.2f} BDT\n"
            f"TrxID: <code>{trx}</code>\n\n"
            f"এডমিন চেক করে ৫-১০ মিনিটের মধ্যে ব্যালেন্স যুক্ত করে দেবে।",
            parse_mode=ParseMode.HTML,
        )

        adm_kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("✅ Approve", callback_data=f"adm_dep_approve_{dep_id}"),
                InlineKeyboardButton("❌ Reject", callback_data=f"adm_dep_reject_{dep_id}"),
            ]
        ])
        try:
            await context.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>নতুন ডিপোজিট রিকোয়েস্ট!</b>\n\n"
                    f"আইডি: #{dep_id}\n"
                    f"👤 ইউজার: {user.full_name} (ID: <code>{user.id}</code>)\n"
                    f"📱 মেথড: {method}\n"
                    f"💰 পরিমাণ: {amount:.2f} BDT\n"
                    f"🔢 TrxID: <code>{trx}</code>"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
        except Exception as e:
            logger.error(f"Failed to alert admin: {e}")
        return

    # Admin Text States
    if not is_admin(user.id):
        return

    if state == "ADMIN_WAITING_STOCK_FILE":
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        prod_id = sdata.get("prod_id")
        if lines and prod_id:
            conn = get_db()
            c = conn.cursor()
            for line in lines:
                c.execute("INSERT INTO stock (product_id, content) VALUES (?, ?)", (prod_id, line))
            conn.commit()
            c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
            total_now = c.fetchone()["cnt"]
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(
                f"✅ সফলভাবে <b>{len(lines)}</b> টি স্টক যুক্ত হয়েছে!\nবর্তমানে মোট স্টক: <b>{total_now}</b> টি।",
                parse_mode=ParseMode.HTML,
            )
        return

    if state == "ADMIN_WAITING_VPN_TEXT":
        order_id = sdata.get("order_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, p.image_file_id FROM orders o JOIN products p ON o.product_id = p.id WHERE o.id = ?", (order_id,))
        order = c.fetchone()
        if order:
            c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (text, order_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ Order #{order_id} সফলভাবে ইউজারের কাছে ডেলিভারি হয়েছে।")

            user_msg = (
                f"🎉 <b>আপনার VPN সার্ভিস প্রস্তুত! (Receipt: {order['receipt_no']})</b>\n"
                f"─────────────────────────────\n"
                f"🛡️ <b>প্রোডাক্ট:</b> {order['product_name']}\n"
                f"🔢 <b>পরিমাণ:</b> {order['quantity']} টি\n\n"
                f"🔑 <b>লগইন অ্যাক্সেস / কি:</b>\n"
                f"<code>{text}</code>\n\n"
                f"📌 <i>ধন্যবাদ আমাদের সাথে থাকার জন্য!</i>"
            )
            if order["image_file_id"]:
                try:
                    await context.bot.send_photo(chat_id=order["user_id"], photo=order["image_file_id"], caption=user_msg, parse_mode=ParseMode.HTML)
                    return
                except Exception:
                    pass
            await context.bot.send_message(chat_id=order["user_id"], text=user_msg, parse_mode=ParseMode.HTML)
        else:
            conn.close()
            await update.message.reply_text("অর্ডার পাওয়া যায়নি।")
        return

    if state == "ADMIN_WAITING_PRICE":
        try:
            p_bdt = float(text)
            rate = float(get_setting("usdt_rate", "125.0"))
            p_usdt = p_bdt / rate
            prod_id = sdata.get("prod_id")
            conn = get_db()
            c = conn.cursor()
            c.execute("UPDATE products SET price_bdt = ?, price_usdt = ? WHERE id = ?", (p_bdt, p_usdt, prod_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ নতুন দাম সফলভাবে সেট হয়েছে: {p_bdt:.2f} BDT ({p_usdt:.4f} USDT)")
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
        return

    if state == "ADMIN_WAITING_NAME":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET name = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ভ্যারিয়েন্টের নাম পরিবর্তন করা হয়েছে: {text}")
        return

    # Admin State: Description edit
    if state == "ADMIN_WAITING_DESC":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET description = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ এই প্রোডাক্টটির কাস্টম টেক্সট/বিবরণ সফলভাবে সেভ করা হয়েছে!")
        return

    # Admin State: Add New Variant
    if state == "ADMIN_WAITING_NEW_PROD_NAME":
        sdata["name"] = text
        set_user_state(user.id, "ADMIN_WAITING_NEW_PROD_PRICE", sdata)
        await update.message.reply_text(f"ভ্যারিয়েন্ট: <b>{text}</b>\nএবার এর প্রতি পিসের BDT দাম কত হবে তা সংখ্যায় লিখুন (যেমন: 25):", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NEW_PROD_PRICE":
        try:
            price_bdt = float(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 25):")
            return
        sub_id = sdata.get("sub_id")
        v_name = sdata.get("name")
        rate = float(get_setting("usdt_rate", "125.0"))
        price_usdt = price_bdt / rate if rate > 0 else 0.0

        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT c.name as cat_name FROM subcategories s JOIN categories c ON s.category_id = c.id WHERE s.id = ?", (sub_id,))
        row = c.fetchone()
        is_vpn = bool(row and "VPN" in row["cat_name"])
        is_on_demand = 1 if is_vpn else 0

        c.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, ?)",
            (sub_id, v_name, price_bdt, price_usdt, is_on_demand),
        )
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"🎉 <b>নতুন ভ্যারিয়েন্ট সফলভাবে তৈরি হয়েছে!</b>\n📦 নাম: {v_name}\n💰 দাম: {price_bdt:.2f} BDT ({price_usdt:.4f} USDT)", parse_mode=ParseMode.HTML)
        return

    # Admin State: Add New Subcategory / Brand
    if state == "ADMIN_WAITING_NEW_SUB_NAME":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (cat_id, text))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ব্র্যান্ড/সেকশন <b>{text}</b> সফলভাবে তৈরি হয়েছে!", parse_mode=ParseMode.HTML)
        return

    # Admin State: Add New Category
    if state == "ADMIN_WAITING_NEW_CAT_NAME":
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO categories (name, icon) VALUES (?, '🛍️')", (text,))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ক্যাটাগরি <b>{text}</b> তৈরি সম্পন্ন হয়েছে!", parse_mode=ParseMode.HTML)
        return

    # Admin State: Rename Subcategory / Category
    if state == "ADMIN_WAITING_REN_SUB":
        sub_id = sdata.get("sub_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE subcategories SET name = ? WHERE id = ?", (text, sub_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ব্র্যান্ডের নাম পরিবর্তন হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_REN_CAT":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE categories SET name = ? WHERE id = ?", (text, cat_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ক্যাটাগরির নাম পরিবর্তন হয়েছে: {text}")
        return

    # Admin State: Interactive Balance Add / Deduct
    if state == "ADMIN_WAITING_ADDBAL_AMT":
        try:
            amt = float(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
            return
        t_id = sdata.get("target_uid")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ইউজার <code>{t_id}</code>-এর অ্যাকাউন্টে <b>{amt:.2f} BDT</b> যোগ করা হয়েছে!", parse_mode=ParseMode.HTML)
        try:
            await context.bot.send_message(chat_id=t_id, text=f"🎉 আপনার অ্যাকাউন্টে {amt:.2f} BDT এডমিন কর্তৃক যোগ করা হয়েছে!")
        except Exception:
            pass
        return

    if state == "ADMIN_WAITING_DEDBAL_AMT":
        try:
            amt = float(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
            return
        t_id = sdata.get("target_uid")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = MAX(0, balance_bdt - ?) WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ইউজার <code>{t_id}</code>-এর অ্যাকাউন্ট থেকে <b>{amt:.2f} BDT</b> কেটে নেওয়া হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_RATE":
        try:
            new_rate = float(text)
            set_setting("usdt_rate", str(new_rate))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ নতুন ডলার রেট সেট হয়েছে: 1 USDT = {new_rate} BDT")
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন:")
        return

    if state == "ADMIN_WAITING_SUPPORT":
        set_setting("support_username", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন সাপোর্ট ইউজারনেম সেট হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_BKASH":
        set_setting("bkash_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন বিকাশ নাম্বার সেট হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_NAGAD":
        set_setting("nagad_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন নগদ নাম্বার সেট হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_BINANCE":
        set_setting("binance_id", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন বাইনান্স পেমেন্ট তথ্য সেট হয়েছে: {text}")
        return

    if state == "ADMIN_WAITING_USER_SEARCH":
        try:
            target_id = int(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক নিউমেরিক আইডি লিখুন:")
            return
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE user_id = ?", (target_id,))
        u = c.fetchone()
        c.execute("SELECT COUNT(*) as cnt, SUM(total_bdt) as spent FROM orders WHERE user_id = ? AND status = 'completed'", (target_id,))
        stats = c.fetchone()
        conn.close()
        set_user_state(user.id, None)

        if not u:
            await update.message.reply_text("❌ এই আইডির কোনো ইউজার পাওয়া যায়নি।")
            return

        spent = stats["spent"] or 0.0
        contact_url = f"https://t.me/{u['username']}" if u['username'] else f"tg://user?id={u['user_id']}"
        info = (
            f"👤 <b>ইউজার প্রোফাইল বিবরণ:</b>\n"
            f"─────────────────────────────\n"
            f"🆔 <b>আইডি:</b> <code>{u['user_id']}</code>\n"
            f"👤 <b>নাম:</b> {u['name']}\n"
            f"💬 <b>ইউজারনেম:</b> @{u['username'] if u['username'] else 'None'}\n"
            f"💰 <b>বর্তমান ব্যালেন্স:</b> <b>{u['balance_bdt']:.2f} BDT</b>\n"
            f"🛍️ <b>মোট কেনাকাটা:</b> {stats['cnt']} টি ({spent:.2f} BDT)\n"
            f"📅 <b>জয়েনিং তারিখ:</b> {u['joined_date']}\n"
        )
        ukb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("➕ Add Balance", callback_data=f"adm_user_addbal_{u['user_id']}"),
                InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm_user_dedbal_{u['user_id']}"),
            ],
            [
                InlineKeyboardButton("💬 Contact User (Direct Chat)", url=contact_url),
            ],
            [
                InlineKeyboardButton("< Back", callback_data="adm_root"),
            ]
        ])
        await update.message.reply_text(info, parse_mode=ParseMode.HTML, reply_markup=ukb)
        return

async def photo_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles photo uploads by Admin for assigning custom product images."""
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_PHOTO_UPLOAD":
        photo_file_id = update.message.photo[-1].file_id
        prod_id = sdata.get("prod_id")

        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET image_file_id = ? WHERE id = ?", (photo_file_id, prod_id))
        conn.commit()
        conn.close()

        set_user_state(user.id, None)
        await update.message.reply_text("✅ এই প্রোডাক্টটির জন্য নতুন ফটো সফলভাবে সেভ করা হয়েছে! কাস্টমার এখন এই ছবিটি দেখতে পাবে।")

async def document_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles Excel (.xlsx) file uploads by Admin for bulk stock upload."""
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_STOCK_FILE":
        doc = update.message.document
        if not doc.file_name.endswith(".xlsx"):
            await update.message.reply_text("❌ দয়া করে একটি সঠিক <code>.xlsx</code> এক্সেল ফাইল আপলোড করুন।", parse_mode=ParseMode.HTML)
            return

        tg_file = await context.bot.get_file(doc.file_id)
        file_bytes = await tg_file.download_as_bytearray()
        try:
            items = read_items_from_excel(file_bytes)
        except Exception as e:
            await update.message.reply_text(f"❌ এক্সেল ফাইল পড়তে সমস্যা হয়েছে: {e}")
            return

        prod_id = sdata.get("prod_id")
        if not items:
            await update.message.reply_text("❌ এক্সেল ফাইলে কোনো স্টক ডেটা পাওয়া যায়নি।")
            return

        conn = get_db()
        c = conn.cursor()
        for item in items:
            c.execute("INSERT INTO stock (product_id, content) VALUES (?, ?)", (prod_id, item))
        conn.commit()
        c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
        total_now = c.fetchone()["cnt"]
        conn.close()

        set_user_state(user.id, None)
        await update.message.reply_text(
            f"🎉 <b>এক্সেল স্টক আপলোড সম্পন্ন!</b>\n\n"
            f"📊 নতুন যুক্ত হয়েছে: <b>{len(items)}</b> টি অ্যাকাউন্ট\n"
            f"📦 স্টকে বর্তমান মোট সংখ্যা: <b>{total_now}</b> টি।",
            parse_mode=ParseMode.HTML,
        )

# Admin balance commands
async def add_balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার: <code>/addbalance <user_id> <amount></code>", parse_mode=ParseMode.HTML)
        return
    try:
        t_id = int(context.args[0])
        amt = float(context.args[1])
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        await update.message.reply_text(f"✅ {t_id} আইডিতে {amt:.2f} BDT যোগ করা হয়েছে।")
        try:
            await context.bot.send_message(chat_id=t_id, text=f"🎉 আপনার অ্যাকাউন্টে {amt:.2f} BDT এডমিন কর্তৃক যোগ করা হয়েছে!")
        except Exception:
            pass
    except Exception as e:
        await update.message.reply_text(f"ভুল হয়েছে: {e}")

async def deduct_balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        return
    if len(context.args) < 2:
        await update.message.reply_text("ব্যবহার: <code>/deductbalance <user_id> <amount></code>", parse_mode=ParseMode.HTML)
        return
    try:
        t_id = int(context.args[0])
        amt = float(context.args[1])
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE users SET balance_bdt = MAX(0, balance_bdt - ?) WHERE user_id = ?", (amt, t_id))
        conn.commit()
        conn.close()
        await update.message.reply_text(f"✅ {t_id} আইডি থেকে {amt:.2f} BDT কেটে নেওয়া হয়েছে।")
    except Exception as e:
        await update.message.reply_text(f"ভুল হয়েছে: {e}")

# ------------------------------------------------------------------------------
# 11. APPLICATION INITIALIZATION & RUNNER
# ------------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("Starting Render Web Service Health Check Server...")
    threading.Thread(target=start_health_server, daemon=True).start()
    print("Initializing Database...")
    init_db()
    print("Database Initialized Successfully.")
    print("Starting Telegram Bot Application...")
    print(f"Configured Super Admin ID: {SUPER_ADMIN_ID}")
    print("=" * 60)

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    # Handlers Registration
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("addbalance", add_balance_command))
    app.add_handler(CommandHandler("deductbalance", deduct_balance_command))

    app.add_handler(CallbackQueryHandler(callback_query_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message_handler))
    app.add_handler(MessageHandler(filters.PHOTO, photo_message_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_message_handler))

    print("Bot is polling for updates...")
    app.run_polling()

if __name__ == "__main__":
    main()
