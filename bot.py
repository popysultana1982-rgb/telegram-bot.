# ==============================================================================
# Telegram Digital Store Bot (Auto-Shop & Request Engine)
# Architecture: Python (python-telegram-bot v21+), SQLite, OpenPyXL, PyOTP
#
# Configured Credentials:
# - BOT_TOKEN: 8656707191:AAEvGFT1_x3C3Tdpd2uV-o38atxxsX7HY54
# - SUPER_ADMIN_ID: 8919985167
# ==============================================================================

import os
import sys
import io
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import time
import json
import logging
import sqlite3
import datetime
from typing import Optional, List, Dict, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pyotp

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
logger = logging.getLogger("DigitalStoreBot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "8656707191:AAEqK_Dn6M_OhFP12ATl6g5WyNCY02Db_sk")
SUPER_ADMIN_ID = int(os.getenv("SUPER_ADMIN_ID", "8919985167"))
DB_FILE = "store_database.db"

# High-resolution placeholder/official brand banners for rich cards
BRAND_BANNERS = {
    "default": "https://images.unsplash.com/photo-1563986768609-322da13575f3?w=800&auto=format&fit=crop&q=80",
    "vpn": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=800&auto=format&fit=crop&q=80",
    "express": "https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=800&auto=format&fit=crop&q=80",
    "nord": "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=800&auto=format&fit=crop&q=80",
    "ipvanish": "https://images.unsplash.com/photo-1614064641938-3bbee52942c7?w=800&auto=format&fit=crop&q=80",
    "cyberghost": "https://images.unsplash.com/photo-1550751827-4bd374c3f58b?w=800&auto=format&fit=crop&q=80",
    "mail": "https://images.unsplash.com/photo-1557200134-90327ee9fafa?w=800&auto=format&fit=crop&q=80",
    "proxy": "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=800&auto=format&fit=crop&q=80",
}

# ------------------------------------------------------------------------------
# 2. DATABASE LAYER (SQLite)
# ------------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    # Users Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY,
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

    # Subcategories (Durations / Packages)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS subcategories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id INTEGER,
        name TEXT NOT NULL,
        FOREIGN KEY (category_id) REFERENCES categories (id) ON DELETE CASCADE
    )
    """)

    # Products Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subcategory_id INTEGER,
        name TEXT NOT NULL,
        price_bdt REAL NOT NULL,
        price_usdt REAL NOT NULL,
        image_url TEXT,
        is_on_demand INTEGER DEFAULT 0,
        description TEXT,
        FOREIGN KEY (subcategory_id) REFERENCES subcategories (id) ON DELETE CASCADE
    )
    """)

    # Stock Items Table (For Auto-Delivery like Mail, Proxy)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        content TEXT NOT NULL,
        is_sold INTEGER DEFAULT 0,
        sold_to INTEGER,
        sold_at TEXT,
        FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
    )
    """)

    # Orders Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        product_id INTEGER,
        product_name TEXT,
        quantity INTEGER,
        total_bdt REAL,
        total_usdt REAL,
        status TEXT DEFAULT 'completed',
        delivered_data TEXT,
        order_date TEXT,
        FOREIGN KEY (user_id) REFERENCES users (user_id)
    )
    """)

    # Deposits Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS deposits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
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

    # Seed Default Settings if missing
    default_settings = {
        "shop_name": "CyberNest Digital Mart",
        "usdt_rate": "125.0",
        "support_username": "@YourSupportHandle",
        "bkash_number": "01XXXXXXXXX (Personal - Send Money)",
        "nagad_number": "01XXXXXXXXX (Personal - Send Money)",
        "rocket_number": "01XXXXXXXXX (Personal - Send Money)",
        "crypto_address": "TQxxxxxxxxx (USDT TRC20 / Binance Pay)",
        "mail_code_url_1": "https://email-otp-viewer.example.com",
        "mail_code_url_2": "https://outlook-code.example.com",
    }
    for k, v in default_settings.items():
        cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
    conn.commit()

    # Seed Default Categories, Subcategories, and Products
    cursor.execute("SELECT COUNT(*) as cnt FROM categories")
    if cursor.fetchone()["cnt"] == 0:
        seed_default_store(conn)

    conn.close()

def seed_default_store(conn):
    cursor = conn.cursor()
    # 1. Categories
    cats = [
        ("🛡️ VPN", "🛡️"),
        ("🌐 Proxy", "🌐"),
        ("✉️ Mail", "✉️"),
        ("✉️ Other Domain Mail", "✉️"),
        ("✉️ Recovery Added Mail", "✉️"),
    ]
    for cname, cicon in cats:
        cursor.execute("INSERT INTO categories (name, icon) VALUES (?, ?)", (cname, cicon))
    conn.commit()

    # Get category IDs
    cursor.execute("SELECT id, name FROM categories")
    cat_map = {row["name"]: row["id"] for row in cursor.fetchall()}

    # 2. Subcategories for VPN
    vpn_id = cat_map.get("🛡️ VPN")
    vpn_subs = ["3 Days", "7 Days", "14 Days", "30 Days"]
    sub_map = {}
    for sub in vpn_subs:
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (vpn_id, sub))
        sub_map[f"vpn_{sub}"] = cursor.lastrowid
    conn.commit()

    # VPN Products
    vpn_plans_3d = [
        ("IP Vanish", 15.0, 0.1200, BRAND_BANNERS["ipvanish"]),
        ("Cyberghost VPN", 15.0, 0.1200, BRAND_BANNERS["cyberghost"]),
        ("Panda VPN", 15.0, 0.1200, BRAND_BANNERS["vpn"]),
        ("Express VPN", 15.0, 0.1200, BRAND_BANNERS["express"]),
        ("Nord VPN", 18.0, 0.1440, BRAND_BANNERS["nord"]),
        ("Surfshark VPN", 18.0, 0.1440, BRAND_BANNERS["vpn"]),
    ]
    for name, bdt, usdt, img in vpn_plans_3d:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 1)",
            (sub_map["vpn_3 Days"], name, bdt, usdt, img),
        )

    vpn_plans_7d = [
        ("Bitdefender VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Surfshark VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Nord VPN", 20.0, 0.1600, BRAND_BANNERS["nord"]),
        ("Express VPN", 20.0, 0.1600, BRAND_BANNERS["express"]),
        ("Hotspot VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("HMA VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("IPVanish VPN", 20.0, 0.1600, BRAND_BANNERS["ipvanish"]),
        ("Vypr VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Cyberghost VPN", 20.0, 0.1600, BRAND_BANNERS["cyberghost"]),
        ("Norton VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("PIA VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Panda VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Potato VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Avast VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
        ("Adguard VPN", 20.0, 0.1600, BRAND_BANNERS["vpn"]),
    ]
    for name, bdt, usdt, img in vpn_plans_7d:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 1)",
            (sub_map["vpn_7 Days"], name, bdt, usdt, img),
        )

    # 14 & 30 Days clones
    for name, bdt, usdt, img in vpn_plans_7d[:6]:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 1)",
            (sub_map["vpn_14 Days"], name, bdt * 1.8, usdt * 1.8, img),
        )
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 1)",
            (sub_map["vpn_30 Days"], name, bdt * 3.2, usdt * 3.2, img),
        )

    # 3. Subcategories for Proxy
    proxy_id = cat_map.get("🌐 Proxy")
    proxy_subs = ["Standard Plans"]
    for psub in proxy_subs:
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (proxy_id, psub))
        psub_id = cursor.lastrowid
        proxies = [
            ("9 Proxy (100 MB)", 12.0, 0.096, BRAND_BANNERS["proxy"]),
            ("9 Proxy (200 MB)", 23.0, 0.184, BRAND_BANNERS["proxy"]),
            ("9 Proxy (500 MB)", 60.0, 0.480, BRAND_BANNERS["proxy"]),
            ("9 Proxy (1 GB)", 110.0, 0.880, BRAND_BANNERS["proxy"]),
            ("Rapid Proxy (100 MB)", 12.0, 0.096, BRAND_BANNERS["proxy"]),
            ("Rapid Proxy (200 MB)", 25.0, 0.200, BRAND_BANNERS["proxy"]),
            ("Rapid Proxy (500 MB)", 60.0, 0.480, BRAND_BANNERS["proxy"]),
            ("Rapid Proxy (1 GB)", 115.0, 0.920, BRAND_BANNERS["proxy"]),
            ("711 Proxy (100 MB)", 10.0, 0.080, BRAND_BANNERS["proxy"]),
            ("711 Proxy (200 MB)", 18.0, 0.144, BRAND_BANNERS["proxy"]),
            ("711 Proxy (500 MB)", 45.0, 0.360, BRAND_BANNERS["proxy"]),
        ]
        for name, bdt, usdt, img in proxies:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 0)",
                (psub_id, name, bdt, usdt, img),
            )

    # 4. Mail Categories
    mail_id = cat_map.get("✉️ Mail")
    cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (mail_id, "General Mail"))
    mid = cursor.lastrowid
    for name, bdt, usdt in [("FR Outlook", 12.0, 0.096), ("Hotmail", 10.0, 0.080), ("Outlook", 10.0, 0.080)]:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 0)",
            (mid, name, bdt, usdt, BRAND_BANNERS["mail"]),
        )

    # 5. Recovery Added Mail
    rec_mail_id = cat_map.get("✉️ Recovery Added Mail")
    cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (rec_mail_id, "Recovery Mail"))
    rmid = cursor.lastrowid
    for name, bdt, usdt in [("Recovery Added FR Outlook", 15.0, 0.120), ("Recovery Added Hotmail", 12.0, 0.096), ("Recovery Added Outlook", 12.0, 0.096)]:
        cursor.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, image_url, is_on_demand) VALUES (?, ?, ?, ?, ?, 0)",
            (rmid, name, bdt, usdt, BRAND_BANNERS["mail"]),
        )

    conn.commit()

# Helper DB queries
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
        is_admin = 1 if user_id == SUPER_ADMIN_ID else 0
        c.execute(
            "INSERT INTO users (user_id, name, username, balance_bdt, joined_date, is_admin) VALUES (?, ?, ?, ?, ?, ?)",
            (user_id, name, username or "", 0.0, joined, is_admin),
        )
        conn.commit()
        c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
        row = c.fetchone()
    else:
        # Keep super admin always 1
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
# 3. KEYBOARD GENERATORS (Dynamic Menus)
# ------------------------------------------------------------------------------
def get_main_reply_keyboard(user_id: int) -> ReplyKeyboardMarkup:
    """Returns the persistent bottom reply keyboard.
    If the user is Super Admin or Admin, includes the '👑 Admin Panel' button."""
    buttons = [
        [KeyboardButton("🛍️ Buy Products")],
        [KeyboardButton("👤 My Profile"), KeyboardButton("💳 Deposit")],
        [KeyboardButton("🔐 Get Code"), KeyboardButton("💬 Support")],
    ]
    if is_admin(user_id):
        buttons.append([KeyboardButton("👑 Admin Panel")])
    return ReplyKeyboardMarkup(buttons, resize_keyboard=True)

# ------------------------------------------------------------------------------
# 4. EXCEL WORKBOOK ENGINE (OpenPyXL for Mail Delivery & Ingestion)
# ------------------------------------------------------------------------------
def generate_excel_delivery(items: List[str], product_name: str, order_id: int) -> io.BytesIO:
    """Generates an aesthetic .xlsx file containing the purchased credentials."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Delivered Credentials"

    # Styling definitions
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Headers
    headers = ["SL", "Product Name", "Account / Credential / Session Data", "Delivered At"]
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

    # Auto-adjust column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 70
    ws.column_dimensions['D'].width = 22

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

def read_items_from_excel(file_bytes: bytes) -> List[str]:
    """Extracts all non-empty account lines from an uploaded .xlsx file."""
    items = []
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if cell is not None:
                val = str(cell).strip()
                # Skip common header strings
                if val.lower() in ("sl", "no", "mail", "email", "password", "account", "credential", "product name", "delivered at"):
                    continue
                if len(val) > 2:
                    items.append(val)
    return items

# ------------------------------------------------------------------------------
# 5. USER HANDLERS (Shop, Profile, Deposit, 2FA, Orders)
# ------------------------------------------------------------------------------
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    shop_name = get_setting("shop_name", "CyberNest Digital Mart")

    welcome_text = (
        f"🎉 <b>স্বাগতম, {user.full_name}!</b> ✨\n\n"
        f"💎 <b>{shop_name}</b>\n"
        f"─────────────────────────────\n"
        f"🛍️ প্রিমিয়াম VPN, Proxy আর Mail সার্ভিস — সবচেয়ে দ্রুত ডেলিভারিতে, ২৪/৭।\n\n"
        f"📌 নিচের মেনু থেকে বেছে নিন 👇"
    )
    reply_kb = get_main_reply_keyboard(user.id)
    await update.message.reply_text(welcome_text, parse_mode=ParseMode.HTML, reply_markup=reply_kb)

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
    # Two categories per row
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
    url1 = get_setting("mail_code_url_1", "https://example.com")
    url2 = get_setting("mail_code_url_2", "https://example.com")

    text = (
        "🔐 <b>Get Code & Authentication Tools</b>\n\n"
        "হটমেইল/আউটলুক ওটিপি কোড এবং লাইভ 2FA কোড পাওয়ার জন্য নিচের অপশন থেকে নির্বাচন করুন:"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🔑 2FA Authenticator (Live Code)", callback_data="tool_2fa_prompt")],
        [InlineKeyboardButton("🌐 Hotmail/Outlook Code Server 1", url=url1)],
        [InlineKeyboardButton("🌐 Hotmail/Outlook Code Server 2", url=url2)],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def support_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    supp = get_setting("support_username", "@YourSupportHandle")
    text = (
        f"💬 <b>কাস্টমার সাপোর্ট</b>\n\n"
        f"যেকোনো সমস্যা, প্রোডাক্ট রিকোয়েস্ট বা ডিপোজিটের সহায়তায় সরাসরি যোগাযোগ করুন:\n"
        f"👉 <b>{supp}</b>\n\n"
        f"আমরা ২৪/৭ সক্রিয় আছি।"
    )
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("📨 Contact Support", url=f"https://t.me/{supp.lstrip('@')}") ]])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def deposit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    bkash = get_setting("bkash_number")
    nagad = get_setting("nagad_number")
    rocket = get_setting("rocket_number")
    crypto = get_setting("crypto_address")
    rate = get_setting("usdt_rate", "125.0")

    text = (
        f"💳 <b>অ্যাড ব্যালেন্স / ডিপোজিট</b>\n"
        f"─────────────────────────────\n"
        f"💰 <b>বর্তমান ব্যালেন্স:</b> {db_user['balance_bdt']:.2f} BDT\n"
        f"💵 <b>ডলার রেট:</b> 1 USDT = {rate} BDT\n\n"
        f"<b>টাকা পাঠানোর নাম্বারসমূহ (Send Money):</b>\n"
        f"📱 <b>bKash:</b> <code>{bkash}</code>\n"
        f"📱 <b>Nagad:</b> <code>{nagad}</code>\n"
        f"📱 <b>Rocket:</b> <code>{rocket}</code>\n"
        f"💎 <b>Crypto:</b> <code>{crypto}</code>\n\n"
        f"⚠️ <i>টাকা পাঠানোর পর নিচের বাটনে চাপ দিয়ে TrxID এবং টাকার পরিমাণ জমা দিন। এডমিন চেক করে ব্যালেন্স যুক্ত করে দেবে।</i>"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📱 Submit Deposit Proof", callback_data="dep_submit_start")],
        [InlineKeyboardButton("< Back", callback_data="menu_home")]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

# ------------------------------------------------------------------------------
# 6. INLINE CALLBACK QUERY DISPATCHER
# ------------------------------------------------------------------------------
async def callback_query_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    db_user = get_or_create_user(user_id, query.from_user.full_name, query.from_user.username)

    # Category Click -> Show Subcategories
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
        text = f"<b>{cat_name} — Select Duration / Plan:</b>"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(f"📜 {s['name']}", callback_data=f"sub_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="buy_menu_root")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Back to Root Categories
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

    # Subcategory Click -> Show Products List with Prices
    if data.startswith("sub_"):
        sub_id = int(data.split("_")[1])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT s.name as sub_name, s.category_id, c.name as cat_name FROM subcategories s JOIN categories c ON s.category_id = c.id WHERE s.id = ?", (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt, price_usdt FROM products WHERE subcategory_id = ?", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} — {meta['sub_name']}" if meta else "Select Plan"
        text = f"🛡️ <b>{title}</b>\n\nSelect a plan:"
        kb = []
        for p in prods:
            btn_text = f"{p['name']} • {p['price_bdt']:.2f} BDT / {p['price_usdt']:.4f} USDT"
            kb.append([InlineKeyboardButton(btn_text, callback_data=f"prod_{p['id']}")])
        back_cat_id = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"cat_{back_cat_id}")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Product Click -> Show Rich Product Card & Quantity Selector
    if data.startswith("prod_"):
        prod_id = int(data.split("_")[1])
        await show_product_card(query, prod_id)
        return

    # Quantity Selection (1x to 10x)
    if data.startswith("qty_"):
        parts = data.split("_")
        prod_id = int(parts[1])
        qty = int(parts[2])
        await process_order(query, prod_id, qty)
        return

    # Custom Quantity Prompt
    if data.startswith("custqty_"):
        prod_id = int(data.split("_")[1])
        set_user_state(user_id, "WAITING_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text(
            "✏️ <b>Custom Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):",
            parse_mode=ParseMode.HTML,
        )
        return

    # 2FA Prompt
    if data == "tool_2fa_prompt":
        set_user_state(user_id, "WAITING_2FA_KEY", {})
        await query.message.reply_text(
            "🔑 <b>2FA Authenticator:</b>\nআপনার 2FA Secret Key টি এখানে পেস্ট করে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    # Profile sub-actions
    if data == "prof_purchases":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        orders = c.fetchall()
        conn.close()
        if not orders:
            await query.message.reply_text("🛒 আপনার পূর্বের কোনো কেনাকাটার হিস্টোরি পাওয়া যায়নি।")
            return
        msg = "🛒 <b>আপনার সাম্প্রতিক অর্ডারসমূহ:</b>\n─────────────────────────────\n"
        for o in orders:
            msg += (
                f"📦 <b>Order #{o['id']}</b> | {o['product_name']}\n"
                f"🔢 পরিমাণ: {o['quantity']} টি | মোট: {o['total_bdt']:.2f} BDT\n"
                f"📅 তারিখ: {o['order_date']}\n"
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
            f"আপনার বন্ধুদের আমন্ত্রণ জানান এবং বোনাস রিওয়ার্ড উপভোগ করুন!\n\n"
            f"🔗 আপনার রেফার লিংক:\n<code>{ref_link}</code>"
        )
        await query.message.reply_text(msg, parse_mode=ParseMode.HTML)
        return

    # Deposit submission start
    if data == "dep_submit_start":
        set_user_state(user_id, "WAITING_DEP_METHOD", {})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("bKash", callback_data="dep_method_bKash"), InlineKeyboardButton("Nagad", callback_data="dep_method_Nagad")],
            [InlineKeyboardButton("Rocket", callback_data="dep_method_Rocket"), InlineKeyboardButton("Crypto (USDT)", callback_data="dep_method_Crypto")],
        ])
        await query.message.reply_text("কোন মাধ্যমে টাকা পাঠিয়েছেন সিলেক্ট করুন:", reply_markup=kb)
        return

    if data.startswith("dep_method_"):
        method = data.split("_")[2]
        set_user_state(user_id, "WAITING_DEP_AMOUNT", {"method": method})
        await query.message.reply_text(f"✅ আপনি <b>{method}</b> নির্বাচন করেছেন।\nকত টাকা পাঠিয়েছেন তা লিখুন (যেমন: 500):", parse_mode=ParseMode.HTML)
        return

    # Admin Deposit Approval / Rejection
    if data.startswith("adm_dep_"):
        if not is_admin(user_id):
            return
        parts = data.split("_")
        action = parts[2]
        dep_id = int(parts[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM deposits WHERE id = ?", (dep_id,))
        dep = c.fetchone()
        if not dep:
            conn.close()
            await query.message.reply_text("❌ ডিপোজিট রিকোয়েস্টটি পাওয়া যায়নি।")
            return

        if action == "approve":
            c.execute("UPDATE deposits SET status = 'approved' WHERE id = ?", (dep_id,))
            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (dep["amount"], dep["user_id"]))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"✅ ডিপোজিট #{dep_id} অনুমোদিত হয়েছে ({dep['amount']} BDT যোগ করা হয়েছে)।")
            # Notify User
            try:
                await context.bot.send_message(
                    chat_id=dep["user_id"],
                    text=f"🎉 <b>ডিপোজিট সফল!</b>\nআপনার অ্যাকাউন্টে {dep['amount']:.2f} BDT যুক্ত করা হয়েছে।",
                    parse_mode=ParseMode.HTML,
                )
            except Exception as e:
                logger.error(f"Error notifying user: {e}")
        elif action == "reject":
            c.execute("UPDATE deposits SET status = 'rejected' WHERE id = ?", (dep_id,))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"❌ ডিপোজিট #{dep_id} বাতিল করা হয়েছে।")
            try:
                await context.bot.send_message(
                    chat_id=dep["user_id"],
                    text=f"❌ আপনার {dep['amount']:.2f} BDT ডিপোজিট রিকোয়েস্টটি বাতিল করা হয়েছে। বিস্তারিত জানতে সাপোর্টে কথা বলুন।",
                    parse_mode=ParseMode.HTML,
                )
            except Exception as e:
                logger.error(f"Error notifying user: {e}")
        return

    # Admin VPN Order Fulfillment
    if data.startswith("adm_fulfill_"):
        if not is_admin(user_id):
            return
        order_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_VPN_TEXT", {"order_id": order_id})
        await query.message.reply_text(
            f"🚀 <b>Order #{order_id} ডেলিভারি করুন:</b>\n"
            f"কাস্টমারকে ডেলিভারি দেওয়ার জন্য VPN এর ইউজারনেম, পাসওয়ার্ড বা কি লিখে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    # Admin Panel Sub-menus (Stock, Products, Broadcast, Stats)
    if data.startswith("admin_"):
        if not is_admin(user_id):
            return
        await handle_admin_callbacks(query, data, context)
        return

# ------------------------------------------------------------------------------
# 7. PRODUCT CARD & PURCHASE ENGINE
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

    # Check available stock
    c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
    avail_stock = c.fetchone()["cnt"]
    conn.close()

    if not p:
        await query.message.reply_text("প্রোডাক্টটি পাওয়া যায়নি।")
        return

    stock_badge = "⚡ Instant Request Delivery (২৪/৭ ডেলিভারি)" if p["is_on_demand"] else f"{avail_stock} টি স্টকে আছে"

    caption = (
        f"🛡️ <b>{p['name']}</b>\n"
        f"🎁 <b>Plan / Duration:</b> {p['sub_name']}\n"
        f"💰 <b>Price:</b> {p['price_bdt']:.2f} BDT / {p['price_usdt']:.4f} USDT per unit\n"
        f"📊 <b>Stock:</b> {stock_badge}\n\n"
        f"Select quantity:"
    )

    # Quantity Buttons
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

    image_url = p["image_url"] or BRAND_BANNERS["default"]
    try:
        await query.message.reply_photo(
            photo=image_url,
            caption=caption,
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
    except Exception as e:
        logger.warning(f"Failed to send photo: {e}, falling back to text.")
        await query.message.reply_text(caption, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))

async def process_order(query, prod_id: int, quantity: int):
    user_id = query.from_user.id
    user_name = query.from_user.full_name
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM products WHERE id = ?", (prod_id,))
    prod = c.fetchone()
    c.execute("SELECT balance_bdt FROM users WHERE user_id = ?", (user_id,))
    user_balance = c.fetchone()["balance_bdt"]

    if not prod:
        conn.close()
        await query.message.reply_text("প্রোডাক্ট পাওয়া যায়নি।")
        return

    total_bdt = prod["price_bdt"] * quantity
    total_usdt = prod["price_usdt"] * quantity

    # 1. Balance Check
    if user_balance < total_bdt:
        conn.close()
        shortage = total_bdt - user_balance
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("💳 Deposit Now", callback_data="dep_submit_start")]])
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

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")

    # 2. On-Demand Products (VPN)
    if prod["is_on_demand"]:
        c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
        c.execute(
            "INSERT INTO orders (user_id, product_id, product_name, quantity, total_bdt, total_usdt, status, order_date) VALUES (?, ?, ?, ?, ?, ?, 'pending_delivery', ?)",
            (user_id, prod_id, prod["name"], quantity, total_bdt, total_usdt, now_str),
        )
        order_id = c.lastrowid
        conn.commit()
        conn.close()

        # Confirmation to User
        await query.message.reply_text(
            f"✅ <b>অর্ডার কনফার্ম হয়েছে! (Order #{order_id})</b>\n"
            f"🛡️ <b>প্রোডাক্ট:</b> {prod['name']}\n"
            f"🔢 <b>পরিমাণ:</b> {quantity} টি\n"
            f"💰 <b>কেটে নেওয়া হয়েছে:</b> {total_bdt:.2f} BDT\n\n"
            f"⏳ আপনার VPN খুব শীঘ্রই প্রস্তুত করে এডমিন ইনবক্সে পাঠিয়ে দেবে। সাথে থাকার জন্য ধন্যবাদ!",
            parse_mode=ParseMode.HTML,
        )

        # Alert Admin
        adm_kb = InlineKeyboardMarkup([[InlineKeyboardButton(f"🚀 Deliver Order #{order_id}", callback_data=f"adm_fulfill_{order_id}")]])
        try:
            await query.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>নতুন VPN অর্ডার এসেছে!</b>\n"
                    f"📦 <b>Order ID:</b> #{order_id}\n"
                    f"👤 <b>কাস্টমার:</b> {user_name} (ID: <code>{user_id}</code>)\n"
                    f"🛡️ <b>প্রোডাক্ট:</b> {prod['name']}\n"
                    f"🔢 <b>পরিমাণ:</b> {quantity} টি\n"
                    f"💰 <b>পেমেন্ট:</b> {total_bdt:.2f} BDT"
                ),
                parse_mode=ParseMode.HTML,
                reply_markup=adm_kb,
            )
        except Exception as e:
            logger.error(f"Failed to alert admin: {e}")
        return

    # 3. Instant Stock Products (Mail, Proxy)
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

    # Save Order Record
    delivered_json = json.dumps(items_content)
    c.execute(
        "INSERT INTO orders (user_id, product_id, product_name, quantity, total_bdt, total_usdt, status, delivered_data, order_date) VALUES (?, ?, ?, ?, ?, ?, 'completed', ?, ?)",
        (user_id, prod_id, prod["name"], quantity, total_bdt, total_usdt, delivered_json, now_str),
    )
    order_id = c.lastrowid
    conn.commit()
    conn.close()

    # Deliver via Excel (.xlsx) file
    excel_bio = generate_excel_delivery(items_content, prod["name"], order_id)
    doc_caption = (
        f"🎉 <b>আপনার অর্ডার সফলভাবে ডেলিভারি হয়েছে! (Order #{order_id})</b>\n\n"
        f"📦 <b>প্রোডাক্ট:</b> {prod['name']}\n"
        f"🔢 <b>পরিমাণ:</b> {quantity} টি\n"
        f"💰 <b>মোট কাটা হয়েছে:</b> {total_bdt:.2f} BDT\n\n"
        f"📂 আপনার ডেলিভারি ডেটা সম্বলিত এক্সেল ফাইলটি উপরে সংযুক্ত করা হয়েছে। ডাউনলোড করে নিন।"
    )
    filename = f"Order_{order_id}_{prod['name'].replace(' ', '_')}.xlsx"
    await query.message.reply_document(
        document=InputFile(excel_bio, filename=filename),
        caption=doc_caption,
        parse_mode=ParseMode.HTML,
    )

# ------------------------------------------------------------------------------
# 8. ADMIN PANEL ENGINE (All Button-Driven)
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
            InlineKeyboardButton("📦 Manage Stock", callback_data="admin_stock_menu"),
            InlineKeyboardButton("🏷️ Manage Products", callback_data="admin_prod_menu"),
        ],
        [
            InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="admin_pending_deposits"),
            InlineKeyboardButton(f"🛡️ VPN Orders ({pend_ord})", callback_data="admin_pending_orders"),
        ],
        [
            InlineKeyboardButton("👥 User Search & Balance", callback_data="admin_user_mgmt"),
            InlineKeyboardButton("📢 Broadcast & Alert", callback_data="admin_broadcast_menu"),
        ],
        [
            InlineKeyboardButton("📊 Shop Statistics", callback_data="admin_statistics"),
            InlineKeyboardButton("⚙️ Settings", callback_data="admin_settings"),
        ]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def handle_admin_callbacks(query, data: str, context: ContextTypes.DEFAULT_TYPE):
    user_id = query.from_user.id

    if data == "admin_stock_menu":
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.id, p.name, COUNT(s.id) as in_stock
        FROM products p
        LEFT JOIN stock s ON p.id = s.product_id AND s.is_sold = 0
        GROUP BY p.id
        HAVING p.is_on_demand = 0
        """)
        rows = c.fetchall()
        conn.close()

        stock_summary = "📦 <b>বর্তমান অটো-স্টক পরিসংখ্যান:</b>\n─────────────────────────────\n"
        for r in rows:
            stock_summary += f"• <b>{r['name']}</b> : {r['in_stock']} টি\n"
        stock_summary += "\nনতুন স্টক যোগ করতে নিচের বাটন চাপুন:"

        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("➕ Add Stock (Excel / Text)", callback_data="admin_addstock_select")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="admin_root")],
        ])
        await query.edit_message_text(stock_summary, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "admin_root":
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
                InlineKeyboardButton("📦 Manage Stock", callback_data="admin_stock_menu"),
                InlineKeyboardButton("🏷️ Manage Products", callback_data="admin_prod_menu"),
            ],
            [
                InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="admin_pending_deposits"),
                InlineKeyboardButton(f"🛡️ VPN Orders ({pend_ord})", callback_data="admin_pending_orders"),
            ],
            [
                InlineKeyboardButton("👥 User Search & Balance", callback_data="admin_user_mgmt"),
                InlineKeyboardButton("📢 Broadcast & Alert", callback_data="admin_broadcast_menu"),
            ],
            [
                InlineKeyboardButton("📊 Shop Statistics", callback_data="admin_statistics"),
                InlineKeyboardButton("⚙️ Settings", callback_data="admin_settings"),
            ]
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "admin_addstock_select":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM products WHERE is_on_demand = 0")
        prods = c.fetchall()
        conn.close()
        kb = []
        for p in prods:
            kb.append([InlineKeyboardButton(p["name"], callback_data=f"adm_uploadstock_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="admin_stock_menu")])
        await query.edit_message_text("কোন প্রোডাক্টের স্টক যোগ করবেন সিলেক্ট করুন:", reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_uploadstock_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_STOCK_FILE", {"prod_id": prod_id})
        await query.message.reply_text(
            f"📥 <b>স্টক আপলোড মোড:</b>\n\n"
            f"• আপনি একটি এক্সেল (<code>.xlsx</code>) ফাইল পাঠাতে পারেন।\n"
            f"• অথবা সরাসরি মেসেজে এক লাইনে একটি করে মেইল/প্রক্সি পেস্ট করে পাঠিয়ে দিন।\n\n"
            f"বট স্বয়ংক্রিয়ভাবে সবগুলো স্টক হিসেবে সেভ করে নেবে।",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "admin_pending_deposits":
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

    if data == "admin_pending_orders":
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
                f"👤 ইউজার: {o['uname']} (ID: <code>{o['user_id']}</code>)\n"
                f"📦 প্রোডাক্ট: {o['product_name']}\n"
                f"🔢 পরিমাণ: {o['quantity']} টি\n"
                f"💰 পরিশোধিত: {o['total_bdt']:.2f} BDT",
                parse_mode=ParseMode.HTML,
                reply_markup=kb,
            )
        return

    if data == "admin_user_mgmt":
        set_user_state(user_id, "ADMIN_WAITING_USER_SEARCH", {})
        await query.message.reply_text("🔍 আপনি যে ইউজারের তথ্য দেখতে বা ব্যালেন্স এডিট করতে চান, তার <b>টেলিগ্রাম নিউমেরিক আইডি</b> লিখে পাঠান:", parse_mode=ParseMode.HTML)
        return

    if data == "admin_statistics":
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

    if data == "admin_broadcast_menu":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔔 Send Stock Alert (Photo Card)", callback_data="adm_bcast_stock_alert")],
            [InlineKeyboardButton("📢 Send General Notice", callback_data="adm_bcast_general")],
            [InlineKeyboardButton("< Back", callback_data="admin_root")],
        ])
        await query.edit_message_text("ব্রডকাস্টের ধরণ বেছে নিন:", reply_markup=kb)
        return

    if data == "adm_bcast_stock_alert":
        set_user_state(user_id, "ADMIN_WAITING_ALERT_DETAILS", {})
        await query.message.reply_text(
            f"🔔 <b>New Stock Alert তৈরি করুন:</b>\n\n"
            f"নিচের ফরম্যাটে লিখে পাঠান:\n"
            f"<code>[Product ID] | [Added Count]</code>\n\n"
            f"উদাহরণ: <code>1 | 50</code>\n"
            f"<i>(বট স্বয়ংক্রিয়ভাবে ব্যানার ছবিসহ Buy Now বাটন তৈরি করে সবার কাছে পাঠাবে)</i>",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "adm_bcast_general":
        set_user_state(user_id, "ADMIN_WAITING_GENERAL_BCAST", {})
        await query.message.reply_text("📢 সকল ইউজারের কাছে যে মেসেজটি পাঠাতে চান, তা লিখে পাঠান:")
        return

    if data == "admin_settings":
        rate = get_setting("usdt_rate", "125.0")
        supp = get_setting("support_username")
        bkash = get_setting("bkash_number")
        nagad = get_setting("nagad_number")
        text = (
            f"⚙️ <b>বট কনফিগারেশন সেটিংস:</b>\n"
            f"─────────────────────────────\n"
            f"💵 ডলার রেট: 1 USDT = {rate} BDT\n"
            f"💬 সাপোর্ট হ্যান্ডেল: {supp}\n"
            f"📱 bKash: {bkash}\n"
            f"📱 Nagad: {nagad}\n\n"
            f"যেকোনো ভ্যালু আপডেট করতে বাটন সিলেক্ট করুন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("💵 Change USDT Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("💬 Change Support Handle", callback_data="adm_set_support")],
            [InlineKeyboardButton("< Back", callback_data="admin_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_set_rate":
        set_user_state(user_id, "ADMIN_WAITING_RATE", {})
        await query.message.reply_text("নতুন USDT এক্সচেঞ্জ রেট লিখুন (যেমন: 128.5):")
        return

    if data == "adm_set_support":
        set_user_state(user_id, "ADMIN_WAITING_SUPPORT", {})
        await query.message.reply_text("নতুন সাপোর্ট ইউজারনেম লিখুন (যেমন: @MyNewSupport):")
        return

# ------------------------------------------------------------------------------
# 9. TEXT & DOCUMENT MESSAGE ROUTER (FSM State Engine)
# ------------------------------------------------------------------------------
async def text_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text.strip()
    state, sdata = get_user_state(user.id)

    # 1. Main Reply Keyboard Navigation
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

    # 2. 2FA Secret Key State
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

    # 3. Custom Quantity State
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
        # Wrap as mock callback query
        class MockQuery:
            def __init__(self, msg, u):
                self.message = msg
                self.from_user = u
                self.bot = context.bot
        await process_order(MockQuery(update.message, user), prod_id, qty)
        return

    # 4. Deposit Submission State Flow
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
            f"✅ <b>ডিপোজিট রিকোয়েস্ট জমা হয়েছে!</b>\n\n"
            f"আইডি: #{dep_id}\n"
            f"মেথড: {method}\n"
            f"পরিমাণ: {amount:.2f} BDT\n"
            f"TrxID: <code>{trx}</code>\n\n"
            f"এডমিন চেক করে ৫-১০ মিনিটের মধ্যে আপনার ব্যালেন্স যুক্ত করে দেবে।",
            parse_mode=ParseMode.HTML,
        )

        # Notify Admin
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

    # 5. Admin Actions State Flow
    if not is_admin(user.id):
        return

    if state == "ADMIN_WAITING_STOCK_FILE":
        # Multi-line text input
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
        c.execute("SELECT o.*, p.image_url FROM orders o JOIN products p ON o.product_id = p.id WHERE o.id = ?", (order_id,))
        order = c.fetchone()
        if order:
            c.execute("UPDATE orders SET status = 'completed', delivered_data = ? WHERE id = ?", (text, order_id))
            conn.commit()
            conn.close()
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ Order #{order_id} সফলভাবে ইউজারের কাছে ডেলিভারি পাঠানো হয়েছে।")

            # Deliver Rich Photo Card to User
            user_msg = (
                f"🎉 <b>আপনার VPN সার্ভিস প্রস্তুত! (Order #{order_id})</b>\n"
                f"─────────────────────────────\n"
                f"🛡️ <b>প্রোডাক্ট:</b> {order['product_name']}\n"
                f"🔢 <b>পরিমাণ:</b> {order['quantity']} টি\n\n"
                f"🔑 <b>লগইন ডিটেইলস / অ্যাক্সেস কি:</b>\n"
                f"<code>{text}</code>\n\n"
                f"📌 <i>কোনো সমস্যা হলে সাপোর্টে যোগাযোগ করুন। আমাদের সাথে থাকার জন্য ধন্যবাদ!</i>"
            )
            banner = order["image_url"] or BRAND_BANNERS["vpn"]
            try:
                await context.bot.send_photo(chat_id=order["user_id"], photo=banner, caption=user_msg, parse_mode=ParseMode.HTML)
            except Exception:
                await context.bot.send_message(chat_id=order["user_id"], text=user_msg, parse_mode=ParseMode.HTML)
        else:
            conn.close()
            await update.message.reply_text("অর্ডার পাওয়া যায়নি।")
        return

    if state == "ADMIN_WAITING_ALERT_DETAILS":
        parts = [p.strip() for p in text.split("|")]
        if len(parts) >= 2:
            try:
                prod_id = int(parts[0])
                added_count = int(parts[1])
            except ValueError:
                await update.message.reply_text("ফরম্যাট ভুল! উদাহরণ: <code>1 | 50</code>", parse_mode=ParseMode.HTML)
                return
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT * FROM products WHERE id = ?", (prod_id,))
            prod = c.fetchone()
            c.execute("SELECT COUNT(*) as cnt FROM stock WHERE product_id = ? AND is_sold = 0", (prod_id,))
            curr_stock = c.fetchone()["cnt"]
            c.execute("SELECT user_id FROM users")
            all_users = [row["user_id"] for row in c.fetchall()]
            conn.close()

            if not prod:
                await update.message.reply_text("প্রোডাক্ট পাওয়া যায়নি।")
                return

            set_user_state(user.id, None)
            alert_caption = (
                f"🔔 <b>New Stock Alert!</b> 🔔\n"
                f"───────────────────────────\n\n"
                f"💥 <b>Item:</b> {prod['name']}\n"
                f"➕ <b>Added:</b> {added_count}\n"
                f"📊 <b>Current stock:</b> {curr_stock}\n"
                f"💰 <b>Price:</b> {prod['price_bdt']:.2f} BDT / {prod['price_usdt']:.4f} USDT\n\n"
                f"───────────────────────────\n"
                f"Click the button below to purchase instantly."
            )
            alert_kb = InlineKeyboardMarkup([[InlineKeyboardButton("💥 Buy Now", callback_data=f"prod_{prod['id']}") ]])
            banner = prod["image_url"] or BRAND_BANNERS["default"]

            sent_cnt = 0
            for uid in all_users:
                try:
                    await context.bot.send_photo(chat_id=uid, photo=banner, caption=alert_caption, parse_mode=ParseMode.HTML, reply_markup=alert_kb)
                    sent_cnt += 1
                except Exception:
                    pass
            await update.message.reply_text(f"✅ সফলভাবে <b>{sent_cnt}</b> জন ইউজারের কাছে Stock Alert পাঠানো হয়েছে!", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_GENERAL_BCAST":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT user_id FROM users")
        all_users = [row["user_id"] for row in c.fetchall()]
        conn.close()
        set_user_state(user.id, None)

        sent = 0
        for uid in all_users:
            try:
                await context.bot.send_message(chat_id=uid, text=f"📢 <b>ঘোষণা:</b>\n\n{text}", parse_mode=ParseMode.HTML)
                sent += 1
            except Exception:
                pass
        await update.message.reply_text(f"✅ <b>{sent}</b> জন ইউজারের কাছে নোটিশ পাঠানো হয়েছে।", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_USER_SEARCH":
        try:
            target_id = int(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক আইডি লিখুন:")
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
        info = (
            f"👤 <b>User Info:</b>\n"
            f"আইডি: <code>{u['user_id']}</code>\n"
            f"নাম: {u['name']}\n"
            f"ইউজারনেম: @{u['username'] if u['username'] else 'None'}\n"
            f"বর্তমান ব্যালেন্স: <b>{u['balance_bdt']:.2f} BDT</b>\n"
            f"মোট কেনাকাটা: {stats['cnt']} টি ({spent:.2f} BDT)\n"
            f"জয়েনিং তারিখ: {u['joined_date']}\n\n"
            f"ব্যালেন্স এডিট করতে চাইলে কমান্ড লিখুন:\n"
            f"যোগ করতে: <code>/addbalance {u['user_id']} 100</code>\n"
            f"কাটতে: <code>/deductbalance {u['user_id']} 50</code>"
        )
        await update.message.reply_text(info, parse_mode=ParseMode.HTML)
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
        await update.message.reply_text(f"✅ নতুন সাপোর্ট হ্যান্ডেল সেট হয়েছে: {text}")
        return

async def document_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles Excel (.xlsx) file uploads by Admin for bulk stock injection."""
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
            await update.message.reply_text("❌ এক্সেল ফাইলের ভেতর কোনো স্টক ডেটা পাওয়া যায়নি।")
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
            f"📊 নতুন যুক্ত হয়েছে: <b>{len(items)}</b> টি অ্যাকাউন্ট/আইটেম\n"
            f"📦 স্টকে বর্তমান মোট সংখ্যা: <b>{total_now}</b> টি।",
            parse_mode=ParseMode.HTML,
        )

# ------------------------------------------------------------------------------
# 10. ADMIN SHORTCUT COMMANDS (Optional power tools)
# ------------------------------------------------------------------------------
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
# RENDER WEB SERVICE HEALTH CHECK SERVER (Keeps Render Free Web Service Active)
# ------------------------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"CyberNest Telegram Bot is Active and Polling!")

    def log_message(self, format, *args):
        # Suppress noisy HTTP polling logs in Render console
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    server.serve_forever()

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
    app.add_handler(MessageHandler(filters.Document.ALL, document_message_handler))

    print("Bot is polling for updates...")
    app.run_polling()

if __name__ == "__main__":
    main()
