# ==============================================================================
# CyberNest Digital Mart — Full-Featured 4-Tier Telegram Auto-Shop Bot
# Architecture: Python (python-telegram-bot v21+), SQLite / Neon PostgreSQL, OpenPyXL, ReportLab, PyOTP
#
# Configured:
# - Super Admin ID: 8919985167
# - Token: Read from BOT_TOKEN env var (fallback to configured token)
# - Embedded Render Web Service Health Check Server included
# ==============================================================================

import os
import sys
import io
import urllib.request
import asyncio
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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing
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


_GLOBAL_BOT_APP = None
_GLOBAL_EVENT_LOOP = None

async def _notify_payment_success(user_id: int, amount: float, invoice_id: str):
    if _GLOBAL_BOT_APP:
        try:
            await _GLOBAL_BOT_APP.bot.send_message(
                chat_id=user_id,
                text=(
                    f"🎉 <b>পেমেন্ট সফলভাবে সম্পন্ন হয়েছে!</b>\n\n"
                    f"💰 যুক্ত হয়েছে: <b>{amount:.2f} BDT</b>\n"
                    f"🧾 ইনভয়েস: <code>{invoice_id}</code>\n\n"
                    f"আপনার ব্যালেন্স স্বয়ংক্রিয়ভাবে আপডেট করা হয়েছে। শপ থেকে কেনাকাটা শুরু করুন!"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass
        try:
            await _GLOBAL_BOT_APP.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>অটোমেটিক পেমেন্ট সম্পন্ন! (Paymently Gateway)</b>\n\n"
                    f"👤 ইউজার আইডি: <code>{user_id}</code>\n"
                    f"💰 পরিমাণ: <b>{amount:.2f} BDT</b>\n"
                    f"🧾 ইনভয়েস: <code>{invoice_id}</code>"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass

# ------------------------------------------------------------------------------
# 2. RENDER WEB SERVICE HEALTH CHECK SERVER (Keeps Render Free Instance Live)
# ------------------------------------------------------------------------------
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"CyberNest Bot is running and healthy!")

    def do_POST(self):
        if "/webhook" in self.path or "/payment" in self.path:
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)
            try:
                data = json.loads(post_data.decode("utf-8"))
                logger.info(f"Received Paymently Webhook: {data}")
                status = str(data.get("status", "")).upper()
                if status in ("COMPLETED", "SUCCESS", "PAID"):
                    invoice_id = str(data.get("invoice_id", ""))
                    amount = float(data.get("amount", 0))
                    metadata = data.get("metadata", {})
                    user_id = int(metadata.get("user_id", 0)) if metadata else 0

                    if user_id > 0 and amount > 0:
                        conn = get_db()
                        c = conn.cursor()
                        c.execute("SELECT id FROM deposits WHERE trx_id = ? AND status = 'approved'", (invoice_id,))
                        if not c.fetchone():
                            now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
                            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amount, user_id))
                            c.execute("INSERT INTO deposits (user_id, method, amount, trx_id, status, created_at) VALUES (?, 'Paymently Auto', ?, ?, 'approved', ?)", (user_id, amount, invoice_id, now_str))
                            conn.commit()
                            if _GLOBAL_EVENT_LOOP and _GLOBAL_BOT_APP:
                                asyncio.run_coroutine_threadsafe(
                                    _notify_payment_success(user_id, amount, invoice_id),
                                    _GLOBAL_EVENT_LOOP
                                )
                        conn.close()

                self.send_response(200)
                self.send_header("Content-type", "application/json")
                self.end_headers()
                self.wfile.write(b'{"status":"ok"}')
                return
            except Exception as e:
                logger.error(f"Error handling webhook: {e}")
                self.send_response(400)
                self.end_headers()
                return

        self.send_response(404)
        self.end_headers()

    def log_message(self, format, *args):
        pass

def start_health_server():
    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), HealthCheckHandler)
    server.serve_forever()

# ------------------------------------------------------------------------------
# 3. HIGH-SPEED PERSISTENT DATABASE LAYER (Neon PostgreSQL + SQLite Fallback)
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
            sql = sql.replace("INTEGER PRIMARY KEY", "BIGINT PRIMARY KEY")
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
            # Normal execution with safe rollback
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
            global _GLOBAL_CONN
            _GLOBAL_CONN = None
            self.conn = get_raw_connection()
            return SmartCursor(self.conn.cursor(), self.is_pg)

    def commit(self):
        if not self.is_pg:
            self.conn.commit()

    def close(self):
        pass  # Persistent connection kept alive

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

    # Categories Table (Level 1)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        icon TEXT DEFAULT '🛍️',
        position INTEGER DEFAULT 0
    )
    """)

    # Subcategories / Brands Table (Level 2)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS subcategories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_id INTEGER,
        name TEXT NOT NULL,
        position INTEGER DEFAULT 0,
        FOREIGN KEY (category_id) REFERENCES categories (id) ON DELETE CASCADE
    )
    """)

    # Products / Variants Table (Level 3)
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

    # Stock Items Table (Email:Pass, Proxy string, Key)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER,
        content TEXT NOT NULL,
        stock_type TEXT DEFAULT 'account',
        is_sold INTEGER DEFAULT 0,
        sold_to BIGINT,
        sold_at TEXT,
        FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
    )
    """)

    # Orders Table
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

    # Ensure position columns exist
    try:
        cursor.execute("ALTER TABLE categories ADD COLUMN position INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE subcategories ADD COLUMN position INTEGER DEFAULT 0")
    except Exception:
        pass

    # Migration: Purge legacy intermediate 'Accounts' and 'Recovery Accounts' subcategories
    try:
        cursor.execute("SELECT id FROM subcategories WHERE name IN ('Accounts', 'Recovery Accounts')")
        bad_subs = cursor.fetchall()
        if bad_subs:
            bad_ids = [r["id"] for r in bad_subs]
            placeholders = ",".join(["?"] * len(bad_ids))
            cursor.execute(f"DELETE FROM stock WHERE product_id IN (SELECT id FROM products WHERE subcategory_id IN ({placeholders}))", bad_ids)
            cursor.execute(f"DELETE FROM products WHERE subcategory_id IN ({placeholders})", bad_ids)
            cursor.execute(f"DELETE FROM subcategories WHERE id IN ({placeholders})", bad_ids)
            conn.commit()
            logger.info(f"Purged legacy subcategories: {bad_ids}")
    except Exception as e:
        logger.warning(f"Migration purge notice: {e}")

    # Ensure direct subcategories exist under Mail
    try:
        cursor.execute("SELECT id FROM categories WHERE name LIKE '%Mail%' AND name NOT LIKE '%Recovery%' LIMIT 1")
        m_row = cursor.fetchone()
        if m_row:
            m_id = m_row["id"]
            for sname, srate in [("Hotmail", 0.65), ("Outlook", 0.68), ("FR Outlook", 0.87)]:
                cursor.execute("SELECT id FROM subcategories WHERE category_id = ? AND name = ?", (m_id, sname))
                if not cursor.fetchone():
                    cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, 0)", (m_id, sname))
                    s_id = cursor.lastrowid
                    cursor.execute("INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)", (s_id, sname, srate, srate / 125.0))
            conn.commit()
    except Exception as e:
        logger.warning(f"Mail subcategory check notice: {e}")

    # Ensure direct subcategories exist under Recovery Added Mail
    try:
        cursor.execute("SELECT id FROM categories WHERE name LIKE '%Recovery%Mail%' LIMIT 1")
        r_row = cursor.fetchone()
        if r_row:
            r_id = r_row["id"]
            for sname, srate in [("Recovery Hotmail", 0.95), ("Recovery Outlook", 0.95), ("Recovery FR Outlook", 1.20)]:
                cursor.execute("SELECT id FROM subcategories WHERE category_id = ? AND name = ?", (r_id, sname))
                if not cursor.fetchone():
                    cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, 0)", (r_id, sname))
                    s_id = cursor.lastrowid
                    cursor.execute("INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)", (s_id, sname, srate, srate / 125.0))
            conn.commit()
    except Exception as e:
        logger.warning(f"Recovery subcategory check notice: {e}")

    conn.commit()

    # Seed Default Settings
    default_settings = {
        "shop_name": "Will be Earn Shop",
        "shop_tagline": "Digital Products & Services",
        "usdt_rate": "125.0",
        "min_deposit_manual": "20.0",
        "min_deposit_instant": "100.0",
        "support_username": "@Talha_juba098",
        "bkash_number": "01XXXXXXXXX (Personal - Send Money)",
        "nagad_number": "01XXXXXXXXX (Personal - Send Money)",
        "binance_id": "Pay ID: 12345678 / USDT (TRC20): TQxxxxxxxxxxxxxxxxxx",
        "paymently_base_url": "https://abutalha.paymently.io/api",
        "paymently_api_key": "DSlm0lfHRXTa6ZaKuBoRfODt6esAIRxhzUWwBacg",
        "render_app_url": "https://telegram-bot-2-cs6b.onrender.com",
        "mail_code_url_1": "https://code.yamin.bd/",
    }
    for k, v in default_settings.items():
        if conn.is_pg:
            cursor.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO NOTHING", (k, v))
        else:
            cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v))
    conn.commit()

    # Auto-healing Seed Check
    cursor.execute("SELECT COUNT(*) as cnt FROM categories")
    row = cursor.fetchone()
    cnt = row["cnt"] if isinstance(row, dict) else row[0]
    cursor.execute("SELECT COUNT(*) as cnt FROM subcategories")
    s_row = cursor.fetchone()
    s_cnt = s_row["cnt"] if isinstance(s_row, dict) else s_row[0]
    if cnt == 0 or s_cnt == 0:
        logger.info("Initializing fresh categories, brands and variants...")
        try:
            cursor.execute("DELETE FROM products")
            cursor.execute("DELETE FROM subcategories")
            cursor.execute("DELETE FROM categories")
            conn.commit()
        except Exception:
            pass
        seed_clean_store(conn)

    conn.close()

def seed_clean_store(conn):
    cursor = conn.cursor()

    # 1. Categories (Level 1)
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

    # 2. VPN Brands (Level 2) & Durations (Level 3)
    vpn_id = cat_map.get("🛡️ VPN")
    vpn_catalog = {
        "Proton VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("10 Days", 25.0), ("15 Days", 35.0), ("30 Days", 60.0)],
        "Express VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("15 Days", 35.0), ("30 Days", 65.0)],
        "Nord VPN": [("3 Days", 18.0), ("7 Days", 20.0), ("15 Days", 35.0), ("30 Days", 60.0)],
        "Surfshark VPN": [("7 Days", 20.0), ("15 Days", 35.0), ("30 Days", 60.0)],
        "Cyberghost VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("30 Days", 60.0)],
        "IPVanish VPN": [("3 Days", 15.0), ("7 Days", 20.0), ("30 Days", 60.0)],
        "Bitdefender VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "Hotspot Shield": [("7 Days", 20.0), ("30 Days", 60.0)],
        "HMA VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "Vypr VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "Norton VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "PIA VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "Panda VPN": [("3 Days", 15.0), ("7 Days", 20.0)],
        "Avast VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
        "Adguard VPN": [("7 Days", 20.0), ("30 Days", 60.0)],
    }

    for brand, variants in vpn_catalog.items():
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (vpn_id, brand))
        sub_id = cursor.lastrowid
        for dur_name, bdt in variants:
            usdt = bdt / 125.0
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 1)",
                (sub_id, dur_name, bdt, usdt),
            )

    # 3. Proxy Brands & Sizes
    proxy_id = cat_map.get("🌐 Proxy")
    proxy_catalog = {
        "9 Proxy": [("100 MB", 12.0), ("200 MB", 23.0), ("500 MB", 60.0), ("1 GB", 110.0)],
        "Rapid Proxy": [("100 MB", 12.0), ("200 MB", 25.0), ("500 MB", 60.0), ("1 GB", 115.0)],
        "711 Proxy": [("100 MB", 10.0), ("200 MB", 18.0), ("500 MB", 45.0)],
    }

    for pbrand, pvariants in proxy_catalog.items():
        cursor.execute("INSERT INTO subcategories (category_id, name) VALUES (?, ?)", (proxy_id, pbrand))
        psub_id = cursor.lastrowid
        for pkg_name, bdt in pvariants:
            usdt = bdt / 125.0
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (psub_id, pkg_name, bdt, usdt),
            )

    # 4. Mail Brands & Plans (Direct: Hotmail, Outlook, FR Outlook)
    mail_id = cat_map.get("✉️ Mail")
    mail_catalog = {
        "Hotmail": [("Hotmail", 0.65)],
        "Outlook": [("Outlook", 0.68)],
        "FR Outlook": [("FR Outlook", 0.87)],
    }
    for mpos, (mbrand, mvariants) in enumerate(mail_catalog.items(), start=1):
        cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, ?)", (mail_id, mbrand, mpos))
        msub_id = cursor.lastrowid
        for vname, bdt in mvariants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (msub_id, vname, bdt, bdt / 125.0),
            )

    # 5. Recovery Added Mail (Direct: Recovery Hotmail, Recovery Outlook, Recovery FR Outlook)
    rec_mail_id = cat_map.get("✉️ Recovery Added Mail")
    rec_catalog = {
        "Recovery Hotmail": [("Recovery Added Hotmail", 0.95)],
        "Recovery Outlook": [("Recovery Added Outlook", 0.95)],
        "Recovery FR Outlook": [("Recovery Added FR Outlook", 1.20)],
    }
    for rpos, (rbrand, rvariants) in enumerate(rec_catalog.items(), start=1):
        cursor.execute("INSERT INTO subcategories (category_id, name, position) VALUES (?, ?, ?)", (rec_mail_id, rbrand, rpos))
        rsub_id = cursor.lastrowid
        for vname, bdt in rvariants:
            cursor.execute(
                "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, 0)",
                (rsub_id, vname, bdt, bdt / 125.0),
            )

    conn.commit()

# DB Helpers
def format_price(val: float) -> str:
    """Formats prices cleanly supporting sub-1 BDT (e.g. 0.87, 0.65) and whole numbers."""
    if val is None:
        return "0"
    if val < 1:
        return f"{val:.2f}"
    if val == int(val):
        return f"{int(val)}"
    return f"{val:.2f}".rstrip('0').rstrip('.')

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
    if conn.is_pg:
        c.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value", (key, value))
    else:
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

    story = []
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

    total_data = [["TOTAL PAID (BDT)", f"{total_bdt:.2f} BDT"]]
    total_table = Table(total_data, colWidths=[370, 150])
    total_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#7B1FA2')),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 13),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(total_table)
    story.append(Spacer(1, 25))

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

    disc = (
        "<i>This is a system-generated receipt and does not require a signature.</i><br/>"
        f"<b>Thank you for shopping with {shop_name}!</b>"
    )
    story.append(Paragraph(disc, ParagraphStyle('Disc', alignment=1, fontSize=9, textColor=colors.HexColor('#555555'))))

    doc.build(story)
    bio.seek(0)
    return bio



def create_paymently_charge(amount: float, user_id: int, user_name: str, bot_username: str) -> Optional[Dict[str, Any]]:
    base_url = get_setting("paymently_base_url", "https://abutalha.paymently.io/api").rstrip("/")
    api_key = get_setting("paymently_api_key", "DSlm0lfHRXTa6ZaKuBoRfODt6esAIRxhzUWwBacg")
    render_url = get_setting("render_app_url", "https://telegram-bot-2-cs6b.onrender.com").rstrip("/")
    
    endpoint = f"{base_url}/checkout-v2"
    payload = {
        "full_name": user_name or "Telegram User",
        "email": f"user{user_id}@telegram.org",
        "amount": str(amount),
        "metadata": {
            "user_id": str(user_id)
        },
        "redirect_url": f"https://t.me/{bot_username.lstrip('@')}",
        "return_type": "GET",
        "cancel_url": f"https://t.me/{bot_username.lstrip('@')}",
        "webhook_url": f"{render_url}/webhook/payment"
    }
    headers = {
        "RT-UDDOKTAPAY-API-KEY": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "CyberNestBot/1.0"
    }
    req = urllib.request.Request(endpoint, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            return res_data
    except Exception as e:
        logger.error(f"Error creating Paymently charge: {e}")
        return None

def verify_paymently_charge(invoice_id: str) -> Optional[Dict[str, Any]]:
    base_url = get_setting("paymently_base_url", "https://abutalha.paymently.io/api").rstrip("/")
    api_key = get_setting("paymently_api_key", "DSlm0lfHRXTa6ZaKuBoRfODt6esAIRxhzUWwBacg")
    endpoint = f"{base_url}/verify-payment"
    payload = {"invoice_id": invoice_id}
    headers = {
        "RT-UDDOKTAPAY-API-KEY": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "CyberNestBot/1.0"
    }
    req = urllib.request.Request(endpoint, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            return res_data
    except Exception as e:
        logger.error(f"Error verifying Paymently charge: {e}")
        return None

def format_credentials_card(items: List[str], prod_name: str, cat_name: str) -> str:
    """Formats credentials with tap-to-copy code blocks matching user's screenshots."""
    is_proxy = "PROXY" in cat_name.upper() or "PROXY" in prod_name.upper()
    lines = []
    
    if is_proxy:
        lines.append("✅ <b>Proxy Delivered!</b>")
        lines.append(f"🌐 <b>{prod_name}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            # Parse server:port:user:pass or server:port@user:pass or similar
            parts = [p.strip() for p in item.replace("@", ":").split(":") if p.strip()]
            if len(parts) >= 4:
                host, port, u_name, pwd = parts[0], parts[1], parts[2], parts[3]
            elif len(parts) == 2:
                host, port = parts[0], parts[1]
                u_name, pwd = "N/A", "N/A"
            else:
                host, port, u_name, pwd = item, "N/A", "N/A", "N/A"

            if len(items) > 1:
                lines.append(f"🌐 <b>Proxy #{idx}</b>")
            lines.append(f"🖥️ <b>Host</b>     : <code>{host}</code>")
            lines.append(f"🔌 <b>Port</b>     : <code>{port}</code>")
            lines.append(f"👤 <b>Username</b> : <code>{u_name}</code>")
            lines.append(f"🔐 <b>Password</b> : <code>{pwd}</code>")
            if idx < len(items):
                lines.append("")
    else:
        lines.append("✅ <b>Service Delivered!</b>")
        lines.append(f"🛡️ <b>{prod_name}</b>")
        lines.append("────────────────────────")
        for idx, item in enumerate(items, start=1):
            parts = [p.strip() for p in item.split(":") if p.strip()]
            if len(parts) >= 2:
                email, pwd = parts[0], parts[1]
                if len(items) > 1:
                    lines.append(f"📦 <b>Item #{idx}</b>")
                lines.append(f"👤 <b>Email / User</b> : <code>{email}</code>")
                lines.append(f"🔐 <b>Password</b>     : <code>{pwd}</code>")
            else:
                if len(items) > 1:
                    lines.append(f"🔑 <b>Key / Code #{idx}</b>")
                lines.append(f"🎟️ <b>Key / Code</b> : <code>{item}</code>")
            if idx < len(items):
                lines.append("")

    lines.append("────────────────────────")
    lines.append("💡 <i>Tap any value to copy</i>")
    return "\n".join(lines)

# ------------------------------------------------------------------------------
# 6. USER HANDLERS (Shop, Profile, Deposit, 2FA, Support)
# ------------------------------------------------------------------------------
async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    shop_name = get_setting("shop_name", "Will be Earn Shop")

    # Deep-Linking verification check
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
        await update.message.reply_text("❌ দুঃখিত! এই রিসিট নম্বরের কোনো অর্ডার পাওয়া যায়নি।")
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
    c.execute("SELECT id, name FROM categories ORDER BY id ASC")
    cats = c.fetchall()
    conn.close()

    text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
    keyboard = []
    for cat in cats:
        keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
    keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])

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

async def deposit_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db_user = get_or_create_user(user.id, user.full_name, user.username)
    min_man = float(get_setting("min_deposit_manual", "20.0"))
    min_ins = float(get_setting("min_deposit_instant", "100.0"))

    text = (
        f"💳 <b>ডিপোজিট মেথড নির্বাচন করুন:</b>\n"
        f"─────────────────────────────\n"
        f"💰 বর্তমান ব্যালেন্স: <b>{db_user['balance_bdt']:.2f} BDT</b>\n\n"
        f"⚡ <b>১. Instant Auto Payment:</b>\n"
        f"• বিকাশ, নগদ ও কার্ড দিয়ে ১ সেকেন্ডে অটোমেটিক ব্যালেন্স যোগ হবে। কোনো TrxID লাগবে না।\n"
        f"• সর্বনিম্ন ডিপোজিট: <b>{min_ins:.0f} BDT</b>\n\n"
        f"📱 <b>২. Manual Send Money:</b>\n"
        f"• পার্সোনাল নাম্বারে সেন্ড মানি করে TrxID পাঠাবেন, এডমিন চেক করে ব্যালেন্স যোগ করবে।\n"
        f"• সর্বনিম্ন ডিপোজিট: <b>{min_man:.0f} BDT</b>\n\n"
        f"নিচের বাটন থেকে আপনার পছন্দের মাধ্যমটি বেছে নিন 👇"
    )
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton(f"⚡ Instant Auto Payment (Min {min_ins:.0f} Tk)", callback_data="dep_opt_instant")],
        [InlineKeyboardButton(f"📱 Manual Send Money (Min {min_man:.0f} Tk)", callback_data="dep_opt_manual")],
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

# ------------------------------------------------------------------------------
# 7. INLINE CALLBACK QUERY DISPATCHER (Users & Admin Nested Flow)
# ------------------------------------------------------------------------------
async def callback_query_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    db_user = get_or_create_user(user_id, query.from_user.full_name, query.from_user.username)

    # ---------------- USER SHOP BROWSING FLOW (4-TIER NESTED) ----------------
    # Step 1 -> Step 2: Category Selected -> Show Brands / Subcategories
    if data.startswith("usr_cat_"):
        cat_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        text = f"<b>{cat_name} — Select Brand / Section:</b>"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(s["name"], callback_data=f"usr_sub_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="usr_root_cats")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data == "usr_cancel_home":
        await query.message.reply_text("❌ কেনাকাটা বাতিল করা হয়েছে। নতুন অর্ডারের জন্য মেনু ব্যবহার করুন।")
        try:
            await query.message.delete()
        except Exception:
            pass
        return

    if data == "usr_root_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY id ASC")
        cats = c.fetchall()
        conn.close()
        text = "🛍️ <b>Buy Products</b>\n\nSelect a category:"
        keyboard = []
        for cat in cats:
            keyboard.append([InlineKeyboardButton(cat["name"], callback_data=f"usr_cat_{cat['id']}")])
        keyboard.append([InlineKeyboardButton("❌ Cancel", callback_data="usr_cancel_home")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(keyboard))
        return

    # Step 2 -> Step 3: Brand Selected -> Show Variants with Real Stock & Exact Decimal Prices
    if data.startswith("usr_sub_"):
        sub_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()
        
        c.execute("""
        SELECT p.id, p.name, p.price_bdt, p.price_usdt, p.is_on_demand, COUNT(st.id) as in_stock
        FROM products p
        LEFT JOIN stock st ON p.id = st.product_id AND st.is_sold = 0
        WHERE p.subcategory_id = ?
        GROUP BY p.id, p.name, p.price_bdt, p.price_usdt, p.is_on_demand
        ORDER BY p.id ASC
        """, (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} ➔ {meta['sub_name']}" if meta else "Select Variant"
        text = f"<b>{title}</b>\n\nনিচের প্ল্যান বা মেয়াদ নির্বাচন করুন:"
        kb = []
        for p in prods:
            stock_str = "⚡ Instant" if p["is_on_demand"] else f"Stock: {p['in_stock']}"
            btn_text = f"{p['name']} | {format_price(p['price_bdt'])} BDT | {stock_str}"
            kb.append([InlineKeyboardButton(btn_text, callback_data=f"usr_prod_{p['id']}")])
        back_cat_id = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"usr_cat_{back_cat_id}")])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 3 -> Step 4: Variant Selected -> Show Product Card & Quantity Selector
    if data.startswith("usr_prod_"):
        prod_id = int(data.split("_")[2])
        await render_purchase_summary(query, prod_id, quantity=1, edit=True)
        return

    # Interactive Purchase Summary Calculator Callbacks
    if data.startswith("calc_add_"):
        parts = data.split("_")
        prod_id = int(parts[2])
        qty = int(parts[3]) + 1
        await render_purchase_summary(query, prod_id, quantity=qty, edit=True)
        return

    if data.startswith("calc_sub_"):
        parts = data.split("_")
        prod_id = int(parts[2])
        qty = max(1, int(parts[3]) - 1)
        await render_purchase_summary(query, prod_id, quantity=qty, edit=True)
        return

    if data.startswith("calc_custom_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "WAITING_CALC_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text("✏️ <b>Enter Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("calc_confirm_"):
        parts = data.split("_")
        prod_id = int(parts[2])
        qty = int(parts[3])
        await process_order(query, prod_id, qty)
        return

    # Step 4 -> Purchase Order
    if data.startswith("qty_"):
        parts = data.split("_")
        prod_id = int(parts[1])
        qty = int(parts[2])
        await process_order(query, prod_id, qty)
        return

    if data.startswith("custqty_"):
        prod_id = int(data.split("_")[1])
        set_user_state(user_id, "WAITING_CUSTOM_QTY", {"prod_id": prod_id})
        await query.message.reply_text(
            "✏️ <b>Custom Quantity:</b>\nআপনি কতগুলো কিনতে চান তা সংখ্যায় লিখে পাঠান (যেমন: 5, 20, 50):",
            parse_mode=ParseMode.HTML,
        )
        return

    # Deposit Choice: Instant Auto Payment vs Manual Send Money
    if data == "dep_opt_instant":
        min_ins = float(get_setting("min_deposit_instant", "100.0"))
        set_user_state(user_id, "WAITING_DEP_AMOUNT_INSTANT", {})
        await query.message.reply_text(
            f"⚡ <b>Instant Auto Payment (কোনো TrxID লাগবে না):</b>\n\n"
            f"আপনি কত টাকা ডিপোজিট করতে চান? টাকার পরিমাণ লিখে পাঠান (সর্বনিম্ন <b>{min_ins:.0f}</b> টাকা):",
            parse_mode=ParseMode.HTML,
        )
        return

    if data == "dep_opt_manual":
        min_man = float(get_setting("min_deposit_manual", "20.0"))
        set_user_state(user_id, "WAITING_DEP_AMOUNT_MANUAL", {})
        await query.message.reply_text(
            f"📱 <b>Manual Send Money (পার্সোনাল নাম্বার):</b>\n\n"
            f"আপনি কত টাকা ডিপোজিট করতে চান? টাকার পরিমাণ লিখে পাঠান (সর্বনিম্ন <b>{min_man:.0f}</b> টাকা):",
            parse_mode=ParseMode.HTML,
        )
        return

    # Auto Payment (Paymently Gateway Link Generation)
    if data == "usr_dep_auto_pay":
        state, sdata = get_user_state(user_id)
        amount = sdata.get("amount", 20.0)
        bot_info = await context.bot.get_me()

        await query.message.reply_text("⏳ আপনার সিকিউর পেমেন্ট লিংক তৈরি করা হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন...")
        charge_res = create_paymently_charge(amount, user_id, query.from_user.full_name, bot_info.username)

        if charge_res and charge_res.get("status") and charge_res.get("payment_url"):
            pay_url = charge_res["payment_url"]
            inv_id = charge_res.get("invoice_id", "")
            set_user_state(user_id, None)

            txt = (
                f"💳 <b>অটোমেটিক পেমেন্ট গেটওয়ে (bKash / Nagad / Cards)</b>\n"
                f"─────────────────────────────\n"
                f"💰 <b>ডিপোজিট পরিমাণ:</b> <b>{amount:.2f} BDT</b>\n"
                f"🧾 <b>ইনভয়েস আইডি:</b> <code>{inv_id}</code>\n\n"
                f"📌 নিচের <b>Pay Now</b> বাটনে চাপ দিয়ে বিকাশ বা নগদে পেমেন্ট সম্পন্ন করুন।\n"
                f"পেমেন্ট শেষ হওয়ামাত্র কোনো TrxID ছাড়াই আপনার ব্যালেন্স স্বয়ংক্রিয়ভাবে যোগ হয়ে যাবে!\n\n"
                f"<i>(পেমেন্ট শেষে অটো যোগ না হলে 'Check Payment Status' বাটনে চাপ দিন)</i>"
            )
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("💳 Pay Now (বিকাশ / নগদ)", url=pay_url)],
                [InlineKeyboardButton("🔄 Check Payment Status", callback_data=f"usr_checkpay_{inv_id}_{amount}")],
                [InlineKeyboardButton("< Cancel", callback_data="usr_root_cats")]
            ])
            await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=kb)
        else:
            await query.message.reply_text("⚠️ গেটওয়ে সার্ভারে সাময়িক সমস্যা হচ্ছে। দয়া করে Manual Send Money অপশনটি ব্যবহার করুন।")
        return

    # Manual payment methods menu
    if data == "usr_dep_manual_menu":
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ])
        await query.edit_message_text("নিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:", reply_markup=kb)
        return

    # Manual verify button (Instant verification fallback)
    if data.startswith("usr_checkpay_"):
        parts = data.split("_")
        inv_id = parts[2]
        amount = float(parts[3])

        await query.message.reply_text(f"🔍 ইনভয়েস <code>{inv_id}</code> যাচাই করা হচ্ছে...")
        v_res = verify_paymently_charge(inv_id)

        if v_res and str(v_res.get("status", "")).upper() in ("COMPLETED", "PAID", "SUCCESS"):
            conn = get_db()
            c = conn.cursor()
            c.execute("SELECT id FROM deposits WHERE trx_id = ? AND status = 'approved'", (inv_id,))
            if not c.fetchone():
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %I:%M %p")
                c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (amount, user_id))
                c.execute("INSERT INTO deposits (user_id, method, amount, trx_id, status, created_at) VALUES (?, 'Paymently Auto', ?, ?, 'approved', ?)", (user_id, amount, inv_id, now_str))
                conn.commit()
                await query.message.reply_text(
                    f"🎉 <b>পেমেন্ট সফলভাবে ভেরিফাইড হয়েছে!</b>\n"
                    f"আপনার অ্যাকাউন্টে <b>{amount:.2f} BDT</b> যোগ করা হয়েছে!",
                    parse_mode=ParseMode.HTML,
                )
            else:
                await query.message.reply_text("✅ এই পেমেন্টের ব্যালেন্স ইতিমধ্যে যোগ করা হয়েছে।")
            conn.close()
        else:
            await query.message.reply_text("⏳ এখনো পেমেন্ট সম্পন্ন হয়নি। পেমেন্ট শেষ করে আবার চেক করুন।")
        return

    # ---------------- DEPOSIT METHOD SELECTION ----------------
    if data.startswith("usr_dep_method_"):
        method = data.split("_")[3]
        state, sdata = get_user_state(user_id)
        amount = sdata.get("amount", 20.0)

        num = ""
        if method == "bKash":
            num = get_setting("bkash_number")
        elif method == "Nagad":
            num = get_setting("nagad_number")
        elif method == "Binance":
            num = get_setting("binance_id")

        sdata["method"] = method
        set_user_state(user_id, "WAITING_DEP_TRX", sdata)

        text = (
            f"📱 <b>{method} পেমেন্ট নির্দেশিকা:</b>\n"
            f"─────────────────────────────\n"
            f"💰 ডিপোজিট পরিমাণ: <b>{amount:.2f} BDT</b>\n"
            f"📌 পেমেন্ট নাম্বার / আইডি (ক্লিক করে কপি করুন):\n"
            f"<code>{num}</code>\n\n"
            f"⚠️ <i>উপরের নাম্বারে <b>{amount:.2f} BDT</b> Send Money করার পর নিচে আপনার Transaction ID (TrxID) টি লিখে পাঠান:</i>"
        )
        await query.edit_message_text(text, parse_mode=ParseMode.HTML)
        return

    # ---------------- 2FA TOOL ----------------
    if data == "tool_2fa_prompt":
        set_user_state(user_id, "WAITING_2FA_KEY", {})
        await query.message.reply_text(
            "🔑 <b>2FA Authenticator:</b>\nআপনার 2FA Secret Key টি এখানে পেস্ট করে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    # ---------------- PROFILE ACTIONS ----------------
    if data == "prof_purchases":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, receipt_no, product_name, total_bdt, order_date FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 10", (user_id,))
        orders = c.fetchall()
        conn.close()
        if not orders:
            await query.message.reply_text("🛒 আপনার পূর্বের কোনো কেনাকাটার হিস্টোরি পাওয়া যায়নি।")
            return

        txt = "🛒 <b>আপনার কেনাকাটার তালিকা:</b>\nবিস্তারিত দেখতে ও রিসিট পেতে অর্ডারের ওপর ক্লিক করুন:"
        kb = []
        for o in orders:
            btn_title = f"🧾 {o['receipt_no']} — {o['product_name']} ({o['total_bdt']:.0f} Tk)"
            kb.append([InlineKeyboardButton(btn_title, callback_data=f"prof_view_order_{o['id']}")])
        kb.append([InlineKeyboardButton("< Back to Profile", callback_data="prof_back_root")])
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # View Specific Purchase Details (matching image_5b8b20.png)
    if data.startswith("prof_view_order_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()

        if not o:
            await query.message.reply_text("❌ অর্ডারটি পাওয়া যায়নি।")
            return

        # Parse credentials
        creds_display = ""
        if o["delivered_data"]:
            try:
                items = json.loads(o["delivered_data"]) if o["delivered_data"].startswith("[") else [o["delivered_data"]]
            except Exception:
                items = [o["delivered_data"]]

            for itm in items:
                parts = [p.strip() for p in itm.split(":") if p.strip()]
                if len(parts) >= 2:
                    creds_display += f"👤 <code>{parts[0]}</code>\n🔐 <code>{parts[1]}</code>\n"
                else:
                    creds_display += f"🔑 <code>{itm}</code>\n"
        else:
            creds_display = "<i>⏳ অর্ডারটি প্রসেসিং হচ্ছে, খুব শীঘ্রই ডেলিভারি দেওয়া হবে।</i>"

        date_only = o["order_date"].split()[0] if o["order_date"] else "N/A"
        txt = (
            f"⚡ <b>Purchase Details</b>\n\n"
            f"Type  : <b>{o['category_name']}</b>\n"
            f"Date  : <code>{date_only}</code>\n"
            f"Total : <b>{o['total_bdt']:.2f} BDT</b>\n"
            f"🧾 Receipt : <code>{o['receipt_no']}</code>\n\n"
            f"<b>Credentials:</b>\n"
            f"{creds_display}"
        )
        kb = [
            [InlineKeyboardButton("🧾 Resend Receipt PDF", callback_data=f"prof_resend_pdf_{o['id']}")],
            [InlineKeyboardButton("< Back", callback_data="prof_purchases")],
        ]
        await query.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Resend Receipt PDF
    if data.startswith("prof_resend_pdf_"):
        order_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ? AND user_id = ?", (order_id, user_id))
        o = c.fetchone()
        conn.close()

        if not o:
            await query.message.reply_text("অর্ডার পাওয়া যায়নি।")
            return

        bot_info = await query.bot.get_me()
        try:
            pdf_bio = generate_pdf_receipt(
                receipt_no=o["receipt_no"],
                date_str=o["order_date"],
                day_name=o["day_of_week"] or "N/A",
                customer_name=query.from_user.full_name,
                telegram_id=user_id,
                username=query.from_user.username or "",
                item_desc=o["product_name"],
                category_name=o["category_name"] or "GENERAL",
                qty=o["quantity"],
                unit_price=o["unit_price"] or o["total_bdt"],
                total_bdt=o["total_bdt"],
                bot_username=bot_info.username,
            )
            await query.message.reply_document(
                document=InputFile(pdf_bio, filename=f"Receipt_{o['receipt_no']}.pdf"),
                caption=f"🧾 <b>Official Payment Receipt:</b> <code>{o['receipt_no']}</code>",
                parse_mode=ParseMode.HTML,
            )
        except Exception as e:
            logger.error(f"Error resending receipt: {e}")
            await query.message.reply_text("❌ রিসিট তৈরি করতে সমস্যা হয়েছে।")
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

    # ---------------- ADMIN PANEL DISPATCHER ----------------
    if data.startswith("adm_"):
        if not is_admin(user_id):
            return
        await handle_admin_callbacks(query, data, context)
        return

# ------------------------------------------------------------------------------
# 8. PRODUCT CARD & PURCHASE ENGINE
# ------------------------------------------------------------------------------

async def render_purchase_summary(query_or_message, prod_id: int, quantity: int = 1, edit: bool = True):
    """Renders the interactive Purchase Summary card with live calculator buttons matching image_2e14ee.png."""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    SELECT p.*, s.name as sub_name, s.category_id, c.name as cat_name
    FROM products p
    JOIN subcategories s ON p.subcategory_id = s.id
    JOIN categories c ON s.category_id = c.id
    WHERE p.id = ?
    """, (prod_id,))
    p = c.fetchone()
    conn.close()

    if not p:
        return

    rate = p["price_bdt"]
    total = rate * quantity
    full_name = f"{p['sub_name']} — {p['name']}"

    text = (
        f"🧾 <b>PURCHASE SUMMARY</b>\n"
        f"─────────────────────────────\n"
        f"📁 <b>Category:</b> {p['cat_name']}\n"
        f"✉️ <b>Package:</b> {full_name}\n"
        f"─────────────────────────────\n"
        f"✅ <b>Quantity:</b> {quantity}\n"
        f"💲 <b>Rate:</b> {format_price(rate)} BDT\n"
        f"🛒 <b>Total Price:</b> {format_price(total)} BDT"
    )

    kb = [
        [
            InlineKeyboardButton("➖", callback_data=f"calc_sub_{p['id']}_{quantity}"),
            InlineKeyboardButton(f"{quantity}", callback_data=f"calc_noop_{p['id']}_{quantity}"),
            InlineKeyboardButton("➕", callback_data=f"calc_add_{p['id']}_{quantity}"),
        ],
        [
            InlineKeyboardButton("✏️ Enter quantity", callback_data=f"calc_custom_{p['id']}")
        ],
        [
            InlineKeyboardButton("Back", callback_data=f"usr_sub_{p['subcategory_id']}"),
            InlineKeyboardButton("Cancel", callback_data="usr_root_cats"),
        ],
        [
            InlineKeyboardButton("Confirm", callback_data=f"calc_confirm_{p['id']}_{quantity}")
        ]
    ]
    reply_markup = InlineKeyboardMarkup(kb)

    if edit and hasattr(query_or_message, "edit_message_text"):
        try:
            await query_or_message.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    if p["image_file_id"]:
        try:
            await query_or_message.reply_photo(photo=p["image_file_id"], caption=text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)
            return
        except Exception:
            pass

    await query_or_message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=reply_markup)

async def show_product_card(query, prod_id: int):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
    SELECT p.*, s.name as sub_name, s.category_id, c.name as cat_name
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
        f"💰 <b>Price:</b> {format_price(p['price_bdt'])} BDT / {p['price_usdt']:.4f} USDT per unit\n"
        f"📊 <b>Stock:</b> {stock_badge}\n\n"
        f"Select quantity:"
    )

    kb = [
        [
            InlineKeyboardButton(f"1× {format_price(p['price_bdt'])} BDT", callback_data=f"qty_{p['id']}_1"),
            InlineKeyboardButton(f"2× {format_price(p['price_bdt']*2)} BDT", callback_data=f"qty_{p['id']}_2"),
        ],
        [
            InlineKeyboardButton(f"3× {format_price(p['price_bdt']*3)} BDT", callback_data=f"qty_{p['id']}_3"),
            InlineKeyboardButton(f"4× {format_price(p['price_bdt']*4)} BDT", callback_data=f"qty_{p['id']}_4"),
        ],
        [
            InlineKeyboardButton(f"5× {format_price(p['price_bdt']*5)} BDT", callback_data=f"qty_{p['id']}_5"),
            InlineKeyboardButton(f"6× {format_price(p['price_bdt']*6)} BDT", callback_data=f"qty_{p['id']}_6"),
        ],
        [
            InlineKeyboardButton(f"7× {format_price(p['price_bdt']*7)} BDT", callback_data=f"qty_{p['id']}_7"),
            InlineKeyboardButton(f"8× {format_price(p['price_bdt']*8)} BDT", callback_data=f"qty_{p['id']}_8"),
        ],
        [
            InlineKeyboardButton(f"9× {format_price(p['price_bdt']*9)} BDT", callback_data=f"qty_{p['id']}_9"),
            InlineKeyboardButton(f"10× {format_price(p['price_bdt']*10)} BDT", callback_data=f"qty_{p['id']}_10"),
        ],
        [InlineKeyboardButton("✏️ Custom Quantity", callback_data=f"custqty_{p['id']}")],
        [InlineKeyboardButton("< Back", callback_data=f"usr_sub_{p['subcategory_id']}")],
    ]

    # If photo exists, send photo! If not, send clean text!
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
            logger.warning(f"Error sending photo: {e}")

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

    if user_balance < total_bdt:
        conn.close()
        shortage = total_bdt - user_balance
        await query.message.reply_text(
            f"❌ <b>অপর্যাপ্ত ব্যালেন্স!</b>\n\n"
            f"প্রয়োজনীয় টাকা: {total_bdt:.2f} BDT\n"
            f"আপনার ব্যালেন্স: {user_balance:.2f} BDT\n"
            f"ঘাটতি: {shortage:.2f} BDT\n\n"
            f"দয়া করে নিচের 💳 Deposit মেনু থেকে ব্যালেন্স রিচার্জ করুন।",
            parse_mode=ParseMode.HTML,
        )
        return

    now = datetime.datetime.now()
    now_str = now.strftime("%Y-%m-%d %I:%M:%S %p")
    days_bn = {"Monday": "সোমবার", "Tuesday": "মঙ্গলবার", "Wednesday": "বুধবার", "Thursday": "বৃহস্পতিবার", "Friday": "শুক্রবার", "Saturday": "শনিবার", "Sunday": "রবিবার"}
    day_name = days_bn.get(now.strftime("%A"), now.strftime("%A"))
    receipt_no = f"RCP-{secrets.token_hex(5).upper()}"
    full_product_name = f"{prod['sub_name']} ({prod['name']})"

    # Check available stock
    c.execute("SELECT id, content, stock_type FROM stock WHERE product_id = ? AND is_sold = 0 LIMIT ?", (prod_id, quantity))
    stock_rows = c.fetchall()

    # If stock is available -> Instant Auto-Delivery
    if len(stock_rows) >= quantity:
        c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
        stock_ids = [r["id"] for r in stock_rows]
        items_content = [r["content"] for r in stock_rows]
        c.execute(
            f"UPDATE stock SET is_sold = 1, sold_to = ?, sold_at = ? WHERE id IN ({','.join(['?']*len(stock_ids))})",
            [user_id, now_str] + stock_ids,
        )
        delivered_json = json.dumps(items_content)
        c.execute("""
        INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, delivered_data, order_date, day_of_week)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'completed', ?, ?, ?)
        """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, delivered_json, now_str, day_name))
        conn.commit()
        conn.close()

        bot_info = await query.bot.get_me()

        # 1. Send Formatted 1-Tap Copy Credentials Card matching screenshot
        card_text = format_credentials_card(items_content, full_product_name, prod["cat_name"])
        await query.message.reply_text(card_text, parse_mode=ParseMode.HTML)

        # 2. If Mail or bulk items, also attach Excel (.xlsx) file
        if "MAIL" in prod["cat_name"].upper() or len(items_content) > 1:
            try:
                excel_bio = generate_excel_delivery(items_content, full_product_name, receipt_no)
                await query.message.reply_document(
                    document=InputFile(excel_bio, filename=f"{receipt_no}_{prod['name'].replace(' ', '_')}.xlsx"),
                    caption=f"📂 <b>Excel Data Backup:</b> <code>{receipt_no}</code>",
                    parse_mode=ParseMode.HTML,
                )
            except Exception as e:
                logger.warning(f"Error sending excel: {e}")

        # Send Official PDF Receipt
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
            await query.message.reply_document(
                document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
                caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
                parse_mode=ParseMode.HTML,
            )
        except Exception as e:
            logger.error(f"Error sending PDF receipt: {e}")

        # Alert Admin
        try:
            await query.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=(
                    f"🔔 <b>নতুন প্রোডাক্ট সেল হয়েছে!</b>\n"
                    f"🧾 Receipt: <code>{receipt_no}</code>\n"
                    f"👤 কাস্টমার: {user_name} (ID: <code>{user_id}</code>)\n"
                    f"📦 প্রোডাক্ট: {full_product_name}\n"
                    f"🔢 পরিমাণ: {quantity} টি | মূল্য: {total_bdt:.2f} BDT"
                ),
                parse_mode=ParseMode.HTML,
            )
        except Exception:
            pass
        return

    # If out of stock / On-demand -> Create Pending Order with Full Notifications & PDF Receipt
    c.execute("UPDATE users SET balance_bdt = balance_bdt - ? WHERE user_id = ?", (total_bdt, user_id))
    c.execute("""
    INSERT INTO orders (receipt_no, user_id, product_id, product_name, category_name, quantity, unit_price, total_bdt, total_usdt, status, order_date, day_of_week)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending_delivery', ?, ?)
    """, (receipt_no, user_id, prod_id, full_product_name, prod["cat_name"], quantity, prod["price_bdt"], total_bdt, total_usdt, now_str, day_name))
    order_id = c.lastrowid
    conn.commit()
    conn.close()

    bot_info = await query.bot.get_me()

    # 1. Send Official Order Confirmation Card to Customer
    confirm_msg = (
        f"🎉 <b>আপনার অর্ডারটি সফলভাবে গ্রহণ করা হয়েছে!</b>\n"
        f"─────────────────────────────\n"
        f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
        f"🛡️ <b>প্রোডাক্ট:</b> {full_product_name}\n"
        f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মোট পরিশোধ:</b> {format_price(total_bdt)} BDT\n"
        f"📅 <b>তারিখ:</b> {now_str} ({day_name})\n"
        f"─────────────────────────────\n"
        f"⏳ <b>ডেলিভারি সময়:</b> আগামী ১৫ থেকে ৩০ মিনিটের মধ্যে এডমিন আপনার ইনবক্সে সার্ভিসটি ডেলিভারি পাঠাবে।"
    )
    await query.message.reply_text(confirm_msg, parse_mode=ParseMode.HTML)

    # 2. Attach Official PDF Receipt with Scannable QR Code to Customer
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
        await query.message.reply_document(
            document=InputFile(pdf_bio, filename=f"Receipt_{receipt_no}.pdf"),
            caption=f"🧾 <b>Official Payment Receipt:</b> <code>{receipt_no}</code>",
            parse_mode=ParseMode.HTML,
        )
    except Exception as e:
        logger.error(f"Error generating PDF receipt for pending order: {e}")

    # 3. Alert Super Admin with Action Buttons
    adm_kb = InlineKeyboardMarkup([
        [InlineKeyboardButton(f"🚀 Deliver Order #{order_id}", callback_data=f"adm_fulfill_{order_id}")],
        [InlineKeyboardButton(f"❌ Cancel & Refund #{order_id}", callback_data=f"adm_refund_{order_id}")]
    ])
    try:
        await query.bot.send_message(
            chat_id=SUPER_ADMIN_ID,
            text=(
                f"🔔 <b>নতুন ম্যানুয়াল অর্ডার প্রয়োজন!</b>\n"
                f"─────────────────────────────\n"
                f"🧾 <b>অর্ডার নম্বর:</b> #{order_id} (Receipt: <code>{receipt_no}</code>)\n"
                f"👤 <b>কাস্টমার:</b> {user_name} (ID: <code>{user_id}</code>)\n"
                f"📦 <b>প্রোডাক্ট:</b> {full_product_name}\n"
                f"🔢 <b>পরিমাণ:</b> {quantity} টি | <b>মূল্য:</b> {format_price(total_bdt)} BDT\n"
                f"📅 <b>সময়:</b> {now_str}\n"
                f"─────────────────────────────\n"
                f"👇 ডেলিভারি দিতে নিচের বাটনে চাপুন:"
            ),
            parse_mode=ParseMode.HTML,
            reply_markup=adm_kb,
        )
    except Exception as e:
        logger.error(f"Error alerting admin for order #{order_id}: {e}")

# ------------------------------------------------------------------------------
# 9. ADMIN PANEL (4-TIER NESTED MENUS FOR STOCK, BUTTONS, PHOTOS, USERS)
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
        f"সব কাজ নেস্টেড বাটন দিয়ে ৩-৪ ধাপে পরিচালনা করুন:\n\n"
        f"💳 পেন্ডিং ডিপোজিট: <b>{pend_dep}</b> টি\n"
        f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
    )
    kb = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("📥 Add Stock", callback_data="adm_stk_step1_cats"),
            InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_btn_step1_cats"),
        ],
        [
            InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_pho_step1_cats"),
            InlineKeyboardButton("💳 Payment Settings", callback_data="adm_pay_settings"),
        ],
        [
            InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
            InlineKeyboardButton(f"🛡️ Orders ({pend_ord})", callback_data="adm_pending_orders"),
        ],
        [
            InlineKeyboardButton("👥 User Management", callback_data="adm_usr_list_0"),
            InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
        ]
    ])
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)

async def handle_admin_callbacks(query, data: str, context: ContextTypes.DEFAULT_TYPE):
    user_id = query.from_user.id

    # Root Admin Dashboard
    if data == "adm_root":
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
            f"🛡️ পেন্ডিং ম্যানুয়াল অর্ডার: <b>{pend_ord}</b> টি"
        )
        kb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("📥 Add Stock", callback_data="adm_stk_step1_cats"),
                InlineKeyboardButton("🏷️ Manage Buttons & Prices", callback_data="adm_btn_step1_cats"),
            ],
            [
                InlineKeyboardButton("🖼️ Manage Product Photos", callback_data="adm_pho_step1_cats"),
                InlineKeyboardButton("💳 Payment Settings", callback_data="adm_pay_settings"),
            ],
            [
                InlineKeyboardButton(f"💳 Deposits ({pend_dep})", callback_data="adm_pending_deposits"),
                InlineKeyboardButton(f"🛡️ Orders ({pend_ord})", callback_data="adm_pending_orders"),
            ],
            [
                InlineKeyboardButton("👥 User Management", callback_data="adm_usr_list_0"),
                InlineKeyboardButton("📊 Shop Statistics", callback_data="adm_statistics"),
            ]
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    # ==================== 1. ADD STOCK FLOW (4-TIER NESTED) ====================
    # Step 1: Select Category
    if data == "adm_stk_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY id ASC")
        cats = c.fetchall()
        conn.close()
        kb = []
        for cat in cats:
            kb.append([InlineKeyboardButton(cat["name"], callback_data=f"adm_stk_step2_sub_{cat['id']}")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("📥 <b>[Add Stock — ধাপ ১]</b>\nযে ক্যাটাগরির স্টকে প্রোডাক্ট যোগ করবেন, সেটি সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 2: Select Brand / Subcategory
    if data.startswith("adm_stk_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        kb = []
        row = []
        for s in subs:
            row.append(InlineKeyboardButton(s["name"], callback_data=f"adm_stk_step3_prod_{s['id']}"))
            if len(row) == 2:
                kb.append(row)
                row = []
        if row:
            kb.append(row)
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_stk_step1_cats")])
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ২]</b>\n<b>{cat_name}</b> ➔ কোন ব্র্যান্ডে স্টক যোগ করবেন সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 3: Select Variant / Duration / Package
    if data.startswith("adm_stk_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} ➔ {meta['sub_name']}" if meta else "Select Variant"
        kb = []
        for p in prods:
            kb.append([InlineKeyboardButton(f"📦 {p['name']} ({format_price(p['price_bdt'])} Tk)", callback_data=f"adm_stk_step4_upload_{p['id']}")])
        back_cat = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back", callback_data=f"adm_stk_step2_sub_{back_cat}")])
        await query.edit_message_text(f"📥 <b>[Add Stock — ধাপ ৩]</b>\n<b>{title}</b> ➔ কোন প্যাকেজ/মেয়াদে স্টক দেবেন সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 4: Prompt Stock Upload (Excel or Text)
    if data.startswith("adm_stk_step4_upload_"):
        prod_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.name as prod_name, s.name as sub_name FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        set_user_state(user_id, "ADMIN_WAITING_STOCK_INPUT", {"prod_id": prod_id})
        await query.message.reply_text(
            f"📥 <b>[Add Stock — ধাপ ৪: ডেটা আপলোড]</b>\n"
            f"📦 নির্বাচিত প্রোডাক্ট: <b>{p['sub_name']} ({p['prod_name']})</b>\n"
            f"─────────────────────────────\n"
            f"• আপনি একটি এক্সেল (<code>.xlsx</code>) ফাইল পাঠাতে পারেন।\n"
            f"• অথবা চ্যাটে এক লাইনে একটি করে অ্যাকাউন্ট/প্রক্সি পেস্ট করে পাঠিয়ে দিন।\n\n"
            f"বট নিজে থেকেই সবগুলো গুনে স্টকে ঢুকিয়ে নেবে।",
            parse_mode=ParseMode.HTML,
        )
        return

    # ==================== 2. MANAGE BUTTONS & PRODUCTS (4-TIER NESTED) ====================
    # Step 1: Categories with Live Reordering (⬆️ / ⬇️)
    if data == "adm_btn_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name, position FROM categories ORDER BY position ASC, id ASC")
        cats = c.fetchall()
        conn.close()
        kb = []
        for idx, cat in enumerate(cats):
            up_btn = InlineKeyboardButton("⬆️", callback_data=f"adm_cat_up_{cat['id']}") if idx > 0 else InlineKeyboardButton("·", callback_data="adm_noop")
            down_btn = InlineKeyboardButton("⬇️", callback_data=f"adm_cat_down_{cat['id']}") if idx < len(cats) - 1 else InlineKeyboardButton("·", callback_data="adm_noop")
            kb.append([
                InlineKeyboardButton(f"📁 {cat['name']}", callback_data=f"adm_btn_step2_sub_{cat['id']}"),
                up_btn,
                down_btn,
            ])
        kb.append([InlineKeyboardButton("➕ Add New Category", callback_data="adm_btn_add_cat")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(
            "🏷️ <b>[বাটন ও প্রোডাক্ট পরিচালনা — ধাপ ১]</b>\n"
            "যে ক্যাটাগরি ম্যানেজ করবেন তাতে চাপুন। পজিশন বদলাতে ⬆️ বা ⬇️ ব্যবহার করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return

    # Step 2: Brands / Subcategories
    if data.startswith("adm_btn_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()

        cat_name = cat_row["name"] if cat_row else "Category"
        c.execute("SELECT id, name, position FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (cat_id,))
        subs = c.fetchall()
        kb = []
        for idx, s in enumerate(subs):
            up_btn = InlineKeyboardButton("⬆️", callback_data=f"adm_sub_up_{s['id']}") if idx > 0 else InlineKeyboardButton("·", callback_data="adm_noop")
            down_btn = InlineKeyboardButton("⬇️", callback_data=f"adm_sub_down_{s['id']}") if idx < len(subs) - 1 else InlineKeyboardButton("·", callback_data="adm_noop")
            kb.append([
                InlineKeyboardButton(f"📁 {s['name']}", callback_data=f"adm_btn_step3_prod_{s['id']}"),
                up_btn,
                down_btn,
            ])
        kb.append([InlineKeyboardButton(f"➕ Add New Brand in {cat_name}", callback_data=f"adm_btn_add_sub_{cat_id}")])
        kb.append([
            InlineKeyboardButton(f"✏️ Rename Category", callback_data=f"adm_ren_cat_{cat_id}"),
            InlineKeyboardButton(f"🗑️ Delete Category", callback_data=f"adm_del_cat_{cat_id}")
        ])
        kb.append([InlineKeyboardButton("< Back to Categories", callback_data="adm_btn_step1_cats")])
        await query.edit_message_text(f"🏷️ <b>[বাটন ও প্রোডাক্ট পরিচালনা — ধাপ ২]</b>\n<b>{cat_name}</b> ➔ ব্র্যান্ড বা সেকশন বেছে নিন (এডিট/ডিলিট করতে ব্র্যান্ডে চাপুন):", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 3: Variants / Plans list
    if data.startswith("adm_btn_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT s.name as sub_name, s.category_id, c.name as cat_name
        FROM subcategories s
        JOIN categories c ON s.category_id = c.id
        WHERE s.id = ?
        """, (sub_id,))
        meta = c.fetchone()
        c.execute("SELECT id, name, price_bdt FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()

        title = f"{meta['cat_name']} ➔ {meta['sub_name']}" if meta else "Section"
        brand_name = meta["sub_name"] if meta else "Brand"
        kb = []
        for p in prods:
            kb.append([InlineKeyboardButton(f"📦 {p['name']} ({format_price(p['price_bdt'])} Tk)", callback_data=f"adm_btn_step4_edit_{p['id']}")])
        kb.append([InlineKeyboardButton(f"➕ Add New Duration / Package in {brand_name}", callback_data=f"adm_btn_add_prod_{sub_id}")])
        kb.append([
            InlineKeyboardButton(f"✏️ Rename Brand ({brand_name})", callback_data=f"adm_ren_sub_{sub_id}"),
            InlineKeyboardButton(f"🗑️ Delete Brand ({brand_name})", callback_data=f"adm_del_sub_{sub_id}")
        ])
        back_cat = meta["category_id"] if meta else 1
        kb.append([InlineKeyboardButton("< Back to Brands", callback_data=f"adm_btn_step2_sub_{back_cat}")])
        await query.edit_message_text(f"🏷️ <b>[বাটন ও প্রোডাক্ট পরিচালনা — ধাপ ৩]</b>\n<b>{title}</b>\n\n• নতুন ভ্যারিয়েন্ট যোগ করতে বা ব্র্যান্ডের নাম পরিবর্তন/ডিলিট করতে নিচের অপশন বেছে নিন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Step 4: Variant Control Card (Price, Rename, Description, Photo, Delete)
    if data.startswith("adm_btn_step4_edit_"):
        prod_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT p.*, s.name as sub_name, s.category_id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE p.id = ?", (prod_id,))
        p = c.fetchone()
        conn.close()

        photo_status = "🖼️ নিজস্ব ছবি সেট করা আছে" if p["image_file_id"] else "⚪ কোনো ছবি নেই (ক্লিন টেক্সট শো করবে)"
        desc_txt = p["description"] if p["description"] else "কোনো কাস্টম বিবরণ দেওয়া নেই"

        txt = (
            f"📦 <b>{p['sub_name']} — {p['name']}</b>\n"
            f"─────────────────────────────\n"
            f"💰 <b>বর্তমান দাম:</b> <b>{p['price_bdt']:.2f} BDT</b> ({p['price_usdt']:.4f} USDT)\n"
            f"📸 <b>ফটো স্ট্যাটাস:</b> {photo_status}\n"
            f"📝 <b>কাস্টম বিবরণ:</b>\n<i>{desc_txt}</i>\n\n"
            f"কী পরিবর্তন করতে চান? নিচের বাটনে চাপ দিন:"
        )
        kb = [
            [InlineKeyboardButton("💰 Change Price (BDT)", callback_data=f"adm_chgprice_{p['id']}")],
            [InlineKeyboardButton("✏️ Rename Variant Name", callback_data=f"adm_rename_{p['id']}")],
            [InlineKeyboardButton("📝 Edit Custom Description", callback_data=f"adm_editdesc_{p['id']}")],
            [InlineKeyboardButton("📸 Change / Set Photo", callback_data=f"adm_setphoto_{p['id']}")],
            [InlineKeyboardButton("🗑️ Delete This Variant", callback_data=f"adm_delprod_{p['id']}")],
            [InlineKeyboardButton("< Back", callback_data=f"adm_btn_step3_prod_{p['subcategory_id']}")],
        ]
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    # Up / Down Reordering Handlers for Categories
    if data.startswith("adm_cat_up_") or data.startswith("adm_cat_down_"):
        is_up = "up" in data
        target_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, position FROM categories ORDER BY position ASC, id ASC")
        all_cats = c.fetchall()
        for idx, ct in enumerate(all_cats):
            if ct["id"] == target_id:
                swap_idx = idx - 1 if is_up else idx + 1
                if 0 <= swap_idx < len(all_cats):
                    other_id = all_cats[swap_idx]["id"]
                    p1 = ct["position"] or idx
                    p2 = all_cats[swap_idx]["position"] or swap_idx
                    if p1 == p2:
                        p1, p2 = idx, swap_idx
                    c.execute("UPDATE categories SET position = ? WHERE id = ?", (p2, target_id))
                    c.execute("UPDATE categories SET position = ? WHERE id = ?", (p1, other_id))
                    conn.commit()
                break
        conn.close()
        # Refresh Step 1
        data = "adm_btn_step1_cats"

    # Up / Down Reordering Handlers for Subcategories (Brands)
    if data.startswith("adm_sub_up_") or data.startswith("adm_sub_down_"):
        is_up = "up" in data
        target_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT category_id FROM subcategories WHERE id = ?", (target_id,))
        p_row = c.fetchone()
        parent_cat_id = p_row["category_id"] if p_row else 1

        c.execute("SELECT id, position FROM subcategories WHERE category_id = ? ORDER BY position ASC, id ASC", (parent_cat_id,))
        all_subs = c.fetchall()
        for idx, sb in enumerate(all_subs):
            if sb["id"] == target_id:
                swap_idx = idx - 1 if is_up else idx + 1
                if 0 <= swap_idx < len(all_subs):
                    other_id = all_subs[swap_idx]["id"]
                    p1 = sb["position"] or idx
                    p2 = all_subs[swap_idx]["position"] or swap_idx
                    if p1 == p2:
                        p1, p2 = idx, swap_idx
                    c.execute("UPDATE subcategories SET position = ? WHERE id = ?", (p2, target_id))
                    c.execute("UPDATE subcategories SET position = ? WHERE id = ?", (p1, other_id))
                    conn.commit()
                break
        conn.close()
        # Refresh Step 2
        data = f"adm_btn_step2_sub_{parent_cat_id}"

    if data == "adm_noop":
        return

    # Brand & Category Rename / Delete Callback Handlers
    if data.startswith("adm_ren_sub_"):
        sub_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_SUB", {"sub_id": sub_id})
        await query.message.reply_text("✏️ <b>ব্র্যান্ডের নতুন নাম লিখে পাঠান:</b>\n(যেমন: FR Outlook বা Nord VPN)", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_del_sub_"):
        sub_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT s.name as sub_name, s.category_id FROM subcategories s WHERE s.id = ?", (sub_id,))
        sub_row = c.fetchone()
        if sub_row:
            cat_id = sub_row["category_id"]
            c.execute("DELETE FROM stock WHERE product_id IN (SELECT id FROM products WHERE subcategory_id = ?)", (sub_id,))
            c.execute("DELETE FROM products WHERE subcategory_id = ?", (sub_id,))
            c.execute("DELETE FROM subcategories WHERE id = ?", (sub_id,))
            conn.commit()
            conn.close()
            await query.message.reply_text(f"🗑️ <b>ব্র্যান্ড '{sub_row['sub_name']}' এবং এর সব প্রোডাক্ট সফলভাবে মুছে ফেলা হয়েছে!</b>", parse_mode=ParseMode.HTML)
        else:
            conn.close()
            await query.message.reply_text("ব্র্যান্ড পাওয়া যায়নি।")
        return

    if data.startswith("adm_ren_cat_"):
        cat_id = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_REN_CAT", {"cat_id": cat_id})
        await query.message.reply_text("✏️ <b>ক্যাটাগরির নতুন নাম লিখে পাঠান:</b>\n(যেমন: ✉️ Hotmail/Outlook Services)", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_del_cat_"):
        cat_id = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT name FROM categories WHERE id = ?", (cat_id,))
        cat_row = c.fetchone()
        if cat_row:
            c.execute("DELETE FROM stock WHERE product_id IN (SELECT p.id FROM products p JOIN subcategories s ON p.subcategory_id = s.id WHERE s.category_id = ?)", (cat_id,))
            c.execute("DELETE FROM products WHERE subcategory_id IN (SELECT id FROM subcategories WHERE category_id = ?)", (cat_id,))
            c.execute("DELETE FROM subcategories WHERE category_id = ?", (cat_id,))
            c.execute("DELETE FROM categories WHERE id = ?", (cat_id,))
            conn.commit()
            conn.close()
            await query.message.reply_text(f"🗑️ <b>ক্যাটাগরি '{cat_row['name']}' এবং এর সব ব্র্যান্ড মুছে ফেলা হয়েছে!</b>", parse_mode=ParseMode.HTML)
        else:
            conn.close()
            await query.message.reply_text("ক্যাটাগরি পাওয়া যায়নি।")
        return

    # Actions: Price, Rename, Description, Delete
    if data.startswith("adm_chgprice_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PRICE", {"prod_id": prod_id})
        await query.message.reply_text("নতুন BDT দাম লিখুন (যেমন: 0.87 বা 0.65 বা 15 — ১ টাকার নিচের যেকোনো দশমিক দিতে পারবেন):")
        return

    if data.startswith("adm_rename_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_NAME", {"prod_id": prod_id})
        await query.message.reply_text("নতুন নাম বা মেয়াদ লিখুন (যেমন: 15 Days বা 1 GB):")
        return

    if data.startswith("adm_editdesc_"):
        prod_id = int(data.split("_")[2])
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
        await query.edit_message_text("✅ প্রোডাক্টটি সফলভাবে মুছে ফেলা হয়েছে।")
        return

    # Add Category / Brand / Variant Prompts
    if data == "adm_btn_add_cat":
        set_user_state(user_id, "ADMIN_WAITING_NEW_CAT_NAME", {})
        await query.message.reply_text("➕ <b>নতুন ক্যাটাগরি তৈরি:</b>\nক্যাটাগরির নাম লিখুন (যেমন: 🎮 Gaming Services):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_btn_add_sub_"):
        cat_id = int(data.split("_")[4])
        set_user_state(user_id, "ADMIN_WAITING_NEW_SUB_NAME", {"cat_id": cat_id})
        await query.message.reply_text("➕ <b>নতুন ব্র্যান্ড / সেকশন যোগ:</b>\nব্র্যান্ডের নাম লিখুন (যেমন: Privado VPN বা Luna Proxy):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_btn_add_prod_"):
        sub_id = int(data.split("_")[4])
        set_user_state(user_id, "ADMIN_WAITING_ADDPROD_NAME", {"sub_id": sub_id})
        await query.message.reply_text("➕ <b>নতুন মেয়াদ / সাইজ যোগ:</b>\nমেয়াদ বা প্যাকেজের নাম লিখুন (যেমন: 10 Days বা 500 MB):", parse_mode=ParseMode.HTML)
        return

    # ==================== 3. MANAGE PHOTOS (4-TIER NESTED) ====================
    if data == "adm_pho_step1_cats":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM categories ORDER BY id ASC")
        cats = c.fetchall()
        conn.close()
        kb = []
        for cat in cats:
            kb.append([InlineKeyboardButton(cat["name"], callback_data=f"adm_pho_step2_sub_{cat['id']}")])
        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ১]</b>\nক্যাটাগরি সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step2_sub_"):
        cat_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name FROM subcategories WHERE category_id = ? ORDER BY id ASC", (cat_id,))
        subs = c.fetchall()
        conn.close()
        kb = []
        for s in subs:
            kb.append([InlineKeyboardButton(f"📁 {s['name']}", callback_data=f"adm_pho_step3_prod_{s['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_pho_step1_cats")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ২]</b>\nব্র্যান্ড সিলেক্ট করুন:", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step3_prod_"):
        sub_id = int(data.split("_")[4])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, name, image_file_id FROM products WHERE subcategory_id = ? ORDER BY id ASC", (sub_id,))
        prods = c.fetchall()
        conn.close()
        kb = []
        for p in prods:
            icon = "🖼️" if p["image_file_id"] else "⚪"
            kb.append([InlineKeyboardButton(f"{icon} {p['name']}", callback_data=f"adm_pho_step4_view_{p['id']}")])
        kb.append([InlineKeyboardButton("< Back", callback_data="adm_pho_step1_cats")])
        await query.edit_message_text("🖼️ <b>[প্রোডাক্ট ফটো ম্যানেজার — ধাপ ৩]</b>\nযে ভ্যারিয়েন্টের ছবি সেট করবেন বেছে নিন (⚪ = ছবি নেই, 🖼️ = নিজস্ব ছবি আছে):", parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_pho_step4_view_"):
        prod_id = int(data.split("_")[4])
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
        kb.append([InlineKeyboardButton("< Back", callback_data=f"adm_pho_step3_prod_{p['subcategory_id']}")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_setphoto_"):
        prod_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_PHOTO_UPLOAD", {"prod_id": prod_id})
        await query.message.reply_text("📸 অনুগ্রহ করে এই প্রোডাক্টের জন্য যে ছবিটি দিতে চান, সেটি টেলিগ্রামের চ্যাটে ফটো আকারে সেন্ড করুন:")
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

    # ==================== 4. PAYMENT SETTINGS ====================
    if data == "adm_pay_settings":
        bkash = get_setting("bkash_number")
        nagad = get_setting("nagad_number")
        binance = get_setting("binance_id")
        rate = get_setting("usdt_rate", "125.0")
        min_man = get_setting("min_deposit_manual", "20.0")
        min_ins = get_setting("min_deposit_instant", "100.0")

        text = (
            f"💳 <b>পেমেন্ট সেটিংস ও কন্ট্রোল প্যানেল:</b>\n"
            f"─────────────────────────────\n"
            f"📱 <b>bKash Personal:</b> <code>{bkash}</code>\n"
            f"📱 <b>Nagad Personal:</b> <code>{nagad}</code>\n"
            f"💎 <b>Binance Info:</b> <code>{binance}</code>\n"
            f"⚡ <b>ইনস্ট্যান্ট অটো মিনিমাম:</b> <b>{float(min_ins):.0f} BDT</b>\n"
            f"📱 <b>ম্যানুয়াল মিনিমাম:</b> <b>{float(min_man):.0f} BDT</b>\n"
            f"💵 <b>USDT রেট:</b> 1 USDT = {rate} BDT\n\n"
            f"যেটি পরিবর্তন করতে চান, নিচের বাটনে চাপ দিন:"
        )
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 Edit bKash Number", callback_data="adm_set_bkash")],
            [InlineKeyboardButton("📱 Edit Nagad Number", callback_data="adm_set_nagad")],
            [InlineKeyboardButton("💎 Edit Binance Info", callback_data="adm_set_binance")],
            [InlineKeyboardButton("⚡ Edit Instant Min Deposit", callback_data="adm_set_mindep_instant")],
            [InlineKeyboardButton("💰 Edit Manual Min Deposit", callback_data="adm_set_mindep_manual")],
            [InlineKeyboardButton("💵 Edit USDT Exchange Rate", callback_data="adm_set_rate")],
            [InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")],
        ])
        await query.edit_message_text(text, parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_set_mindep_instant":
        set_user_state(user_id, "ADMIN_WAITING_MINDEP_INSTANT", {})
        await query.message.reply_text("⚡ ইনস্ট্যান্ট অটো পেমেন্টে সর্বনিম্ন ডিপোজিট কত টাকা হবে লিখুন (যেমন: 100):")
        return

    if data == "adm_set_mindep_manual":
        set_user_state(user_id, "ADMIN_WAITING_MINDEP_MANUAL", {})
        await query.message.reply_text("📱 ম্যানুয়াল সেন্ড মানিতে সর্বনিম্ন ডিপোজিট কত টাকা হবে লিখুন (যেমন: 20):")
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

    if data == "adm_set_mindep":
        set_user_state(user_id, "ADMIN_WAITING_MINDEP", {})
        await query.message.reply_text("সর্বনিম্ন ডিপোজিট কত টাকা হবে তা সংখ্যায় লিখুন (যেমন: 10, 20, 50):")
        return

    if data == "adm_set_rate":
        set_user_state(user_id, "ADMIN_WAITING_RATE", {})
        await query.message.reply_text("নতুন USDT এক্সচেঞ্জ রেট লিখুন (যেমন: 128.5):")
        return

    # ==================== 5. USER MANAGEMENT (PAGINATED & INTERACTIVE) ====================
    if data.startswith("adm_usr_list_"):
        page = int(data.split("_")[3])
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
            f"👥 <b>ইউজার ম্যানেজমেন্ট ও প্রোফাইল তালিকা</b>\n"
            f"─────────────────────────────\n"
            f"মোট রেজিস্টার্ড ইউজার: <b>{total_u}</b> জন\n"
            f"নিচের ইউজারের উপর ক্লিক করে প্রোফাইল ও ব্যালেন্স কন্ট্রোল করুন:"
        )

        kb = []
        for u in u_list:
            u_name = u["name"][:16] if u["name"] else "User"
            btn_title = f"👤 {u_name} (ID: {u['user_id']}) - {u['balance_bdt']:.0f} Tk"
            kb.append([InlineKeyboardButton(btn_title, callback_data=f"adm_usr_view_{u['user_id']}")])

        nav_row = []
        if page > 0:
            nav_row.append(InlineKeyboardButton("◀️ Previous", callback_data=f"adm_usr_list_{page - 1}"))
        if (offset + limit) < total_u:
            nav_row.append(InlineKeyboardButton("Next ▶️", callback_data=f"adm_usr_list_{page + 1}"))
        if nav_row:
            kb.append(nav_row)

        kb.append([InlineKeyboardButton("< Back to Dashboard", callback_data="adm_root")])
        await query.edit_message_text(txt, parse_mode=ParseMode.HTML, reply_markup=InlineKeyboardMarkup(kb))
        return

    if data.startswith("adm_usr_view_"):
        target_uid = int(data.split("_")[3])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE user_id = ?", (target_uid,))
        u = c.fetchone()
        c.execute("SELECT COUNT(*) as cnt, SUM(total_bdt) as spent FROM orders WHERE user_id = ? AND status = 'completed'", (target_uid,))
        stats = c.fetchone()
        c.execute("SELECT receipt_no, product_name, quantity, total_bdt, order_date FROM orders WHERE user_id = ? ORDER BY id DESC LIMIT 10", (target_uid,))
        all_orders = c.fetchall()
        conn.close()

        if not u:
            await query.message.reply_text("❌ এই ইউজারকে পাওয়া যায়নি।")
            return

        spent = stats["spent"] or 0.0
        contact_url = f"tg://user?id={u['user_id']}"

        orders_txt = "\n📦 <b>অর্ডারের বিস্তারিত তালিকা (সর্বশেষ ১০টি):</b>\n"
        if all_orders:
            for ro in all_orders:
                orders_txt += f"• <code>{ro['receipt_no']}</code> | <b>{ro['product_name']}</b> ({ro['quantity']} টি) - <b>{ro['total_bdt']:.0f} Tk</b> [{ro['order_date']}]\n"
        else:
            orders_txt += "<i>এখনো কোনো অর্ডার করেনি।</i>\n"

        info = (
            f"👤 <b>ইউজার প্রোফাইল ও পূর্ণাঙ্গ হিসাব:</b>\n"
            f"─────────────────────────────\n"
            f"🆔 <b>আইডি:</b> <code>{u['user_id']}</code>\n"
            f"👤 <b>নাম:</b> {u['name']}\n"
            f"💬 <b>ইউজারনেম:</b> @{u['username'] if u['username'] else 'None'}\n"
            f"💰 <b>বর্তমান ব্যালেন্স:</b> <b>{u['balance_bdt']:.2f} BDT</b>\n"
            f"🛍️ <b>মোট সফল অর্ডার:</b> <b>{stats['cnt']} টি</b> (মোট খরচ: <b>{spent:.2f} BDT</b>)\n"
            f"📅 <b>জয়েনিং তারিখ:</b> {u['joined_date']}\n"
            f"─────────────────────────────"
            f"{orders_txt}"
        )
        ukb = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("➕ Add Balance", callback_data=f"adm_usr_addbal_{u['user_id']}"),
                InlineKeyboardButton("➖ Deduct Balance", callback_data=f"adm_usr_dedbal_{u['user_id']}"),
            ],
            [
                InlineKeyboardButton("💬 Contact User (Direct Chat)", url=contact_url),
            ],
            [
                InlineKeyboardButton("< Back to User List", callback_data="adm_usr_list_0"),
            ]
        ])
        await query.edit_message_text(info, parse_mode=ParseMode.HTML, reply_markup=ukb)
        return

    if data.startswith("adm_usr_addbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_ADDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💰 ইউজার <code>{target_uid}</code>-এর অ্যাকাউন্টে কত টাকা যোগ করবেন লিখুন (যেমন: 100):", parse_mode=ParseMode.HTML)
        return

    if data.startswith("adm_usr_dedbal_"):
        target_uid = int(data.split("_")[3])
        set_user_state(user_id, "ADMIN_WAITING_DEDBAL_AMT", {"target_uid": target_uid})
        await query.message.reply_text(f"💸 ইউজার <code>{target_uid}</code>-এর অ্যাকাউন্ট থেকে কত টাকা কাটবেন লিখুন (যেমন: 50):", parse_mode=ParseMode.HTML)
        return

    # ==================== 6. DEPOSITS & ORDERS FULFILLMENT ====================
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
            await query.message.reply_text("❌ ডিপোজিট পাওয়া যায়নি।")
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

    if data == "adm_pending_orders":
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT o.*, u.name as uname FROM orders o JOIN users u ON o.user_id = u.user_id WHERE o.status = 'pending_delivery'")
        ords = c.fetchall()
        conn.close()
        if not ords:
            await query.message.reply_text("✅ বর্তমানে কোনো পেন্ডিং অর্ডার নেই।")
            return
        for o in ords:
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"🚀 Deliver Order #{o['id']}", callback_data=f"adm_fulfill_{o['id']}")],
                [InlineKeyboardButton(f"❌ Cancel & Refund #{o['id']}", callback_data=f"adm_refund_{o['id']}")]
            ])
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

    if data.startswith("adm_fulfill_"):
        order_id = int(data.split("_")[2])
        set_user_state(user_id, "ADMIN_WAITING_MANUAL_DELIVERY", {"order_id": order_id})
        await query.message.reply_text(
            f"🚀 <b>Order #{order_id} ম্যানুয়াল ডেলিভারি:</b>\n"
            f"কাস্টমারকে পাঠানোর জন্য ভিপিএন/প্রক্সি/অ্যাকাউন্টের ডিটেইলস লিখে পাঠান:",
            parse_mode=ParseMode.HTML,
        )
        return

    if data.startswith("adm_refund_"):
        order_id = int(data.split("_")[2])
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM orders WHERE id = ?", (order_id,))
        o = c.fetchone()
        if o and o["status"] == "pending_delivery":
            c.execute("UPDATE orders SET status = 'refunded' WHERE id = ?", (order_id,))
            c.execute("UPDATE users SET balance_bdt = balance_bdt + ? WHERE user_id = ?", (o["total_bdt"], o["user_id"]))
            conn.commit()
            conn.close()
            await query.edit_message_text(f"❌ Order #{order_id} বাতিল করা হয়েছে এবং {o['total_bdt']:.2f} BDT ইউজারের অ্যাকাউন্টে রিফান্ড করা হয়েছে।")
            try:
                await context.bot.send_message(
                    chat_id=o["user_id"],
                    text=f"⚠️ আপনার Order #{order_id} বাতিল করা হয়েছে এবং {o['total_bdt']:.2f} BDT ব্যালেন্সে রিফান্ড করা হয়েছে।",
                )
            except Exception:
                pass
        else:
            conn.close()
            await query.message.reply_text("অর্ডার পাওয়া যায়নি।")
        return

    # Batch Stock Alert Broadcast Menu
    if data == "adm_bcast_stock_menu":
        conn = get_db()
        c = conn.cursor()
        c.execute("""
        SELECT p.name as prod_name, s.name as sub_name, p.price_bdt, COUNT(st.id) as cnt
        FROM products p
        JOIN subcategories s ON p.subcategory_id = s.id
        JOIN stock st ON p.id = st.product_id AND st.is_sold = 0
        GROUP BY p.id, s.name, p.name, p.price_bdt
        HAVING COUNT(st.id) > 0
        """)
        rows = c.fetchall()
        conn.close()

        if not rows:
            await query.message.reply_text("⚠️ বর্তমানে কোনো স্টকে নতুন প্রোডাক্ট নেই। আগে স্টক যোগ করুন।")
            return

        lines = ["🔔 <b>NEW STOCK ALERT!</b> 🔔", "────────────────────────"]
        for r in rows:
            lines.append(f"• <b>{r['sub_name']} {r['prod_name']}</b> — {r['cnt']} টি স্টকে আছে (দাম: {format_price(r['price_bdt'])} Tk)")
        lines.append("────────────────────────")
        lines.append("🛍️ দ্রুত কেনাকাটা করতে নিচে /start দিয়ে শপ ওপেন করুন!")
        alert_preview = "\n".join(lines)

        set_user_state(user_id, "ADMIN_CONFIRM_BCAST_TEXT", {"msg": alert_preview})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🚀 Send This Alert to All Users", callback_data="adm_bcast_execute_now")],
            [InlineKeyboardButton("< Cancel", callback_data="adm_root")],
        ])
        await query.message.reply_text(f"📢 <b>ইউজারদের কাছে এই নোটিশটি যাবে:</b>\n\n{alert_preview}", parse_mode=ParseMode.HTML, reply_markup=kb)
        return

    if data == "adm_bcast_execute_now":
        state, sdata = get_user_state(user_id)
        alert_text = sdata.get("msg", "")
        set_user_state(user_id, None)

        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT user_id FROM users")
        all_u = [r["user_id"] for r in c.fetchall()]
        conn.close()

        sent = 0
        for uid in all_u:
            try:
                await context.bot.send_message(chat_id=uid, text=alert_text, parse_mode=ParseMode.HTML)
                sent += 1
            except Exception:
                pass
        await query.message.reply_text(f"✅ সফলভাবে <b>{sent}</b> জন ইউজারের কাছে Stock Alert পৌঁছে দেওয়া হয়েছে!", parse_mode=ParseMode.HTML)
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

# ------------------------------------------------------------------------------
# 10. TEXT, PHOTO & DOCUMENT ROUTER (STATE MACHINE)
# ------------------------------------------------------------------------------
async def text_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text.strip()
    state, sdata = get_user_state(user.id)

    # Main Persistent Reply Keyboard Routing
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

    # 1. User Deposit Flow: Instant Auto Payment Handler
    if state == "WAITING_DEP_AMOUNT_INSTANT":
        min_ins = float(get_setting("min_deposit_instant", "100.0"))
        try:
            amt = float(text)
            if amt < min_ins:
                await update.message.reply_text(
                    f"❌ <b>ইনস্ট্যান্ট পেমেন্টে সর্বনিম্ন ডিপোজিট হলো {min_ins:.0f} টাকা।</b>\n"
                    f"{min_ins:.0f} বা তার বেশি টাকার পরিমাণ লিখুন:",
                    parse_mode=ParseMode.HTML,
                )
                return
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যায় টাকার পরিমাণ লিখুন (যেমন: 100):")
            return

        bot_info = await context.bot.get_me()
        await update.message.reply_text("⏳ আপনার সিকিউর পেমেন্ট লিংক তৈরি করা হচ্ছে, অনুগ্রহ করে অপেক্ষা করুন...")
        charge_res = create_paymently_charge(amt, user.id, user.full_name, bot_info.username)

        if charge_res and charge_res.get("status") and charge_res.get("payment_url"):
            pay_url = charge_res["payment_url"]
            inv_id = charge_res.get("invoice_id", "")
            set_user_state(user.id, None)

            txt = (
                f"💳 <b>ইনস্ট্যান্ট অটো পেমেন্ট (bKash / Nagad)</b>\n"
                f"─────────────────────────────\n"
                f"💰 <b>ডিপোজিট পরিমাণ:</b> <b>{amt:.2f} BDT</b>\n"
                f"🧾 <b>ইনভয়েস আইডি:</b> <code>{inv_id}</code>\n\n"
                f"📌 নিচের <b>Pay Now</b> বাটনে চাপ দিয়ে বিকাশ বা নগদে পেমেন্ট সম্পন্ন করুন।\n"
                f"পেমেন্ট শেষ হওয়ামাত্র কোনো TrxID দেওয়া ছাড়াই একা একা ব্যালেন্স যুক্ত হয়ে যাবে!\n\n"
                f"<i>(পেমেন্ট শেষে অটো যোগ না হলে 'Check Payment Status' বাটনে চাপ দিন)</i>"
            )
            kb = InlineKeyboardMarkup([
                [InlineKeyboardButton("💳 Pay Now (বিকাশ / নগদ)", url=pay_url)],
                [InlineKeyboardButton("🔄 Check Payment Status", callback_data=f"usr_checkpay_{inv_id}_{amt}")],
                [InlineKeyboardButton("< Cancel", callback_data="usr_root_cats")]
            ])
            await update.message.reply_text(txt, parse_mode=ParseMode.HTML, reply_markup=kb)
        else:
            await update.message.reply_text("⚠️ গেটওয়ে সার্ভারে সাময়িক সমস্যা হচ্ছে। দয়া করে Manual Send Money অপশনটি ব্যবহার করুন।")
        return

    # 2. User Deposit Flow: Manual Send Money Handler
    if state == "WAITING_DEP_AMOUNT_MANUAL":
        min_man = float(get_setting("min_deposit_manual", "20.0"))
        try:
            amt = float(text)
            if amt < min_man:
                await update.message.reply_text(
                    f"❌ <b>ম্যানুয়াল সেন্ড মানিতে সর্বনিম্ন ডিপোজিট হলো {min_man:.0f} টাকা।</b>\n"
                    f"{min_man:.0f} বা তার বেশি টাকার পরিমাণ লিখুন:",
                    parse_mode=ParseMode.HTML,
                )
                return
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যায় টাকার পরিমাণ লিখুন (যেমন: 20 বা 50):")
            return

        set_user_state(user.id, "WAITING_DEP_TRX", {"amount": amt})
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("📱 bKash", callback_data="usr_dep_method_bKash"), InlineKeyboardButton("📱 Nagad", callback_data="usr_dep_method_Nagad")],
            [InlineKeyboardButton("💎 Binance", callback_data="usr_dep_method_Binance")],
        ])
        await update.message.reply_text(
            f"✅ আপনি <b>{amt:.2f} BDT</b> ম্যানুয়াল ডিপোজিট করতে চেয়েছেন।\n\n"
            f"নিচের কোন মাধ্যমে সেন্ড মানি করবেন সিলেক্ট করুন:",
            parse_mode=ParseMode.HTML,
            reply_markup=kb,
        )
        return

    if state == "WAITING_DEP_TRX":
        trx = text
        method = sdata.get("method", "Manual")
        amount = sdata.get("amount", 20.0)
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
            f"✅ <b>ডিপোজিট রিকোয়েস্ট সাবমিটেড!</b>\n\n"
            f"আইডি: #{dep_id}\n"
            f"মেথড: {method}\n"
            f"পরিমাণ: {amount:.2f} BDT\n"
            f"TrxID: <code>{trx}</code>\n\n"
            f"এডমিন ভেরিফিকেশনের পর ব্যালেন্স যোগ হবে।",
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
        except Exception:
            pass
        return

    # 2. 2FA Secret Key Generator
    if state == "WAITING_2FA_KEY":
        secret = text.replace(" ", "").strip()
        try:
            totp = pyotp.TOTP(secret)
            code_2fa = totp.now()
            time_remaining = 30 - (int(time.time()) % 30)
            set_user_state(user.id, None)
            res_text = (
                f"🔑 <b>2FA Live Code</b>\n\n"
                f"🔢 কোড: <code>{code_2fa}</code> (ক্লিক করে কপি করুন)\n"
                f"⏳ মেয়াদ বাকি: <b>{time_remaining} সেকেন্ড</b>\n\n"
                f"📌 <i>প্রতি ৩০ সেকেন্ড পর পর কোড পরিবর্তিত হবে।</i>"
            )
            await update.message.reply_text(res_text, parse_mode=ParseMode.HTML)
        except Exception:
            await update.message.reply_text("❌ অবৈধ 2FA Secret Key! দয়া করে সঠিক কি দিন।")
        return

    # Interactive Calculator Custom Quantity
    if state == "WAITING_CALC_CUSTOM_QTY":
        try:
            qty = int(text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 5 বা 20):")
            return
        prod_id = sdata.get("prod_id")
        set_user_state(user.id, None)
        await render_purchase_summary(update.message, prod_id, quantity=qty, edit=False)
        return

    # 3. Custom Quantity
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

    # 4. Admin Action States
    if not is_admin(user.id):
        return

    # Stock Text Upload
    if state == "ADMIN_WAITING_STOCK_INPUT":
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

    # Manual Order Fulfillment
    if state == "ADMIN_WAITING_MANUAL_DELIVERY":
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
            await update.message.reply_text(f"✅ Order #{order_id} সফলভাবে ইউজারের কাছে ডেলিভারি পাঠানো হয়েছে।")

            user_msg = format_credentials_card([text], order['product_name'], order['category_name'])
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

    # Rename Subcategory / Brand
    if state == "ADMIN_WAITING_REN_SUB":
        sub_id = sdata.get("sub_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE subcategories SET name = ? WHERE id = ?", (text, sub_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ব্র্যান্ডের নাম সফলভাবে পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    # Rename Category
    if state == "ADMIN_WAITING_REN_CAT":
        cat_id = sdata.get("cat_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE categories SET name = ? WHERE id = ?", (text, cat_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ক্যাটাগরির নাম সফলভাবে পরিবর্তন হয়েছে: <b>{text}</b>", parse_mode=ParseMode.HTML)
        return

    # Add Category
    if state == "ADMIN_WAITING_NEW_CAT_NAME":
        conn = get_db()
        c = conn.cursor()
        c.execute("INSERT INTO categories (name, icon) VALUES (?, '🛍️')", (text,))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন ক্যাটাগরি <b>{text}</b> তৈরি সম্পন্ন হয়েছে!", parse_mode=ParseMode.HTML)
        return

    # Add Subcategory / Brand
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

    # Add Variant Name -> Prompt Price
    if state == "ADMIN_WAITING_ADDPROD_NAME":
        sdata["name"] = text
        set_user_state(user.id, "ADMIN_WAITING_ADDPROD_PRICE", sdata)
        await update.message.reply_text(
            f"মেয়াদ/প্যাকেজের নাম: <b>{text}</b>\n\n"
            f"এবার এর প্রতি পিসের BDT দাম লিখুন (যেমন: 0.87 বা 0.65 বা 15 — ১ টাকার নিচের দশমিক সংখ্যাও দিতে পারবেন):",
            parse_mode=ParseMode.HTML,
        )
        return

    if state == "ADMIN_WAITING_ADDPROD_PRICE":
        try:
            price_bdt = float(text)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 25):")
            return
        sub_id = sdata.get("sub_id")
        p_name = sdata.get("name")
        rate = float(get_setting("usdt_rate", "125.0"))
        price_usdt = price_bdt / rate if rate > 0 else 0.0

        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT c.name as cat_name FROM subcategories s JOIN categories c ON s.category_id = c.id WHERE s.id = ?", (sub_id,))
        row = c.fetchone()
        is_vpn = bool(row and "VPN" in row["cat_name"].upper())
        is_on_demand = 1 if is_vpn else 0

        c.execute(
            "INSERT INTO products (subcategory_id, name, price_bdt, price_usdt, is_on_demand) VALUES (?, ?, ?, ?, ?)",
            (sub_id, p_name, price_bdt, price_usdt, is_on_demand),
        )
        conn.commit()
        conn.close()

        set_user_state(user.id, None)
        await update.message.reply_text(
            f"🎉 <b>নতুন ভ্যারিয়েন্ট সফলভাবে তৈরি হয়েছে!</b>\n\n"
            f"📦 নাম: {p_name}\n"
            f"💰 দাম: {price_bdt:.2f} BDT ({price_usdt:.4f} USDT)",
            parse_mode=ParseMode.HTML,
        )
        return

    # Edit Description
    if state == "ADMIN_WAITING_DESC":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET description = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text("✅ এই প্রোডাক্টটির কাস্টম বিবরণ সফলভাবে সেভ করা হয়েছে!")
        return

    # Edit Price
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

    # Rename Variant
    if state == "ADMIN_WAITING_NAME":
        prod_id = sdata.get("prod_id")
        conn = get_db()
        c = conn.cursor()
        c.execute("UPDATE products SET name = ? WHERE id = ?", (text, prod_id))
        conn.commit()
        conn.close()
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ ভ্যারিয়েন্টের নাম পরিবর্তন হয়েছে: {text}")
        return

    # Edit Payment Settings
    if state == "ADMIN_WAITING_BKASH":
        set_setting("bkash_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন বিকাশ নাম্বার সেট হয়েছে: <code>{text}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_NAGAD":
        set_setting("nagad_number", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন নগদ নাম্বার সেট হয়েছে: <code>{text}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_BINANCE":
        set_setting("binance_id", text)
        set_user_state(user.id, None)
        await update.message.reply_text(f"✅ নতুন বাইনান্স তথ্য সেট হয়েছে: <code>{text}</code>", parse_mode=ParseMode.HTML)
        return

    if state == "ADMIN_WAITING_MINDEP_INSTANT":
        try:
            new_min = float(text)
            if new_min < 0:
                raise ValueError
            set_setting("min_deposit_instant", str(new_min))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ ইনস্ট্যান্ট অটো পেমেন্টে সর্বনিম্ন ডিপোজিট সীমা সেট হয়েছে: <b>{new_min:.0f} BDT</b>", parse_mode=ParseMode.HTML)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 100):")
        return

    if state == "ADMIN_WAITING_MINDEP_MANUAL":
        try:
            new_min = float(text)
            if new_min < 0:
                raise ValueError
            set_setting("min_deposit_manual", str(new_min))
            set_user_state(user.id, None)
            await update.message.reply_text(f"✅ ম্যানুয়াল সেন্ড মানিতে সর্বনিম্ন ডিপোজিট সীমা সেট হয়েছে: <b>{new_min:.0f} BDT</b>", parse_mode=ParseMode.HTML)
        except ValueError:
            await update.message.reply_text("দয়া করে সঠিক সংখ্যা লিখুন (যেমন: 20):")
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

    # Balance Add / Deduct
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

async def photo_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
        await update.message.reply_text("✅ এই প্রোডাক্টটির জন্য নতুন ফটো সফলভাবে সেভ করা হয়েছে! কাস্টমার এখন কেনার সময় এই ছবিটি দেখতে পাবে।")

async def document_message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_admin(user.id):
        return

    state, sdata = get_user_state(user.id)
    if state == "ADMIN_WAITING_STOCK_INPUT":
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
            f"📊 নতুন যুক্ত হয়েছে: <b>{len(items)}</b> টি অ্যাকাউন্ট/আইটেম\n"
            f"📦 স্টকে বর্তমান মোট সংখ্যা: <b>{total_now}</b> টি।",
            parse_mode=ParseMode.HTML,
        )

# ------------------------------------------------------------------------------
# 11. APPLICATION INITIALIZATION & RUNNER
# ------------------------------------------------------------------------------
async def post_init(application):
    global _GLOBAL_BOT_APP, _GLOBAL_EVENT_LOOP
    _GLOBAL_BOT_APP = application
    try:
        _GLOBAL_EVENT_LOOP = asyncio.get_running_loop()
    except Exception:
        _GLOBAL_EVENT_LOOP = asyncio.get_event_loop_policy().get_event_loop()
    logger.info("CyberNest Bot post_init completed successfully.")

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

    global _GLOBAL_BOT_APP, _GLOBAL_EVENT_LOOP
    try:
        loop = asyncio.get_event_loop_policy().get_event_loop()
    except Exception:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    _GLOBAL_EVENT_LOOP = loop

    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()
    _GLOBAL_BOT_APP = app

    # Handlers Registration
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CallbackQueryHandler(callback_query_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_message_handler))
    app.add_handler(MessageHandler(filters.PHOTO, photo_message_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_message_handler))

    print("Bot is polling for updates...")
    app.run_polling()

if __name__ == "__main__":
    main()
