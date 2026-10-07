import os
import json
import threading
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

from telegram import (
    Update,
    ReplyKeyboardMarkup,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)

# =========================================================
# CONFIGURATION
# =========================================================

BOT_TOKEN = os.getenv("BOT_TOKEN")

# এখানে তোমার Telegram User ID বসাবে
ADMIN_ID = int(os.getenv("ADMIN_ID", "123456789"))

SUPPORT_USERNAME = "@your_support"

DATA_FILE = "bot_data.json"

# Demo products
PRODUCTS = {
    "VPN 1 Month": 100,
    "VPN 3 Months": 250,
    "Digital Service 1": 150,
    "Digital Service 2": 200,
}

# =========================================================
# DATABASE
# =========================================================

def load_data():
    if not os.path.exists(DATA_FILE):
        return {}

    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def get_user(user_id, name="User"):
    data = load_data()
    uid = str(user_id)

    if uid not in data:
        data[uid] = {
            "id": uid,
            "name": name,
            "balance": 0,
            "orders": []
        }
        save_data(data)

    return data[uid]


# =========================================================
# KEYBOARDS
# =========================================================

def main_keyboard():
    keyboard = [
        ["🛍️ প্রোডাক্ট", "💰 ব্যালেন্স"],
        ["📦 আমার অর্ডার", "👤 প্রোফাইল"],
        ["📞 সাপোর্ট"],
    ]

    return ReplyKeyboardMarkup(
        keyboard,
        resize_keyboard=True
    )


def product_keyboard():
    buttons = []

    for product, price in PRODUCTS.items():
        buttons.append([
            InlineKeyboardButton(
                f"{product} - {price} BDT",
                callback_data=f"product:{product}"
            )
        ])

    return InlineKeyboardMarkup(buttons)


# =========================================================
# START
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.effective_user

    get_user(
        user.id,
        user.first_name or "User"
    )

    await update.message.reply_text(
        f"স্বাগতম {user.first_name or 'User'}! 👋\n\n"
        "এটি একটি ডেমো ডিজিটাল সার্ভিস শপ।\n"
        "নিচের মেনু থেকে একটি অপশন নির্বাচন করুন।",
        reply_markup=main_keyboard()
    )


# =========================================================
# MESSAGE HANDLER
# =========================================================

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):

    user = update.effective_user
    text = update.message.text

    user_data = get_user(
        user.id,
        user.first_name or "User"
    )

    # -----------------------------------------------------
    # PRODUCTS
    # -----------------------------------------------------

    if text == "🛍️ প্রোডাক্ট":

        await update.message.reply_text(
            "🛍️ আমাদের প্রোডাক্ট:\n\n"
            "নিচের যেকোনো একটি প্রোডাক্ট নির্বাচন করুন।",
            reply_markup=product_keyboard()
        )

    # -----------------------------------------------------
    # BALANCE
    # -----------------------------------------------------

    elif text == "💰 ব্যালেন্স":

        balance = user_data["balance"]

        await update.message.reply_text(
            f"💰 আপনার বর্তমান ব্যালেন্স:\n\n"
            f"{balance} BDT"
        )

    # -----------------------------------------------------
    # ORDERS
    # -----------------------------------------------------

    elif text == "📦 আমার অর্ডার":

        orders = user_data.get("orders", [])

        if not orders:
            await update.message.reply_text(
                "📦 আপনার এখনো কোনো অর্ডার নেই।"
            )
            return

        message = "📦 আপনার অর্ডার হিস্ট্রি:\n\n"

        for order in orders[-10:]:
            message += (
                f"🛍️ {order['product']}\n"
                f"💰 Price: {order['price']} BDT\n"
                f"📊 Status: {order['status']}\n"
                f"🕒 {order['time']}\n\n"
            )

        await update.message.reply_text(message)

    # -----------------------------------------------------
    # PROFILE
    # -----------------------------------------------------

    elif text == "👤 প্রোফাইল":

        await update.message.reply_text(
            f"👤 আপনার প্রোফাইল\n\n"
            f"🆔 User ID: {user.id}\n"
            f"📛 Name: {user_data['name']}\n"
            f"💰 Balance: {user_data['balance']} BDT\n"
            f"📦 Orders: {len(user_data['orders'])}"
        )

    # -----------------------------------------------------
    # SUPPORT
    # -----------------------------------------------------

    elif text == "📞 সাপোর্ট":

        await update.message.reply_text(
            f"📞 সাপোর্টের জন্য যোগাযোগ করুন:\n\n"
            f"{SUPPORT_USERNAME}"
        )


# =========================================================
# PRODUCT CALLBACK
# =========================================================

async def handle_product(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    product_name = query.data.replace("product:", "")

    if product_name not in PRODUCTS:
        await query.edit_message_text(
            "❌ প্রোডাক্ট পাওয়া যায়নি।"
        )
        return

    price = PRODUCTS[product_name]

    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "✅ অর্ডার করুন",
                callback_data=f"order:{product_name}"
            )
        ],
        [
            InlineKeyboardButton(
                "❌ বাতিল",
                callback_data="cancel"
            )
        ]
    ])

    await query.edit_message_text(
        f"🛍️ প্রোডাক্ট:\n"
        f"{product_name}\n\n"
        f"💰 মূল্য: {price} BDT\n\n"
        "আপনি কি এই প্রোডাক্টটি অর্ডার করতে চান?",
        reply_markup=keyboard
    )


# =========================================================
# ORDER
# =========================================================

async def handle_order(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    user = query.from_user

    product_name = query.data.replace("order:", "")
    price = PRODUCTS.get(product_name)

    if price is None:
        await query.edit_message_text(
            "❌ প্রোডাক্ট পাওয়া যায়নি।"
        )
        return

    data = load_data()
    uid = str(user.id)

    user_data = data.get(uid)

    if not user_data:
        user_data = get_user(
            user.id,
            user.first_name or "User"
        )
        data = load_data()

    balance = user_data["balance"]

    # -----------------------------------------------------
    # BALANCE CHECK
    # -----------------------------------------------------

    if balance < price:

        await query.edit_message_text(
            f"❌ পর্যাপ্ত ব্যালেন্স নেই।\n\n"
            f"💰 আপনার ব্যালেন্স: {balance} BDT\n"
            f"💵 প্রয়োজন: {price} BDT\n\n"
            f"সাপোর্ট: {SUPPORT_USERNAME}"
        )

        return

    # -----------------------------------------------------
    # CREATE ORDER
    # -----------------------------------------------------

    order_id = len(user_data["orders"]) + 1

    order = {
        "id": order_id,
        "product": product_name,
        "price": price,
        "status": "Pending",
        "time": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )
    }

    user_data["balance"] -= price
    user_data["orders"].append(order)

    data[uid] = user_data
    save_data(data)

    await query.edit_message_text(
        f"✅ আপনার অর্ডার গ্রহণ করা হয়েছে!\n\n"
        f"🛍️ Product: {product_name}\n"
        f"💰 Price: {price} BDT\n"
        f"🆔 Order ID: #{order_id}\n"
        f"📊 Status: Pending\n\n"
        "অ্যাডমিন অর্ডারটি রিভিউ করবে।"
    )

    # -----------------------------------------------------
    # ADMIN NOTIFICATION
    # -----------------------------------------------------

    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "✅ Approve",
                callback_data=f"approve:{user.id}:{order_id}"
            ),
            InlineKeyboardButton(
                "❌ Reject",
                callback_data=f"reject:{user.id}:{order_id}"
            )
        ]
    ])

    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text=(
            "🔔 নতুন অর্ডার!\n\n"
            f"👤 User: {user.first_name or 'User'}\n"
            f"🆔 User ID: {user.id}\n"
            f"🛍️ Product: {product_name}\n"
            f"💰 Price: {price} BDT\n"
            f"🆔 Order ID: #{order_id}"
        ),
        reply_markup=keyboard
    )


# =========================================================
# ADMIN APPROVE / REJECT
# =========================================================

async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    if query.from_user.id != ADMIN_ID:

        await query.answer(
            "❌ আপনার এই অ্যাকশন করার অনুমতি নেই।",
            show_alert=True
        )

        return

    parts = query.data.split(":")

    action = parts[0]
    user_id = parts[1]
    order_id = int(parts[2])

    data = load_data()

    if user_id not in data:
        await query.edit_message_text(
            "❌ ইউজার পাওয়া যায়নি।"
        )
        return

    user_data = data[user_id]

    target_order = None

    for order in user_data["orders"]:
        if order["id"] == order_id:
            target_order = order
            break

    if target_order is None:
        await query.edit_message_text(
            "❌ অর্ডার পাওয়া যায়নি।"
        )
        return

    # -----------------------------------------------------
    # APPROVE
    # -----------------------------------------------------

    if action == "approve":

        target_order["status"] = "Approved"

        save_data(data)

        await context.bot.send_message(
            chat_id=int(user_id),
            text=(
                "🎉 আপনার অর্ডার Approved হয়েছে!\n\n"
                f"🛍️ {target_order['product']}\n"
                f"🆔 Order ID: #{order_id}"
            )
        )

        await query.edit_message_text(
            query.message.text +
            "\n\n✅ STATUS: APPROVED"
        )

    # -----------------------------------------------------
    # REJECT
    # -----------------------------------------------------

    elif action == "reject":

        target_order["status"] = "Rejected"

        # টাকা ফেরত
        user_data["balance"] += target_order["price"]

        save_data(data)

        await context.bot.send_message(
            chat_id=int(user_id),
            text=(
                "❌ আপনার অর্ডার Reject করা হয়েছে।\n\n"
                f"🛍️ {target_order['product']}\n"
                f"💰 {target_order['price']} BDT আপনার ব্যালেন্সে ফেরত দেওয়া হয়েছে।"
            )
        )

        await query.edit_message_text(
            query.message.text +
            "\n\n❌ STATUS: REJECTED\n"
            "💰 টাকা ফেরত দেওয়া হয়েছে।"
        )


# =========================================================
# CANCEL
# =========================================================

async def cancel_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):

    query = update.callback_query
    await query.answer()

    await query.edit_message_text(
        "❌ অর্ডার বাতিল করা হয়েছে।"
    )


# =========================================================
# RENDER HEALTH CHECK SERVER
# =========================================================

class HealthCheckHandler(BaseHTTPRequestHandler):

    def do_GET(self):

        self.send_response(200)
        self.end_headers()

        self.wfile.write(
            b"Telegram Shop Bot is running!"
        )

    def log_message(self, format, *args):
        return


def run_web_server():

    port = int(
        os.environ.get("PORT", 10000)
    )

    server = HTTPServer(
        ("0.0.0.0", port),
        HealthCheckHandler
    )

    server.serve_forever()


# =========================================================
# MAIN
# =========================================================

def main():

    if not BOT_TOKEN:
        print("❌ BOT_TOKEN পাওয়া যায়নি!")
        return

    print("🚀 Bot starting...")

    # Render web server
    threading.Thread(
        target=run_web_server,
        daemon=True
    ).start()

    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .build()
    )

    app.add_handler(
        CommandHandler("start", start)
    )

    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            handle_message
        )
    )

    app.add_handler(
        CallbackQueryHandler(
            handle_product,
            pattern=r"^product:"
        )
    )

    app.add_handler(
        CallbackQueryHandler(
            handle_order,
            pattern=r"^order:"
        )
    )

    app.add_handler(
        CallbackQueryHandler(
            admin_callback,
            pattern=r"^(approve|reject):"
        )
    )

    app.add_handler(
        CallbackQueryHandler(
            cancel_callback,
            pattern=r"^cancel$"
        )
    )

    print("✅ Bot is running!")

    app.run_polling()


if __name__ == "__main__":
    main()
