import os
import re
import json
from datetime import datetime
from io import BytesIO
import openpyxl

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

# ==================== কনফিগারেশন ====================
BOT_TOKEN = os.getenv("BOT_TOKEN", "8433977940:AAFAfplsbl5BROvS0k_yJ1Yy_nTvY-hYv2Q")
BOT_USERNAME = "NEW_FRESH_GMAILACCOUNTSELL50_bot"
ADMIN_ID = 8919985167  # আপনার টেলিগ্রাম ইউআইডি
SUPPORT_USERNAME = "@Talha_juba098"

DATA_FILE = "bot_users_data.json"
# ====================================================

def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def get_user_data(user_id, name="User"):
    data = load_data()
    uid = str(user_id)
    today = datetime.now().strftime("%Y-%m-%d")

    if uid not in data:
        data[uid] = {
            "id": uid,
            "name": name,
            "balance": 0.0,
            "state": "idle",
            "today_date": today,
            "today_count": 0,
            "history": [],
        }
    else:
        if data[uid].get("today_date") != today:
            data[uid]["today_date"] = today
            data[uid]["today_count"] = 0

    save_data(data)
    return data[uid]

def update_user_data(uid, key, value):
    data = load_data()
    uid = str(uid)
    if uid in data:
        data[uid][key] = value
        save_data(data)

def get_main_keyboard():
    return ReplyKeyboardMarkup(
        [
            ["নিউ ফ্রেশ জিমেইল অ্যাকাউন্ট সেল"],
            ["ব্যালেন্স", "উইথড্র"],
            ["প্রোফাইল", "হিস্ট্রি"],
            ["সাপোর্ট"],
        ],
        resize_keyboard=True,
    )

# --- ইউজার /start কমান্ড ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    get_user_data(user.id, user.first_name)
    update_user_data(user.id, "state", "idle")

    await update.message.reply_text(
        f"স্বাগতম {user.first_name}! 👋\nজিমেইল অ্যাকাউন্ট সেল করতে নিচের মেনু বাটন ব্যবহার করুন।",
        reply_markup=get_main_keyboard(),
    )

# --- অ্যাডমিন প্যানেল কমান্ড (/admin) ---
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("❌ আপনি এই বটের অ্যাডমিন নন।")
        return

    admin_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 বটের মোট তথ্য (Stats)", callback_data="adm_stats")],
        [InlineKeyboardButton("📢 সব ইউজারকে মেসেজ পাঠান (Broadcast)", callback_data="adm_broadcast")],
    ])

    await update.message.reply_text(
        "👑 **অ্যাডমিন কন্ট্রোল প্যানেল**\n\n"
        "💰 ব্যালেন্স যোগ করতে লিখুন:\n"
        "`/addbalance <User_ID> <পরিমাণ>`\n"
        "যেমন: `/addbalance 123456789 50`\n\n"
        "নিচের বাটন চেপে বাকি কাজ করুন:",
        reply_markup=admin_keyboard,
        parse_mode="Markdown"
    )

# --- অ্যাডমিন দ্বারা ব্যালেন্স অ্যাড করার কমান্ড ---
async def add_balance_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        return

    try:
        args = context.args
        if len(args) < 2:
            await update.message.reply_text("ব্যবহারের নিয়ম: `/addbalance <User_ID> <টাকা>`", parse_mode="Markdown")
            return

        target_uid = str(args[0])
        amount = float(args[1])

        all_data = load_data()
        if target_uid not in all_data:
            await update.message.reply_text("❌ এই ইউজার আইডিটি বটের ডেটাবেজে খুঁজে পাওয়া যায়নি।")
            return

        all_data[target_uid]["balance"] += amount
        save_data(all_data)

        # ইউজারকে নোটিফিকেশন পাঠানো
        try:
            await context.bot.send_message(
                chat_id=int(target_uid),
                text=f"🎉 আপনার অ্যাকাউন্টে {amount:.2f} BDT ব্যালেন্স যোগ করা হয়েছে!\nবর্তমান ব্যালেন্স: {all_data[target_uid]['balance']:.2f} BDT"
            )
        except Exception:
            pass

        await update.message.reply_text(f"✅ ইউজার `{target_uid}`-এর অ্যাকাউন্টে {amount:.2f} BDT যোগ করা হয়েছে।", parse_mode="Markdown")

    except ValueError:
        await update.message.reply_text("❌ টাকার পরিমাণ সঠিক সংখ্যায় লিখুন।")

# --- সাধারণ মেসেজ বাটন হ্যান্ডলার ---
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = update.message.text
    u_data = get_user_data(user.id, user.first_name)

    # ব্রডকাস্ট মেসেজ পাঠানো
    if user.id == ADMIN_ID and u_data.get("state") == "awaiting_broadcast":
        all_data = load_data()
        count = 0
        for uid in all_data.keys():
            try:
                await context.bot.send_message(chat_id=int(uid), text=f"📢 **অ্যাডমিন নোটিশ:**\n\n{text}", parse_mode="Markdown")
                count += 1
            except Exception:
                continue
        update_user_data(user.id, "state", "idle")
        await update.message.reply_text(f"✅ মোট {count} জন ইউজারের কাছে নোটিশ পাঠানো হয়েছে!")
        return

    if text == "নিউ ফ্রেশ জিমেইল অ্যাকাউন্ট সেল":
        if u_data["today_count"] >= 5:
            await update.message.reply_text("⚠️ আজকের জন্য আপনার সাবমিশন লিমিট শেষ! আপনি একদিনে সর্বোচ্চ ৫টি জিমেইল জমা দিতে পারবেন।")
            return

        update_user_data(user.id, "state", "waiting_file")
        remaining = 5 - u_data["today_count"]
        await update.message.reply_text(
            f"📁 অনুগ্রহ করে আপনার এক্সেল (.xlsx, .xls) অথবা টেক্সট (.txt, .csv) ফাইলটি পাঠান।\n\n"
            f"📌 নিয়মাবলী:\n"
            f"- ফাইলে শুধুমাত্র ভ্যালিড @gmail.com অ্যাকাউন্ট থাকতে হবে। অন্য কোনো মেইল গ্রহণযোগ্য নয়।\n"
            f"- আজ আপনি আর সর্বোচ্চ {remaining}টি জিমেইল জমা দিতে পারবেন।"
        )

    elif text == "ব্যালেন্স":
        await update.message.reply_text(f"💰 আপনার বর্তমান ব্যালেন্স: {u_data['balance']:.2f} BDT")

    elif text == "উইথড্র":
        await update.message.reply_text(f"💳 সর্বনিম্ন উইথড্র লিমিট ১০০ BDT। ব্যালেন্স পর্যাপ্ত হলে সরাসরি অ্যাডমিনের সাথে যোগাযোগ করুন: {SUPPORT_USERNAME}")

    elif text == "প্রোফাইল":
        info = (
            f"👤 ইউজার প্রোফাইল:\n\n"
            f"🆔 User ID: {user.id}\n"
            f"📛 Name: {u_data['name']}\n"
            f"💰 Balance: {u_data['balance']} BDT\n"
            f"📊 আজকের জমা: {u_data['today_count']}/5 টি"
        )
        await update.message.reply_text(info)

    elif text == "হিস্ট্রি":
        history = u_data.get("history", [])
        history_text = "\n".join(history[-10:]) if history else "আপনার কোনো পূর্ববর্তী হিস্ট্রি পাওয়া যায়নি।"
        await update.message.reply_text(f"📜 সাম্প্রতিক হিস্ট্রি:\n\n{history_text}")

    elif text == "সাপোর্ট":
        await update.message.reply_text(f"যেকোনো সাহায্য বা সাপোর্টের জন্য যোগাযোগ করুন: {SUPPORT_USERNAME}")

# --- ফাইল স্ক্যান ও ফিল্টারিং হ্যান্ডলার ---
async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    u_data = get_user_data(user.id, user.first_name)

    if u_data.get("state") != "waiting_file":
        await update.message.reply_text("দয়া করে আগে মেনু থেকে 'নিউ ফ্রেশ জিমেইল অ্যাকাউন্ট সেল' সিলেক্ট করুন।")
        return

    doc = update.message.document
    file_name = doc.file_name.lower()
    ext = file_name.split(".")[-1] if "." in file_name else ""

    if ext not in ["xlsx", "xls", "csv", "txt"]:
        await update.message.reply_text("❌ ভুল ফাইল ফরম্যাট! দয়া করে .xlsx, .xls, .csv অথবা .txt ফাইল দিন।")
        return

    file = await context.bot.get_file(doc.file_id)
    file_bytes = BytesIO()
    await file.download_to_memory(file_bytes)
    file_bytes.seek(0)

    raw_text = ""
    if ext == "xlsx":
        try:
            wb = openpyxl.load_workbook(file_bytes, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    row_vals = [str(v) for v in row if v is not None]
                    raw_text += " ".join(row_vals) + "\n"
        except Exception:
            await update.message.reply_text("❌ এক্সেল ফাইলটি পড়তে সমস্যা হয়েছে।")
            return
    else:
        try:
            raw_text = file_bytes.read().decode("utf-8", errors="ignore")
        except Exception:
            await update.message.reply_text("❌ ফাইলটি পড়া সম্ভব হয়নি।")
            return

    # শুধুমাত্র @gmail.com ও পাসওয়ার্ড ফিল্টার
    pattern = r"[a-zA-Z0-9._%+-]+@gmail\.com(?:[:\s,]+[^\s\r\n]+)?"
    matches = re.findall(pattern, raw_text, re.IGNORECASE)
    valid_accounts = list(set(matches))

    if not valid_accounts:
        await update.message.reply_text("❌ ফাইলে কোনো সঠিক @gmail.com একাউন্ট পাওয়া যায়নি! আলতো ফালতো মেইল গ্রহণ করা হবে না।")
        return

    total_found = len(valid_accounts)
    remaining_quota = 5 - u_data["today_count"]

    if total_found > remaining_quota:
        await update.message.reply_text(f"❌ ফাইলে মোট {total_found}টি জিমেইল রয়েছে! আপনি আজ সর্বোচ্চ {remaining_quota}টি সাবমিট করতে পারবেন।")
        return

    u_data["today_count"] += total_found
    u_data["state"] = "idle"
    save_data(load_data() | {str(user.id): u_data})

    await update.message.reply_text(f"✅ আপনার ফাইলটি স্ক্যান করে মোট {total_found}টি ভ্যালিড Gmail পাওয়া গেছে। ফাইলটি অ্যাডমিনের কাছে জমা হয়েছে।")

    # অ্যাডমিনকে (আপনাকে) মেসেজ পাঠানো
    preview = "\n".join(valid_accounts[:5])
    admin_msg = (
        f"🔔 নতুন জিমেইল সাবমিশন এসেছে!\n\n"
        f"👤 User ID: `{user.id}`\n"
        f"📛 Name: {user.first_name}\n"
        f"📊 মোট: {total_found} টি\n\n"
        f"প্রিভিউ:\n`{preview}`"
    )

    keyboard = [
        [
            InlineKeyboardButton("✅ Approve", callback_data=f"adm_app_{user.id}_{total_found}"),
            InlineKeyboardButton("❌ Reject", callback_data=f"adm_rej_{user.id}_{total_found}"),
        ]
    ]

    await context.bot.send_message(
        chat_id=ADMIN_ID,
        text=admin_msg,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode="Markdown"
    )

# --- অ্যাডমিন ইনলাইন বাটন হ্যান্ডলার ---
async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    admin_id = query.from_user.id
    data = query.data

    if admin_id != ADMIN_ID:
        await query.answer("পারমিশন নেই।", show_alert=True)
        return

    if data == "adm_stats":
        all_data = load_data()
        total_users = len(all_data)
        await query.answer()
        await query.message.reply_text(f"📊 **স্ট্যাটিস্টিকস:**\n\nমোট ইউজার সংখ্যা: {total_users} জন")
        return

    if data == "adm_broadcast":
        update_user_data(admin_id, "state", "awaiting_broadcast")
        await query.answer()
        await query.message.reply_text("📢 আপনি সকল ইউজারকে যে নোটিশটি পাঠাতে চান তা লিখে চ্যাটে সেন্ড করুন:")
        return

    if data.startswith("adm_app_") or data.startswith("adm_rej_"):
        _, action, target_user_id, count_str = data.split("_")
        count = int(count_str)
        all_data = load_data()
        target_user = all_data.get(target_user_id)

        if not target_user:
            await query.answer("ইউজার পাওয়া যায়নি।")
            return

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        if action == "app":
            target_user["history"].append(f"{now_str} - {count} টি জিমেইল Approved ✅")
            save_data(all_data)

            await context.bot.send_message(
                chat_id=int(target_user_id),
                text=f"🎉 আপনার সাবমিট করা {count}টি জিমেইল ভেরিফাই করে সফলভাবে Approved করা হয়েছে!",
            )
            await query.edit_message_text(query.message.text + f"\n\n───────────────\n✅ Status: APPROVED By Admin")

        elif action == "rej":
            target_user["today_count"] = max(0, target_user["today_count"] - count)
            target_user["history"].append(f"{now_str} - {count} টি জিমেইল Rejected ❌")
            save_data(all_data)

            await context.bot.send_message(
                chat_id=int(target_user_id),
                text="❌ আপনার সাবমিট করা জিমেইলগুলো অ্যাডমিন Reject করেছে। সঠিক ফরম্যাটে আবার চেষ্টা করুন।",
            )
            await query.edit_message_text(query.message.text + f"\n\n───────────────\n❌ Status: REJECTED By Admin")

        await query.answer("সম্পন্ন হয়েছে!")

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("addbalance", add_balance_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(CallbackQueryHandler(handle_callback))

    print(f"✅ @{BOT_USERNAME} সফলভাবে চালু হয়েছে...")
    app.run_polling()

if __name__ == "__main__":
    main()
