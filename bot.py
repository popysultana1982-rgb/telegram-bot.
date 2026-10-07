import os
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
import telebot


# Telegram Bot Token
TOKEN = os.getenv("BOT_TOKEN")

bot = telebot.TeleBot(TOKEN)


# /start command
@bot.message_handler(commands=["start"])
def start(message):
    bot.reply_to(message, "Hello! Bot is working.")


# Reply to normal messages
@bot.message_handler(func=lambda message: True)
def reply(message):
    bot.reply_to(message, "I received your message!")


# Simple web server for Render
class HealthCheckHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Telegram Bot is running!")

    def log_message(self, format, *args):
        return


def run_web_server():
    port = int(os.environ.get("PORT", 10000))

    server = HTTPServer(
        ("0.0.0.0", port),
        HealthCheckHandler
    )

    server.serve_forever()


# Start web server
threading.Thread(
    target=run_web_server,
    daemon=True
).start()


# Start Telegram bot
bot.infinity_polling()
