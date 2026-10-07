import os
import telebot

TOKEN = os.getenv("BOT_TOKEN")

bot = telebot.TeleBot(TOKEN)

@bot.message_handler(commands=["start"])
def start(message):
    bot.reply_to(message, "Hello! Bot is working.")

@bot.message_handler(func=lambda message: True)
def reply(message):
    bot.reply_to(message, "I received your message!")

bot.infinity_polling()
