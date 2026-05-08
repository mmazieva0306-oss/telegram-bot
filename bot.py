import logging
import os
from datetime import datetime
from flask import Flask, request
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, ConversationHandler, CallbackQueryHandler, ContextTypes, filters
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# ==== ЭТА ЧАСТЬ КОДА (ОБРАБОТЧИКИ КОМАНД) ОСТАЕТСЯ БЕЗ ИЗМЕНЕНИЙ ====
# ==== ВАШИ СТАНДАРТНЫЕ ФУНКЦИЯ: start, get_region, get_product и т.д. ====
# ==== (Вставьте сюда ваш существующий код от start до cancel) ====
# ============================================================
# ... (Здесь должен быть весь ваш код с async def start, get_region и т.д.) ...
# ============================================================

# ============================================================
# ЗАПУСК С ВЕБХУКОМ (ЭТО НОВАЯ ЧАСТЬ)
# ============================================================

# Создаем Flask-приложение для приема сигналов от Render
flask_app = Flask(__name__)
# Глобальная переменная для бота
telegram_app = None

@flask_app.route(f'/{BOT_TOKEN}', methods=['POST'])
async def webhook():
    """Точка входа для обновлений от Telegram"""
    try:
        # Получаем данные от Telegram
        update = Update.de_json(request.get_json(force=True), telegram_app.bot)
        # Отправляем их в нашего бота
        await telegram_app.process_update(update)
        return 'OK', 200
    except Exception as e:
        logging.error(f"Ошибка вебхука: {e}")
        return 'Error', 500

@flask_app.route('/')
def health_check():
    """Стандартная проверка для Render"""
    return 'Бот работает', 200

def setup_webhook():
    """Настраивает связь между Telegram и нашим Flask"""
    # Render сам подставляет свой внешний адрес в переменную окружения
    render_url = os.environ.get('RENDER_EXTERNAL_URL')
    if not render_url:
        logging.warning("RENDER_EXTERNAL_URL не найден")
        return False
    
    # Формируем полный адрес для вебхука
    webhook_url = f"{render_url}/{BOT_TOKEN}"
    
    # Говорим Telegram: "Присылай все обновления сюда"
    telegram_app.bot.set_webhook(url=webhook_url, drop_pending_updates=True)
    logging.info(f"✅ Вебхук установлен: {webhook_url}")
    return True

def main():
    global telegram_app
    
    # Готовим файл Excel
    init_excel()
    
    # Создаем приложение Telegram
    telegram_app = Application.builder().token(BOT_TOKEN).build()
    
    # Регистрируем обработчики (они у вас уже есть в коде выше)
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            REGION: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_region)],
            PRODUCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_product)],
            PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_price)],
            VOLUME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_volume)],
            CONTACT: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_contact)],
            CONFIRM: [CallbackQueryHandler(confirm_order)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    telegram_app.add_handler(conv_handler)
    telegram_app.add_handler(CommandHandler("excel", send_excel))
    
    # Настраиваем вебхук вместо polling
    if setup_webhook():
        # Запускаем Flask сервер для прослушивания порта
        port = int(os.environ.get('PORT', 8080))
        logging.info(f"🚀 Запуск Flask сервера на порту {port}")
        flask_app.run(host='0.0.0.0', port=port)
    else:
        # Если вебхук не настроился (локальная отладка), запускаем polling
        logging.warning("Запуск в режиме polling")
        telegram_app.run_polling()

if __name__ == "__main__":
    main()
