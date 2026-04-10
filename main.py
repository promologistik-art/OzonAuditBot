from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from config import BOT_TOKEN
from parser import merge_all
from generator import generate_output_excel
from io import BytesIO
import tempfile
import os

# Хранилище файлов (user_id -> список BytesIO)
user_files = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Приветствие и инструкция."""
    text = """
👋 Привет! Я бот-аудитор для селлеров Ozon.

📤 **Загрузи 3 отчёта ОДНИМ СООБЩЕНИЕМ:**

1️⃣ **Отчёт о начислениях**
   Финансы → Начисления → Скачать

2️⃣ **Управление остатками**
   FBO → Управление остатками → Скачать

3️⃣ **Аналитика продвижения**
   Продвижение → Аналитика → Скачать

⚠️ Важно: все отчёты должны быть за ОДИНАКОВЫЙ период.

После загрузки я пришлю Excel-файл с аналитикой.
В нём нужно будет заполнить колонку «Себестоимость».
"""
    await update.message.reply_text(text)


async def handle_files(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Приём файлов и обработка."""
    user_id = update.effective_user.id
    
    if not update.message.document:
        await update.message.reply_text("❌ Отправь файлы как документы (Excel).")
        return
    
    file = update.message.document
    
    # Проверка расширения
    if not file.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text(f"❌ {file.file_name} — не Excel-файл.")
        return
    
    # Скачиваем
    file_obj = await context.bot.get_file(file.file_id)
    file_bytes = BytesIO()
    await file_obj.download_to_memory(file_bytes)
    file_bytes.seek(0)
    
    # Сохраняем
    if user_id not in user_files:
        user_files[user_id] = []
    
    user_files[user_id].append(file_bytes)
    
    count = len(user_files[user_id])
    await update.message.reply_text(f"✅ Получен файл {count} из 3: {file.file_name}")
    
    # Если получены все 3 файла — обрабатываем
    if count == 3:
        await update.message.reply_text("🔄 Обрабатываю файлы...")
        
        try:
            f1, f2, f3 = user_files[user_id]
            df = merge_all(f1, f2, f3)
            output = generate_output_excel(df)
            
            await update.message.reply_document(
                document=output,
                filename="OzonAudit_Report.xlsx",
                caption="""
📊 Отчёт готов!

1. Открой файл в Excel
2. Заполни колонку «Себестоимость»
3. Колонка «Чистая_прибыль» пересчитается автоматически

Формула: Чистая прибыль = Выручка - Расход на рекламу - Себестоимость
"""
            )
            
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при обработке: {e}")
        
        finally:
            # Очищаем
            del user_files[user_id]


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена загрузки."""
    user_id = update.effective_user.id
    if user_id in user_files:
        del user_files[user_id]
    await update.message.reply_text("🔄 Загрузка отменена. Можешь начать заново с /start")


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_files))
    
    print("Бот запущен...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()