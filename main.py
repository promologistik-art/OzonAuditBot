from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from config import BOT_TOKEN
from parser import merge_all
from generator import generate_excel
from io import BytesIO

user_files = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = """
👋 Привет! Я бот-аудитор Ozon.

📤 Загрузи 3 отчёта ОДНИМ сообщением:

1️⃣ Отчёт о начислениях (Финансы → Начисления)
2️⃣ Управление остатками (FBO → Управление остатками)
3️⃣ Аналитика продвижения (Продвижение → Аналитика)

⚠️ Все отчёты за ОДИНАКОВЫЙ период.
"""
    await update.message.reply_text(text)


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    doc = update.message.document
    
    if not doc.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Отправьте Excel-файл")
        return
    
    file = await context.bot.get_file(doc.file_id)
    data = BytesIO()
    await file.download_to_memory(data)
    data.seek(0)
    
    if user_id not in user_files:
        user_files[user_id] = []
    
    user_files[user_id].append(data)
    count = len(user_files[user_id])
    await update.message.reply_text(f"✅ Получен файл {count} из 3: {doc.file_name}")
    
    if count == 3:
        await update.message.reply_text("🔄 Обрабатываю...")
        try:
            df = merge_all(user_files[user_id][0], user_files[user_id][1], user_files[user_id][2])
            if df is None:
                await update.message.reply_text("❌ Ошибка: не удалось обработать файлы")
            else:
                out = generate_excel(df)
                await update.message.reply_document(
                    document=out,
                    filename="OzonAudit_Report.xlsx",
                    caption="📊 Отчёт готов! Заполните колонку «Себестоимость»."
                )
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка: {e}")
        finally:
            user_files.pop(user_id, None)


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    print("Бот запущен")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()