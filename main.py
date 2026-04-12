from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler
from config import BOT_TOKEN
from parser import parse_accruals, parse_stock, parse_ads, merge_three
from generator import generate_excel
from io import BytesIO

# Храним состояние: user_id -> {accruals: BytesIO, stock: BytesIO, ads: BytesIO}
user_files = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("1️⃣ Загрузить отчёт о начислениях", callback_data="upload_accruals")],
        [InlineKeyboardButton("2️⃣ Загрузить управление остатками", callback_data="upload_stock")],
        [InlineKeyboardButton("3️⃣ Загрузить аналитику продвижения", callback_data="upload_ads")],
        [InlineKeyboardButton("📊 Сформировать отчёт", callback_data="generate")],
    ]
    await update.message.reply_text(
        "👋 Загрузите 3 отчёта Ozon.\n\n"
        "Нажимайте кнопки по очереди и прикрепляйте файл.\n"
        "Когда все 3 файла будут загружены — нажмите «Сформировать отчёт».",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    data = query.data
    
    if data.startswith("upload_"):
        file_type = data.replace("upload_", "")
        context.user_data["awaiting_file"] = file_type
        
        names = {
            "accruals": "Отчёт о начислениях",
            "stock": "Управление остатками",
            "ads": "Аналитика продвижения"
        }
        await query.edit_message_text(
            f"📎 Прикрепите файл: **{names[file_type]}**\n\n"
            "Отправьте Excel-файл (.xlsx или .xls).",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("◀️ Назад", callback_data="back")]
            ])
        )
    
    elif data == "back":
        await start(update, context)
    
    elif data == "generate":
        if user_id not in user_files or len(user_files[user_id]) < 3:
            missing = []
            if user_id not in user_files:
                missing = ["начисления", "остатки", "реклама"]
            else:
                if "accruals" not in user_files[user_id]:
                    missing.append("начисления")
                if "stock" not in user_files[user_id]:
                    missing.append("остатки")
                if "ads" not in user_files[user_id]:
                    missing.append("реклама")
            
            await query.edit_message_text(
                f"❌ Не хватает файлов: {', '.join(missing)}",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data="back")]
                ])
            )
            return
        
        await query.edit_message_text("🔄 Обрабатываю файлы...")
        
        try:
            files = user_files[user_id]
            df = merge_three(
                files["accruals"],
                files["stock"],
                files["ads"]
            )
            
            if df is None:
                await query.edit_message_text(
                    "❌ Ошибка при обработке файлов. Проверьте, что загружены правильные отчёты.",
                    reply_markup=InlineKeyboardMarkup([
                        [InlineKeyboardButton("◀️ Назад", callback_data="back")]
                    ])
                )
                return
            
            out = generate_excel(df)
            await query.message.reply_document(
                document=out,
                filename="OzonAudit_Report.xlsx",
                caption="📊 Отчёт готов! Заполните колонку «Себестоимость»."
            )
            
            # Очищаем
            user_files.pop(user_id, None)
            
            await query.message.reply_text(
                "✅ Готово! Можете загрузить новые файлы через /start",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("🔄 Новый анализ", callback_data="back")]
                ])
            )
            
        except Exception as e:
            await query.edit_message_text(
                f"❌ Ошибка: {e}",
                reply_markup=InlineKeyboardMarkup([
                    [InlineKeyboardButton("◀️ Назад", callback_data="back")]
                ])
            )


async def file_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    doc = update.message.document
    
    if not doc.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Только Excel-файлы (.xlsx или .xls)")
        return
    
    file_type = context.user_data.get("awaiting_file")
    if not file_type:
        await update.message.reply_text("Сначала нажмите кнопку, чтобы выбрать тип файла.")
        return
    
    file = await context.bot.get_file(doc.file_id)
    data = BytesIO()
    await file.download_to_memory(data)
    data.seek(0)
    
    if user_id not in user_files:
        user_files[user_id] = {}
    
    user_files[user_id][file_type] = data
    
    names = {
        "accruals": "начислений",
        "stock": "остатков",
        "ads": "рекламы"
    }
    
    loaded = len(user_files[user_id])
    await update.message.reply_text(
        f"✅ Файл загружен: {names[file_type]}\n"
        f"📊 Загружено {loaded}/3 файлов."
    )
    
    # Очищаем ожидание
    context.user_data.pop("awaiting_file", None)
    
    # Показываем меню снова
    keyboard = [
        [InlineKeyboardButton("1️⃣ Загрузить отчёт о начислениях", callback_data="upload_accruals")],
        [InlineKeyboardButton("2️⃣ Загрузить управление остатками", callback_data="upload_stock")],
        [InlineKeyboardButton("3️⃣ Загрузить аналитику продвижения", callback_data="upload_ads")],
        [InlineKeyboardButton("📊 Сформировать отчёт", callback_data="generate")],
    ]
    await update.message.reply_text(
        "Что загружаем дальше?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, file_handler))
    print("Бот запущен")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()