from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from config import BOT_TOKEN
from parser import merge_all, identify_report, parse_accruals, parse_stock, parse_ads
from generator import generate_excel
from io import BytesIO

user_files = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = """
👋 Привет! Я бот-аудитор Ozon.

📤 Загрузи 3 отчёта ОДНИМ сообщением (каждый файл отдельно):

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
    
    user_files[user_id].append((doc.file_name, data))
    count = len(user_files[user_id])
    await update.message.reply_text(f"✅ Получен файл {count} из 3: {doc.file_name}")
    
    if count == 3:
        await update.message.reply_text("🔄 Проверяю типы файлов...")
        
        logs = []
        types = []
        
        for name, f in user_files[user_id]:
            f.seek(0)
            t = identify_report(f)
            types.append(t)
            logs.append(f"📄 {name} → {t}")
        
        await update.message.reply_text("\n".join(logs))
        
        if 'accruals' not in types or 'stock' not in types or 'ads' not in types:
            await update.message.reply_text(f"❌ Не все типы отчётов найдены. Получено: {types}")
            user_files.pop(user_id, None)
            return
        
        await update.message.reply_text("🔄 Паршу отчёт о начислениях...")
        
        # Парсим по очереди
        for name, f in user_files[user_id]:
            f.seek(0)
            if identify_report(f) == 'accruals':
                f.seek(0)
                df_acc = parse_accruals(f)
                if df_acc is None:
                    await update.message.reply_text("❌ Ошибка парсинга отчёта о начислениях")
                    user_files.pop(user_id, None)
                    return
                await update.message.reply_text(f"✅ Начисления: {len(df_acc)} товаров")
                break
        
        await update.message.reply_text("🔄 Паршу остатки...")
        for name, f in user_files[user_id]:
            f.seek(0)
            if identify_report(f) == 'stock':
                f.seek(0)
                df_stock = parse_stock(f)
                if df_stock is None:
                    await update.message.reply_text("❌ Ошибка парсинга остатков")
                else:
                    await update.message.reply_text(f"✅ Остатки: {len(df_stock)} товаров")
                break
        else:
            df_stock = None
        
        await update.message.reply_text("🔄 Паршу рекламу...")
        for name, f in user_files[user_id]:
            f.seek(0)
            if identify_report(f) == 'ads':
                f.seek(0)
                df_ads = parse_ads(f)
                if df_ads is None:
                    await update.message.reply_text("❌ Ошибка парсинга рекламы")
                else:
                    await update.message.reply_text(f"✅ Реклама: {len(df_ads)} записей")
                break
        else:
            df_ads = None
        
        await update.message.reply_text("🔄 Свожу данные...")
        
        try:
            # Сводим
            df = df_acc.copy()
            
            if df_stock is not None:
                df = df.merge(df_stock, on='Артикул', how='left')
            else:
                df['Остаток'] = 0
                df['Продаж_в_день'] = 0
            
            if df_ads is not None:
                df = df.merge(df_ads, on='SKU', how='left')
            else:
                df['Расход_на_рекламу'] = 0
            
            df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
            df['Себестоимость'] = ''
            df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
            
            await update.message.reply_text(f"✅ Сведено: {len(df)} товаров")
            
            out = generate_excel(df)
            await update.message.reply_document(
                document=out,
                filename="OzonAudit_Report.xlsx",
                caption="📊 Отчёт готов! Заполните колонку «Себестоимость»."
            )
            
        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка при сведении: {e}")
            import traceback
            print(traceback.format_exc())
        
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