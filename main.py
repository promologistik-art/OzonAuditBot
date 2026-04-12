from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from config import BOT_TOKEN
from parser import merge_all, identify_report, parse_accruals, parse_stock, parse_ads
from generator import generate_excel
from io import BytesIO

user_files = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Загрузите 3 отчёта Ozon:\n"
        "1. Отчёт о начислениях\n"
        "2. Управление остатками\n"
        "3. Аналитика продвижения"
    )


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    doc = update.message.document
    
    if not doc.file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("❌ Только Excel-файлы")
        return
    
    file = await context.bot.get_file(doc.file_id)
    data = BytesIO()
    await file.download_to_memory(data)
    data.seek(0)
    
    if user_id not in user_files:
        user_files[user_id] = []
    user_files[user_id].append((doc.file_name, data))
    
    cnt = len(user_files[user_id])
    await update.message.reply_text(f"✅ Файл {cnt}/3: {doc.file_name}")
    
    if cnt == 3:
        await update.message.reply_text("🔄 Определяю типы файлов...")
        
        # Определяем типы
        reports = {'accruals': None, 'stock': None, 'ads': None}
        for name, f in user_files[user_id]:
            f.seek(0)
            t = identify_report(f)
            await update.message.reply_text(f"📄 {name} → {t}")
            
            f.seek(0)
            if t == 'accruals':
                reports['accruals'] = (f, name)
            elif t == 'stock':
                reports['stock'] = (f, name)
            elif t == 'ads':
                reports['ads'] = (f, name)
        
        if not all(reports.values()):
            await update.message.reply_text("❌ Не все типы файлов найдены")
            user_files.pop(user_id, None)
            return
        
        await update.message.reply_text("🔄 Паршу начисления...")
        f, name = reports['accruals']
        f.seek(0)
        df_acc = parse_accruals(f)
        if df_acc is None:
            await update.message.reply_text("❌ Ошибка парсинга начислений")
            user_files.pop(user_id, None)
            return
        await update.message.reply_text(f"✅ Начисления: {len(df_acc)} товаров\nSKU: {df_acc['SKU'].tolist()[:3]}...")
        
        await update.message.reply_text("🔄 Паршу остатки...")
        f, name = reports['stock']
        f.seek(0)
        df_stock = parse_stock(f)
        if df_stock is not None:
            await update.message.reply_text(f"✅ Остатки: {len(df_stock)} товаров\nАртикулы: {df_stock['Артикул'].tolist()[:3]}...")
        else:
            await update.message.reply_text("⚠️ Остатки не распарсились")
            df_stock = None
        
        await update.message.reply_text("🔄 Паршу рекламу...")
        f, name = reports['ads']
        f.seek(0)
        df_ads = parse_ads(f)
        if df_ads is not None:
            await update.message.reply_text(f"✅ Реклама: {len(df_ads)} записей\nSKU: {df_ads['SKU'].tolist()[:3]}...\nРасходы: {df_ads['Расход_на_рекламу'].tolist()[:3]}...")
        else:
            await update.message.reply_text("⚠️ Реклама не распарсилась")
            df_ads = None
        
        await update.message.reply_text("🔄 Свожу данные...")
        
        # СВЕДЕНИЕ ВРУЧНУЮ здесь же, для диагностики
        df = df_acc.copy()
        
        # Приводим SKU к строке
        df['SKU'] = df['SKU'].astype(str)
        
        if df_stock is not None:
            await update.message.reply_text(f"Артикулы в начислениях: {df['Артикул'].tolist()}")
            await update.message.reply_text(f"Артикулы в остатках: {df_stock['Артикул'].tolist()}")
            df = df.merge(df_stock, on='Артикул', how='left')
        else:
            df['Остаток'] = 0
            df['Продаж_в_день'] = 0
        
        if df_ads is not None:
            df_ads['SKU'] = df_ads['SKU'].astype(str)
            await update.message.reply_text(f"SKU в начислениях: {df['SKU'].tolist()}")
            await update.message.reply_text(f"SKU в рекламе: {df_ads['SKU'].tolist()}")
            df = df.merge(df_ads, on='SKU', how='left')
        else:
            df['Расход_на_рекламу'] = 0
        
        df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
        df['Себестоимость'] = ''
        df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
        
        await update.message.reply_text(f"✅ Сведено: {len(df)} товаров")
        await update.message.reply_text(f"Расходы на рекламу: {df['Расход_на_рекламу'].tolist()}")
        await update.message.reply_text(f"Остатки: {df['Остаток'].tolist()}")
        
        out = generate_excel(df)
        await update.message.reply_document(out, filename="OzonAudit_Report.xlsx")
        
        user_files.pop(user_id, None)


def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    print("Бот запущен")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()