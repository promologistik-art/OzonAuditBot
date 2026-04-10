import pandas as pd
from io import BytesIO
from typing import Dict, Tuple, Optional


def identify_report(file_bytes: BytesIO) -> str:
    """Определяет тип отчёта по содержимому."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    for sheet_name in xl.sheet_names[:1]:
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=sheet_name, nrows=10)
        columns_str = ' '.join(df.columns.astype(str)).lower()
        
        if 'группа услуг' in columns_str and 'тип начисления' in columns_str:
            return 'accruals'
        elif 'среднесуточные продажи' in columns_str and 'ликвидность' in columns_str:
            return 'stock'
        elif 'инструмент' in columns_str and 'расход' in columns_str:
            return 'ads'
    
    return 'unknown'


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта о начислениях."""
    try:
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=0, header=None)
        
        # Ищем строку с заголовком "Артикул"
        header_row = None
        for i, row in df.iterrows():
            if 'Артикул' in row.values:
                header_row = i
                break
        
        if header_row is None:
            print("❌ Не найдена строка с 'Артикул' в отчёте о начислениях")
            return None
        
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=0, header=header_row)
        
        # Проверяем наличие нужных колонок
        required = ['Артикул', 'Сумма итого, руб.']
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"❌ Отсутствуют колонки в отчёте о начислениях: {missing}")
            print(f"Доступные колонки: {df.columns.tolist()}")
            return None
        
        # Группируем по артикулу
        agg = df.groupby('Артикул').agg({
            'Сумма итого, руб.': 'sum',
            'Количество': 'sum' if 'Количество' in df.columns else pd.Series.sum,
            'Название товара': 'first',
            'SKU': 'first'
        }).reset_index()
        
        # Переименовываем
        agg.rename(columns={
            'Сумма итого, руб.': 'Выручка',
            'Количество': 'Продано_шт'
        }, inplace=True)
        
        print(f"✅ Отчёт о начислениях обработан. Найдено товаров: {len(agg)}")
        return agg[['Артикул', 'SKU', 'Название товара', 'Продано_шт', 'Выручка']]
        
    except Exception as e:
        print(f"❌ Ошибка парсинга отчёта о начислениях: {e}")
        return None


def parse_stock(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта об остатках (лист 'Товары')."""
    try:
        file_bytes.seek(0)
        xl = pd.ExcelFile(file_bytes)
        
        # Ищем лист с остатками
        sheet = None
        for s in xl.sheet_names:
            if 'Товары' in s:
                sheet = s
                break
        
        if sheet is None:
            print(f"❌ Лист 'Товары' не найден. Доступные листы: {xl.sheet_names}")
            return None
        
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=sheet, header=None)
        
        # Ищем строку с заголовком "Артикул"
        header_row = None
        for i, row in df.iterrows():
            if 'Артикул' in row.values:
                header_row = i
                break
        
        if header_row is None:
            print(f"❌ Не найдена строка с 'Артикул' в листе '{sheet}'")
            return None
        
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=sheet, header=header_row)
        
        # Проверяем наличие нужных колонок
        required = ['Артикул', 'Доступно к продаже']
        missing = [c for c in required if c not in df.columns]
        if missing:
            print(f"❌ Отсутствуют колонки в отчёте об остатках: {missing}")
            print(f"Доступные колонки: {df.columns.tolist()}")
            return None
        
        # Выбираем нужные колонки
        cols = ['Артикул', 'Доступно к продаже']
        if 'Среднесуточные продажи за 28 дней' in df.columns:
            cols.append('Среднесуточные продажи за 28 дней')
        
        result = df[cols].copy()
        result.rename(columns={
            'Доступно к продаже': 'Остаток',
            'Среднесуточные продажи за 28 дней': 'Продаж_в_день'
        }, inplace=True)
        
        # Добавляем колонку Продаж_в_день если её нет
        if 'Продаж_в_день' not in result.columns:
            result['Продаж_в_день'] = 0
        
        # Агрегируем по артикулу
        result = result.groupby('Артикул').agg({
            'Остаток': 'sum',
            'Продаж_в_день': 'sum'
        }).reset_index()
        
        print(f"✅ Отчёт об остатках обработан. Найдено товаров: {len(result)}")
        return result
        
    except Exception as e:
        print(f"❌ Ошибка парсинга отчёта об остатках: {e}")
        return None


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта по рекламе."""
    try:
        file_bytes.seek(0)
        xl = pd.ExcelFile(file_bytes)
        
        # Ищем лист Statistics
        sheet = None
        for s in xl.sheet_names:
            if 'Statistics' in s:
                sheet = s
                break
        
        if sheet is None:
            # Пробуем первый лист
            sheet = xl.sheet_names[0]
            print(f"⚠️ Лист 'Statistics' не найден, использую '{sheet}'")
        
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=sheet, header=0)
        
        # Проверяем наличие нужных колонок
        if 'SKU' not in df.columns or 'Расход, ₽' not in df.columns:
            print(f"❌ Отсутствуют колонки SKU или Расход, ₽ в отчёте о рекламе")
            print(f"Доступные колонки: {df.columns.tolist()}")
            return None
        
        result = df[['SKU', 'Расход, ₽']].copy()
        result.rename(columns={'Расход, ₽': 'Расход_на_рекламу'}, inplace=True)
        
        # Убираем пустые SKU
        result = result[result['SKU'].notna()]
        
        # Группируем по SKU
        result = result.groupby('SKU').agg({'Расход_на_рекламу': 'sum'}).reset_index()
        
        print(f"✅ Отчёт о рекламе обработан. Найдено товаров: {len(result)}")
        return result
        
    except Exception as e:
        print(f"❌ Ошибка парсинга отчёта о рекламе: {e}")
        return None


def merge_all(file1: BytesIO, file2: BytesIO, file3: BytesIO) -> Optional[pd.DataFrame]:
    """Сведение всех трёх отчётов."""
    
    files = [file1, file2, file3]
    reports = {'accruals': None, 'stock': None, 'ads': None}
    
    # Определяем типы файлов
    for i, f in enumerate(files):
        f.seek(0)
        rtype = identify_report(f)
        print(f"Файл {i+1} определён как: {rtype}")
        
        if rtype == 'accruals' and reports['accruals'] is None:
            f.seek(0)
            reports['accruals'] = parse_accruals(f)
        elif rtype == 'stock' and reports['stock'] is None:
            f.seek(0)
            reports['stock'] = parse_stock(f)
        elif rtype == 'ads' and reports['ads'] is None:
            f.seek(0)
            reports['ads'] = parse_ads(f)
        else:
            print(f"⚠️ Файл {i+1} не распознан или дублирует тип: {rtype}")
    
    # Проверяем, что все отчёты загружены
    missing = [name for name, val in reports.items() if val is None]
    if missing:
        print(f"❌ Не хватает отчётов: {missing}")
        return None
    
    # Сводим
    df = reports['accruals']
    
    if reports['stock'] is not None:
        df = df.merge(reports['stock'], on='Артикул', how='left')
    
    if reports['ads'] is not None:
        df = df.merge(reports['ads'], left_on='SKU', right_on='SKU', how='left')
    
    # Заполняем пропуски
    df['Остаток'] = df['Остаток'].fillna(0)
    df['Продаж_в_день'] = df['Продаж_в_день'].fillna(0)
    df['Расход_на_рекламу'] = df['Расход_на_рекламу'].fillna(0)
    
    # Добавляем пустую колонку для себестоимости
    df['Себестоимость'] = ''
    
    # Добавляем колонку с формулой прибыли
    df['Чистая_прибыль'] = df['Выручка'] - df['Расход_на_рекламу']
    
    print(f"✅ Все отчёты сведены. Итоговая таблица: {len(df)} строк")
    return df