import pandas as pd
from io import BytesIO
from typing import Optional


def identify_report(file_bytes: BytesIO) -> str:
    """Определяет тип отчёта по наличию ключевых колонок в любом месте файла."""
    file_bytes.seek(0)
    
    # Пробуем разные листы
    xl = pd.ExcelFile(file_bytes)
    
    for sheet_name in xl.sheet_names:
        file_bytes.seek(0)
        # Читаем без заголовков, чтобы найти ключевые слова
        df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
        
        # Превращаем весь датафрейм в строку для поиска
        all_text = ' '.join(df_raw.astype(str).values.flatten()).lower()
        
        if 'группа услуг' in all_text and 'тип начисления' in all_text:
            return 'accruals'
        elif 'среднесуточные продажи' in all_text and 'доступно к продаже' in all_text:
            return 'stock'
        elif 'инструмент' in all_text and 'расход, ₽' in all_text:
            return 'ads'
    
    return 'unknown'


def find_header_row(file_bytes: BytesIO, sheet_name: str, keyword: str) -> int:
    """Находит строку с заголовками по ключевому слову."""
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    
    for i, row in df.iterrows():
        row_str = ' '.join(str(v) for v in row.values if pd.notna(v))
        if keyword in row_str:
            return i
    return 0


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта о начислениях."""
    file_bytes.seek(0)
    
    # Ищем лист с данными
    xl = pd.ExcelFile(file_bytes)
    sheet_name = xl.sheet_names[0]  # Обычно первый лист
    
    # Находим строку с заголовками
    header_row = find_header_row(file_bytes, sheet_name, 'Артикул')
    if header_row == 0:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row)
    
    # Удаляем строки, где все значения NaN
    df = df.dropna(how='all')
    
    # Проверяем наличие нужных колонок
    if 'Артикул' not in df.columns or 'Сумма итого, руб.' not in df.columns:
        return None
    
    # Разделяем на товары и общие расходы
    df_goods = df[df['Артикул'].notna()].copy()
    df_extra = df[df['Артикул'].isna()].copy()
    
    # Агрегируем по товарам
    agg = df_goods.groupby('Артикул').agg({
        'Сумма итого, руб.': 'sum',
        'Название товара': 'first',
        'SKU': 'first'
    }).reset_index()
    
    agg.columns = ['Артикул', 'Выручка', 'Название_товара', 'SKU']
    
    # Количество продаж (если есть колонка Количество)
    if 'Количество' in df_goods.columns:
        qty = df_goods.groupby('Артикул')['Количество'].sum().reset_index()
        qty.columns = ['Артикул', 'Продано_шт']
        agg = agg.merge(qty, on='Артикул', how='left')
    else:
        agg['Продано_шт'] = 0
    
    # Сумма общих расходов
    total_extra = df_extra['Сумма итого, руб.'].sum() if len(df_extra) > 0 else 0
    total_revenue = agg['Выручка'].sum()
    
    # Размазываем общие расходы пропорционально выручке
    if total_revenue > 0:
        agg['Общие_расходы'] = (agg['Выручка'] / total_revenue) * total_extra
    else:
        agg['Общие_расходы'] = 0
    
    agg['Чистая_прибыль'] = agg['Выручка'] + agg['Общие_расходы']
    
    return agg


def parse_stock(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта об остатках."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    # Ищем лист "Товары"
    sheet_name = None
    for s in xl.sheet_names:
        if 'Товары' in s:
            sheet_name = s
            break
    
    if sheet_name is None:
        sheet_name = xl.sheet_names[0]
    
    # Находим строку с заголовками
    header_row = find_header_row(file_bytes, sheet_name, 'Артикул')
    if header_row == 0:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=header_row)
    df = df.dropna(how='all')
    
    if 'Артикул' not in df.columns:
        return None
    
    result = df[['Артикул', 'Доступно к продаже']].copy()
    result.columns = ['Артикул', 'Остаток']
    
    if 'Среднесуточные продажи за 28 дней' in df.columns:
        result['Продаж_в_день'] = df['Среднесуточные продажи за 28 дней']
    else:
        result['Продаж_в_день'] = 0
    
    # Удаляем строки без артикула
    result = result[result['Артикул'].notna()]
    
    # Группируем (на случай дублей)
    result = result.groupby('Артикул').agg({
        'Остаток': 'sum',
        'Продаж_в_день': 'sum'
    }).reset_index()
    
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта по рекламе."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    # Ищем лист Statistics
    sheet_name = None
    for s in xl.sheet_names:
        if 'Statistics' in s:
            sheet_name = s
            break
    
    if sheet_name is None:
        sheet_name = xl.sheet_names[0]
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=0)
    df = df.dropna(how='all')
    
    # Ищем колонки с SKU и расходом (могут быть с разными названиями)
    sku_col = None
    cost_col = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'sku' in col_lower:
            sku_col = col
        if 'расход' in col_lower:
            cost_col = col
    
    if sku_col is None or cost_col is None:
        return None
    
    result = df[[sku_col, cost_col]].copy()
    result.columns = ['SKU', 'Расход_на_рекламу']
    result = result[result['SKU'].notna()]
    result = result.groupby('SKU')['Расход_на_рекламу'].sum().reset_index()
    
    return result


def merge_all(file1: BytesIO, file2: BytesIO, file3: BytesIO) -> Optional[pd.DataFrame]:
    """Сведение всех трёх отчётов."""
    files = [file1, file2, file3]
    reports = {'accruals': None, 'stock': None, 'ads': None}
    
    for f in files:
        f.seek(0)
        t = identify_report(f)
        f.seek(0)
        
        if t == 'accruals' and reports['accruals'] is None:
            reports['accruals'] = parse_accruals(f)
        elif t == 'stock' and reports['stock'] is None:
            reports['stock'] = parse_stock(f)
        elif t == 'ads' and reports['ads'] is None:
            reports['ads'] = parse_ads(f)
    
    if reports['accruals'] is None:
        return None
    
    df = reports['accruals']
    
    # Добавляем остатки
    if reports['stock'] is not None:
        df = df.merge(reports['stock'], on='Артикул', how='left')
    else:
        df['Остаток'] = 0
        df['Продаж_в_день'] = 0
    
    # Добавляем рекламу
    if reports['ads'] is not None:
        df = df.merge(reports['ads'], on='SKU', how='left')
    else:
        df['Расход_на_рекламу'] = 0
    
    # Заполняем пропуски
    df['Остаток'] = df['Остаток'].fillna(0)
    df['Продаж_в_день'] = df['Продаж_в_день'].fillna(0)
    df['Расход_на_рекламу'] = df['Расход_на_рекламу'].fillna(0)
    
    # Добавляем пустую колонку для себестоимости
    df['Себестоимость'] = ''
    
    # Итоговая прибыль
    df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
    
    return df