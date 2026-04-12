import pandas as pd
from io import BytesIO
from typing import Optional


def identify_report(file_bytes: BytesIO) -> str:
    """Определяет тип отчёта по колонкам."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    sheet = xl.sheet_names[0]
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, nrows=5)
    cols = ' '.join(df.columns.astype(str)).lower()
    
    if 'группа услуг' in cols and 'тип начисления' in cols:
        return 'accruals'
    elif 'доступно к продаже' in cols and 'ликвидность' not in cols:
        return 'stock'
    elif 'инструмент' in cols and 'расход' in cols:
        return 'ads'
    return 'unknown'


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=0)
    
    # Ищем строку с реальными заголовками
    header_row = None
    for i, row in df.iterrows():
        if 'Артикул' in str(row.values):
            header_row = i
            break
    
    if header_row is None:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=0, header=header_row)
    
    # Разделяем
    df_goods = df[df['Артикул'].notna()].copy()
    df_extra = df[df['Артикул'].isna()].copy()
    
    # Агрегируем по товарам
    agg = df_goods.groupby('Артикул').agg({
        'Сумма итого, руб.': 'sum',
        'Количество': 'sum' if 'Количество' in df.columns else lambda x: 0,
        'Название товара': 'first',
        'SKU': 'first'
    }).reset_index()
    
    agg.columns = ['Артикул', 'Выручка', 'Продано_шт', 'Название_товара', 'SKU']
    
    # Сумма общих расходов
    total_extra = df_extra['Сумма итого, руб.'].sum()
    total_revenue = agg['Выручка'].sum()
    
    # Размазываем общие расходы
    if total_revenue > 0:
        agg['Общие_расходы'] = (agg['Выручка'] / total_revenue) * total_extra
    else:
        agg['Общие_расходы'] = 0
    
    agg['Чистая_прибыль'] = agg['Выручка'] + agg['Общие_расходы']
    
    return agg


def parse_stock(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    sheet = 'Товары' if 'Товары' in xl.sheet_names else 0
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet)
    
    header_row = None
    for i, row in df.iterrows():
        if 'Артикул' in str(row.values):
            header_row = i
            break
    
    if header_row is None:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header_row)
    
    result = df[['Артикул', 'Доступно к продаже']].copy()
    result.columns = ['Артикул', 'Остаток']
    result['Продаж_в_день'] = df['Среднесуточные продажи за 28 дней'] if 'Среднесуточные продажи за 28 дней' in df.columns else 0
    
    result = result.groupby('Артикул').agg({
        'Остаток': 'sum',
        'Продаж_в_день': 'sum'
    }).reset_index()
    
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    sheet = 'Statistics' if 'Statistics' in xl.sheet_names else 0
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=0)
    
    if 'SKU' not in df.columns or 'Расход, ₽' not in df.columns:
        return None
    
    result = df[['SKU', 'Расход, ₽']].copy()
    result.columns = ['SKU', 'Расход_на_рекламу']
    result = result[result['SKU'].notna()]
    result = result.groupby('SKU')['Расход_на_рекламу'].sum().reset_index()
    
    return result


def merge_all(file1: BytesIO, file2: BytesIO, file3: BytesIO) -> Optional[pd.DataFrame]:
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
    
    if reports['stock'] is not None:
        df = df.merge(reports['stock'], on='Артикул', how='left')
    else:
        df['Остаток'] = 0
        df['Продаж_в_день'] = 0
    
    if reports['ads'] is not None:
        df = df.merge(reports['ads'], on='SKU', how='left')
    else:
        df['Расход_на_рекламу'] = 0
    
    df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
    df['Себестоимость'] = ''
    df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
    
    return df