import pandas as pd
from io import BytesIO
from typing import Optional


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта о начислениях — заголовки в строке 3 (индекс 2)."""
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=0, header=2)
    df = df.dropna(how='all')
    
    # Проверяем, что колонка 'Артикул' существует
    if 'Артикул' not in df.columns:
        # Пробуем найти строку с 'Артикул' автоматически
        file_bytes.seek(0)
        df_raw = pd.read_excel(file_bytes, sheet_name=0, header=None)
        header_row = 0
        for i, row in df_raw.iterrows():
            if 'Артикул' in row.values:
                header_row = i
                break
        file_bytes.seek(0)
        df = pd.read_excel(file_bytes, sheet_name=0, header=header_row)
        df = df.dropna(how='all')
    
    df_goods = df[df['Артикул'].notna()].copy()
    df_extra = df[df['Артикул'].isna()].copy()
    
    agg = df_goods.groupby('Артикул').agg({
        'Сумма итого, руб.': 'sum',
        'Название товара': 'first',
        'SKU': 'first',
        'Количество': 'sum' if 'Количество' in df_goods.columns else lambda x: 0
    }).reset_index()
    
    agg.columns = ['Артикул', 'Выручка', 'Название_товара', 'SKU', 'Продано_шт']
    
    total_extra = df_extra['Сумма итого, руб.'].sum() if len(df_extra) > 0 else 0
    total_revenue = agg['Выручка'].sum()
    
    if total_revenue > 0:
        agg['Общие_расходы'] = (agg['Выручка'] / total_revenue) * total_extra
    else:
        agg['Общие_расходы'] = 0
    
    agg['Чистая_прибыль'] = agg['Выручка'] + agg['Общие_расходы']
    return agg


def parse_stock(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта об остатках — лист 'Товары', заголовки в строке 2 (индекс 1)."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    sheet = None
    for s in xl.sheet_names:
        if s == 'Товары':
            sheet = s
            break
    if sheet is None:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=1)
    df = df.dropna(how='all')
    
    result = df[['Артикул', 'Доступно к продаже']].copy()
    result.columns = ['Артикул', 'Остаток']
    
    if 'Среднесуточные продажи за 28 дней' in df.columns:
        result['Продаж_в_день'] = df['Среднесуточные продажи за 28 дней']
    else:
        result['Продаж_в_день'] = 0
    
    result = result[result['Артикул'].notna()]
    result = result.groupby('Артикул').agg({'Остаток': 'sum', 'Продаж_в_день': 'sum'}).reset_index()
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта по рекламе — лист 'Statistics', заголовки в строке 2 (индекс 1)."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    sheet = None
    for s in xl.sheet_names:
        if s == 'Statistics':
            sheet = s
            break
    if sheet is None:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=1)
    df = df.dropna(how='all')
    
    result = df[['SKU', 'Расход, ₽']].copy()
    result.columns = ['SKU', 'Расход_на_рекламу']
    result = result[result['SKU'].notna()]
    result = result.groupby('SKU')['Расход_на_рекламу'].sum().reset_index()
    return result


def merge_three(acc: BytesIO, stock: BytesIO, ads: BytesIO) -> Optional[pd.DataFrame]:
    acc.seek(0)
    stock.seek(0)
    ads.seek(0)
    
    df_acc = parse_accruals(acc)
    df_stock = parse_stock(stock)
    df_ads = parse_ads(ads)
    
    if df_acc is None:
        return None
    
    df = df_acc
    df['SKU'] = df['SKU'].astype(str)
    
    if df_stock is not None:
        df = df.merge(df_stock, on='Артикул', how='left')
    else:
        df['Остаток'] = 0
        df['Продаж_в_день'] = 0
    
    if df_ads is not None:
        df_ads['SKU'] = df_ads['SKU'].astype(str)
        df = df.merge(df_ads, on='SKU', how='left')
    else:
        df['Расход_на_рекламу'] = 0
    
    df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
    df['Себестоимость'] = ''
    df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
    
    return df