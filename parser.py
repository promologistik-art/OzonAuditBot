import pandas as pd
from io import BytesIO
from typing import Dict, Tuple


def identify_report(file_bytes: BytesIO) -> str:
    """Определяет тип отчёта по содержимому."""
    xl = pd.ExcelFile(file_bytes)
    
    for sheet_name in xl.sheet_names[:1]:
        df = pd.read_excel(file_bytes, sheet_name=sheet_name, nrows=5)
        columns_str = ' '.join(df.columns.astype(str)).lower()
        
        if 'группа услуг' in columns_str and 'тип начисления' in columns_str:
            return 'accruals'
        elif 'среднесуточные продажи' in columns_str or 'ликвидность' in columns_str:
            return 'stock'
        elif 'инструмент' in columns_str and 'расход' in columns_str:
            return 'ads'
    
    return 'unknown'


def parse_accruals(file_bytes: BytesIO) -> pd.DataFrame:
    """Парсинг отчёта о начислениях."""
    df = pd.read_excel(file_bytes, sheet_name=0, header=14)
    
    # Группируем по артикулу
    agg = df.groupby('Артикул').agg({
        'Сумма итого, руб.': 'sum',
        'Количество': 'sum',
        'Название товара': 'first',
        'SKU': 'first'
    }).reset_index()
    
    agg.rename(columns={'Сумма итого, руб.': 'Выручка', 'Количество': 'Продано_шт'}, inplace=True)
    return agg[['Артикул', 'SKU', 'Название товара', 'Продано_шт', 'Выручка']]


def parse_stock(file_bytes: BytesIO) -> pd.DataFrame:
    """Парсинг отчёта об остатках (лист 'Товары')."""
    xl = pd.ExcelFile(file_bytes)
    sheet = 'Товары' if 'Товары' in xl.sheet_names else 0
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=2)
    
    result = df[['Артикул', 'Доступно к продаже', 'Среднесуточные продажи за 28 дней']].copy()
    result.rename(columns={
        'Доступно к продаже': 'Остаток',
        'Среднесуточные продажи за 28 дней': 'Продаж_в_день'
    }, inplace=True)
    
    # Агрегируем по артикулу (если есть дубли)
    result = result.groupby('Артикул').agg({
        'Остаток': 'sum',
        'Продаж_в_день': 'sum'
    }).reset_index()
    
    return result


def parse_ads(file_bytes: BytesIO) -> pd.DataFrame:
    """Парсинг отчёта по рекламе."""
    xl = pd.ExcelFile(file_bytes)
    sheet = 'Statistics' if 'Statistics' in xl.sheet_names else 0
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=0)
    
    result = df[['SKU', 'Расход, ₽']].copy()
    result.rename(columns={'Расход, ₽': 'Расход_на_рекламу'}, inplace=True)
    result = result.groupby('SKU').agg({'Расход_на_рекламу': 'sum'}).reset_index()
    
    return result


def merge_all(file1: BytesIO, file2: BytesIO, file3: BytesIO) -> pd.DataFrame:
    """Сведение всех трёх отчётов."""
    
    files = [file1, file2, file3]
    reports = {'accruals': None, 'stock': None, 'ads': None}
    
    for f in files:
        f.seek(0)
        rtype = identify_report(f)
        f.seek(0)
        if rtype == 'accruals':
            reports['accruals'] = parse_accruals(f)
        elif rtype == 'stock':
            reports['stock'] = parse_stock(f)
        elif rtype == 'ads':
            reports['ads'] = parse_ads(f)
    
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
    
    return df