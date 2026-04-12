import pandas as pd
from io import BytesIO
from typing import Optional


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта о начислениях."""
    file_bytes.seek(0)
    
    # Читаем без заголовков, чтобы найти строку с 'Артикул'
    df_raw = pd.read_excel(file_bytes, sheet_name=0, header=None)
    
    header_row = 0
    for i, row in df_raw.iterrows():
        if 'Артикул' in row.values:
            header_row = i
            break
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=0, header=header_row)
    
    # Находим колонки по названиям (ищем частичное совпадение)
    sku_col = None
    name_col = None
    amount_col = None
    qty_col = None
    
    for col in df.columns:
        col_str = str(col)
        if 'Артикул' == col_str:
            art_col = col
        if 'SKU' in col_str:
            sku_col = col
        if 'Название товара' in col_str:
            name_col = col
        if 'Сумма итого' in col_str:
            amount_col = col
        if 'Количество' in col_str:
            qty_col = col
    
    if art_col is None or amount_col is None:
        return None
    
    # Разделяем на товары и общие расходы
    df_goods = df[df[art_col].notna()].copy()
    df_extra = df[df[art_col].isna()].copy()
    
    # Группируем по артикулу
    agg = df_goods.groupby(art_col)[amount_col].sum().reset_index()
    agg.columns = ['Артикул', 'Выручка']
    
    # Добавляем SKU и название
    if sku_col and name_col:
        info = df_goods.groupby(art_col).agg({sku_col: 'first', name_col: 'first'}).reset_index()
        info.columns = ['Артикул', 'SKU', 'Название_товара']
        agg = agg.merge(info, on='Артикул', how='left')
    else:
        agg['SKU'] = ''
        agg['Название_товара'] = ''
    
    # Количество
    if qty_col:
        qty = df_goods.groupby(art_col)[qty_col].sum().reset_index()
        qty.columns = ['Артикул', 'Продано_шт']
        agg = agg.merge(qty, on='Артикул', how='left')
    else:
        agg['Продано_шт'] = 0
    
    # Общие расходы
    total_extra = df_extra[amount_col].sum() if len(df_extra) > 0 else 0
    total_revenue = agg['Выручка'].sum()
    
    if total_revenue > 0:
        agg['Общие_расходы'] = (agg['Выручка'] / total_revenue) * total_extra
    else:
        agg['Общие_расходы'] = 0
    
    agg['Чистая_прибыль'] = agg['Выручка'] + agg['Общие_расходы']
    
    return agg


def parse_stock(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта об остатках — лист 'Товары'."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    # Ищем лист "Товары"
    sheet = None
    for s in xl.sheet_names:
        if s == 'Товары':
            sheet = s
            break
    if sheet is None:
        return None
    
    # Читаем без заголовков
    file_bytes.seek(0)
    df_raw = pd.read_excel(file_bytes, sheet_name=sheet, header=None)
    
    # Ищем строку с 'Артикул'
    header_row = 0
    art_col_idx = None
    stock_col_idx = None
    sales_col_idx = None
    
    for i, row in df_raw.iterrows():
        row_values = row.astype(str).tolist()
        if 'Артикул' in row_values:
            header_row = i
            # Находим индексы колонок
            for j, val in enumerate(row_values):
                if 'Артикул' in val:
                    art_col_idx = j
                elif 'Доступно к продаже' in val:
                    stock_col_idx = j
                elif 'Среднесуточные продажи' in val:
                    sales_col_idx = j
            break
    
    if header_row == 0 or art_col_idx is None or stock_col_idx is None:
        return None
    
    # Читаем с заголовками
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header_row)
    
    # Получаем названия колонок по индексам
    art_col = df.columns[art_col_idx]
    stock_col = df.columns[stock_col_idx]
    
    result = df[[art_col, stock_col]].copy()
    result.columns = ['Артикул', 'Остаток']
    
    if sales_col_idx is not None:
        sales_col = df.columns[sales_col_idx]
        result['Продаж_в_день'] = df[sales_col]
    else:
        result['Продаж_в_день'] = 0
    
    result = result[result['Артикул'].notna()]
    result = result.groupby('Артикул').agg({'Остаток': 'sum', 'Продаж_в_день': 'sum'}).reset_index()
    
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    """Парсинг отчёта по рекламе — лист 'Statistics'."""
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
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=0)
    
    # Ищем колонки
    sku_col = None
    cost_col = None
    
    for col in df.columns:
        col_str = str(col)
        if 'SKU' in col_str:
            sku_col = col
        if 'Расход' in col_str:
            cost_col = col
    
    if sku_col is None or cost_col is None:
        return None
    
    result = df[[sku_col, cost_col]].copy()
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