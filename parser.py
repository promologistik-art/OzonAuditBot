import pandas as pd
from io import BytesIO
from typing import Optional


def identify_report(file_bytes: BytesIO) -> str:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    for sheet_name in xl.sheet_names:
        file_bytes.seek(0)
        df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
        all_text = ' '.join(df_raw.astype(str).values.flatten()).lower()
        
        if 'группа услуг' in all_text and 'тип начисления' in all_text:
            return 'accruals'
        elif 'среднесуточные продажи' in all_text and 'ликвидность' in all_text:
            return 'stock'
        elif 'инструмент' in all_text and 'расход' in all_text:
            return 'ads'
    
    return 'unknown'


def find_header_row(file_bytes: BytesIO, sheet_name: str, keyword: str) -> int:
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
    for i, row in df.iterrows():
        if keyword in str(row.values):
            return i
    return 0


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    sheet = xl.sheet_names[0]
    
    header = find_header_row(file_bytes, sheet, 'Артикул')
    if header == 0:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header)
    df = df.dropna(how='all')
    
    df_goods = df[df['Артикул'].notna()].copy()
    df_extra = df[df['Артикул'].isna()].copy()
    
    # Группируем по Артикул, но сохраняем SKU (цифровой)
    agg = df_goods.groupby('Артикул').agg({
        'Сумма итого, руб.': 'sum',
        'Название товара': 'first',
        'SKU': 'first',  # <-- ВАЖНО: сохраняем цифровой SKU
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
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    sheet = None
    for s in xl.sheet_names:
        if s == 'Товары':
            sheet = s
            break
    if sheet is None:
        return None
    
    header = find_header_row(file_bytes, sheet, 'Артикул')
    if header == 0:
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header)
    df = df.dropna(how='all')
    
    # Берём Артикул (строковый) и остатки
    result = df[['Артикул', 'Доступно к продаже']].copy()
    result.columns = ['Артикул', 'Остаток']
    result['Продаж_в_день'] = df['Среднесуточные продажи за 28 дней'] if 'Среднесуточные продажи за 28 дней' in df.columns else 0
    
    result = result[result['Артикул'].notna()]
    result = result.groupby('Артикул').agg({'Остаток': 'sum', 'Продаж_в_день': 'sum'}).reset_index()
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
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
    df = df.dropna(how='all')
    
    if 'SKU' not in df.columns or 'Расход, ₽' not in df.columns:
        return None
    
    result = df[['SKU', 'Расход, ₽']].copy()
    result.columns = ['SKU', 'Расход_на_рекламу']
    result = result[result['SKU'].notna()]
    result = result.groupby('SKU')['Расход_на_рекламу'].sum().reset_index()
    return result


def merge_all(f1: BytesIO, f2: BytesIO, f3: BytesIO) -> Optional[pd.DataFrame]:
    files = [f1, f2, f3]
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
    
    # Сводим остатки по Артикулу (строковому)
    if reports['stock'] is not None:
        df = df.merge(reports['stock'], on='Артикул', how='left')
    else:
        df['Остаток'] = 0
        df['Продаж_в_день'] = 0
    
    # Сводим рекламу по SKU (цифровому)
    if reports['ads'] is not None:
        # Приводим SKU к строке для надёжности
        df['SKU'] = df['SKU'].astype(str)
        reports['ads']['SKU'] = reports['ads']['SKU'].astype(str)
        df = df.merge(reports['ads'], on='SKU', how='left')
    else:
        df['Расход_на_рекламу'] = 0
    
    df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
    df['Себестоимость'] = ''
    df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
    
    return df