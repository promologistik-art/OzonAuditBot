import pandas as pd
from io import BytesIO
from typing import Optional


def identify_report(file_bytes: BytesIO) -> str:
    """Определяет тип отчёта по ключевым словам во всех листах."""
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    for sheet_name in xl.sheet_names:
        file_bytes.seek(0)
        df_raw = pd.read_excel(file_bytes, sheet_name=sheet_name, header=None)
        all_text = ' '.join(df_raw.astype(str).values.flatten()).lower()
        
        if 'группа услуг' in all_text and 'тип начисления' in all_text:
            return 'accruals'
        elif 'доступно к продаже' in all_text and 'среднесуточные продажи' in all_text:
            return 'stock'
        elif 'инструмент' in all_text and 'расход' in all_text:
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
    return -1


def parse_accruals(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    sheet = xl.sheet_names[0]
    
    header = find_header_row(file_bytes, sheet, 'Артикул')
    if header < 0:
        print("Не найдена строка 'Артикул' в начислениях")
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header)
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
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    # Ищем лист "Товары"
    sheet = None
    for s in xl.sheet_names:
        if s == 'Товары':
            sheet = s
            break
    if sheet is None:
        print("Лист 'Товары' не найден")
        return None
    
    header = find_header_row(file_bytes, sheet, 'Артикул')
    if header < 0:
        print(f"Не найдена строка 'Артикул' в листе {sheet}")
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=header)
    df = df.dropna(how='all')
    
    if 'Артикул' not in df.columns or 'Доступно к продаже' not in df.columns:
        print(f"Нет нужных колонок в листе {sheet}")
        print(f"Колонки: {df.columns.tolist()}")
        return None
    
    result = df[['Артикул', 'Доступно к продаже']].copy()
    result.columns = ['Артикул', 'Остаток']
    
    if 'Среднесуточные продажи за 28 дней' in df.columns:
        result['Продаж_в_день'] = df['Среднесуточные продажи за 28 дней']
    else:
        result['Продаж_в_день'] = 0
    
    result = result[result['Артикул'].notna()]
    result = result.groupby('Артикул').agg({'Остаток': 'sum', 'Продаж_в_день': 'sum'}).reset_index()
    
    print(f"Остатки распарсены: {len(result)} товаров")
    return result


def parse_ads(file_bytes: BytesIO) -> Optional[pd.DataFrame]:
    file_bytes.seek(0)
    xl = pd.ExcelFile(file_bytes)
    
    # Ищем лист Statistics
    sheet = None
    for s in xl.sheet_names:
        if s == 'Statistics':
            sheet = s
            break
    if sheet is None:
        print("Лист 'Statistics' не найден")
        return None
    
    file_bytes.seek(0)
    df = pd.read_excel(file_bytes, sheet_name=sheet, header=0)
    df = df.dropna(how='all')
    
    if 'SKU' not in df.columns or 'Расход, ₽' not in df.columns:
        print(f"Нет нужных колонок в листе {sheet}")
        print(f"Колонки: {df.columns.tolist()}")
        return None
    
    result = df[['SKU', 'Расход, ₽']].copy()
    result.columns = ['SKU', 'Расход_на_рекламу']
    result = result[result['SKU'].notna()]
    result = result.groupby('SKU')['Расход_на_рекламу'].sum().reset_index()
    
    print(f"Реклама распарсена: {len(result)} записей")
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
    df['SKU'] = df['SKU'].astype(str)
    
    if reports['stock'] is not None:
        df = df.merge(reports['stock'], on='Артикул', how='left')
    else:
        df['Остаток'] = 0
        df['Продаж_в_день'] = 0
    
    if reports['ads'] is not None:
        reports['ads']['SKU'] = reports['ads']['SKU'].astype(str)
        df = df.merge(reports['ads'], on='SKU', how='left')
    else:
        df['Расход_на_рекламу'] = 0
    
    df.fillna({'Остаток': 0, 'Продаж_в_день': 0, 'Расход_на_рекламу': 0}, inplace=True)
    df['Себестоимость'] = ''
    df['Итог_прибыль'] = df['Чистая_прибыль'] - df['Расход_на_рекламу']
    
    return df