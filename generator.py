import pandas as pd
from io import BytesIO


def generate_excel(df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    cols = ['Артикул', 'SKU', 'Название_товара', 'Продано_шт', 'Выручка',
            'Общие_расходы', 'Расход_на_рекламу', 'Остаток', 'Продаж_в_день',
            'Себестоимость', 'Итог_прибыль']
    
    available = [c for c in cols if c in df.columns]
    
    with pd.ExcelWriter(output, engine='openpyxl') as w:
        df[available].to_excel(w, sheet_name='Анализ', index=False)
        pd.DataFrame({'Инструкция': [
            '1. Заполните колонку "Себестоимость"',
            '2. Итоговая прибыль = Выручка + Общие расходы - Расход на рекламу - Себестоимость'
        ]}).to_excel(w, sheet_name='Инструкция', index=False)
    
    output.seek(0)
    return output