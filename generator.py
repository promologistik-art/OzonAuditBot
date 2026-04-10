import pandas as pd
from io import BytesIO


def generate_output_excel(df: pd.DataFrame) -> BytesIO:
    """Генерация выходного Excel-файла."""
    output = BytesIO()
    
    # Порядок колонок
    columns_order = [
        'Артикул', 'SKU', 'Название товара',
        'Продано_шт', 'Выручка', 'Расход_на_рекламу',
        'Остаток', 'Продаж_в_день',
        'Себестоимость', 'Чистая_прибыль'
    ]
    
    available_cols = [col for col in columns_order if col in df.columns]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df[available_cols].to_excel(writer, sheet_name='Анализ', index=False)
        
        # Инструкция на втором листе
        instructions = pd.DataFrame({
            'Инструкция': [
                '1. Заполни колонку "Себестоимость" для каждого товара',
                '2. Колонка "Чистая_прибыль" = Выручка - Расход_на_рекламу - Себестоимость',
                '3. Для пересчёта замени формулу в колонке "Чистая_прибыль"',
                '',
                'Формула для Excel:',
                '=E2-F2-I2  (где E=Выручка, F=Расход_на_рекламу, I=Себестоимость)'
            ]
        })
        instructions.to_excel(writer, sheet_name='Инструкция', index=False)
    
    output.seek(0)
    return output