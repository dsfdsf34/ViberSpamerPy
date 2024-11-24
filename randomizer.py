import pandas as pd
import os

# Запрос пути к файлу
file_path = input("Введите путь к файлу Excel: ")

# Проверка, существует ли файл по указанному пути
if not os.path.exists(file_path):
    print("Файл не найден. Пожалуйста, проверьте путь.")
else:
    # Загрузка файла Excel
    df = pd.read_excel(file_path)

    # Рандомизируем строки
    df_shuffled = df.sample(frac=1).reset_index(drop=True)

    # Сохранение рандомизированного файла
    output_file = input("Введите название для сохранения нового файла (с расширением .xlsx): ")
    df_shuffled.to_excel(output_file, index=False)

    print(f"Рандомизация завершена. Результат сохранен в файл: {output_file}")
