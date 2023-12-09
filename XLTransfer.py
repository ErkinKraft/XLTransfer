import os
import pandas as pd
from art import tprint
tprint('XLTransfer V1.0')
print('EK SoftWare')
print('Version for institutions')
print('')
print('')


while True:
    # Ввод пути к исходному файлу Excel
    input_file = input("Введите путь к исходному файлу Excel> ")

    # Считывание исходного файла Excel
    df = pd.read_excel(input_file)

    # Ввод названий строк для выборки
    row_names = input("Укажите названия строк через запятую> ").split(",")

    # Ввод названий столбцов для выборки
    column_names = input("Укажите названия столбцов через запятую> ").split(",")

    # Получение соответствующих номеров строк и столбцов
    rows = [df.index[df.iloc[:, 0] == name].tolist()[0] for name in row_names]
    columns = [df.columns.get_loc(name) for name in column_names]

    # Создание нового DataFrame с выбранными строками и столбцами
    new_df = df.iloc[rows, columns].reset_index(drop=True)

    # Установка названий строк в новом DataFrame
    new_df.index = row_names

    # Ввод пути для сохранения нового файла Excel
    output_dir = input("Введите путь, по которому сохранить новый файл Excel> ")

    # Проверка и создание директории, если она не существует
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Установка пути и имени нового файла Excel
    namesave = input('Название конечного файла> ')
    output_file = os.path.join(output_dir, namesave + ".xlsx")

    # Сохранение нового файла Excel
    new_df.to_excel(output_file, index=True)

    print("Новый файл Excel успешно создан и сохранен.")
    print('')
    allow = input('Повторить? (y/n)> ')
    if allow != 'n':
        print('')
    else:
        break