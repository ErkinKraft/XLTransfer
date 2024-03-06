import os
import pandas as pd
from art import tprint
tprint('XLTransfer V1.1')
print('EK SoftWare')
print('Version for institutions')
print('')
print('')

#косяк с выносом названия строки с поиском по названию

# Просим пользователя ввести путь до исходного файла Excel
input_path = input("Введите путь до исходного файла Excel: ")

# Загружаем исходный файл Excel
df = pd.read_excel(input_path)

# Просим пользователя выбрать действие
action = input("Выберите действие (1 - вынести строки по названию, 2 - вынести строки по параметрам): ")


def convert_param(param):
    print('def')
    try:
        return int(param)
    except ValueError:
        return str(param)


if action == "1":
    # Просим пользователя ввести названия строк и столбцов
    row_names = input("Введите названия строк через запятую: ").split(",")
    column_names = input("Введите названия столбцов через запятую: ").split(",")

    # Просим пользователя ввести путь и название файла для сохранения
    output_path = input("Введите путь для сохранения файла: ")
    output_filename = input("Введите название файла: ")

    # Выбираем только указанные строки и столбцы
    filtered_df = df.loc[df.iloc[:, 0].isin(row_names), column_names]

    # Сохраняем в новый файл
    filtered_df.to_excel(f"{output_path}/{output_filename}.xlsx", index=False)

elif action == "2":
    # Просим пользователя выбрать действие
    sub_action = input("Выберите действие (1 - все параметры, 2 - хотя бы один параметр): ")

    # Просим пользователя ввести параметры
    parameters = input("Введите параметры через запятую: ").split(",")

    # Просим пользователя ввести путь и название файла для сохранения
    output_path = input("Введите путь для сохранения файла: ")
    output_filename = input("Введите название файла: ")

    if sub_action == "1":
        # Выбираем строки, где все указанные параметры присутствуют
        converted_parameters = [convert_param(param) for param in parameters]
        filtered_df = df[df.isin(converted_parameters).any(axis=1)]


    elif sub_action == "2":
        # Выбираем строки, где хотя бы один из указанных параметров присутствует
        filtered_df = df[df.isin(parameters).any(axis=1)]

    # Сохраняем в новый файл
    filtered_df.to_excel(f"{output_path}/{output_filename}.xlsx", index=False)



else:
    print("Неверно выбрано действие. Пожалуйста, выберите 1 или 2.")



