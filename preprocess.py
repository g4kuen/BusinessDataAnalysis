import csv
import math
import os
import re
import numpy as np
import pandas as pd
#updated,           high,           low,        open,       close,      volume, average,U/R
#когда обновлено    наивысшее      наименьшее   открытие    закрытие    объем   среднее
column_names = ['updated', 'high', 'low', 'open', 'close', 'volume', 'average', 'U/R']
#идеи восстановления nan
#high  посмотреть что выше, open или Close
#low посмотреть что меньше, open или close
#open посмотреть закрытие прошлого дня
#close посмотреть открытие следующего дня
#volume никак
#average среднее между высшим и средним
#nan,nan,nan,nan,nan,nan,,



df = pd.read_csv('Выгрузка для ЛАБ5-6.csv', encoding='cp1251', delimiter=';')

# Запись DataFrame в новый файл с кодировкой UTF-8 и делиметером ";"
df.to_csv('new_file.csv', encoding='utf-8', sep=';', index=False)


df = pd.read_csv('new_file.csv', encoding='Utf-8', names=column_names,delimiter=';',skiprows=1)
print(df)



input_file = 'new_file.csv'
output_file = 'new_file.csv'

if not os.path.exists(input_file):
    print(f"Файл {input_file} не существует.")
    exit()

with open(input_file, 'r', newline='') as infile:
    reader = csv.reader(infile)
    next(reader)
    rows = [row for row in reader]

with open(output_file, 'w', newline='') as outfile:
    writer = csv.writer(outfile)
    writer.writerows(rows)

print("Первая строка удалена из файла и сохранена в", output_file)
#df.to_csv('new_file.csv', encoding='Utf-8', index=False)

def preprocess(df):
    #удаление запятых
    for col in df.columns:
        df[col] = df[col].astype(str).str.replace(',', '', regex=True)
    return df


def replace_months_all_cells(df):
    # исправление мар июн июл
    month_map = {
        'янв': '01', 'фев': '02', 'мар': '03', 'апр': '04', 'май': '05', 'июн': '06',
        'июл': '07', 'авг': '08', 'сен': '09', 'окт': '10', 'ноя': '11', 'дек': '12'
    }
    for col in df.columns:
        for month, number in month_map.items():
            df[col] = df[col].str.replace(rf'\b{month}\b', number, regex=True)
    return df

def round_df(df):
    #округление до 4 знаков столбца average
    df['average'] = pd.to_numeric(df['average'], errors='coerce')
    df['average'] = df['average'].apply(lambda x: round(x, 4))
    return df
def fill_high(df):
    #nan для большего
    for i in range(1, len(df)):
        if df.loc[i, 'high'] == 'nan':
            if df.loc[i, 'open'] != 'nan' and df.loc[i, 'close'] != 'nan':
                df.loc[i, 'high'] = max(float(df.loc[i, 'open']), float(df.loc[i, 'close']))
            elif df.loc[i-1, 'high'] != 'nan':
                df.loc[i, 'high'] = df.loc[i-1, 'high']
            else:
                df.loc[i, 'high'] = np.nan
    return df

def fill_low(df):
    #nan для самого маленького
    for i in range(1, len(df)):
        if df.loc[i, 'low'] == 'nan':
            if df.loc[i, 'open'] != 'nan' and df.loc[i, 'close'] != 'nan':
                df.loc[i, 'low'] = min(float(df.loc[i, 'open']), float(df.loc[i, 'close']))
            elif df.loc[i-1, 'low'] != 'nan':
                df.loc[i, 'low'] = df.loc[i-1, 'low']
            else:
                df.loc[i, 'low'] = np.nan
    return df

def fill_open(df):
    #nan для открытия
    for i in range(1, len(df)):
        if df.loc[i, 'open'] == 'nan':
            if df.loc[i-1, 'close'] != 'nan':
                df.loc[i, 'open'] = df.loc[i-1, 'close']
            else:
                df.loc[i, 'open'] = np.nan
    return df

def fill_close(df):
    #nan для закрытия
    for i in range(len(df)-2, -1, -1):
        if df.loc[i, 'close'] == 'nan':
            if df.loc[i+1, 'open'] != 'nan':
                df.loc[i, 'close'] = df.loc[i+1, 'open']
            else:
                df.loc[i, 'close'] = np.nan
    return df

def fill_volume(df):
    #nan для объема
    df = df.dropna(subset=['volume'])
    return df



def fill_average(df):
    for i in range(len(df)):
        if pd.isnull(df.loc[i, 'average']) or df.loc[i, 'average'] == 'nan':
            high = df.loc[i, 'high']
            low = df.loc[i, 'low']
            if high != '' and low != '':  # Проверяем, что строки не пустые
                try:
                    df.loc[i, 'average'] = (float(high) + float(low)) / 2
                except ValueError:
                    # Если значение не может быть преобразовано в число, пропускаем его
                    pass
    return df
def remove_rows_with_commas(df):
    #rofl, не нужно
    df = df[df['volume'].str.count(',') < 2]
    return df

#drop na для volume
def drop_na_rows_volume(df):
    df = df[df['volume'] != 'nan']
    return df
def drop_na_rows_volume_second(df):
    df = df[df['volume'] != '']
    return df

def drop_na_rows_volume_third(df):
    df = df.dropna(subset=['volume'])
    return df

def fill_date(df):
    #nan для дат
    for i, row in df.iterrows():
        if pd.isna(row['updated']) or row['updated'] == 'nan':
            if i == 0:
                df.loc[i, 'updated'] = '01.01.2000 0:00'
            else:
                prev_date = pd.to_datetime(df.loc[i-1, 'updated'], format='%d.%m.%Y %H:%M')
                next_date = prev_date + pd.Timedelta(days=1)
                df.loc[i, 'updated'] = next_date.strftime('%d.%m.%Y %H:%M')
    return df

def format_date(df):
    #форматирование дат
    day_of_week_dict = {
        'Monday': 'Понедельник',
        'Tuesday': 'Вторник',
        'Wednesday': 'Среда',
        'Thursday': 'Четверг',
        'Friday': 'Пятница',
        'Saturday': 'Суббота',
        'Sunday': 'Воскресенье'
    }
    for i, row in df.iterrows():
        date = pd.to_datetime(row['updated'], format='%d.%m.%Y %H:%M')
        day_of_week = date.strftime('%A')
        formatted_date = f"{day_of_week_dict[day_of_week]} {date.day:02d}.{date.month:02d}.{date.year}г."
        df.loc[i, 'updated'] = formatted_date
    return df

def remove_saturday_rows(df):
    df = df[~df['updated'].str.contains('Суббота')]
    return df
def drop_zero_volume_rows(df):
    df = df[df['volume'] != '0']
    return df

def drop_zero_volume_rows1(df):
    df = df[df['volume'] != '0.0']
    return df

def drop_zero_volume_rows2(df):
    df = df[df['volume'] != '0,0']
    return df

def clean_text(text):
    if isinstance(text, str):  # Check if the value is a string
        # Remove non-Russian alphabets and non-space characters
        cleaned_text = re.sub(r'[^а-яА-Я\s]', '', text)
        # Remove extra spaces
        cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
        return cleaned_text.strip()
    else:
        return np.nan


df= preprocess(df.copy())
print(df.head())
df=replace_months_all_cells(df.copy())
print(df.head())
df = df.dropna(how='all')
print(df.head())
df=round_df(df.copy())
print(df.head())
df = fill_open(df)
df = fill_close(df)
df = fill_high(df)
df = fill_low(df)
#df = df.applymap(clean_text)
df = fill_average(df)
df = drop_na_rows_volume(df)
df = drop_na_rows_volume_second(df)
df=drop_na_rows_volume_third(df)
df=fill_date(df)
df=format_date(df)
df=remove_saturday_rows(df)
df=drop_zero_volume_rows(df)
df=drop_zero_volume_rows1(df)
df=drop_zero_volume_rows2(df)
print(df[df['volume']=='0'])
df.to_csv('preprocessed_data.csv', encoding='utf-8', index=False)
