#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Разработать программу на языке программирования Python, которая выделяет ключевые слова в заданном
русскоязычном тексте. Использовать частотный анализ слов в тексте. За приведение слов к нормальной форме
и исключение стоп-слов можно получить дополнительные баллы. Разрешается использования сторонних библиотек.
Входные данные – текстовый файл с расширением txt с заданным текстом (путь к файлу задается в консоли).
Выходные данные – таблица Excel с упорядоченным списком ключевых слов.

Решение необходимо выложить на GitHub и предоставить ссылку на просмотр.
"""
import nltk
from nltk.corpus import stopwords
import pymorphy2
import re
import csv
from pandas import DataFrame

morph = pymorphy2.MorphAnalyzer() # модуль для выполнения морфологического анализа слов,
                                  # будем использовать для привдения слов к нормальной форме

nltk.download('stopwords') # подгружаем список стоп-слов

# файлы куда будем записывать результат, при необходимости код можно дополнить динамическим их заданием
result_csv = 'result.csv'
result_xlsx = 'result.xlsx'

while True:
    try:
        # открываем заданный исходный файл и файл для записи результатов
        file_name = input('Введите имя или путь к txt файлу: ') #file_name = '0_sample1.txt'
        with open(file_name, encoding='utf-8') as input_file, open(result_csv, mode='w') as output_file:
            result = {} # создаем пустой словарь, куда будем писать число вхождений каждого слова 
                        # (ключ - слово, значение - число вхождений слова в тексте)
            for line in input_file: # выполняем построчный анализ строк в исходном файле
                for word in re.findall(r'\w+', line): # с помощью регулярного выражения отсекаем спецсимволы
                    word = word.lower() # для дальнейшего сравнения слова с уже найденными уникальными словами
                                        # преобразуем слово в нижний регистр 
                    if not word.isdigit() and word not in stopwords.words('russian'): # пропускаем цифры и стоп-слова
                        word = morph.parse(word)[0].normal_form # преобразуем найденное слово к нормальной форме
                                                                # перед сравнением с тем, что уже нашли ранее
                        if word not in result: # если слова нет в словаре - помещаем его туда как первое вхождение
                            result[word] = 1
                        else:                  # если слово уже есть в словаре - увеличивем счетчик на 1
                            result[word] += 1
            # сортируем список элементов по числу вхождений от большего к меньшему и преобразуем результат в кортеж 
            result = [(value, key) for key,value in result.items()]
            result.sort(reverse=True)
            result = tuple(result)
            # записываем результат в csv
            csv_writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(['число вхождений','слово'])
            for item in result:
                csv_writer.writerow(item)
            # дополнительно записываем результат в xlsx
            df = DataFrame(result)
            df.to_excel(result_xlsx, sheet_name='result', index=False, startrow=1, header=['число вхождений','слово'])
            print('Результаты обработки записаны в {} и {}'.format(result_xlsx, result_csv))
        break
    except (FileNotFoundError, IOError):
        print('Неверный пуль или имя файла')

