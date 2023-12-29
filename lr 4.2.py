import fitz
import re
from typing import Dict


def extract_info_from_pdf(filename:str) ->Dict:
    # Открываем PDF-файл
    with open(filename, 'rb') as file:
        pdf = fitz.open(file)

        # Извлекаем текст из документа
        text = ''
        for page_num in range(len(pdf)):
            page = pdf[page_num]
            text = page.get_text("text")

        # Извлекаем нужные строки с помощью регулярных выражений
        name_match = re.search(r'Название программы для ЭВМ:\s*(.*)', text)
        name = name_match.group(1).strip() if name_match else ''

        registration_number_match = re.search(r'Номер регистрации \(свидетельства\):\s*(.*)', text, re.IGNORECASE)
        registration_number = registration_number_match.group(1).strip() if registration_number_match else ''

        registration_date_match = re.search(r'Дата регистрации:\s*(.*)', text, re.IGNORECASE)
        registration_date = registration_date_match.group(1).strip() if registration_date_match else ''

        authors = ''
        authors_match = re.findall(r'Автор\(ы\):\s*(.*\))|\s*([а-яА-Я]+.[а-яА-Я]+.[а-яА-Я]+.\(ru\))', text, re.IGNORECASE)

        for i in range(len(authors_match)):
            for j in range(len(authors_match[i])):
                if authors_match[i][j] != '':
                    authors += authors_match[i][j] + '\n'

        right_holder_match = re.search(r'Правообладатель\(и\):\s*(\D*\))', text, re.IGNORECASE)
        right_holder = right_holder_match.group(1).strip() if authors_match else ''

        # Возвращаем извлеченную информацию
        return {
            'Название': name,
            'Номер регистрации': registration_number,
            'Дата регистрации': registration_date,
            'Авторы': authors,
            'Правообладатель': right_holder
        }


# Пример использования
filename = '../../Downloads/Telegram Desktop/Виртуальная_установка_«Поршневой_компрессор».PDF'
info = extract_info_from_pdf(filename)

for k, v in info.items():
    print(k+': \t', v, '\n ................')
