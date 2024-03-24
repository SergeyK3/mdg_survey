# Опрос файлов протоколов МДГ
# перед запуском скрипта нужно активировать свое виртуальное окружение, 
# запустив в терминале виртуальное окружение: venv\Scripts\activate 
# Не забыть запустить в консоли. Это нужно делать каждый раз?
# pip install pdfplumber openpyxl xlsxwriter tabula-py pandas tqdm
# в 52-53-й строке корректировать маршрут папки

import os
import pdfplumber
import re
import pandas as pd
from tqdm import tqdm
import time

# Функция для извлечения значения между двумя фразами
def extract_value_between_phrases(text, start_phrase, end_phrase):
    start_index = text.find(start_phrase) + len(start_phrase)
    end_index = text.find(end_phrase, start_index)
    return text[start_index:end_index].strip()

# Функция для извлечения информации о пациенте из текста
def extract_patient_info(text):
    protocol_number = extract_value_between_phrases(text, "Заключение мультидисциплинарной группы (МДГ)*", "1. Қай медициналық ұйымда").replace(' ', '')
    full_name = extract_value_between_phrases(text, "его наличии) пациента) ", "3. ИИН/ЖСН")
    iin = extract_value_between_phrases(text, "3. ИИН/ЖСН ", "Жасы")
    gender = extract_value_between_phrases(text, "4. Жынысы (Пол) - ", "5. Науқастың тұрақты мекен")
    address = extract_value_between_phrases(text, " местожительства пациента) ", "6. МДТ жолдамасы")
    egok = extract_value_between_phrases(text, "11. Науқастың жағдайы (общее состояние) ", "12. Қосымша")
    mkb = extract_value_between_phrases(text, "Диагноз: ", "Жасы")
    recommend = extract_value_between_phrases(text, "лечение, химиолучевое лечение) ", "лечение, химиолучевое лечение) ")
    clin_group = extract_value_between_phrases(text, "группе (Iб), (II), (III)) ", "4) Симптоматикалық")
    doctor = extract_value_between_phrases(text, "заключение МДГ) онколог: ", "Хаттаманыңтолтырылғанкүнi")
    date_prot = extract_value_between_phrases(text, "(Дата составления заключения)", "Хаттаманыңтолтырылғанкүнi")
    link = extract_value_between_phrases(text, "https://doc.ast", "")
    
    return {
        'протокол ': protocol_number,
        'Дата протокола': date_prot,
        'Пациент': full_name,
        'ИИН': iin,
        'Пол': gender,
        'Адрес': address,
        'мкб10': mkb,
        'ECOG': egok,
        'Рекомендации': recommend,
        'Клин группа': clin_group,
        'Врач': doctor,
        'Ссылка': link
    }

# Задаем путь к папке с PDF-файлами
pdf_folder_path = "F:\\Элдок pdf общие\\ПротокМДГ\\"

# Проверяем существование введенной папки
if not os.path.exists(pdf_folder_path):
    print("Указанной папки не существует.")
    exit()

# Создаем столбцы будущего DataFrame
columns = ['протокол ', 'Дата протокола', 'Пациент', 'ИИН', 'Пол', 'Адрес', 'мкб10', 'ECOG', 'Рекомендации', 'Клин группа', 'Врач', 'Ссылка']
all_data = pd.DataFrame(columns=columns)

# Проходим по всем PDF-файлам в указанной папке
for file_name in os.listdir(pdf_folder_path):
    if file_name.endswith('.pdf'):
        file_path = os.path.join(pdf_folder_path, file_name)
        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text()
        patient_info = extract_patient_info(text)
        all_data = all_data.append(patient_info, ignore_index=True)

# Сохраняем DataFrame в Excel-файл

# Создаем общий DataFrame для всех файлов
all_data = pd.DataFrame(columns=columns)

# Измеряем количество файлов для отображения прогресса
num_files = len([name for name in os.listdir(pdf_folder_path) if name.endswith('.pdf')])

# Используем tqdm для отслеживания прогресса
with tqdm(total=num_files, desc="Обработка файлов") as pbar:
    # Пример использования функции extract_patient_info для каждого PDF-файла в указанной директории
    for file_name in os.listdir(pdf_folder_path):
        if file_name.endswith('.pdf'):
            file_path = os.path.join(pdf_folder_path, file_name)
            data, num_pages = extract_patient_info(file_path)
            if data is not None:
                # Обновляем словарь data данными о количестве страниц
                data['Количество стр'] = num_pages
                # Создаем DataFrame из словаря data
                data_df = pd.DataFrame([data])
                # Объединяем DataFrame с общим DataFrame all_data
                all_data = pd.concat([all_data, data_df], ignore_index=True)

            # # Измеряем время выполнения для каждого файла для видео
            # # После видео этот блок следует закомментировать
            # start_time = time.time()

            # # Обрабатываем файл и измеряем прогресс с помощью tqdm
            # data = extract_patient_info(file_path)
            # if data is not None:
            #     all_data = pd.concat([all_data, pd.DataFrame([data])], ignore_index=True)

            # # Выводим время выполнения для каждого файла
            # print(f"Время обработки файла {file_name}: {time.time() - start_time} сек.")
            # pbar.update(1)  # Увеличиваем прогресс на 1

# Обработка столбца 'Фамилия, имя, отчество пациента'
all_data['пациент'] = all_data['пациент'].apply(lambda x: ' '.join(str(x).strip().split()) if isinstance(x, str) else x)

# Извлекаем последнее слово из пути к папке
output_folder_name = os.path.basename(os.path.normpath(pdf_folder_path))
# Запрашиваем у пользователя имя файла для сохранения данных (используем последнее слово из пути к папке)
excel_file_name = input(f"Введите имя Excel-файла (без расширения, по умолчанию {output_folder_name}.xlsx): ")
if not excel_file_name:
    excel_file_name = output_folder_name

# Проверяем расширение файла
if not excel_file_name.endswith('.xlsx'):
    excel_file_name += '.xlsx'

# Сохраняем данные в Excel-файл с использованием xlsxwriter
output_file_path = os.path.join(pdf_folder_path, excel_file_name)
if not all_data.empty:
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        all_data.to_excel(writer, index=False)
        
        # Получение объекта workbook и worksheet
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Проверка на наличие дубликатов по указанным полям
        duplicate_rows = all_data.duplicated(subset=['протокол ', 'пациент'], keep=False)

        # Сортировка строк по указанным полям
        all_data.sort_values(by=['протокол ', 'пациент'], inplace=True)

        # Создание формата для подсветки дубликатов
        duplicate_format = workbook.add_format({'bg_color': 'yellow'})

        # Применение подсветки к дубликатам строк
        for idx, value in enumerate(duplicate_rows):
            if value:
                worksheet.set_row(idx + 1, cell_format=duplicate_format)

        print(f"\nДанные успешно сохранены в файл: {os.path.abspath(output_file_path)}")
else:
    print("\nНет данных для сохранения в Excel-файл.")
