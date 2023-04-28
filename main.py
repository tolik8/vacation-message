"""
Повідомлення про відпустку - кадри (Шамрик)
"""

import datetime
import os
# pip install python-docx
from docx import Document

# константи, початкові змінні
date_format = '%d.%m.%Y'
data_list = 'list.csv'
template = 'template.docx'

# очищаємо екран
os.system('cls' if os.name == 'nt' else 'clear')

# отримуємо робочий каталог
folder_path = os.path.dirname(os.path.realpath(__file__))
data_path = folder_path + os.sep + 'data' + os.sep
result_path = folder_path + os.sep + 'result' + os.sep

if not os.path.exists(result_path):
    os.makedirs(result_path, exist_ok=True)

# видаляємо всі файли з папки result
for file_name in os.listdir(result_path):
    os.remove(os.path.join(result_path, file_name))

lines = []
# відкриваємо текстовий CSV файл для читання
with open(data_path + data_list, 'r') as f:
    # проходимо по кожному рядку в файлі
    for line in f:
        lines.append(line)

pib_remember = ''
position_remember = ''

for line in lines:
    # розділяємо значення з рядка
    values = line.strip().split(';')
    if values[1] == '':
        values[1] = pib_remember
    if values[2] == '':
        values[2] = position_remember

    # присвоюємо значення
    date2 = values[0]
    date1_obj = datetime.datetime.strptime(date2, date_format) - datetime.timedelta(weeks=2)
    date1 = date1_obj.strftime(date_format)
    date1_values = date1.strip().split('.')
    mydate1 = date1_values[2] + '-' + date1_values[1] + '-' + date1_values[0]
    pib = values[1]
    position = values[2]
    pib_remember = pib
    position_remember = position

    date2_values = date2.strip().split('.')
    mydate2 = date2_values[2] + '-' + date2_values[1] + '-' + date2_values[0]

    pib_values = pib.strip().split(' ')
    pib1 = pib_values[0]
    pib2_short = pib_values[1][0]
    pib3_short = pib_values[2][0]
    pib_short = pib1 + ' ' + pib2_short + '.' + pib3_short + '.'

    # відкриваємо шаблон документа Word
    document = Document(data_path + template)

    # проходимо по всіх параграфах документу
    for paragraph in document.paragraphs:
        # замінюємо частину тексту, якщо вона містить потрібне значення
        if '%DATE2%' in paragraph.text:
            paragraph.text = paragraph.text.replace('%DATE2%', date2)

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = cell.text.replace("%NAME%", pib_short)
                cell.text = cell.text.replace("%POSITION%", position)
                cell.text = cell.text.replace("%DATE1%", date1)

    # зберігаємо новий документ Word зі значеннями з рядка
    print(f'result{os.sep}{mydate1} {pib1} {pib2_short}{pib3_short} {mydate2}.docx')
    document.save(f'result{os.sep}{mydate1} {pib1} {pib2_short}{pib3_short} {mydate2}.docx')
