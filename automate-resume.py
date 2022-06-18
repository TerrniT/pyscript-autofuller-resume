# pip install docxtpl
import json
import locale
import os

from datetime import datetime as dt
from docxtpl import DocxTemplate

# 
def filling_doc():
    locale.setlocale(locale.LC_ALL, '')
    doc = DocxTemplate("data.docx")
    user = json.load(open(os.path.join(os.getcwd(), 'user_info.json'), encoding='utf-8'))
    if not os.path.isdir(os.path.join(os.getcwd(), 'personal')):
        os.mkdir(os.path.join(os.getcwd(), 'personal'))
    for usr in user:
        print(f'\r[+] Заполняю: {usr["last_name"]}', end='')
        data = {'manager': 'И. С. Иванов', 'reason': 'реструктуризацией и оптимизацией закваски',
                'date': dt.strftime(dt.now(), '%d %B %Y'), 'fio': f'{usr["last_name"]} {usr["first_name"]} '
                                                                  f'{usr["middle_name"]}', 'post': usr['post'],
                'first_middle': f'{usr["first_name"]} {usr["middle_name"]}',
                'contract_date': dt.strftime(dt.strptime(usr['contract_date'], "%d.%m.%Y"), '%d %B %Y'),
                'contract_num': usr['contract_num'], 'day_x': '01.07.2021'}
        doc.render(data)
        doc.save(os.path.join(os.getcwd(), 'personal',
                              f'{usr["last_name"]} {usr["first_name"]} {usr["middle_name"]}.docx'))


def main():
    filling_doc()
    print('\n[+] Все записи обработаны')


if __name__ == "__main__":
    main()
