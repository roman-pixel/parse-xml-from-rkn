import datetime
import shutil
import csv
import openpyxl
from progress.spinner import Spinner
from itertools import groupby


def write_xlsx_file(data):

    work_book = openpyxl.Workbook()
    work_sheet = work_book.active
    work_sheet.title = "Свод по операторам"

    # Заполняем шапку таблицы
    work_sheet['A1'] = 'Населенный пункт'
    work_sheet['B1'] = 'МТС GSM'
    work_sheet['C1'] = 'МТС UMTS'
    work_sheet['D1'] = 'МТС LTE'
    work_sheet['E1'] = 'Билайн GSM'
    work_sheet['F1'] = 'Билайн UMTS'
    work_sheet['G1'] = 'Билайн LTE'
    work_sheet['H1'] = 'Мегафон GSM'
    work_sheet['I1'] = 'Мегафон UMTS'
    work_sheet['J1'] = 'Мегафон LTE'
    work_sheet['K1'] = 'Теле2 GSM'
    work_sheet['L1'] = 'Теле2 UMTS'
    work_sheet['M1'] = 'Теле2 LTE'
    work_sheet['N1'] = 'Доп. информация'

    count = 0
    count_row = 2
    mas = []
    # file_non_doubles = open('non_doubles.txt', 'w')

    for i in data:
        if ('МегаФон' in i[1]) or ('МТС' in i[1]) or ('мтс' in i[1]) or ('Т2' in i[1]) or ('Вымпел' in i[1]):
            mas.append(i)

    mas.sort()

    # added = set()
    # new_arr = []
    # for el in mas:
    #     if el[0] not in added:
    #         new_arr.append(el)
    #         added.add(el[0])
    #         file_non_doubles.write(str(el)+'\n')

    # file_non_doubles.close()

    # print(*new_arr, sep='\n')
    
    spinner = Spinner('Обработка данных: ')

    for i in mas:
        spinner.next()

        work_sheet['A' + str(count_row)] = i[0]

        if 'МТС' in i[1] or 'мтс' in i[1]:
            work_sheet['B' + str(count_row)] = i[2]
            work_sheet['C' + str(count_row)] = i[3]
            work_sheet['D' + str(count_row)] = i[4]
        if 'Вымпелк' in i[1]:
            work_sheet['E' + str(count_row)] = i[2]
            work_sheet['F' + str(count_row)] = i[3]
            work_sheet['G' + str(count_row)] = i[4]
        if 'МегаФон' in i[1]:
            work_sheet['H' + str(count_row)] = i[2]
            work_sheet['I' + str(count_row)] = i[3]
            work_sheet['J' + str(count_row)] = i[4]
        if 'Т2' in i[1]:
            work_sheet['K' + str(count_row)] = i[2]
            work_sheet['L' + str(count_row)] = i[3]
            work_sheet['M' + str(count_row)] = i[4]

        if i[0] != mas[count+1][0]:
            for row in work_sheet.iter_rows(min_row=count_row, max_row=count_row, min_col=2, max_col=13):
                for cell in row:
                    # print(cell.value)
                    if cell.value is None:
                        cell.value = 'нет'
            count_row += 1

        if count < len(mas)-2:
            count += 1

    spinner.finish()

    try:
        work_book.save('./processed_file/%s' % 'Свод по операторам.xlsx')
    except PermissionError:
        print("Файл открыт в другой программе. Закройте файл и повторите попытку.")
        return


def read_csv_file(file_name):
    data = []
    count = 0
    try:
        with open(file_name, 'r', newline='', encoding='Windows-1251') as r_file:
            reader = csv.reader(r_file, delimiter=';')
            # добавляем в список город, операторов и наличие типов связи
            for row in reader:
                if count == 0:
                    count += 1
                    continue
                data.append([row[7], row[8], row[13], row[14], row[15]])
    except IOError:
        print("Файл открыт в другой программе. Закройте файл и повторите попытку.")
        return

    return data


def main():
    bd = datetime.datetime.now()

    csv_file = './processed_file/%s' % 'itog.csv'
    csv_file_dest = './source_file/%s' % 'itog.csv'

    try:
        shutil.copyfile(csv_file, csv_file_dest)
    except IOError:
        print("Файл itog.csv не найден")
        return

    write_xlsx_file(read_csv_file(csv_file_dest))

    ed = datetime.datetime.now()
    total_time = ed - bd
    print('Время выполения программы: ')
    print(total_time)
    return


main()
