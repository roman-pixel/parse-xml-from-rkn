import datetime
import shutil
import csv
import openpyxl
from progress.spinner import Spinner
from operator import itemgetter
import numpy as np


def write_xlsx_file(data):

    work_book = openpyxl.Workbook()
    work_sheet = work_book.active
    work_sheet.title = "Свод по операторам"

    # Заполняем шапку таблицы
    work_sheet['A1'] = 'Уникальный идентификатор населенного пункта в информационной системе Роскомнадзора'
    work_sheet['B1'] = 'Код ФИАС'
    work_sheet['C1'] = 'Код региона'
    work_sheet['D1'] = 'Наименование региона'
    work_sheet['E1'] = 'Город'
    work_sheet['F1'] = 'Район'
    work_sheet['G1'] = 'Населенный пункт'
    work_sheet['H1'] = 'МТС GSM'
    work_sheet['I1'] = 'МТС UMTS'
    work_sheet['J1'] = 'МТС LTE'
    work_sheet['K1'] = 'МегаФон GSM'
    work_sheet['L1'] = 'МегаФон UMTS'
    work_sheet['M1'] = 'МегаФон LTE'
    work_sheet['N1'] = 'Билайн GSM'
    work_sheet['O1'] = 'Билайн UMTS'
    work_sheet['P1'] = 'Билайн LTE'
    work_sheet['Q1'] = 'Теле2 GSM'
    work_sheet['R1'] = 'Теле2 UMTS'
    work_sheet['S1'] = 'Теле2 LTE'
    work_sheet['T1'] = 'Услуги метсной ТС'
    work_sheet['U1'] = 'Телематические услуги связи'
    # телематические услуги связи (скорость) и единицы измероение в одну ячейку
    work_sheet['V1'] = 'Телематические услуги связи (максимальная скорость передачи данных)'
    # сумма
    work_sheet['W1'] = 'Эфирное телевидение (количество цифровых каналов)'
    work_sheet['X1'] = 'Количество таксофонов'  # сумма
    work_sheet['Y1'] = 'Количество точек доступа'  # сумма
    work_sheet['Z1'] = 'URL'  # url для проверки
    work_sheet['AA1'] = 'Доп. информация'

    # дополнительная информация в виде:
    # 4 основных оператора:
    # МТС- GSM-900/1800; МТС- UMTS- нет; МТС- LTE - нет;
    # Билайн - GSM-900; Билайн - UMTS- нет; Билайн - LTE - нет;
    # Мегафон - GSM-нет; Мегафон - UMTS- нет; Мегафон - LTE - нет;
    # Теле2 - GSM-нет; Теле2 - UMTS- нет; Теле2 - LTE - нет

    # region переменные
    count = 0
    count_row = 2
    data_with_four_operators = []
    check_new_fias = []
    uslugi_TS = 'нет'
    telematicheskie_uslugi = 'нет'
    # sum_tv = 0
    # sum_tax = 0
    # sum_point = 0
    # file_non_doubles = open('non_doubles.txt', 'w')
    # file_data_with_four_operators = open('data_with_four_operators.txt', 'w')
    # endregion

    for i in data:
        # if i[8] == 'да':
        #     uslugi_TS = 'да'
        # if i[9] == 'да':
        #     telematicheskie_uslugi = 'да'

        if ('МегаФон' in i[7]) or ('МТС' in i[7]) or ('мтс' in i[7]) or ('Т2' in i[7]) or ('Вымпел' in i[7]):
            data_with_four_operators.append(i)

    # for i in data_with_four_operators:
    #     file_data_with_four_operators.write(str(i) + '\n')
    # file_data_with_four_operators.close()

    # n = []
    # for i in data_with_four_operators:
    #     if i[1] not in n:
    #         n.append(i[1])
    #         file_non_doubles.write(str(i) + '\n')
    # file_non_doubles.close()

    #############################################

    spinner = Spinner('\033[43mОбработка данных:\033[0m')

    for i in data_with_four_operators:
        spinner.next()

        if count == 0:
            check_new_fias.append(i[1])

        if i[1] not in check_new_fias:
            for row in work_sheet.iter_rows(min_row=count_row, max_row=count_row, min_col=8, max_col=19):
                for cell in row:
                    # print(cell.value)
                    if cell.value is None:
                        cell.value = 'нет'

            svod_mts = 'МТС - GSM - '+str(work_sheet['H' + str(count_row)].value)+'; МТС - UMTS - '+str(
                work_sheet['I' + str(count_row)].value)+'; МТС - LTE - '+str(work_sheet['J' + str(count_row)].value)+';'
            svod_biline = 'Билайн - GSM - '+str(work_sheet['K' + str(count_row)].value)+'; Билайн - UMTS - '+str(
                work_sheet['L' + str(count_row)].value)+'; Билайн - LTE - '+str(work_sheet['M' + str(count_row)].value)+';'
            svod_megafon = 'Мегафон - GSM - '+str(work_sheet['N' + str(count_row)].value)+'; Мегафон - UMTS - '+str(
                work_sheet['O' + str(count_row)].value)+'; Мегафон - LTE - '+str(work_sheet['P' + str(count_row)].value)+';'
            
            work_sheet['AA' + str(count_row)] = svod_mts + \
                svod_biline + svod_megafon

            check_new_fias.append(i[1])
            count_row += 1

        work_sheet['A' + str(count_row)] = i[0]  # ID
        work_sheet['B' + str(count_row)] = i[1]  # ФИАС
        work_sheet['C' + str(count_row)] = i[2]  # код региона
        work_sheet['D' + str(count_row)] = i[3]  # наименование региона
        work_sheet['E' + str(count_row)] = i[4]  # город
        work_sheet['F' + str(count_row)] = i[5]  # район
        work_sheet['G' + str(count_row)] = i[6]  # населенный пункт
        work_sheet['Z' + str(count_row)] = i[17]  # url для проверки'

        if 'МТС' in i[7] or 'мтс' in i[7]:
            work_sheet['H' + str(count_row)] = i[11]
            work_sheet['I' + str(count_row)] = i[12]
            work_sheet['J' + str(count_row)] = i[13]
        if 'МегаФон' in i[7]:
            work_sheet['K' + str(count_row)] = i[11]
            work_sheet['L' + str(count_row)] = i[12]
            work_sheet['M' + str(count_row)] = i[13]
        if 'Вымпелк' in i[7]:
            work_sheet['N' + str(count_row)] = i[11]
            work_sheet['O' + str(count_row)] = i[12]
            work_sheet['P' + str(count_row)] = i[13]
        if 'Т2' in i[7]:
            work_sheet['Q' + str(count_row)] = i[11]
            work_sheet['R' + str(count_row)] = i[12]
            work_sheet['S' + str(count_row)] = i[13]

        count += 1

        # if count < len(data_with_four_operators)-2:
        #     count += 1

    spinner.finish()

    try:
        work_book.save('./processed_file/%s' % 'Свод по операторам.xlsx')
    except PermissionError:
        print(
            "\033[31mФайл открыт в другой программе. Закройте файл и повторите попытку \033[0m")
        return


def read_csv_file(file_name):
    data = []
    count = 0
    file = open('sorted.txt', 'w', encoding='utf-8')
    try:
        with open(file_name, 'r', newline='', encoding='Windows-1251') as r_file:
            reader = csv.reader(r_file, delimiter=';')
            # добавляем в список город, операторов и наличие типов связи
            for row in reader:
                if count == 0:
                    count += 1
                    continue
                # объединяем показания скорости и единицы измерения
                merge_rows = row[11] + ' ' + row[12]
                data.append([row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10],
                             merge_rows, row[13], row[14], row[15], row[16], row[17], row[18], row[19]])

            # сортируем по населенному пункту и району
            data = sorted(data, key=itemgetter(6, 5, 1))

            for i in data:
                file.write(str(i)+'\n')
            file.close()

    except IOError:
        print(
            "\033[31mФайл открыт в другой программе. Закройте файл и повторите попытку \033[0m")
        return

    return data


def main():
    bd = datetime.datetime.now()

    csv_file = './processed_file/%s' % 'itog.csv'
    csv_file_dest = './source_file/%s' % 'itog.csv'

    try:
        shutil.copyfile(csv_file, csv_file_dest)
    except IOError:
        print("\033[31mФайл itog.csv не найден \033[0m")
        return

    write_xlsx_file(read_csv_file(csv_file_dest))

    ed = datetime.datetime.now()
    total_time = ed - bd
    print('\033[32mОбработка завершена\033[0m')
    print('Время выполения программы: ')
    print(total_time)
    return


main()
