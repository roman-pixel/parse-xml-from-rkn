import datetime
import shutil
import csv
import openpyxl
from progress.spinner import Spinner
from operator import itemgetter
import numpy as np

spinner = Spinner('\033[43mОбработка данных:\033[0m')


def data_filter_max_speed(max_speed_data):
    max_speed_for_sort = []

    max_speed = ''
    count = 0
    is_convert = False

    for i in max_speed_data:
        max_speed_split = i.split(' ')

        if max_speed_split[1] == 'Мб/с':
            mb_to_kb = int(max_speed_split[0]) * 1000
            max_speed_split[0] = mb_to_kb
            max_speed_split[1] = 'Кб/с'
            is_convert = True

        max_speed_for_sort.append(
            [int(max_speed_split[0]), max_speed_split[1]])

        count += 1

        if count == len(max_speed_data):
            count = 0
            break

    max_speed_for_sort.sort(key=lambda x: x[0], reverse=True)

    if is_convert:
        speed = round(max_speed_for_sort[0][0] / 1000)
        max_speed = str(speed) + ' ' + 'Мб/с'
    else:
        max_speed = str(max_speed_for_sort[0][0]) + \
            ' ' + max_speed_for_sort[0][1]

    return max_speed

# def data_count_columns(data):


# region этот код вибирал из файла только данные по мтс, мегафону, билайну и теле2
# def data_filter_operators(data):
#     data = []

#     for i in data:
#         if ('МегаФон' in i[7]) or ('МТС' in i[7]) or ('мтс' in i[7]) or ('Т2' in i[7]) or ('Вымпел' in i[7]):
#             data.append(i)

#     return data
# endregion


def data_writer(work_sheet, data):
    count = 0
    count_row = 2

    check_new_fias = []
    max_speed = []

    is_ts = False
    is_uslugi_svyazi = False

    for i in data:
        spinner.next()

        # region
        if count == 0:
            check_new_fias.append(i[1])

        if i[1] not in check_new_fias:
            ts_str = 'нет'
            uslugi_svyazi_str = 'нет'

            if is_ts:
                ts_str = 'да'
            if is_uslugi_svyazi:
                uslugi_svyazi_str = 'да'

            is_ts = False
            is_uslugi_svyazi = False

            work_sheet['T' + str(count_row)] = ts_str  # ТС
            # Услуги связи
            work_sheet['U' + str(count_row)] = uslugi_svyazi_str

            for row in work_sheet.iter_rows(min_row=count_row, max_row=count_row, min_col=8, max_col=19):
                for cell in row:
                    # print(cell.value)
                    if cell.value is None:
                        cell.value = 'нет'

            # этот говнокод ничто иное, как суммарная информация по всем операторам
            svod_mts = 'МТС - GSM - '+str(work_sheet['H' + str(count_row)].value)+'; МТС - UMTS - '+str(
                work_sheet['I' + str(count_row)].value)+'; МТС - LTE - '+str(work_sheet['J' + str(count_row)].value)+'; '
            svod_megafon = 'МегаФон - GSM - '+str(work_sheet['K' + str(count_row)].value)+'; МегаФон - UMTS - '+str(
                work_sheet['L' + str(count_row)].value)+'; МегаФон - LTE - '+str(work_sheet['M' + str(count_row)].value)+'; '
            svod_beeline = 'Билайн - GSM - '+str(work_sheet['N' + str(count_row)].value)+'; Билайн - UMTS - '+str(
                work_sheet['O' + str(count_row)].value)+'; Билайн - LTE - '+str(work_sheet['P' + str(count_row)].value)+'; '
            svod_tele2 = 'Т2 - GSM - '+str(work_sheet['Q' + str(count_row)].value)+'; Т2 - UMTS - '+str(
                work_sheet['R' + str(count_row)].value)+'; Т2 - LTE - '+str(work_sheet['S' + str(count_row)].value)+';'

            work_sheet['AA' + str(count_row)] = svod_mts + \
                svod_megafon + svod_beeline + svod_tele2

            # запись максимальной скорости
            work_sheet['V' + str(count_row)] = data_filter_max_speed(max_speed)
            max_speed.clear()

            check_new_fias.append(i[1])
            count_row += 1
        # endregion

        max_speed.append(i[10])

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

        if i[8] == 'да':
            is_ts = True
        if i[9] == 'да':
            is_uslugi_svyazi = True

        count += 1

    spinner.finish()
    return


def write_xlsx_file(data):

    work_book = openpyxl.Workbook()
    work_sheet = work_book.active
    work_sheet.title = "Свод по операторам"

    # region Заполняем шапку таблицы
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
    # endregion

    file_data = open('data.txt', 'w')

    for i in data:
        file_data.write(str(i[1]) + ' - ' + str(i[6]) + ' - ' + str(i[8]) +
                        ' - ' + str(i[9]) + ' - ' + str(i[10]) + ' - ' + str(i[7]) + '\n')
    file_data.close()

    data_writer(work_sheet, data)

    try:
        work_book.save('./processed_file/%s' % 'Свод по операторам.xlsx')
    except PermissionError:
        print(
            "\033[31mФайл открыт в другой программе. Закройте файл и повторите попытку\033[0m")
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
            "\033[31mФайл открыт в другой программе. Закройте файл и повторите попытку\033[0m")
        return

    return data


def main():
    bd = datetime.datetime.now()

    csv_file = './processed_file/%s' % 'itog.csv'
    csv_file_dest = './source_file/%s' % 'itog.csv'

    try:
        shutil.copyfile(csv_file, csv_file_dest)
    except IOError:
        print("\033[31mФайл itog.csv не найден\033[0m")
        return

    write_xlsx_file(read_csv_file(csv_file_dest))

    ed = datetime.datetime.now()
    total_time = ed - bd
    print('\033[32mОбработка завершена\033[0m')
    print('Время выполения программы: ')
    print(total_time)
    return


main()
