import datetime
import xml.etree.ElementTree as ET
import os
import fnmatch
##############################################################################################


def node_processing(xml):
    mas = dict()
    for table in xml.iter('place'):
        for child in table:
            mas[child.tag] = child.text

    if mas['region_code'] != '43':
        return

    for table in xml.iter('record'):
        for child in table:
            if child.tag != 'place':
                mas[child.tag] = child.text
    # print(mas)
    mas['check_url'] = 'https://reestr-svyaz.rkn.gov.ru/place/' + \
        mas['place_id'] + '.htm'
    return mas
##############################################################################################


def RightField(fname):
    mas = ['place_id', 'fias_guid', 'region_code', 'region_name', 'city', 'rayon', 'place', 'os_name',
           'is_tm', 'tm_max_access_speed', 'tm_type',
           'gsm_type', 'is_umts', 'is_lte',
           'etv_d_channel_cnt',
           'is_local_station',
           'payphone_count', 'ap_cnt', 'check_url']
    for x in mas:
        # print(x)
        if fname == x:
            return True
    return False
##############################################################################################


def form_record(mas, count_iter):
    line = ''
    sep = ';'
    for x in mas:
        if RightField(x):
            if mas[x] == None or mas[x] == '':
                tmp = 0
            else:
                tmp = mas[x]

            # замена названий городов, районов и населенных пунктов на пустое значение
            if x == 'city' and tmp == 0 or x == 'rayon' and tmp == 0 or x == 'place' and tmp == 0:
                tmp = ''

            # замена названий типов сетей, ТС и телематических услуг связи на да или нет
            if x == 'is_tm' or x == 'is_local_station' or x == 'is_umts' or x == 'is_lte':
                if tmp == '0' or tmp == 0:
                    tmp = 'нет'
                elif tmp == '1' or tmp == 1:
                    tmp = 'да'

            line = line + sep + str(tmp)
    line = str(count_iter) + line + '\n'
    return line
##############################################################################################


def sum_mas(mas, tmp):
    if mas['is_tm'] == None or mas['is_tm'] == '':
        mas['is_tm'] = 0

    if mas['is_umts'] == None or mas['is_umts'] == '':
        mas['is_umts'] = 0

    if mas['is_lte'] == None or mas['is_lte'] == '':
        mas['is_lte'] = 0

    if mas['etv_d_channel_cnt'] == None or mas['etv_d_channel_cnt'] == '':
        mas['etv_d_channel_cnt'] = 0

    if mas['is_local_station'] == None or mas['is_local_station'] == '':
        mas['is_local_station'] = 0

    if mas['payphone_count'] == None or mas['payphone_count'] == '':
        mas['payphone_count'] = 0

    if mas['ap_cnt'] == None or mas['ap_cnt'] == '':
        mas['ap_cnt'] = 0
 ########
    if tmp['is_tm'] == None or tmp['is_tm'] == '':
        tmp['is_tm'] = 0

    if tmp['is_umts'] == None or tmp['is_umts'] == '':
        tmp['is_umts'] = 0

    if tmp['is_lte'] == None or tmp['is_lte'] == '':
        tmp['is_lte'] = 0

    if tmp['etv_d_channel_cnt'] == None or tmp['etv_d_channel_cnt'] == '':
        tmp['etv_d_channel_cnt'] = 0

    if tmp['is_local_station'] == None or tmp['is_local_station'] == '':
        tmp['is_local_station'] = 0

    if tmp['payphone_count'] == None or tmp['payphone_count'] == '':
        tmp['payphone_count'] = 0

    if mas['ap_cnt'] == None or mas['ap_cnt'] == '':
        mas['ap_cnt'] = 0

  ############
    if tmp['is_tm'] == 0:
        tmp['is_tm'] = mas['is_tm']

    if tmp['is_umts'] == 0:
        tmp['is_umts'] = mas['is_umts']

    if tmp['is_lte'] == 0:
        tmp['is_lte'] = mas['is_lte']

    tmp['etv_d_channel_cnt'] = int(
        mas['etv_d_channel_cnt'])+int(tmp['etv_d_channel_cnt'])

    if tmp['is_local_station'] == 0:
        tmp['is_local_station'] = mas['is_local_station']

    tmp['payphone_count'] = int(
        mas['payphone_count'])+int(tmp['payphone_count'])

    if tmp['gsm_type'] == None or tmp['gsm_type'] == '':
        tmp['gsm_type'] = 'нет'

    if mas['gsm_type'] == None or mas['gsm_type'] == '':
        mas['gsm_type'] = 'нет'
    else:
        if tmp['gsm_type'].find('900') == -1 or tmp['gsm_type'].find('1800') == -1:
            tmp['gsm_type'] = mas['gsm_type']

    return tmp

##############################################################################################


def shapka_record():
    line = '№;Уникальный идентификатор населенного пункта в информационной системе Роскомнадзора;Код ФИАС;Код региона;Наименование региона;Город;Район;Населенный пункт;Оператор связи;Услуги местной ТС;Телематические услуги связи;Телематические услуги связи: максимальная скорость передачи;Единица измерения;Наличие GSM;Наличие UMTS;Наличие LTE;Эфирное телевидение: количество цифровых каналов;Количество таксофонов;Количество точек доступа;URL\n'
    return line

##############################################################################################


def main():
    if os.path.exists('./processed_file') == False:
        os.mkdir("processed_file")
        print('\033[34mСоздана папка processed_file\033[0m')
    if os.path.exists('./source_file') == False:
        os.mkdir("source_file")
        print('\033[34mСоздана папка source_file. Поместите в нее файлы xml и xlsx для обработки и запустите программу еще раз\033[0m')
        return

    # ищем тот самый файл с расширением xml
    xml_file = ''
    count_of_xml_files = 0
    listOfFiles = os.listdir('./source_file')
    pattern = "*.xml"
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern):
            count_of_xml_files += 1

            if count_of_xml_files > 1:
                print('\033[31mНайдено более одного файла с расширением xml\033[0m')
                return

            xml_file = entry

    if xml_file == '':
        print('\033[31mФайл c расширением .xml не найден\033[0m')
        return

    bd = datetime.datetime.now()
    count = 1
    stat = False
    node = ''
    fias_guid = ''
    count_iter = 0
    lines = 0

    try:
        with open('./source_file/%s' % xml_file, 'r', encoding='utf-8') as f, open('./processed_file/%s' % 'Выгрузка.csv', 'w', encoding='windows-1251') as fw:

            fw.write(shapka_record())

            for line in f:
                line = line.lstrip('\t ')
                line = line.rstrip('\n ')
                line = line.replace('rkn:', '')

                if line == '<record>':
                    stat = True

                if line == '</record>':
                    node = node + line

                    try:
                        xml = ET.fromstring(node)

                    except SyntaxError as e:
                        print(format(e.message))

                    mas = node_processing(xml)

                    if (mas):
                        count_iter += 1
                        # if (mas['is_tm']!=None):
                        if fias_guid != mas['fias_guid']:
                            print("Количество записей: " + str(count_iter))

                            fw.write(form_record(mas, count_iter))
                            fias_guid = mas['fias_guid']

                        else:
                            fw.write(form_record(mas, count_iter))

                    count += 1
                    node = ''
                    stat = False

                if stat:
                    node = node + line + '\n'
    except IOError:
        print("\033[31mФайл открыт в другой программе. Закройте файл и повторите попытку\033[0m")

    ed = datetime.datetime.now()
    # print(bd)
    # print(ed)
    print('\033[32mОбработка завершена\033[0m')
    total_time = ed - bd
    print('Время выполения программы: ')
    print(total_time)
    return


##############################################################################################
main()
