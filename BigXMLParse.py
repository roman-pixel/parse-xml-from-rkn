from sys import stdout
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
 #   print(mas)
    return mas
##############################################################################################


def RightField(fname):
    mas = ['place_id', 'fias_guid', 'region_code', 'region_name', 'city', 'rayon', 'place', 'os_name',
           'is_tm', 'tm_max_access_speed', 'tm_type',
           'gsm_type', 'is_umts', 'is_lte',
           'etv_d_channel_cnt',
           'is_local_station',
           'payphone_count']
    for x in mas:
        # print(x,fname)
        if fname == x:
            return True
    return False
##############################################################################################


def form_record(mas, number):
    line = ''
    line1 = ''
    sep = ';'
    for x in mas:
        if RightField(x):
            if mas[x] == None or mas[x] == '':
                tmp = 0
            else:
                tmp = mas[x]
            line = line + sep + str(tmp)
            # line1 = line1 + sep + x
    line = str(number) + line + '\n'
    # print(line1)
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


def main():
    listOfFiles = os.listdir('.')
    pattern = "*.xml"
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern):
            xml_file = entry

    bd = datetime.datetime.now()
    count = 1
    stat = False
    node = ''
    fias_guid = ''
    number = 0
    with open(xml_file, 'r', encoding='utf-8') as f, open('43reg.csv', 'w', encoding='utf-8') as fw:
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
                    # if (mas['is_tm']!=None):
                    if fias_guid != mas['fias_guid']:
                        number += 1
                        print(number, count)

                        fw.write(form_record(mas, number))
                        fias_guid = mas['fias_guid']

                    else:
                        fw.write(form_record(mas, number))

                count = count + 1
                node = ''
                stat = False

            if stat:
                node = node + line + '\n'
    ed = datetime.datetime.now()
    print(bd)
    print(ed)
    return


##############################################################################################
main()
