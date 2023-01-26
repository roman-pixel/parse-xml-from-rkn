with open('43reg.csv', 'r', encoding='utf-8') as f, open('43reg obr.csv', 'w', encoding='windows-1251') as fw:
    count = 1
    sep = ';'
    i=1
    is_tm_sum=0
    line1=''
    stat = True
    is_local_station_sum=0    
    gsm_type = 'нет'
    is_umts_sum = 0
    is_lte_sum = 0
    etv_d_channel_cnt_sum =0	
    payphone_count_sum =0	
		
    strtr='num;place_id;fias_guid;region_code;region_name;city;rayon;place;is_local_station;is_tm;tm_max_access_speed;tm_type;gsm_type;is_umts;is_lte;etv_d_channel_cnt;payphone_count;url\n'
    fw.write(strtr)
    tm_max_access_speed	= 0
    tm_type ='Кб/с'	

    for line in f:
        if stat:
            mas1=line.split(';')
            stat=False
        else:    
            

            mas=line.split(';')
            stat=False
            if mas1[0] != mas[0]:
                is_tm_sum = is_tm_sum + int(mas1[10])
                is_local_station_sum = is_local_station_sum + int(mas1[9])
                

                if tm_type == 'Кб/с' and mas1[12]!= 'Кб/с':
                    tm_type = mas1[12]     
                    tm_max_access_speed = int(mas1[11])
                    
                if tm_type == 'Мб/с' and mas1[12] == 'Гб/с':
                    tm_type = mas1[12]     
                    tm_max_access_speed = int(mas1[11])
                    
                if tm_type == mas1[12]:
                    if tm_max_access_speed < int(mas1[11]):
                        tm_max_access_speed = int(mas1[11])

                if mas1[13] != 'нет':    
                    gsm_type ='2g'
    
                if isinstance(is_umts_sum, str) == False:
                    if is_umts_sum > 0:
                        is_umts_sum = '3g'    
                    else:
                        is_umts_sum = is_umts_sum + int(mas1[14])
                        
                if isinstance(is_lte_sum, str) == False:                    
                    if is_lte_sum >0:
                        is_lte_sum = '4g'  
                    else:
                        is_lte_sum = is_lte_sum + int(mas1[15])
                    
                etv_d_channel_cnt_sum = etv_d_channel_cnt_sum + int(mas1[16])
                payphone_count_sum = payphone_count_sum + int(mas1[17])            
                
                line1 = mas1[0]+sep+mas1[1]+sep+mas1[2]+sep+mas1[3]+sep+mas1[4]+sep+mas1[5]+sep+mas1[6]+sep+mas1[7]+sep
                line1 = line1 + str(is_local_station_sum)+sep+str(is_tm_sum)+sep
                line1 = line1 + str(tm_max_access_speed)+ sep+str(tm_type+sep)
                line1 = line1 + gsm_type +sep+str(is_umts_sum) +sep+str(is_lte_sum)+sep
                line1 = line1 + str(etv_d_channel_cnt_sum) +sep+str(payphone_count_sum)+'\n'  
                fw.write(line1)
                
                print(line1)
                count = count + 1
                i = 1
                is_tm_sum = 0
                is_local_station_sum = 0
                tm_max_access_speed	= 0
                tm_type ='Кб/с'	
                
                gsm_type = 'нет'
                is_umts_sum = 0
                is_lte_sum = 0
                etv_d_channel_cnt_sum =0	
                payphone_count_sum =0	
                
                is_tm_sum = is_tm_sum + int(mas[10])
                is_local_station_sum = is_local_station_sum + int(mas[9])
                

                if tm_type == 'Кб/с' and mas[12]!= 'Кб/с':
                    tm_type = mas[12]     
                    tm_max_access_speed = int(mas[11])
                    
                if tm_type == 'Мб/с' and mas[12] == 'Гб/с':
                    tm_type = mas[12]     
                    tm_max_access_speed = int(mas[11])
                    
                if tm_type == mas[12]:
                    if tm_max_access_speed < int(mas[11]):
                        tm_max_access_speed = int(mas[11])
                        
                if mas[13] != 'нет':    
                    gsm_type ='2g'
    
                if isinstance(is_umts_sum, str) == False:
                    if is_umts_sum > 0:
                        is_umts_sum = '3g'    
                    else:
                        is_umts_sum = is_umts_sum + int(mas[14])
                        
                if isinstance(is_lte_sum, str) == False:                    
                    if is_lte_sum >0:
                        is_lte_sum = '4g'  
                    else:
                        is_lte_sum = is_lte_sum + int(mas[15])
                    
                etv_d_channel_cnt_sum = etv_d_channel_cnt_sum + int(mas[16])
                payphone_count_sum = payphone_count_sum + int(mas[17])            
                
                line1 = mas[0]+sep+mas[1]+sep+mas[2]+sep+mas[3]+sep+mas[4]+sep+mas[5]+sep+mas[6]+sep+mas[7]+sep
                line1 = line1 + str(is_local_station_sum)+sep+str(is_tm_sum)+sep
                line1 = line1 + str(tm_max_access_speed)+ sep+str(tm_type+sep)
                line1 = line1 + gsm_type +sep+str(is_umts_sum) +sep+str(is_lte_sum)+sep
                line1 = line1 + str(etv_d_channel_cnt_sum) +sep+str(payphone_count_sum)+'\n'
                 
            else:    
                is_tm_sum = is_tm_sum + int(mas[10])
                is_local_station_sum = is_local_station_sum + int(mas[9])
                

                if tm_type == 'Кб/с' and mas[12]!= 'Кб/с':
                    tm_type = mas[12]     
                    tm_max_access_speed = int(mas[11])
                    
                if tm_type == 'Мб/с' and mas[12] == 'Гб/с':
                    tm_type = mas[12]     
                    tm_max_access_speed = int(mas[11])
                    
                if tm_type == mas[12]:
                    if tm_max_access_speed < int(mas[11]):
                        tm_max_access_speed = int(mas[11])
                        
                           
                if mas[13] != 'нет':    
                    gsm_type ='2g'
    
                if isinstance(is_umts_sum, str) == False:
                    if is_umts_sum > 0:
                        is_umts_sum = '3g'    
                    else:
                        is_umts_sum = is_umts_sum + int(mas[14])
                        
                if isinstance(is_lte_sum, str) == False:                    
                    if is_lte_sum >0:
                        is_lte_sum = '4g'  
                    else:
                        is_lte_sum = is_lte_sum + int(mas[15])
                    
                etv_d_channel_cnt_sum = etv_d_channel_cnt_sum + int(mas[16])
                payphone_count_sum = payphone_count_sum + int(mas[17])            
                
                line1 = mas[0]+sep+mas[1]+sep+mas[2]+sep+mas[3]+sep+mas[4]+sep+mas[5]+sep+mas[6]+sep+mas[7]+sep
                line1 = line1 + str(is_local_station_sum)+sep+str(is_tm_sum)+sep
                line1 = line1 + str(tm_max_access_speed)+ sep+str(tm_type+sep)
                line1 = line1 + gsm_type +sep+str(is_umts_sum) +sep+str(is_lte_sum)+sep
                line1 = line1 + str(etv_d_channel_cnt_sum) +sep+str(payphone_count_sum)+'\n'
                #print(line1)
                
                z=mas[0]
                i = i + 1
            mas1=mas    
                
    # laststr = '5185;167214;b3dfd4c8-ff6c-4901-93c2-6bb2e38f65ff;43;Кировская обл;0;Нолинский р-н;сдт Надежда (Рябиновщина), тер;1;1;2;Мб/с;нет;0;0;0;2'
    # fw.write(laststr)