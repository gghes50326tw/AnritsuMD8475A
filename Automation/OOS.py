# -*- coding: utf-8 -*-
"""
Created on Fri Nov 26 16:08:15 2021

@author: Vivian
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Nov 15 16:01:54 2021

@author: Vivian
"""
from colorama import init
import pyvisa as visa
import logging
import time
import csv
import configparser
import paramiko
import os
path = os.getcwd()
rslt_path= os.path.join(path, os.pardir,'output','CASE_34','final_result')
if os.path.isdir(rslt_path):
    pass
else:
    os.makedirs(rslt_path)
rsltcsv_path = os.path.join(path, os.pardir,'output','CASE_34','final_result','CASE_34.csv')
rsltlog_path = os.path.join(path, os.pardir,'output','CASE_34','final_result','CASE_34.log')
sorcelte_path = os.path.join(path, 'Comb_Crossover_L.csv')
sorcewcdma_path = os.path.join(path, 'Comb_Crossover_W.csv')
config_path = os.path.join(path, os.pardir,'config','CASE_34.ini')
init(autoreset=True)

config = configparser.ConfigParser()
config.read(config_path)
timeout=int(config.get('timeout','timeout'))

try:
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy)
    ssh.connect(config['SSH']['IP'],'22',config['SSH']['NAME'],config['SSH']['PASSWORD'])
except Exception as e:
    print('[ERRPR]:' + str(e))
    print('Can not establish a SSH connection')
    print("Please make sure you have the correct acess\
          right and the repository exists")
          
cellular_off = 'netsh mbn set powerstate interface=\"Cellular\" state=off'
cellular_on = 'netsh mbn set powerstate interface=\"Cellular\" state=on'

logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%m-%d %H:%M',
                    filename= rsltlog_path,filemode="w")
logger =logging.getLogger("mylogger")
logger.setLevel(logging.INFO)

if not logger.handlers:
    
    fh = logging.FileHandler("CASE_34.log" , encoding="utf-8")
    ch = logging.StreamHandler()
    
    formatter = logging.Formatter(fmt="%(message)s",
                                  datefmt="%Y/%m/%d %X")
    
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    
    logger.addHandler(fh)
    logger.addHandler(ch)

#版號
logger.info('CASE_34')
logger.info('Ver : 1.1.0.2')

#GPIB地址
rm = visa.ResourceManager()
gpib_ad1=config['GPIB']['GPIB1']
try:
    LTE1=rm.open_resource(gpib_ad1)
except visa.VisaIOError :
    print("GPIB NOT CONNECT")
    time.sleep(1)

start = time.time()
while time.time() < start + timeout: 
    try:
        rslt=LTE1.query('*IDN?').rstrip()
        logger.info('GPIB CONNECTION SUECSS')
        break
    except:
        time.sleep(1)

def LTE_CMD(cmd):
    start = time.time()
    while time.time() < start + timeout:
        try:
            LTE1.write(cmd)
            time.sleep(1)
            break
        except Exception as e:
            err = e
            print(err)
logger.info('==========START==========')
#設定參數
studio_start=time.time()   
#開啟csv檔案
with open(rsltcsv_path,'w',newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['UMTS','BAND A','BAND B','Regist Time','Result'])
    
    with open(sorcelte_path, 'r',encoding='UTF-8-sig')as csv_file:
        rows = csv.DictReader(csv_file,delimiter=",")
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.write('LOADSIMPARAM "C:\LTEtoLTE.wnssp2"' )
                logger.info('LOADSIMPARAM,PLEASE WAIT')
                break
            except Exception as e:
                err = e
        
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.query('stat?')
                logger.info('LOAD LTEtoLTE.wnssp2 : SUCESS')
                break
            except Exception as e:
                err = e
                
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.write('start')
                logger.info('START SIMULATION,PLEASE WAIT')
                break
            except Exception as e:
                err = e
        
        LTE_CMD(cmd="stat?")
        logger.info('********************************')        
        for row in rows:
            studio_end=time.time()
            studio_timing = studio_end - studio_start
            if studio_timing >= 18000:
                studio_start=time.time()
                
                logger.info('******RESTART SMARTSTUDIO*******')
                LTE_CMD(cmd="EXIT")
               
                LTE_CMD(cmd="stat?") 
              
                LTE_CMD(cmd="RUN")         
                       
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt=LTE1.query('stat?').rstrip()
                        if rslt == "NOTRUN":
                            break   
                    except Exception as e:
                        err = e
                
                LTE_CMD(cmd="start")        
                
                logger.info('START SIMULATION,PLEASE WAIT')
                LTE_CMD(cmd="stat?")
                
                logger.info('********************************')
            
            ssh.exec_command(cellular_off)
            print('\033[33mcellular_off\033[0m')
            time.sleep(3)
            
            LTE_CMD(cmd="OUTOFSERVICE OUT,BTS1")
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt =='OUT':
                        logger.info('SET BTS1 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["LTE",row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                continue
                
            LTE_CMD(cmd="OUTOFSERVICE OUT,BTS2")
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS2').rstrip()
                    time.sleep(2)
                    if rslt =='OUT':
                        logger.info('SET BTS2 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE OUT,BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 OUT OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["LTE",row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 OUT OF SERVICE : FAIL'])
                continue
            
            LTE_CMD(cmd='DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] +',BTS1')
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('DUPLEXMODE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt ==row['BTS1_DUPLEXMODE']:
                        logger.info('SET BTS1 DUPLEXMODE ' + row['BTS1_DUPLEXMODE']  + ': SUCESS')
                        break
                    else:
                        LTE1.write('DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] +',BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 DUPLEXMODE ' + row['BTS1_DUPLEXMODE']  + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 DUPLEXMODE SETTING FAIL'])
                continue
            
            
            LTE_CMD(cmd='BAND ' + row["BTS1_Band"] + ',BTS1')
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('BAND? BTS1').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt ==  row["BTS1_Band"] :
                logger.info('SET BTS1 BAND ' +  row["BTS1_Band"] + ': SUCESS')
            else:
                logger.error('SET BTS1 BAND ' +  row["BTS1_Band"] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 BAND SETTING FAIL'])
                continue
              
            LTE_CMD(cmd='BANDWIDTH ' + row['BTS1_Bandwidth'] + ' ,BTS1')
        
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('BANDWIDTH? BTS1').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS1_Bandwidth']  :
                logger.info('SET BTS1 BANDWIDTH ' + row['BTS1_Bandwidth'] + ': SUCESS') 
            else:
                logger.error('SET BTS1 BANDWIDTH ' + row['BTS1_Bandwidth'] + ': FAIL') 
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 BANDWIDTH SETTING FAIL'])
                continue
            
            LTE_CMD(cmd='OLVL ' + row['BTS1_Output Level'] + ',BTS1')
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OLVL? BTS1').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS1_Output Level']+'.0':
                logger.info('SET BTS1 OUTPUTLEVEL ' + row['BTS1_Output Level'] + ': SUCESS')
            else:
                logger.info('SET BTS1 OUTPUTLEVEL ' + row['BTS1_Output Level'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 OUTPUTLEVEL SETTING FAIL'])
                continue
                    
            LTE_CMD(cmd='DLCHAN ' + row['BTS1_DL Channel'] + ',BTS1')
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('DLCHAN? BTS1').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS1_DL Channel']:
                logger.info('SET BTS1 DLCHAN ' + row['BTS1_DL Channel'] + ': SUCESS')
            else:
                logger.error('SET BTS1 DLCHAN ' + row['BTS1_DL Channel'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 DLCHAN SETTING FAIL'])
                continue
                    
            LTE_CMD(cmd='DUPLEXMODE ' + row['BTS2_DUPLEXMODE']+ ',BTS2')
            
            cnt=3
            while cnt > 0:
                try:
                    rslt=LTE1.query('DUPLEXMODE? BTS2').rstrip()
                    time.sleep(2)
                    if rslt ==row['BTS2_DUPLEXMODE']:
                        logger.info('SET BTS2 DUPLEXMODE ' + row['BTS2_DUPLEXMODE']  + ': SUCESS')
                        break
                    else:
                        LTE1.write('DUPLEXMODE ' + row['BTS2_DUPLEXMODE'] +',BTS2')
                        time.sleep(5)
                        cnt-=1
                except Exception as e:
                    err = e
            else:
                logger.error('SET BTS2 DUPLEXMODE ' + row['BTS2_DUPLEXMODE'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 DUPLEXMODE SETTING FAIL'])
                continue        
                    
            LTE_CMD(cmd='BAND ' + row['BTS2_Band'] + ',BTS2')

            LTE_CMD(cmd='BAND? BTS2')            

            if rslt == row['BTS2_Band']:
                logger.info('SET BTS2 BAND ' + row['BTS2_Band'] + ': SUCESS')
            else:
                logger.info('SET BTS2 BAND ' + row['BTS2_Band'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 BAND SETTING FAIL'])
                continue
               
            LTE_CMD(cmd='DLCHAN ' + row['BTS2_DL Channel'] + ',BTS2')
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('DLCHAN? BTS2').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS2_DL Channel']:
                logger.info('SET BTS2 DLCHAN ' + row['BTS2_DL Channel'] + ': SUCESS')
            else:
                logger.info('SET BTS2 DLCHAN ' + row['BTS2_DL Channel'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 DLCHAN SETTING FAIL'])
                continue
                    
            LTE_CMD(cmd='BANDWIDTH ' + row['BTS2_Bandwidth']+',BTS2')
             
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('BANDWIDTH? BTS2').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS2_Bandwidth']:
                logger.info('SET BTS2 BANDWIDTH ' + row['BTS2_Bandwidth'] + ': SUCESS')
            else:
                logger.info('SET BTS2 BANDWIDTH ' + row['BTS2_Bandwidth'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 BANDWIDTH SETTING FAIL'])
                continue
                    
            LTE_CMD(cmd='OLVL '  + row['BTS2_Output Level'] + ',BTS2')
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OLVL? BTS2').rstrip()
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            if rslt == row['BTS2_Output Level']+'.0':
                logger.info('SET BTS2 OUTPUTLEVEL ' + row['BTS2_Output Level'] + ': SUCESS')
            else:
                logger.info('SET BTS2 OUTPUTLEVEL ' + row['BTS2_Output Level'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['LTE',row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 OUTPUTLEVEL SETTING FAIL'])
                continue
            
            ssh.exec_command(cellular_on)
            print('\033[33mcellular_on\033[0m')
            time.sleep(3)
            
            LTE_CMD(cmd='OUTOFSERVICE IN,BTS1')
            
            cnt=3
            while cnt > 0:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt =='IN':
                        logger.info('SET BTS1 IN OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE IN,BTS1')
                        time.sleep(5)
                        cnt-=1
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info('SET BTS1 IN OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["LTE",row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 IN OF SERVICE : FAIL'])
                continue
        
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt = LTE1.query('CALLStat? BTS1').rstrip()
                    if rslt == 'NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE':
                        logger.info('BTS1 CONNECTION SUCESS')
                        break
                    else:
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info('OVER-TIME :BTS1 CONNECTION FAIL')
                writer.writerow(['LTE',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 CONNECTION FAIL'])
                logger.info('********************************')
                continue       
            
            time.sleep(5)
            
            LTE_CMD(cmd='OUTOFSERVICE OUT,BTS1')
            
            cnt=3
            while cnt > 0:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt =='OUT':
                        logger.info('SET BTS1 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE IN,BTS1')
                        time.sleep(5)
                        cnt-=1
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["LTE",row['BTS1_Band'],row['BTS2_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                continue
            
            LTE_CMD(cmd='OUTOFSERVICE IN ,BTS2')
        
            cnt=3
            while cnt > 0:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS2').rstrip()
                    time.sleep(2)
                    if rslt =='IN':
                        logger.info('SET BTS2 IN OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE IN,BTS2')
                        time.sleep(5)
                        cnt-=1
                        continue
                except Exception as e:
                    err = e
            else:
                logger.error('SET BTS2 IN OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["LTE",row['BTS1_Band'],row['BTS2_Band'],' X ','BTS2 IN OF SERVICE : FAIL'])
                continue
            
            while time.time() < start + timeout:
                try:
                    rslt = LTE1.query('CALLStat? BTS2').rstrip()
                    if rslt=='NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE':
                        break
                    else:
                        continue
                except Exception as e:
                    err = e
                
            end = time.time()
            if(end-start)<=timeout:
                logger.info('BTS2 CONNETION TIME :' + str(end-start) + ' PASS')
                writer.writerow(['LTE',row["BTS1_Band"],row['BTS2_Band'],str(end-start),'PASS'])
            else:
                logger.error('OVER-TIME :BTS2 CONNECTION FAIL')
                writer.writerow(['LTE',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 CONNECTION FAIL'])
                logger.info('********************************')
                continue
            
            logger.info('********************************')    
                 
    LTE_CMD(cmd='stop')
    
    logger.info('STOP SIMULATION,PLEASE WAIT')
    LTE_CMD(cmd='stat?')
    
    with open(sorcewcdma_path, 'r',encoding='UTF-8-sig')as csv_file:
        rows = csv.DictReader(csv_file,delimiter=",")
        
        start = time.time()
        while time.time() < start + timeout:
            try:
                LTE1.write('LOADSIMPARAM "C:\WCDMAtoWCDMA.wnssp2"' )
                time.sleep(3)
                break
            except Exception as e:
                err = e
        logger.info('LOADSIMPARAM,PLEASE WAIT')
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.query('stat?')
                logger.info('LOAD WCDMAtoWCDMA.wnssp2 : SUCESS')
                break
            except Exception as e:
                err = e
        
        start = time.time()
        while time.time() < start + timeout:
            try:
                LTE1.write('start')
                time.sleep(3)
                break
            except Exception as e:
                err = e
        
        logger.info('START SIMULATION,PLEASE WAIT')
        start = time.time()
        while time.time() < start + timeout:
            try:
                LTE1.query('stat?')
                break
            except Exception as e:
                err = e
        logger.info('********************************')        
        for row in rows:
            studio_end=time.time()
            studio_timing = studio_end - studio_start
            if studio_timing >= 18000:
                studio_start=time.time()
                
                logger.info('******RESTART SMARTSTUDIO*******')
                while time.time() < start + timeout:    
                    try:
                        LTE1.write('EXIT')
                        time.sleep(3)
                        break
                    except Exception as e:
                        err = e
                        
                start = time.time()
                while time.time() < start + timeout:    
                    try:
                        LTE1.query('stat?')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                start = time.time()
                while time.time() < start + timeout:    
                    try:
                        LTE1.write('RUN')
                        time.sleep(3)
                        break
                    except Exception as e:
                        err = e
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt=LTE1.query('stat?').rstrip()
                        if rslt == "NOTRUN":
                            break   
                    except Exception as e:
                        err = e
                        
                start = time.time()
                while time.time() < start + timeout:    
                    try:
                        LTE1.write('start')
                        time.sleep(3)
                        break
                    except Exception as e:
                        err = e
                
                logger.info('START SIMULATION,PLEASE WAIT')
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        LTE1.query('stat?')
                        break
                    except Exception as e:
                        err = e
                        
                logger.info('********************************')            
            ssh.exec_command(cellular_off)
            print('\033[33mcellular_off\033[0m')
            time.sleep(3)
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OUTOFSERVICE OUT,BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
        
            err ="SETTING FAIL" 
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(1)
                    if rslt =='OUT':
                        logger.info('SET BTS1 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["WCDMA",row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                continue
                
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OUTOFSERVICE OUT,BTS2')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err ="SETTING FAIL" 
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS2').rstrip()
                    time.sleep(1)
                    if rslt =='OUT':
                        logger.info('SET BTS2 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE OUT,BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 OUT OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(["WCDMA",row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 OUT OF SERVICE : FAIL'])
                continue
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('BAND ' + row["BTS1_Band"] + ',BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
        
            err ="SETTING FAIL"    
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('BAND? BTS1').rstrip()
                    time.sleep(1)
                    if rslt ==  row["BTS1_Band"] :
                        logger.info('SET BTS1 BAND ' +  row["BTS1_Band"] + ': SUCESS')
                        break
                    else:
                        LTE1.write('BAND ' + row["BTS1_Band"] + ',BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 BAND ' +  row["BTS1_Band"] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 BAND SETTING FAIL'])
                continue
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OLVL ' + row['BTS1_Output Level'] + ',BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OLVL? BTS1').rstrip()
                    time.sleep(1)
                    if rslt == row['BTS1_Output Level']+'.0':
                        logger.info('SET BTS1 OUTPUTLEVEL ' + row['BTS1_Output Level'] + ': SUCESS')
                        break
                    else:
                        LTE1.write('OLVL ' + row['BTS1_Output Level'] + ',BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 OUTPUTLEVEL ' + row['BTS1_Output Level'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 OUTPUTLEVEL SETTING FAIL'])
                continue
                 
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('DLCHAN ' + row['BTS1_DL Channel'] + ',BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('DLCHAN? BTS1').rstrip()
                    time.sleep(1)
                    if rslt == row['BTS1_DL Channel']:
                        logger.info('SET BTS1 DLCHAN ' + row['BTS1_DL Channel'] + ': SUCESS')
                        break
                    else:
                        LTE1.write('DLCHAN ' + row['BTS1_DL Channel'] + ',BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 DLCHAN ' + row['BTS1_DL Channel'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 DLCHAN SETTING FAIL'])
                continue
                 
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('BAND ' + row['BTS2_Band'] + ',BTS2')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('BAND? BTS2').rstrip()
                    time.sleep(1)
                    if rslt == row['BTS2_Band']:
                        logger.info('SET BTS2 BAND ' + row['BTS2_Band'] + ': SUCESS')
                        break
                    else:
                        LTE1.write('BAND ' + row['BTS2_Band'] + ',BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 BAND ' + row['BTS2_Band'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 BAND SETTING FAIL'])
                continue
                    
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('DLCHAN ' + row['BTS2_DL Channel'] + ',BTS2')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('DLCHAN? BTS2').rstrip()
                    time.sleep(1)
                    if rslt == row['BTS2_DL Channel']:
                        logger.info('SET BTS2 DLCHAN ' + row['BTS2_DL Channel'] + ': SUCESS')
                        break
                    else:
                        LTE1.write('DLCHAN ' + row['BTS2_DL Channel'] + ',BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 DLCHAN ' + row['BTS2_DL Channel'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 DLCHAN SETTING FAIL'])
                continue
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OLVL '  + row['BTS2_Output Level'] + ',BTS2')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OLVL? BTS2').rstrip()
                    time.sleep(1)
                    if rslt == row['BTS2_Output Level']+'.0' :
                        logger.info('SET BTS2 OUTPUTLEVEL ' + row['BTS2_Output Level'] + ': SUCESS')
                        break
                    else:
                        LTE1.write('OLVL '  + row['BTS2_Output Level'] + ',BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 OUTPUTLEVEL ' + row['BTS2_Output Level'] + ': FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 OUTPUTLEVEL SETTING FAIL'])
                continue
            
            ssh.exec_command(cellular_on)
            print('\033[33mcellular_on\033[0m')
            time.sleep(3)
        
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OUTOFSERVICE IN,BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e
            
            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt =='IN':
                        logger.info('SET BTS1 IN OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE IN,BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 IN OF SERVICE : FAIL')
                logger.info('********************************')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','SET BTS1 IN OF SERVICE : FAIL'])
                continue
    
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt = LTE1.query('CALLStat? BTS1').rstrip()
                    if rslt == 'NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE':
                        logger.info('BTS1 CONNECTION SUCESS')
                        break
                except Exception as e:
                    err = e
            else:
                logger.info('OVER-TIME :BTS1 CONNECTION FAIL')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 CONNECTION FAIL'])
                logger.info('********************************')
                continue
            
            time.sleep(5)
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OUTOFSERVICE OUT,BTS1')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e

            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                    time.sleep(2)
                    if rslt =='OUT':
                        logger.info('SET BTS1 OUT OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                logger.info('********************************')
                continue
            
            start = time.time()
            while time.time() < start + timeout:
                try:
                    LTE1.write('OUTOFSERVICE IN ,BTS2')
                    time.sleep(1)
                    break
                except Exception as e:
                    err = e

            err = "SETTING FAIL"
            start = time.time()
            while time.time() < start + timeout:
                try:
                    rslt=LTE1.query('OUTOFSERVICE? BTS2').rstrip()
                    time.sleep(2)
                    if rslt =='IN':
                        logger.info('SET BTS2 IN OF SERVICE : SUCESS')
                        break
                    else:
                        LTE1.write('OUTOFSERVICE IN,BTS2')
                        time.sleep(5)
                        continue
                except Exception as e:
                    err = e
            else:
                logger.info("ERROR:",str(err))
                logger.info('SET BTS2 IN OF SERVICE : FAIL')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 IN OF SERVICE : FAIL'])
                logger.info('********************************')
                continue
                
            while time.time() < start + timeout:
                try:
                    rslt = LTE1.query('CALLStat? BTS2').rstrip()
                    if rslt=='NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE':
                        break
                except Exception as e:
                    err = e
                
            end = time.time()
            if(end-start)<=timeout:
                logger.info('BTS2 CONNETION TIME :' + str(end-start) + ' PASS')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],str(end-start),'PASS'])
            else:
                logger.info('OVER-TIME :BTS2 CONNECTION FAIL')
                writer.writerow(['WCDMA',row["BTS1_Band"],row['BTS2_Band'],' X ','BTS2 CONNECTION FAIL'])
                logger.info('********************************')
                continue
            logger.info('********************************')    
            
    start = time.time()
    while time.time() < start + timeout:
        try:
            LTE1.write('stop')
            time.sleep(3)
            break
        except Exception as e:
            err = e
    
    logger.info('STOP SIMULATION,PLEASE WAIT')
    start = time.time()
    while time.time() < start + timeout:
        try:
            LTE1.query('stat?')
            break
        except Exception as e:
            err = e
    logger.info('==========STOP==========')
        
