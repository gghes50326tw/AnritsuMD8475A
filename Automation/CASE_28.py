# -*- coding: utf-8 -*-
"""
Created on Fri Jan  7 16:37:49 2022

@author: Vivian
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Dec 29 11:23:36 2021

@author: Vivian
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 13:52:29 2021

@author: Vivian
"""

from colorama import init
import pyvisa as visa
import logging
import time
import csv
import configparser

import os
import subprocess
import shutil
if os.path.isdir('C:\ScreenCapture'):
    shutil.rmtree('C:\ScreenCapture')
path = os.getcwd()
rslt_path= os.path.join(path, os.pardir,'output','CASE_28','final_result')
if os.path.isdir(rslt_path):
    shutil.rmtree(rslt_path)
rsltdontroam_path=os.path.join(path, os.pardir,'output','CASE_28','final_result','DontRoam')
rsltroam_path=os.path.join(path, os.pardir,'output','CASE_28','final_result','Roam')
if os.path.isdir(rslt_path):
    pass
else:
    os.makedirs(rslt_path)

if os.path.isdir(rslt_path):
    shutil.rmtree(rslt_path)
if os.path.exists(rsltroam_path):
    shutil.rmtree(rsltroam_path)
os.makedirs(rsltdontroam_path)
os.makedirs(rsltroam_path)

rsltcsv_path = os.path.join(path, os.pardir,'output','CASE_28','final_result','CASE_28.csv')
rsltlog_path = os.path.join(path, os.pardir,'output','CASE_28','final_result','CASE_28.log')
sorcelte_path = os.path.join(path, 'CASE28LTE.csv')
sorcewcdma_path = os.path.join(path, 'Comb_Single_W.csv')
config_path = os.path.join(path, os.pardir,'config','CASE_27_28.ini')
screenshotd_path = os.path.join(path, 'screenshot.ps1')
screenshot_path = os.path.join(path, 'screenshot1.ps1')
screenshotd3_path = os.path.join(path, 'screenshot3g.ps1')
screenshot3_path = os.path.join(path, 'screenshot13g.ps1')
init(autoreset=True)
#讀取config檔
config = configparser.ConfigParser()
config.read(config_path)
timeout=int(config.get('timeout','timeout'))
BTS=config._sections
cellular_off = 'netsh mbn set powerstate interface=\"Cellular\" state=off'
cellular_on = 'netsh mbn set powerstate interface=\"Cellular\" state=on'
dont_roam = 'netsh mbn set dataroamcontrol interface="Cellular" profileset=all state=none'
roam = 'netsh mbn set dataroamcontrol interface="Cellular" profileset=all state=all'
check_roam = 'netsh mbn show interfaces'
argsd=[r'powershell.exe',screenshotd_path]
argsd3=[r'powershell.exe',screenshotd3_path]
args=[r'powershell.exe',screenshot_path]
args3=[r'powershell.exe',screenshot3_path]
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%m-%d %H:%M',
                    filename= rsltlog_path,filemode="w")
logger =logging.getLogger("mylogger")
logger.setLevel(logging.INFO)

if not logger.handlers:
    
    fh = logging.FileHandler("CASE_27.log" , encoding="utf-8")
    ch = logging.StreamHandler()
    
    formatter = logging.Formatter(fmt="%(message)s",
                                  datefmt="%Y/%m/%d %X")
    
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    
    logger.addHandler(fh)
    logger.addHandler(ch)

logger.info('CASE_28')
logger.info('Ver : 2.0.0.2')  

#GPIB地址
rm = visa.ResourceManager()
gpib_ad1=config['GPIB']['GPIB']
LTE1=rm.open_resource(gpib_ad1)
start = time.time()
while time.time() < start + timeout:
    try:
        rslt=LTE1.query('*IDN?').rstrip()
        logger.info('GPIB CONNECTION SUECSS')
        break
    except Exception as e:
        err = e

start = time.time()
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

logger.info('RESTART SIMULATION,PLEASE WAIT')
start = time.time()
while time.time() < start + timeout:
    try:
        rslt=LTE1.query('stat?').rstrip()
        if rslt == "NOTRUN":
            break   
    except Exception as e:
        err = e

logger.info('==============START==============')

with open(rsltcsv_path,'w',newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['UMTS','BAND','RESULT','REMARK'])
    with open(sorcelte_path, 'r',encoding='UTF-8-sig')as csv_file:
        rows = csv.DictReader(csv_file,delimiter=",")
        
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.write('LOADSIMPARAM "C:\LTE.wnssp2"' )
                logger.info('LOADSIMPARAM,PLEASE WAIT')
                break
            except Exception as e:
                err = e
                
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.query('stat?')
                logger.info('LOAD LTE.wnssp2 : SUCESS')
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
  
        start = time.time()
        while time.time() < start + timeout:
            try:
                LTE1.query('stat?')
                break
            except Exception as e:
                err = e
        logger.info('********************************')
        for row in rows:
            if config['B1'][row['BTS1_Band']] == config['YN']['SET']:
                
                subprocess.Popen(cellular_off, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                print('\033[33mcellular_off\033[0m')
                time.sleep(5)
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                start = time.time()
                err="SETTING FAIL"
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] +',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
              
                err = "SETTIMG FAIL"                        
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('DUPLEXMODE? BTS1').rstrip()
                        time.sleep(2)
                        if rslt == row['BTS1_DUPLEXMODE']:
                            logger.info('SET DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] +',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET DUPLEXMODE ' + row['BTS1_DUPLEXMODE'] +  ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 DUPLEXMODE : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('BAND ' + row['BTS1_Band'] +',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
              
                err="SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('BAND? BTS1').rstrip()
                        if rslt == row['BTS1_Band']:
                            logger.info('SET BAND ' + row['BTS1_Band'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('BAND ' + row['BTS1_Band'] +',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET BAND ' + row['BTS1_Band'] +  ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 BAND : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('BANDWIDTH ' + row['BTS1_Bandwidth'] + ',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        rslt = LTE1.query('BANDWIDTH? BTS1').rstrip()
                        time.sleep(2)
                        if rslt ==row['BTS1_Bandwidth']:
                            logger.info('SET BANDWIDTH ' + row['BTS1_Bandwidth'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('BANDWIDTH ' + row['BTS1_Bandwidth'] + ',BTS1')
                            time.sleep(8)
                            err = "SETTING FAIL"
                            continue
                    except Exception as e:
                        err = e    
                else:
                    logger.info('ERROR:',err)
                    logger.info('SET BANDWIDTH ' + row['BTS1_Bandwidth'] +  ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 BANDWIDTH : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OLVL ' + row['BTS1_Output Level'] + ',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                
                err="SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('OLVL? BTS1').rstrip()
                        if rslt == row['BTS1_Output Level']+'.0':
                            logger.info('SET OLVL ' + row['BTS1_Output Level'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('OLVL ' + row['BTS1_Output Level'] + ',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET OLVL ' + row['BTS1_Output Level'] + ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 OUTPUTLEVEL : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('MCC ' + row['MCC'] +',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                
                err="SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MCC? BTS1').rstrip()
                        time.sleep(3)
                        if rslt == row['MCC']:
                            logger.info('SET MCC '+ row['MCC'] +' : SUCESS')
                            break
                        else:
                            LTE1.write('MCC ' + row['MCC'] +',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MCC '+ row['MCC'] +' : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 MCC SETTING FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        LTE1.write('MNC '+ row['MNC'] +',BTS1')
                        break
                    except Exception as e:
                        err = e
                
                err = "SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MNC? BTS1').rstrip()
                        if rslt == row['MNC'] + 'F':
                            logger.info('SET MNC '+ row['MNC'] +' : SUCESS')
                            break
                        else:
                           LTE1.write('MNC '+ row['MNC'] +',BTS1')
                           time.sleep(5)
                           continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MNC '+ row['MNC'] +' : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 MNC SETTING FAIL'])
                    continue
                
                subprocess.Popen(cellular_on, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                print('\033[33mcellular_on\033[0m')
                time.sleep(5)
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OUTOFSERVICE IN,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                err="SETTING FAIL"
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 IN OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 IN OF SERVICE : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('CALLStat? BTS1').rstrip()
                        if rslt == 'NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE':
                            break
                        else:
                            continue
                    except Exception as e:
                        err = e 
                else:
                    logger.info('OVER-TIME :BTS1 ROAMING CONNECTION FAIL')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 ROAMING CONNECTION FAIL '])
                    logger.info('********************************')
                    continue       
                
                subprocess.Popen(roam, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                logger.info('<ROAMING>')
                time.sleep(5)
                
                checkroam=subprocess.Popen(check_roam, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                result = checkroam.stdout.read()
                result_str = str(result,encoding = 'utf-8')
                
                if "Roaming                : Yes" in result_str :
                    logger.info('ROAMING INFORMATION : PASS')
                    time.sleep(3)
                    subprocess.Popen(args, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                    time.sleep(8)
                    writer.writerow(["LTE",row['BTS1_Band'],'PASS',' '])
                else:
                    subprocess.Popen(args, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                    time.sleep(8)
                    logger.info('ROAMING INFORMATION : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],'FAIL','ROAMING-NOT CONNECTED : FAIL'])
                    continue
                    
                subprocess.Popen(cellular_off, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                print('\033[33mcellular_off\033[0m')
                time.sleep(5)
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                err="SETTING FAIL"
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],'X',' X ','BTS1 OUT OF SERVICE : FAIL'])
                    continue
                
                logger.info('*PRESET BAND' + row['BTS1_Band'] +'*')
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('MCC 001 ,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                
                err = "SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MCC? BTS1').rstrip()
                        if rslt == "001":
                            logger.info('SET MCC 001 : SUCESS')
                            break
                        else:
                            LTE1.write('MCC 001 ,BTS1')
                            time.sleep(5)
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MCC 001 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','X','BTS1 MCC SETTING FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        LTE1.write('MNC 01 ,BTS1')
                        break
                    except Exception as e:
                        err = e
  
                err = "SETTING FAIL"
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MNC? BTS1').rstrip()
                        if rslt == '01F':
                            logger.info('SET MNC 01 : SUCESS')
                            break
                        else:
                            LTE1.write('MNC 01 ,BTS1')
                            time.sleep(5)
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MNC 01 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],'X', 'X ','BTS1 MNC SETTING FAIL'])
                    continue
                
                subprocess.Popen(cellular_on, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                print('\033[33mcellular_on\033[0m')
                time.sleep(5)
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OUTOFSERVICE IN,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                err="SETTING FAIL"        
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 IN OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 IN OF SERVICE : FAIL'])
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
                    writer.writerow(["LTE",row['BTS1_Band'],' X ','BTS1 CONNECTION FAIL '])
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
    
    with open(sorcewcdma_path, 'r',encoding='UTF-8-sig')as csv_file:
        rows = csv.DictReader(csv_file,delimiter=",")
        
        start = time.time()
        while time.time() < start + timeout:
            try:
                LTE1.write('LOADSIMPARAM "C:\WCDMA.wnssp2"' )
                time.sleep(3)
                break
            except Exception as e:
                err = e
                
        logger.info('LOADSIMPARAM,PLEASE WAIT')
        start = time.time()
        while time.time() < start + timeout: 
            try:
                LTE1.query('stat?')
                logger.info('LOAD WCDMA.wnssp2 :SUCESS')
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
            
            if config['WCDMA'][row['BTS1_Band']] == config['YN']['SET']: 
                subprocess.Popen(cellular_off, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','BTS1 OUT OF SERVICE : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('BAND ' + row['BTS1_Band'] +',BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
              
                start = time.time()
                err="setting FAIL"
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('BAND? BTS1').rstrip()
                        if rslt == row['BTS1_Band']:
                            logger.info('SET BAND ' + row['BTS1_Band'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('BAND ' + row['BTS1_Band'] +',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET BAND ' + row['BTS1_Band'] +  ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','BTS1 BAND : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write("MCC 310 ,BTS1")
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
                start = time.time()
                err = "SETTING FAIL"
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MCC? BTS1').rstrip()
                        if rslt == '310':
                            logger.info('SET MCC 310 : SUCESS')
                            break
                        else:
                            LTE1.write("MCC 310 ,BTS1")
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MCC 310 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','SET MCC : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write("MNC 08 ,BTS1")
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                start = time.time()
                err="SETTING FAIL"
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MNC? BTS1').rstrip()
                        if rslt == '08F':
                            logger.info('SET MNC 08 : SUCESS')
                            break
                        else:
                            LTE1.write("MNC 08 ,BTS1")
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET MNC 08 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','SET MNC SETTING FAIL'])
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
                err = "SETTING FAIL"
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('OLVL? BTS1').rstrip()
                        if rslt == row['BTS1_Output Level']+'.0':
                            logger.info('SET OLVL ' + row['BTS1_Output Level'] + ': SUCESS')
                            break
                        else:
                            LTE1.write('OLVL ' + row['BTS1_Output Level'] + ',BTS1')
                            time.sleep(5)
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET OLVL ' + row['BTS1_Output Level'] + ': FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','SET BANDWIDTH : FAIL'])
                    continue
                
                subprocess.Popen(cellular_on, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
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
                            err="SETTING FAIL"
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info('ERROR:',err)
                    logger.info('SET BTS1 IN OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','BTS1 IN OF SERVICE : FAIL'])
                    continue
                
                start = time.time()
                err = "CONNECTION FAIL"
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('CALLStat? BTS1').rstrip()
                        if rslt == 'NONE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE' or rslt == 'IDLE,COMMUNICATION,NONE,NONE,NONE,NONE,NONE' :
                            logger.info('BTS1 ROAMING CONNECTION SUCESS')
                            break
                        else:
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('OVER-TIME :BTS1 CONNECTION FAIL')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','BTS1 ROAMING CONNECTION FAIL '])
                    logger.info('********************************')
                    continue 
                
                subprocess.Popen(roam, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                logger.info('<ROAMING INFORMATION CONNECTED>')
                time.sleep(5)
                
                checkroam=subprocess.Popen(check_roam, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                result = checkroam.stdout.read()
                result_str = str(result,encoding = 'utf-8')
                if "Roaming                : Yes" in result_str :
                    logger.info('ROAMING INFORMATION : PASS')
                    time.sleep(5)
                    subprocess.Popen(args3, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                    time.sleep(8)
                    writer.writerow(["WCDMA",row['BTS1_Band'],'PASS',' '])
                else:
                    subprocess.Popen(args3, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                    time.sleep(8)
                    logger.info('ROAMING-NOT CONNECTED : FAIL')
                    writer.writerow(["WCDMA",row['BTS1_Band'],'FAIL','ROAMING INFORMATION : FAIL'])
                    continue
                    
                subprocess.Popen(cellular_off, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
                print('\033[33mcellular_off\033[0m')
                logger.info('*PRESET BAND' + row['BTS1_Band'] +'*')
                time.sleep(5)
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('OUTOFSERVICE OUT,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                        
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
                            err="SETTING FAIL"
                            continue
                    except Exception as e:
                        err = e
                else:
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 OUT OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],'X','BTS1 OUT OF SERVICE : FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout: 
                    try:
                        LTE1.write('MCC 001 ,BTS1')
                        time.sleep(1)
                        break
                    except Exception as e:
                        err = e
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MCC? BTS1').rstrip()
                        break
                    except Exception as e:
                        err = e
                if rslt == "001":
                    logger.info('SET MCC 001 : SUCESS')
                else:
                    logger.info('SET MCC 001 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],' X ','BTS1 MCC SETTING FAIL'])
                    continue
                
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        LTE1.write('MNC 01 ,BTS1')
                        break
                    except Exception as e:
                        err = e
  
                start = time.time()
                while time.time() < start + timeout:
                    try:
                        rslt = LTE1.query('MNC? BTS1').rstrip()
                        break
                    except Exception as e:
                        err = e
                if rslt == '01F':
                    logger.info('SET MNC 01 : SUCESS')
                else:
                    logger.info('SET MNC 01 : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],'X','BTS1 MNC SETTING FAIL'])
                    continue
                
                subprocess.Popen(cellular_on, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
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
                        
                start = time.time()
                err="SETTING FAIL"
                while time.time() < start + timeout: 
                    try:
                        rslt=LTE1.query('OUTOFSERVICE? BTS1').rstrip()
                        time.sleep(1)
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
                    logger.info("ERROR:",err)
                    logger.info('SET BTS1 IN OF SERVICE : FAIL')
                    logger.info('********************************')
                    writer.writerow(["WCDMA",row['BTS1_Band'],'X','BTS1 IN OF SERVICE : FAIL'])
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
                    writer.writerow(["WCDMA",row['BTS1_Band'],'X','BTS1 CONNECTION FAIL '])
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

screen = os.walk('C:\ScreenCapture')

for root, dirs, files  in screen:
    for i in files:
        fileA = os.path.join(root, str(i))
        print(fileA+'\n')
        shutil.copyfile(fileA, os.path.join(rslt_path,str(i)))
         
logger.info('===============END===============')    
