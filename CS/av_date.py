#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - av_date.py
#@Alexandre Buissé - 2019

import winreg

def reg(HIVE, key, value):
    '''
    **FR**
    Interroge le registre Windows en fonction de la ruche, la clé et du nom de la valeur à rechercher
    dans la clé passés en paramètre
    Retourne la valeur de la clé
    **EN**
    Ask regedit with the hive, the key and the value to search, return the value of the key
    '''
    registry_key = winreg.OpenKey(HIVE, key, 0, winreg.KEY_READ)
    value, regtype = winreg.QueryValueEx(registry_key, value)
    winreg.CloseKey(registry_key)
    return value
    
def getMcAfee():
    '''
    **FR**
    Interroge le registre Windows à propos des informations qu'il a sur McAfee
    **EN**
    Ask regedit for McAfee informations
    '''
    hive = winreg.HKEY_LOCAL_MACHINE
    
    key = r"SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Application Plugins\VIRUSCAN8800"
    version = "DATVersion"
    date = "DatDate"
    version = reg(hive,key,version).split('.')[0]
    date = reg(hive,key,date)
    date = (date[0:4],date[4:6],date[6:8])
    
    keySrvLst = r"SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Agent"
    epoLst = "ePOServerList"
    srvList = reg(hive, keySrvLst, epoLst)
    lstmsg = "`\n--Liste des serveurs EPO--\n"
    # print("{d[2]}/{d[1]}/{d[0]} ({v})".format(d=date,v=version))
    return "Date & version : {d[2]}/{d[1]}/{d[0]} ({v})".format(d=date,v=version), lstmsg + srvList.replace(';', '\n')

def getWsus():
    '''
    **FR**
    Interroge le registre Windows à propos des informations qu'il a sur WSUS
    **EN**
    Ask regedit for WSUS informations
    '''
    hive = winreg.HKEY_LOCAL_MACHINE
    
    keySrv = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    srv = "WUServer"
    
    srvWSUS = reg(hive, keySrv, srv)
    lstmsg = "Serveur : "
    # print(srvWSUS)
    return lstmsg + srvWSUS      

def initMcAfee(logFilePath):
    return str(logFilePath)+"McAfee.txt"

if __name__ == '__main__':
    date, epoLst = getMcAfee()
    print(date)
    print(epoLst)
    srvWSUS = getWsus()
    print(srvWSUS)