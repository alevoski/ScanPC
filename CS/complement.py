#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - main.py
#@Alexandre Buissé - 2019

#Standard imports
import os
import time

# Project modules imports
import writer
import ask_dismount
import av_date

username = os.getlogin()
computername = os.environ['COMPUTERNAME']

def genEicar(log):
    '''
    **FR**
    Génére un fichier Eicar pour tester l'antivirus
    **EN**
    Generate an Eicar file to test antivirus
    '''
    #Création du virus
    logFilePath = os.path.dirname(log)
    # print(logFilePath)
    # sys.exit()
    msg0 = "\n--Antivirus testing--\nAn Eicar virus test will be generated to test the antivirus."
    writer.write(msg0)
    eicarFile = str(logFilePath)+"/eicar_testFile"
    element = 'X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*'
    writer.writeLog(eicarFile, element)
    time.sleep(3)
    #Demande à l'utilisateur
    msg0 = "\nDid the antivirus alert you ? (y = yes, n = no)\n"
    writer.write(msg0)
    avEtat = "\n***** Antivirus status *****\n"
    writer.writeLog(log, avEtat)
    userRep = "FAILED - The antivirus does not alert/detect viruses.\n"
    if ask_dismount.reponse() == 'y':
        userRep = "OK - The antivirus does alert/detect viruses.\n"
    writer.writeLog(log, userRep)

def elemInLog(mode, thelist, elem):
    '''
    **FR**
    Trouve un élément dans une liste, 
    **EN**
    Find an element in a list
    '''
    if mode == 1: #Applications
        resultKO = 'Not installed\n'
        resultOK = 'Installed\n'
    elif mode == 2: #Services
        resultKO = 'Not started\n'
        resultOK = 'Started\n'
    else:
        return '?'
    result = resultKO

    if elem in thelist:
        result = resultOK
        # print(elem, " found !")
    return result
    
def mcAfee(compleLog):
    '''
    **FR**
    Récupere une information complémentaire sur McAfee
    **EN**
    Get more informations about McAfee
    '''
    restemp = "N/A"
    dateVer = "Date & version : " + restemp
    epoLst = "\n--EPO servers list--\n" + restemp
    try:
        dateVer, epoLst = av_date.getMcAfee()
    except Exception:
        # print(Exception)
        pass
    infoMcAfee = "***** McAfee informations *****\n"
    writer.writeLog(compleLog, infoMcAfee)
    writer.writeLog(compleLog, dateVer)
    writer.writeLog(compleLog, epoLst + '\n')
    
def wsus(log, servicesList):
    '''
    **FR**
    Récupere une information complémentaire sur WSUS/BranchCache
    **EN**
    Get more informations about WSUS/BranchCache
    ''' 
    infoBC = "\n***** WSUS/BranchCache service state *****\n"
    writer.writeLog(log, infoBC)
    res = elemInLog(2, servicesList, 'PeerDistSvc')
    writer.writeLog(log, res)

    restemp = "N/A\n"
    srvWSUS = "Server : " + restemp
    try:
        srvWSUS = av_date.getWsus()
    except Exception:
        # print(Exception)
        pass
    writer.writeLog(log, srvWSUS)
    
def init(log, softwareDict, servicesList):
    '''
    **FR**
    Initialisation de la recherche d'infos complémentaires
    **EN**
    Init search
    '''
    # 1 - Ecriture début de log (à la fin du log fourni)  
    elem = "************* Other tests *************"
    writer.prepaLogScan(log, elem)

    # 2 - Obtention des informations complémentaires
    #McAfee
    mcAfee(log)

    #LAPS
    infoLaps = "\n***** LAPS state *****\n"
    writer.writeLog(log, infoLaps)
    res = elemInLog(1, list(softwareDict.keys()), 'Local Admin Password Solution')
    writer.writeLog(log, res)

    ###Services

    ##BranchCache/WSUS
    wsus(log, servicesList)

    ##AppLocker
    infoAL = "\n***** AppLocker service state *****\n"
    writer.writeLog(log, infoAL)
    res = elemInLog(2, servicesList, 'AppIDSvc')
    writer.writeLog(log, res)

    #TestAV
    msg0 = "\nDo you want to perform an antivirus detection test ? (y = yes, n = no)\n"
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
        genEicar(log)