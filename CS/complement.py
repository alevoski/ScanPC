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
    avEtat = "<br>***** Antivirus status *****<br>"
    writer.writeLog(log, avEtat)
    userRep = "FAILED - The antivirus does not alert/detect viruses.<br>\n"
    if ask_dismount.reponse() == 'y':
        userRep = "OK - The antivirus does alert/detect viruses.<br>\n"
    writer.writeLog(log, userRep)

def elemInList(mode, thelist, elem):
    '''
    **FR**
    Trouve un élément dans une liste
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
    return elem + ' : ' + result
    
def mcAfee(compleLog):
    '''
    **FR**
    Récupere une information complémentaire sur McAfee
    **EN**
    Get more informations about McAfee
    '''
    restemp = "N/A"
    dateVer = "Date & version : " + restemp
    epoLst = "\n--EPO servers list--<br>" + restemp
    try:
        dateVer, epoLst = av_date.getMcAfee()
    except Exception:
        # print(Exception)
        pass
    infoMcAfee = "***** McAfee informations *****<br>"
    writer.writeLog(compleLog, infoMcAfee)
    writer.writeLog(compleLog, dateVer + '<br>')
    writer.writeLog(compleLog, epoLst + '<br>\n')
    
def wsus(log, servicesDictRunning):
    '''
    **FR**
    Récupere une information complémentaire sur WSUS/BranchCache
    **EN**
    Get more informations about WSUS/BranchCache
    ''' 
    infoBC = "<br>***** WSUS/BranchCache service state *****<br>"
    writer.writeLog(log, infoBC)
    res = elemInList(2, list(servicesDictRunning.keys()), 'PeerDistSvc')
    writer.writeLog(log, res  + '<br>')

    restemp = "N/A<br>\n"
    srvWSUS = "WSUS server : " + restemp
    try:
        srvWSUS = av_date.getWsus() + '<br>'
    except Exception:
        # print(Exception)
        pass
    writer.writeLog(log, srvWSUS)
    
def init(logFilePath, softwareDict, servicesDictRunning):
    '''
    **FR**
    Initialisation de la recherche d'infos complémentaires
    **EN**
    Init search
    '''
    # 1 - Ecriture début de log (à la fin du log fourni)
    log = logFilePath + "final.html"
    elem = "************* Other tests *************"
    writer.prepaLogScan(log, elem)
    writer.writeLog(log, '<div>\n')

    # 2 - Obtention des informations complémentaires
    #McAfee
    mcAfee(log)

    #LAPS
    infoLaps = "<br>***** LAPS state *****<br>"
    writer.writeLog(log, infoLaps)
    res = elemInList(1, list(softwareDict.keys()), 'Local Administrator Password Solution')
    writer.writeLog(log, res + '<br>')

    ###Services

    ##BranchCache/WSUS
    wsus(log, servicesDictRunning)

    ##AppLocker
    infoAL = "<br>***** AppLocker service state *****<br>"
    writer.writeLog(log, infoAL)
    res = elemInList(2, list(servicesDictRunning.keys()), 'AppIDSvc')
    writer.writeLog(log, res  + '<br>')

    #TestAV
    msg0 = "\nDo you want to perform an antivirus detection test ? (y = yes, n = no)\n"
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
        genEicar(log)