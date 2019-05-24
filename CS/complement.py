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
    # 1 - Création du virus
    logFilePath = os.path.dirname(log)
    msg0 = '\n--Antivirus testing--\nAn Eicar virus test will be generated to test the antivirus.'
    writer.write(msg0)
    eicarFile = str(logFilePath) + '/eicar_testFile'
    element = 'X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*'
    writer.writeLog(eicarFile, element)
    time.sleep(3)
    
    # 2 - Demande à l'utilisateur
    msg0 = '\nDid the antivirus alert you ? (y = yes, n = no)\n'
    writer.write(msg0)
    avEtat = 'Antivirus status'
    userRep = 'FAILED - The antivirus does not alert/detect viruses'
    if ask_dismount.reponse() == 'y':
        userRep = 'OK - The antivirus does alert/detect viruses'
    return avEtat, userRep

def elemInList(mode, thelist, elem):
    '''
    **FR**
    Trouve un élément dans une liste
    **EN**
    Find an element in a list
    '''
    if mode == 1: #Applications
        resultKO = 'Not installed'
        resultOK = 'Installed'
    elif mode == 2: #Services
        resultKO = 'Not started'
        resultOK = 'Started'
    else:
        return '?'
    result = resultKO

    if elem in thelist:
        result = resultOK

    return result
    
def mcAfee(compleLog):
    '''
    **FR**
    Récupere une information complémentaire sur McAfee
    **EN**
    Get more informations about McAfee
    '''
    restemp = 'N/A'
    dateVer = 'Date & version : ' + restemp
    epoLst = '\n--EPO servers list--<br>' + restemp
    try:
        dateVer, epoLst = av_date.getMcAfee()
    except Exception:
        pass
    infoMcAfee = '***** McAfee informations *****<br>'
    writer.writeLog(compleLog, infoMcAfee)
    writer.writeLog(compleLog, dateVer + '<br>')
    writer.writeLog(compleLog, epoLst + '<br>\n\n')

def init(logFilePath, softwareDict, servicesDictRunning):
    '''
    **FR**
    Initialisation de la recherche d'infos complémentaires
    **EN**
    Init search
    '''
    # 1 - Ecriture début de log (à la fin du log fourni)
    log = logFilePath + 'final.html'
    elem = '<h2>Additional scans</h2>'
    writer.prepaLogScan(log, elem)
    writer.writeLog(log, '<div>\n')

    # 2 - Obtention des informations complémentaires
    # McAfee
    mcAfee(log)

    complementDict = {} #Name Result
    #LAPS
    infoLaps = 'LAPS'
    toFind = 'Local Administrator Password Solution'
    res = elemInList(1, list(softwareDict.keys()), toFind)
    complementDict[infoLaps] = {'Name':infoLaps, 'Description':toFind, 'Result':res, 'Type':'Software test'}

    ### Services

    ## BranchCache/WSUS
    # BranchCache
    infoBC = 'BranchCache'
    toFind = 'PeerDistSvcs'
    res = elemInList(2, list(servicesDictRunning.keys()), toFind)
    complementDict[infoBC] = {'Name':infoBC, 'Description':toFind, 'Result':res, 
                                'Type':'Service test'}

    # WSUS
    infoWSUS = 'WSUS'
    toFind = 'WSUS server'
    restemp = 'N/A'
    srvWSUS = restemp
    try:
        srvWSUS = av_date.getWsus()
    except Exception:
        pass
    complementDict[infoWSUS] = {'Name':infoWSUS, 'Description':toFind, 'Result':srvWSUS, 
                                'Type':'Service test'}

    ## AppLocker
    infoAL = 'AppLocker'
    toFind = 'AppIDSvc'
    res = elemInList(2, list(servicesDictRunning.keys()), toFind)
    complementDict[infoAL] = {'Name':infoAL, 'Description':toFind, 'Result':res, 'Type':'Service test'}

    # TestAV
    msg0 = '\nDo you want to perform an antivirus detection test ? (y = yes, n = no)\n'
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
       avEtat, userRep = genEicar(log)
       toFind = 'Antivirus alert'
       complementDict[avEtat] = {'Name':avEtat, 'Description':toFind, 'Result':userRep, 'Type':'Service test'}

    # 3 - Ecrire du fichier CSV
    header = ['Type', 'Name', 'Description', 'Result']
    csvFile = logFilePath + 'additional_scans.csv'
    writer.writeCSV(csvFile, header, complementDict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Additional Scans')

    # 5 - Ecriture de la fin du log
    writer.writeLog(log, str(len(complementDict)) + ' additional scans :\n')
    writer.writeLog(log, htmltxt)
    writer.writeLog(log, '\n</div>\n')