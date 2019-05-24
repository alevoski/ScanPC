#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - main.py
#@Alexandre Buissé - 2019

#Standard imports
import os
import time
import webbrowser
from datetime import datetime

#Project modules imports
import writer
import verif
import ask_dismount
import scans
import software_scan
import complement

computername = os.environ['COMPUTERNAME']
styleCssFile = os.getcwd() + '/style.css'

def scanPart():
    ''' 
    **FR**
    Retourne la lettre du lecteur où est exécuté l'app
    **EN**
    Return the drive letter where the app is launched
    ''' 
    key = os.getcwd() #On prend le chemin du répertoire du script/exe
    key = str(key[:2])
    # print("key => ", str(key)) #test ok
    return key

def scanPc(key, logFilePath):
    ''' 
    **FR**
    Exécute le scan de l'ordinateur en appelant diverses fonctions,
    écrit les logs sur le périphérique qui a lancé l'exécutable,
    Retourne le chemin du log final
    **EN**
    Process the computer scan with several other functions,
    write logs on the device where the app has been launched,
    Return the final log path 
    '''
    # print("logFilePath : "+str(logFilePath))

    texte3 = "Computer scanning in progress... Please wait...\n"
    writer.write(texte3)
    
    logFile = str(logFilePath) + "FINAL.html"
    
    #write first html tags
    elementRaw = """<!DOCTYPE html>
    <html>
        <head>
            <title>ScanPC result</title>
            <meta charset="utf-8">
            <link rel="stylesheet" type="text/css" href="style.css">
        </head>
        <body>
            <p>
            {0}
            {1} <br>
            </p>"""
            # </body>
        # </html>"""
    
    #texte d'intro 
    texte02 = '<h1>Scanning of computer "' + computername + '"</h1>'
    element2 = '<time>' + datetime.now().strftime("%A %d %B %Y %H:%M:%S") + '</time>\n'
    element = elementRaw.format(texte02, element2)
    writer.writeLog(logFile, element)

    #Scans de bases
    logAllUsers = scans.userInfo(logFilePath)
    logSF = scans.sharedFolders(logFilePath)
    hotfixDict = scans.hotFixesInfo(logFilePath)
    logfilefull = scans.systemInfo(logFilePath)
    securityProductDict = scans.securityProductInfo(logFilePath)
    procDict = scans.processInfo(logFilePath)
    servicesDictRunning = scans.servicesInfo(logFilePath)
    portsDict = scans.portsInfo(logFilePath)

    texte4 = "\nBasic scan ended.\n"
    writer.write(texte4)
    
    #Scan des logiciels
    softwareDict = software_scan.softwareInit(logFilePath)

    return logFile, softwareDict, servicesDictRunning

def readandcopy(concatenateBasesFiles1):
    '''
    **FR**
    Proposer à l'utilisateur d'ouvrir le log passé en paramètre et de le copier sur l'ordinateur scanné
    **EN**
    Ask user if he wants to read the log and if he wants a copy of the log on the scanned computer
    '''
    msg0 = "\nDo you want to read the scan report ? (y = yes, n = no)\n"
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
        openers = webbrowser._tryorder
        opened = 0
        for browsers in openers:
            # print(browsers)
            if 'firefox' in browsers.lower() or 'iexplore' in browsers.lower():
                # print('Opening with ' + browsers)
                browser = webbrowser.get(webbrowser._tryorder[openers.index(browsers)])
                browser.open('file:' + str(concatenateBasesFiles1))
                opened = 1
                break
        if opened ==0:
            webbrowser.open(str(concatenateBasesFiles1))

    # Copie de fichier vers un emplacement défini
    msg0 = "\nDo you want a copy the scan report on the computer (C:\ drive) ? (y = yes, n = no)\n"
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
        #Copy log
        fileName = os.path.basename(concatenateBasesFiles1)
        fileToWrite = "C:/" + fileName
        writer.copyFile(concatenateBasesFiles1, fileToWrite)
        #Copy CSS file
        if os.path.isfile(styleCssFile):
            fileName = os.path.basename(styleCssFile)
            fileToWrite = "C:/" + fileName
            writer.copyFile(styleCssFile, fileToWrite)

def fin(key):
    '''
    **FR**
    Si le périphérique est amovible, proposer à l'utilisateur de le retirer proprement
    **EN**
    If the device is removable, ask the user to safely remove (dismount) his device
    '''
    #Test si lecteur amovible    
    texte2 = "Computer scanning ended, details have been saved on " + str(key) + ".\n"
    writer.write(texte2)
    if verif.amovible(key):
        #Proposer la déconnexion de la clé
        msg0 = "\nDo you want to dismount your device and quit ? (y = yes, n = no)\n"
        writer.write(msg0)
        if ask_dismount.reponse() == 'y':
            ask_dismount.dismount(key)
            sys.exit(0)
    ask_dismount.quitter()

while True:
    chose = 0
    print('\t\t\t---------- ScanPC ---------')
    print('\t***************************************************************')
    print('\t*\t\t    Computer scanning                         *')
    print('\t*\t\t          Disclaimer !                        *')
    print('\t*\t This sofware only scan your computer                 *')
    print('\t*\t for users, various system information and software.  *')
    print('\t*\t No copy of your data will be performed.              *')
    print('\t***************************************************************')
    print()
    while chose == 0:
        #Begin timer
        beginScan = time.time()
        
        #Début du programme
        key = scanPart()
        uniqueDir = str(computername) + '_' + datetime.now().strftime('%d%m%y%H%M%S')
        logFilePath = str(key) + "/logScanPC/" + datetime.now().strftime('%Y/%m/%d/') + uniqueDir + '/' + uniqueDir
        #Appelle de scanPC
        logFile, softwareDict, servicesDictRunning = scanPc(key, logFilePath)
        #Tests complémentaires
        complement.init(logFilePath, softwareDict, servicesDictRunning)
        writer.writeLog(logFile, '\n</div>\n')
        
        #Ending timer
        endScan = time.time()
        totalTimeScan = endScan - beginScan
        elem = "----------------------- Scan ended ------------------------"
        writer.prepaLogScan(logFile, elem)
        totalTime = 'Computer analyzed in {} seconds.'.format(round(totalTimeScan, 2))
        writer.writeLog(logFile, totalTime)
        writer.write('Computer analyzed in ' + str(round(totalTimeScan, 2)) + ' seconds.\n')
        
        #write last html tags
        elementRaw = """
            </body>
        </html>"""
        writer.writeLog(logFile, elementRaw)
        
        #copy css file in the log directory
        if os.path.isfile(styleCssFile):
            fileName = os.path.basename(styleCssFile)
            fileToWrite = str(key) + "/logScanPC/" + datetime.now().strftime('%Y/%m/%d/') + uniqueDir + '/'  + fileName
            writer.copyFile(styleCssFile, fileToWrite)

        #Fin du programme
        readandcopy(logFile)
        fin(key)