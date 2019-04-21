#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - main.py
#@Alexandre Buissé - 2019

#Standard imports
import os
import time
import shutil
from datetime import datetime

#Project modules imports
import writer
import verif
import ask_dismount
import scans
import software_scan
import complement

computername = os.environ['COMPUTERNAME']

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
    
    logFile = str(logFilePath) + "FINAL.txt"
    
    #texte d'intro 
    texte01 = '*'*114 + '\n'
    texte02 = '-' * 26 + " Scanning of computer ''" + computername + '" ' + '-' *26
    element2 = '\n' + '-' * 26 + datetime.now().strftime("%A %d %B %Y %H:%M:%S") + '-' *26 + '\n'
    # logIntro = str(logFilePath)+"intro.txt"
    element0 = texte01 + texte02 + element2 + texte01
    writer.writeLog(logFile, element0)
    
    #Scans de bases
    logAllUsers = scans.userInfo(logFile)
    logSF = scans.sharedFolders(logFile)
    logHotFix = scans.hotFixesInfo(logFile)
    logfilefull = scans.systemInfo(logFile)
    logProc = scans.processInfo(logFile)
    servicesList = scans.servicesInfo(logFile)
    logPorts = scans.portsInfo(logFile)

    texte4 = "\nBasic scan ended.\n"
    writer.write(texte4)
    
    #Scan des logiciels
    softwareDict = software_scan.softwareInit(logFile)

    return logFile, softwareDict, servicesList

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
        os.system("notepad " + str(concatenateBasesFiles1))
        
    # Copie de fichier vers un emplacement défini
    msg0 = "\nDo you want a copy the scan report on the computer (C:\ drive) ? (y = yes, n = no)\n"
    writer.write(msg0)
    if ask_dismount.reponse() == 'y':
        fileName = os.path.basename(concatenateBasesFiles1)
        fileToWrite = "C:/"+fileName
        try:
            shutil.copy2(concatenateBasesFiles1, fileToWrite)
            if os.path.isfile(fileToWrite):
                msg = "\n"+fileName+" has been copied on C:\\ !\n"
                writer.write(msg)
            else:
                raise IOError
        except (IOError, PermissionError) as e:
            msg = "\nUnable to copy " + fileName + " on C:\\ ! :\nPermission denied !"
            writer.write(msg)

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
        logFilePath = str(key) + "/logScanPC/" + datetime.now().strftime('%Y/%m/%d/%d%m%y%H%M%S') + "Log" + str(computername)
        #Appelle de scanPC
        logFile, softwareDict, servicesList = scanPc(key, logFilePath)
        #Tests complémentaires
        complement.init(logFile, softwareDict, servicesList)
        
        #Ending timer
        endScan = time.time()
        totalTimeScan = endScan - beginScan
        elem = "----------------------- Scan ended ------------------------"
        writer.prepaLogScan(logFile, elem)
        totalTime = 'Computer analyzed in {} seconds.'.format(round(totalTimeScan, 2))
        writer.writeLog(logFile, totalTime)
        writer.write('Computer analyzed in ' + str(round(totalTimeScan, 2)) + ' seconds.\n')

        #Fin du programme
        readandcopy(logFile)
        fin(key)