#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - software_scan.py
#@Alexandre Buissé - 2019

#Standard imports
import winreg
import os
import json
import re

#Project modules imports
import writer

hive = winreg.HKEY_LOCAL_MACHINE
computername = os.environ['COMPUTERNAME']
otherSoftLst = ['Java Auto Updater']

def reg(HIVE, thekey, value):
    '''
    **FR**
    Interroge le registre Windows en fonction de la ruche, la clé et du nom de la valeur à rechercher
    dans la clé passés en paramètre (32 et 64 bits)
    Retourne la valeur de la clé
    **EN**
    Ask regedit with the hive, the key and the value to search (32 and 64 bits)
    Return the value of the key
    '''
    registry_key = None
    try:
        registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ)
    except FileNotFoundError:
        try:
            registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
        except FileNotFoundError:
            try:
                registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
            except FileNotFoundError:
                pass
    # print(registry_key)
    if registry_key != None:
        try:
            value, regtype = winreg.QueryValueEx(registry_key, value)
        except Exception:
            value = None
            pass
        winreg.CloseKey(registry_key)
    return value
    
def softwareInit(logFilePath):
    '''
    **FR**
    Initialisation et fonction principale de la recherche des logiciels
    **EN**
    Init the main function to search installed software
    '''
    texte5 = "Software scanning in progress...\n"
    writer.write(texte5)

    # 1 - Début de l'analyse des logiciels
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = "***************************Software informations of computer ''" + computername + "''***************************"
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des logiels de l'ordinateur
    regkey = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" #32 bits
    regkey2 = r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" #64 bits
    softwareList = softwareLst(regkey)
    softwareDict1 = searchSoftware(softwareList, regkey)
    softwareList2 = softwareLst(regkey2)
    softwareDict2 = searchSoftware(softwareList2, regkey2)
    
    # 3 - Concaténation des deux dictionnaires de valeurs des clés de registre (32b et 64b)
    finalDict = softwareDict1
    for item, value in softwareDict2.items():
        # print(item, value)
        if item not in finalDict.items():
            # print(item)
            finalDict.update({item:value})

    # 4 - Ecriture du fichier CSV
    header = ['Name', 'Version', 'Last Update', 'Last Version', 'Up to date', 'Publisher', 'Location']
    csvFile = logFilePath + "software.csv"
    writer.writeCSV(csvFile, header, finalDict)
    
    # 5 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Installed software')

    # 6 - Placer des id sur certains tags du tableau
    # htmltxt = writer.modHTML(htmltxt, '<TD>ko</TD>', '<TD id="ko">ko</TD>')
    htmltxt = htmltxt.replace('<TD>ko</TD>', '<TD id="ko">ko</TD>')
    # htmltxt = writer.modHTML(htmltxt, '<TD>ok</TD>', '<TD id="ok">ok</TD>')
    htmltxt = htmltxt.replace('<TD>ok</TD>', '<TD id="ok">ok</TD>')
    
    # 6 - Ecriture fin du log
    writer.writeLog(logFile, str(len(finalDict)) + ' software found :\n')
    writer.writeLog(logFile, htmltxt)
    elem = '-' * 25 + ' Software scanning ended ' + '-' * 25
    writer.prepaLogScan(logFile, elem)
    writer.writeLog(logFile, '\n</div>\n')
    texte5bis = "Software scanning ended.\n"
    writer.write(texte5bis)

    return finalDict

def softwareLst(thekey):
    '''
    **FR**
    Utiliser le registre de Windows pour obtenir les logiciels installés
    Retourne la liste des logiciels installés
    **EN**
    Use regedit to get all installed software
    Return a list of them
    '''
    registry_key = None
    software = None
    try:
        # registry_key = winreg.OpenKey(hive, thekey, 0, winreg.KEY_READ)
        registry_key = winreg.OpenKey(hive, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
    except FileNotFoundError:
        try:
            registry_key = winreg.OpenKey(hive, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        except FileNotFoundError:
            try:
                registry_key = winreg.OpenKey(hive, thekey, 0, winreg.KEY_READ)
            except FileNotFoundError:
                pass
    # print(registry_key)
    if registry_key != None:
        i = 0
        software = []
        while True:
            try:
                software.append(winreg.EnumKey(registry_key, i)) #ok pour 1 clé
                i+=1
                # print("value :", str(registry_key))
            except Exception:
                # print(Exception)
                winreg.CloseKey(registry_key)
                break
        # print(software)
    return software

def isUptodate(nameSoft, versionSoft):
    '''
    Compare version of the installed software with the last version known in software_list.json
    Return 'ok' if the software is up to date or 'ko' if not
    Return also the last known update and the last known version
    '''
    #Voir si les logiciels sont à jour
    #open json software list
    softwareList = os.getcwd() + '/software_list.json'
    with open(softwareList, mode='r', encoding='utf-8') as f:
         data = json.load(f)
    #compare software version in the json list with the dict
    for soft in data['software']:
        #match software
        if soft['name'].lower() in nameSoft[:len(soft['name'])].lower():
            lastUpdate = soft['last_update']
            lastVersion = soft['last_version']
            #match version
            if soft['last_version'] in versionSoft:
                stat = 'ok'
                break
            else:
                stat = 'ko'
                break
        else:
            stat = ''
            lastUpdate = ''
            lastVersion = ''
    return stat, lastUpdate, lastVersion
    
def searchVersionInName(nameSoft):
    '''
    Search and return a version number from a software name
    '''
    try:
        # print(str(re.findall("(\d{0,2}(\.\d{1,2}).*)", nameSoft)[0][0]))
        return str(re.findall("(\d{0,2}(\.\d{1,2}).*)", nameSoft)[0][0])
    except IndexError:
        return ''
    
def searchSoftware(software, thekey):
    '''
    **FR**
    Utiliser la liste des logiciels installés pour interroger le registre de Windows
    Retourne le dictionnaire des logiciels installés contenant la version, l'éditeur
    et le chemin d'installation
    **EN**
    Use the list of installed software to ask regedit
    Return the dict of installed software with version, publisher and install location
    '''
    name = "DisplayName"
    version = "DisplayVersion"
    publisher = "Publisher"
    path = "InstallLocation"
    softwareDict = {}
    try:
        for soft in software:
            softkey = thekey+'\\'+soft
            nameSoft = reg(hive, softkey, name)
            versionSoft = reg(hive, softkey, version)
            publisherSoft = reg(hive, softkey, publisher)
            pathSoft = reg(hive, softkey, path)
            # Conditions : some software does not have version & avoid Microsoft Updates
            if nameSoft != None and nameSoft != '' and 'Update for' not in nameSoft and publisherSoft != None:
                # Is version empty then look for it in name
                if versionSoft == None:
                    versionSoft = searchVersionInName(nameSoft)
                # Is software up to date ?
                stat = ''
                lastUpdate = ''
                lastVersion = ''
                if nameSoft not in otherSoftLst: #do nothing if software is in otherSoftLst
                    stat, lastUpdate, lastVersion = isUptodate(nameSoft, versionSoft)
                # Write in dict
                softwareDict[nameSoft.title()] = {'Name':nameSoft.title(),
                                            'Version':versionSoft,
                                            'Last Update':lastUpdate,
                                            'Last Version':lastVersion,
                                            'Up to date':stat,
                                            'Publisher':publisherSoft,
                                            'Location':pathSoft}
    except TypeError:
        pass
    return softwareDict
  
if __name__ == '__main__':
    softwareInit(r"C:\STOCKAGE\logScanPC\2019\05\7/")