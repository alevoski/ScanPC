#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - software_scan.py
#@Alexandre Buissé - 2019

#Standard imports
import winreg
import os

#Project modules imports
import writer

hive = winreg.HKEY_LOCAL_MACHINE
computername = os.environ['COMPUTERNAME']

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
    header = ['Name', 'Version', 'Publisher', 'Location']
    csvFile = logFilePath + "software.csv"
    writer.writeCSV(csvFile, header, finalDict)
    
    # 5 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Installed software')

    # 5 - Ecriture fin du log
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
            # print(soft)
            softkey = thekey+'\\'+soft
            # print(softkey) # test ok
            # try:
            nameSoft = reg(hive, softkey, name)
            versionSoft = reg(hive, softkey, version)
            publisherSoft = reg(hive, softkey, publisher)
            pathSoft = reg(hive, softkey, path)
            # print(nameSoft)
            # Conditions : some software does not have version & avoid Microsoft Updates
            if nameSoft != None and nameSoft != '' and 'Update for' not in nameSoft and publisherSoft != None:
                # print(nameSoft.title())
                # print(versionSoft)
                # print(publisherSoft)
                softwareDict[nameSoft.title()] = {'Name':nameSoft.title(),
                                                'Version':versionSoft,
                                                'Publisher':publisherSoft,
                                                'Location':pathSoft}
        # print(softwareDict)
    except TypeError:
        pass
    return softwareDict
  
if __name__ == '__main__':
    softwareInit(r"C:\STOCKAGE\logScanPC\2019\05\6/")