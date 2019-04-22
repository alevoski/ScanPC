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
    
def softwareInit(logFile):
    '''
    **FR**
    Initialisation et fonction principale de la recherche des logiciels
    **EN**
    Init the main function to search installed software
    '''
    texte5 = "Software scanning in progress...\n"
    writer.write(texte5)

    # 1 - Début de l'analyse des logiciels
    # logSoftBegin = logFilePath + "begin_Soft.txt"
    elem = "***************************Software informations of computer ''" + computername + "''***************************"
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des logiels de l'ordinateur
    # logfilesoft = str(logFilePath)+"TEMP_SOFTWARE.txt"
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
    # print(finalDict)
    
    # 4 - Vérification si première lettre des logiciels en majuscule
    for item, value in finalDict.items():
        if item[0].islower() == True:
            newItem = item.title()
            del finalDict[item]
            finalDict[newItem] = value   
    
    # 5 - Tri des logiciels par ordre alphabétique + calcul pour création tableau
    maxSoftName = 0 #Longueur colonne nom logiciel
    maxLocationName = 0 #Longueur colonne path
    maxPublisherName = 0 #Longueur colonne éditeur
    maxVersionName = 0 #Longueur colonne version

    for item, value in sorted(finalDict.items()):       
        # Plus grand nom 
        if len(item) > maxSoftName:
            maxSoftName = len(item)
        for valuename, subvalue in value.items():
            #Plus grand path
            if valuename == 'installLocation' and subvalue != None:
                if len(subvalue) > maxLocationName: 
                    maxLocationName = len(subvalue)
            #Plus grand éditeur
            if valuename == 'publisher' and subvalue != None:
                if len(subvalue) > maxPublisherName: 
                    maxPublisherName = len(subvalue)
            #Plus grande version
            if valuename == 'version' and subvalue != None:
                if len(subvalue) > maxVersionName:
                    maxVersionName = len(subvalue)
   
    # 6 - Ecriture dans un fichier sous forme de tableau
    #Header du tableau
    diffLogiciel = maxSoftName - len("Logiciel")
    logicielHeader = "Logiciel" + " " * diffLogiciel + "|"
    
    diffPath = maxLocationName - len("Chemin")
    pathHeader = "Chemin" + " " * diffPath + "|"
   
    diffEditor = maxPublisherName - len("Editeur")
    editHeader = "Editeur" + " " * diffEditor + "|"
    
    diffVersion = maxVersionName - len("Version")
    versionHeader = "Version" + " " * diffVersion + "|"
    
    header = logicielHeader + versionHeader + editHeader + pathHeader
    limHeader = "_" * maxSoftName + "_" * maxLocationName + "_" * maxPublisherName + "_" * maxVersionName
    
    writer.writeLog(logFile, header + "\n" + limHeader +"\n")
    
    #Valeurs du tableau
    for item, value in sorted(finalDict.items()):
        diffLogiciel = maxSoftName - len(item)
        logicielValue= item + " " * diffLogiciel + "|"
        versionValue = " " * maxVersionName + "|"
        editorValue = " " * maxPublisherName + "|"
        pathValue = " " * maxLocationName + "|"
        for valuename, subvalue in value.items():
            if valuename == 'installLocation' and subvalue != None:
                diffPath = maxLocationName - len(subvalue)
                pathValue= subvalue + " " * diffPath + "|"
            if valuename == 'publisher' and subvalue != None:
                diffEditor = maxPublisherName - len(subvalue)
                editorValue= subvalue + " " * diffEditor + "|"
            if valuename == 'version' and subvalue != None:
                diffVersion = maxVersionName - len(subvalue)
                versionValue = subvalue + " " * diffVersion + "|"
        all = logicielValue + versionValue + editorValue + pathValue
        writer.writeLog(logFile, all + '\n')

    # 7 - Ecriture de la fin du log
    elem = '-' * 25 + ' Software scanning ended ' + '-' * 25
    writer.prepaLogScan(logFile, elem)
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
            if nameSoft != None and publisherSoft != None: #some software does not have version
                # print(nameSoft)
                # print(versionSoft)
                # print(publisherSoft)
                softwareDict[nameSoft] = {'version':versionSoft,'publisher':publisherSoft, 'installLocation':pathSoft}
        # print(softwareDict)
    except TypeError:
        pass
    return softwareDict
  
if __name__ == '__main__':
    softwareInit(r"<PATH>")