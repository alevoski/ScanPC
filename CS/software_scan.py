#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - software_scan.py
#@Alexandre Buissé - 2019/2020

'''
Scanning module : searches installed software in registry
'''

#Standard imports
import winreg
import os
import json
import re
from unicodedata import normalize

#Project modules imports
import writer

HIVE = winreg.HKEY_LOCAL_MACHINE
COMPUTERNAME = os.environ['COMPUTERNAME']
OTHER_SOFT_LIST = ['Java Auto Updater']

def reg(thekey, value):
    '''
    **FR**
    Interroge le registre Windows en fonction de la ruche, la clé et du nom de la valeur à rechercher
    dans la clé passés en paramètre (32 et 64 bits)
    Retourne la valeur de la clé
    **EN**
    Ask regedit with the HIVE, the key and the value to search (32 and 64 bits)
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
    if registry_key is not None:
        try:
            value, _ = winreg.QueryValueEx(registry_key, value)
        except Exception:
            value = None
        winreg.CloseKey(registry_key)
    return value

def software_init(log_file_path):
    '''
    **FR**
    Initialisation et fonction principale de la recherche des logiciels
    **EN**
    Init the main function to search installed software
    '''
    texte5 = "Software scanning in progress...\n"
    writer.writer(texte5)

    # 1 - Début de l'analyse des logiciels
    log_file = log_file_path + "final.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Software informations of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtenir la liste des logiels de l'ordinateur
    regkey = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" #32 bits
    regkey2 = r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" #64 bits
    software_list1 = software_lst(regkey)
    software_dict1 = search_software(software_list1, regkey)
    software_list2 = software_lst(regkey2)
    software_dict2 = search_software(software_list2, regkey2)

    # 3 - Concaténation des deux dictionnaires de valeurs des clés de registre (32b et 64b)
    final_dict = software_dict1
    for item, value in software_dict2.items():
        # print(item, value)
        if item not in final_dict.items():
            # print(item)
            final_dict.update({item:value})

    # 4 - Ecriture du fichier CSV
    header = ['Name', 'Version', 'Last Update', 'Last Version', 'Up to date', 'Publisher', 'Location']
    csv_file = log_file_path + "software.csv"
    writer.write_csv(csv_file, header, final_dict)

    # 5 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Installed software')

    # 6 - Placer des id sur certains tags du tableau
    htmltxt = htmltxt.replace('<TD>ko</TD>', '<TD id="ko">ko</TD>')
    htmltxt = htmltxt.replace('<TD>ok</TD>', '<TD id="ok">ok</TD>')

    # 7 - Bilan
    soft_ok = 0
    soft_ko = 0
    others = 0
    for _, value in final_dict.items():
        stat = value['Up to date']
        if stat == 'ok':
            soft_ok += 1
        elif stat == 'ko':
            soft_ko += 1
        else:
            others += 1

    soft_summary_dict = {}
    soft_summary_dict[str(len(final_dict))] = {'Total Software':str(len(final_dict)),
                                               'Up to date':str(soft_ok), '% up to date':str(round(soft_ok/len(final_dict)*100)),
                                               'Outdated':str(soft_ko), '% outdated':str(round(soft_ko/len(final_dict)*100)),
                                               'others':str(others), '% others':str(round(others/len(final_dict)*100))}

    header = ['Total Software', 'Up to date', '% up to date', 'Outdated', '% outdated', 'others', '% others']
    csv_file = log_file_path + "software_summary.csv"
    writer.write_csv(csv_file, header, soft_summary_dict)
    htmltxt2 = writer.csv2html(csv_file, 'Software summary')
    htmltxt2 = htmltxt2.replace('<TH>Up to date</TH>', '<TH id="ok">Up to date</TH>')
    htmltxt2 = htmltxt2.replace('<TH>% up to date</TH>', '<TH id="ok">% up to date</TH>')
    htmltxt2 = htmltxt2.replace('<TH>Outdated</TH>', '<TH id="ko">Outdated</TH>')
    htmltxt2 = htmltxt2.replace('<TH>% outdated</TH>', '<TH id="ko">% outdated</TH>')

    # 8 - Ecriture fin du log
    writer.writelog(log_file, str(len(final_dict)) + ' software found :\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '<br>Software summary :')
    writer.writelog(log_file, htmltxt2)
    writer.writelog(log_file, '\n</div>\n')
    texte5bis = "Software scanning ended.\n"
    writer.writer(texte5bis)

    return final_dict

def software_lst(thekey):
    '''
    **FR**
    Utiliser le registre de Windows pour obtenir les logiciels installés
    Retourne la liste des logiciels installés
    **EN**
    Use regedit to get all installed software
    Return a list of them
    '''
    registry_key = None
    software_list = None
    try:
        # registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ)
        registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
    except FileNotFoundError:
        try:
            registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ | winreg.KEY_WOW64_32KEY)
        except FileNotFoundError:
            try:
                registry_key = winreg.OpenKey(HIVE, thekey, 0, winreg.KEY_READ)
            except FileNotFoundError:
                pass
    # print(registry_key)
    if registry_key is not None:
        i = 0
        software_list = []
        while True:
            try:
                software_list.append(winreg.EnumKey(registry_key, i)) # ok pour 1 clé
                i += 1
                # print("value :", str(registry_key))
            except Exception:
                # print(Exception)
                winreg.CloseKey(registry_key)
                break
        # print(software_list)
    return software_list

def is_up_to_date(name_soft, version_soft):
    '''
    Compare version of the installed software with the last version known in software_list.json
    Return 'ok' if the software is up to date or 'ko' if not
    Return also the last known update and the last known version
    '''
    # Voir si les logiciels sont à jour
    # Open json software list
    software_list_file = os.getcwd() + '/software_list.json'
    with open(software_list_file, mode='r', encoding='utf-8') as softfile:
        data = json.load(softfile)
    # Compare software version in the json list with the dict
    for soft in data['software']:
        stat = ''
        last_update = ''
        last_version = ''
        # Match software
        if soft['name'].lower() in name_soft[:len(soft['name'])].lower():
            stat = 'ko'
            last_update = soft['last_update']
            last_version = soft['last_version']
            # Match version
            if soft['last_version'] in version_soft:
                stat = 'ok'
            break
    return stat, last_update, last_version

def search_version_in_name(name_soft):
    '''
    Search and return a version number from a software name
    '''
    try:
        return str(re.findall(r"(\d{0,2}(\.\d{1,2}).*?)", name_soft)[0][0])
    except IndexError:
        return ''

def search_software(software, thekey):
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
    software_dict = {}
    try:
        for soft in software:
            softkey = thekey+'\\'+soft
            name_soft = reg(softkey, name)
            version_soft = reg(softkey, version)
            publisher_soft = reg(softkey, publisher)
            path_soft = reg(softkey, path)
            # Conditions : some software does not have version & avoid Microsoft Updates
            if name_soft is not None and name_soft != '' and 'Update for' not in name_soft and publisher_soft is not None:
                if normalize('NFKD', name_soft.title()) not in software_dict:
                    # Is version empty then look for it in name
                    if version_soft is None:
                        version_soft = search_version_in_name(name_soft)
                    # Is software up to date ?
                    stat = ''
                    last_update = ''
                    last_version = ''
                    if name_soft not in OTHER_SOFT_LIST: # Do nothing if software is in OTHER_SOFT_LIST
                        stat, last_update, last_version = is_up_to_date(name_soft, version_soft)
                    # Write in dict
                    software_dict[normalize('NFKD', name_soft.title())] = {'Name':normalize('NFKD', name_soft.title()),
                                                                           'Version':version_soft,
                                                                           'Last Update':last_update,
                                                                           'Last Version':last_version,
                                                                           'Up to date':stat,
                                                                           'Publisher':publisher_soft,
                                                                           'Location':path_soft}
    except TypeError:
        pass
    return software_dict
