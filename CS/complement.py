#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - complement.py
#@Alexandre Buissé - 2019/2020

'''
Scanning module : performs additional scans on the computer
'''

#Standard imports
import time
import winreg

# Project modules imports
import writer
import verif

def reg(hive, key, value):
    '''
    **FR**
    Interroge le registre Windows en fonction de la ruche, la clé et du nom de la valeur à rechercher
    dans la clé passés en paramètre
    Retourne la valeur de la clé
    **EN**
    Ask regedit with the hive, the key and the value to search, return the value of the key
    '''
    registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ)
    value, _ = winreg.QueryValueEx(registry_key, value)
    winreg.CloseKey(registry_key)
    return value

def get_mc_afee():
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
    version = reg(hive, key, version).split('.')[0]
    date = reg(hive, key, date)
    date = (date[0:4], date[4:6], date[6:8])

    key_srv = r"SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Agent"
    srvs = reg(hive, key_srv, 'ePOServerList')
    lstmsg = "`\n--EPO servers list--<br>"
    # print("{d[2]}/{d[1]}/{d[0]} ({v})".format(d=date,v=version))
    return "Date & version : {d[2]}/{d[1]}/{d[0]} ({v})".format(d=date, v=version), lstmsg + srvs.replace(';', '\n')

def get_wsus():
    '''
    **FR**
    Interroge le registre Windows à propos des informations qu'il a sur WSUS
    **EN**
    Ask regedit for WSUS informations
    '''
    hive = winreg.HKEY_LOCAL_MACHINE
    key_srv = r'SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    srv = 'WUServer'
    srv_wsus = reg(hive, key_srv, srv)
    return srv_wsus

def gen_eicar(log_file_path):
    '''
    **FR**
    Génére un fichier Eicar pour tester l'antivirus
    **EN**
    Generate an Eicar file to test antivirus
    '''
    # 1 - Création du virus
    msg0 = '\n--Antivirus testing--\nAn Eicar virus test will be generated to test the antivirus.'
    writer.writer(msg0)
    eicar_file = str(log_file_path) + 'eicar_testFile'
    element = r'X5O!P%@AP[4\PZX54(P^)7CC)7}$EICAR-STANDARD-ANTIVIRUS-TEST-FILE!$H+H*'
    writer.writelog(eicar_file, element)
    time.sleep(3)

    # 2 - Demande à l'utilisateur
    msg0 = '\nDid the antivirus alert you ? (y = yes, n = no)\n'
    writer.writer(msg0)
    user_rep = 'FAILED - The antivirus does not alert/detect viruses'
    if verif.reponse() == 'y':
        user_rep = 'OK - The antivirus does alert/detect viruses'
    return user_rep

def elem_in_list(mode, thelist, elem):
    '''
    **FR**
    Trouve un élément dans une liste
    **EN**
    Find an element in a list
    '''
    if mode == 1: # Applications
        result_ko = 'Not installed'
        result_ok = 'Installed'
    elif mode == 2: # Services
        result_ko = 'Not started'
        result_ok = 'Started'
    else:
        return '?'
    result = result_ko

    if elem in thelist:
        result = result_ok

    return result

def mc_afee(log_file):
    '''
    **FR**
    Récupere une information complémentaire sur McAfee
    **EN**
    Get more informations about McAfee
    '''
    restemp = 'N/A'
    date_ver = 'Date & version : ' + restemp
    epo_lst = '\n--EPO servers list--<br>' + restemp
    try:
        date_ver, epo_lst = get_mc_afee()
    except Exception:
        pass
    info_mc_afee = '***** McAfee informations *****<br>'
    writer.writelog(log_file, info_mc_afee)
    writer.writelog(log_file, date_ver + '<br>')
    writer.writelog(log_file, epo_lst + '<br>\n\n')

def complement_init(log_file_path, software_dict, services_running_dict):
    '''
    **FR**
    Initialisation de la recherche d'infos complémentaires
    **EN**
    Init search
    '''
    # 1 - Ecriture début de log (à la fin du log fourni)
    log_file = log_file_path + 'FINAL.html'
    elem = '<h2>Additional scans</h2>'
    writer.prepa_log_scan(log_file, elem)
    writer.writelog(log_file, '<div>\n')

    # 2 - Obtention des informations complémentaires
    # McAfee
    mc_afee(log_file)

    complement_dict = {} # Name Result
    #LAPS
    info_laps = 'LAPS'
    to_find = 'Local Administrator Password Solution'
    res = elem_in_list(1, list(software_dict.keys()), to_find)
    complement_dict[info_laps] = {'Name':info_laps,
                                  'Description':to_find,
                                  'Result':res,
                                  'Type':'Software test'}

    ### Services

    ## BranchCache/WSUS
    # BranchCache
    info_bc = 'BranchCache'
    to_find = 'PeerDistSvcs'
    res = elem_in_list(2, list(services_running_dict.keys()), to_find)
    complement_dict[info_bc] = {'Name':info_bc,
                                'Description':to_find,
                                'Result':res,
                                'Type':'Service test'}

    # WSUS
    info_wsus = 'WSUS'
    to_find = 'WSUS server'
    restemp = 'N/A'
    srv_wsus = restemp
    try:
        srv_wsus = get_wsus()
    except Exception:
        pass
    complement_dict[info_wsus] = {'Name':info_wsus,
                                  'Description':to_find,
                                  'Result':srv_wsus,
                                  'Type':'Service test'}

    ## AppLocker
    info_al = 'AppLocker'
    to_find = 'AppIDSvc'
    res = elem_in_list(2, list(services_running_dict.keys()), to_find)
    complement_dict[info_al] = {'Name':info_al,
                                'Description':to_find,
                                'Result':res,
                                'Type':'Service test'}

    # TestAV
    msg0 = '\nDo you want to perform an antivirus detection test ? (y = yes, n = no)\n'
    writer.writer(msg0)
    if verif.reponse() == 'y':
        user_rep = gen_eicar(log_file_path)
        av_etat = 'Antivirus status'
        to_find = 'Antivirus alert'
        complement_dict[av_etat] = {'Name':av_etat,
                                    'Description':to_find,
                                    'Result':user_rep,
                                    'Type':'Service test'}

    # 3 - Ecrire du fichier CSV
    header = ['Type', 'Name', 'Description', 'Result']
    csv_file = log_file_path + 'additional_scans.csv'
    writer.write_csv(csv_file, header, complement_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Additional Scans')

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, str(len(complement_dict)) + ' additional scans :\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')
