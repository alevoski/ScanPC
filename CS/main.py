#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - main.py
#@Alexandre Buissé - 2019/2020

#Standard imports
import os
import time
import webbrowser
from datetime import datetime

#Project modules imports
import writer
import verif
import scans
from software_scan import software_init
from complement import complement_init

COMPUTERNAME = os.environ['COMPUTERNAME']
STYLECSSFILE = os.getcwd() + '/style.css'

def scanpart():
    '''
    **FR**
    Retourne la lettre du lecteur où est exécuté l'app
    **EN**
    Return the drive letter where the app is launched
    '''
    key = os.getcwd() # On prend le chemin du répertoire du script/exe
    key = str(key[:2])
    return key

def scanpc(log_file_path):
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
    texte = "Computer scanning in progress... Please wait...\n"
    writer.writer(texte)

    log_file = str(log_file_path) + "FINAL.html"

    # Write first html tags
    elementraw = """<!DOCTYPE html>
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

    # Texte d'intro
    texte = '<h1>Scanning of computer "' + COMPUTERNAME + '"</h1>'
    element2 = '<time>' + datetime.now().strftime("%A %d %B %Y %H:%M:%S") + '</time>\n'
    element = elementraw.format(texte, element2)
    writer.writelog(log_file, element)

    # Scans de base
    scans.user_info(log_file_path)
    scans.shared_folders_info(log_file_path)
    scans.hotfixes_info(log_file_path)
    scans.system_info(log_file_path)
    scans.security_product_info(log_file_path)
    scans.process_info(log_file_path)
    services_running_dict = scans.services_info(log_file_path)
    scans.ports_info(log_file_path)
    scans.persistence_info(log_file_path)

    texte4 = "\nBasic scan ended.\n"
    writer.writer(texte4)

    # Scan des logiciels
    software_dict = software_init(log_file_path)

    return log_file, software_dict, services_running_dict

def readandcopy(log_file):
    '''
    **FR**
    Proposer à l'utilisateur d'ouvrir le log passé en paramètre
    et de le copier sur l'ordinateur scanné
    **EN**
    Ask user if he wants to read the log
    and if he wants a copy on the scanned computer
    '''
    msg0 = "\nDo you want to read the scan report ? (y = yes, n = no)\n"
    writer.writer(msg0)
    if verif.reponse() == 'y':
        openers = webbrowser._tryorder
        opened = 0
        for browsers in openers:
            # print(browsers)
            if 'firefox' in browsers.lower() or 'iexplore' in browsers.lower():
                # print('Opening with ' + browsers)
                browser = webbrowser.get(webbrowser._tryorder[openers.index(browsers)])
                browser.open('file:' + str(log_file))
                opened = 1
                break
        if opened == 0:
            webbrowser.open(str(log_file))

    # Copie de fichier vers un emplacement défini
    msg0 = "\nDo you want a copy the scan report on the computer (C:\\ drive) ? (y = yes, n = no)\n"
    writer.writer(msg0)
    if verif.reponse() == 'y':
        # Copy log
        file_name = os.path.basename(log_file)
        file_to_write = "C:/" + file_name
        writer.copy_file(log_file, file_to_write)
        # Copy CSS file
        if os.path.isfile(STYLECSSFILE):
            file_name = os.path.basename(STYLECSSFILE)
            file_to_write = "C:/" + file_name
            writer.copy_file(STYLECSSFILE, file_to_write)

def fin(key):
    '''
    **FR**
    Si le périphérique est amovible, proposer à l'utilisateur de le retirer proprement
    **EN**
    If the device is removable, ask the user to safely remove (dismount) his device
    '''
    # Test si lecteur amovible
    texte = "Computer scanning ended, details have been saved on " + str(key) + ".\n"
    writer.writer(texte)
    if verif.amovible(key):
        # Proposer la déconnexion de la clé
        msg0 = "\nDo you want to dismount your device and quit ? (y = yes, n = no)\n"
        writer.writer(msg0)
        if verif.reponse() == 'y':
            verif.dismount(key)
    verif.quitter()

def init():
    '''
    **FR**
    Initialisation du programme
    **EN**
    Software initialization
    '''
    print('\t\t\t---------- ScanPC ---------')
    print('\t***************************************************************')
    print('\t*\t\t    Computer scanning                         *')
    print('\t*\t\t          Disclaimer !                        *')
    print('\t*\t This sofware only scan your computer                 *')
    print('\t*\t for users, various system information and software.  *')
    print('\t*\t No copy of your data will be performed.              *')
    print('\t***************************************************************')
    print()

    # Begin timer
    begin_scan = time.time()

    # Début du programme
    key = scanpart()
    unique_dir = str(COMPUTERNAME) + '_' + datetime.now().strftime('%d%m%y%H%M%S')
    log_file_path = str(key) + "/logScanPC/" + datetime.now().strftime('%Y/%m/%d/') + unique_dir + '/' + unique_dir
    # Appel de scanPC
    log_file, software_dict, services_running_dict = scanpc(log_file_path)
    # Tests complémentaires
    complement_init(log_file_path, software_dict, services_running_dict)
    writer.writelog(log_file, '\n</div>\n')

    # Ending timer
    end_scan = time.time()
    total_time_scan = end_scan - begin_scan
    elem = "----------------------- Scan ended ------------------------"
    writer.prepa_log_scan(log_file, elem)
    total_time = 'Computer analyzed in {} seconds.'.format(round(total_time_scan, 2))
    writer.writelog(log_file, total_time)
    writer.writer('Computer analyzed in ' + str(round(total_time_scan, 2)) + ' seconds.\n')

    # Write last html tags
    elementraw = """
        </body>
    </html>"""
    writer.writelog(log_file, elementraw)

    # Copy css file in the log directory
    if os.path.isfile(STYLECSSFILE):
        file_name = os.path.basename(STYLECSSFILE)
        file_to_write = str(key) + "/logScanPC/" + datetime.now().strftime('%Y/%m/%d/') + unique_dir + '/'  + file_name
        writer.copy_file(STYLECSSFILE, file_to_write)

    #Fin du programme
    readandcopy(log_file)
    fin(key)

if __name__ == '__main__':
    init()
