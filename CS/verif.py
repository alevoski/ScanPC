#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - verif.py
#@Alexandre Buissé - 2019/2020

#Standard imports
import msvcrt
import os
import sys
import subprocess
import win32file

#Project modules imports
from writer import writer

def quitter():
    '''
    **FR**
    Quitter
    **EN**
    Quit
    '''
    msg = "Press any key to quit.\n"
    writer(msg)
    msvcrt.getch()
    os.system('cls')
    sys.exit(0)

def reponse():
    '''
    **FR**
    Réponse
    **EN**
    Response
    '''
    while True:
        rep = msvcrt.getch()
        if 'y' in str(rep):
            # print("machin ok")
            return 'y'
        elif 'n' in str(rep):
            return 'n'

def dismount(key):
    '''
    **FR**
    Démonter le périphérique passé en paramètre
    **EN**
    Dismount device passed in parameter
    '''
    my_app = "RemoveDrive.exe"
    remover, _ = verif_prog(my_app, key)
    if remover != '':
        params = key + ' -L -b' #-L =>boucle tant que la clé n'a pa pu être éjectée et -b pour afficher le msg Windows
        cmdline_user = my_app + ' ' + params
        #Obtenir le chemin du dossier du programme removedrive
        directory = os.path.dirname(os.path.abspath(remover))
        # print("test : ",cmdline_user)
        # print("test directory : ",directory)
        subprocess.call(cmdline_user, cwd=directory, shell=True, stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    else:
        quitter()

def verif_prog(prog, key):
    '''
    **FR**
    Recherche d'un programme
    **EN**
    Search an app
    '''
    prog_find = ''
    prog_find_path = ''
    compteur = -1

    # Recherche
    prog_find, prog_find_path, compteur = find_that(prog, key)

    if compteur == -1:
        print("\nThis app needs " + str(prog) + " !")

    return prog_find, prog_find_path

def find_that(prog, folder):
    '''
    **FR**
    Recherche récursive
    **EN**
    Recursive search
    '''
    prog_find = ""
    prog_find_path = ""
    compteur = -1
    # print("Recherche dans => "+str(folder))
    folder = folder+"/"
    for root, _, filenames in os.walk(folder):
        for filename in filenames:
            # print("recherche de "+str(prog)+" dans "+str(folder))#"ok
            # print("filename :",filename)#ok
            if prog == filename:
                prog_find_path = os.path.abspath(root)
                # print("path prog :"+str(prog_find_path)) #test ok
                prog_find = os.path.abspath(root)+"/"+prog
                # root+"/"+prog
                # print (prog_find+" trouvé !\n") #ok
                compteur = 1
                return prog_find, prog_find_path, compteur
    return prog_find, prog_find_path, compteur

def amovible(key):
    '''
    **FR**
    Test si le périphérique passé en paramètre est amovible
    **EN**
    Test if device in parameter is removable
    '''
    return win32file.GetDriveType(key) == win32file.DRIVE_REMOVABLE
