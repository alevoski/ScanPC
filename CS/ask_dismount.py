#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - ask_dismount.py
#@Alexandre Buissé - 2019

#Standard imports
import msvcrt
import os
import sys

#Project modules imports
import verif
import writer

def quitter(): 
    '''
    **FR**
    Quitter 
    **EN**
    Quit
    '''
    msg = "Press any key to quit.\n"
    writer.write(msg)
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
            break
        elif 'n' in str(rep):
            return 'n'
            break     

def dismount(key):
    '''
    **FR**
    Démonter le périphérique passé en paramètre
    **EN**
    Dismount device passed in parameter
    '''
    myApp = "RemoveDrive.exe"
    remover, removerPath = verif.verifProg(myApp, key)
    if remover != "":
        params = key+" -L -b" #-L =>boucle tant que la clé n'a pa pu être éjectée et -b pour afficher le msg Windows
        cmdlineUser = myApp+" "+params
        #Obtenir le chemin du dossier du programme removedrive
        dir = os.path.dirname(os.path.abspath(remover))
        # print("test : ",cmdlineUser)
        # print("test dir : ",dir)
        subprocess.call(cmdlineUser, cwd=dir, shell=True, stdout=subprocess.PIPE, stdin=subprocess.PIPE, stderr=subprocess.PIPE)
    else:
       quitter() 