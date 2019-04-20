#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - verif.py
#@Alexandre Buissé - 2019

#Standard imports
import win32file

#Project modules imports (depuis la version 2.48)
import ask_dismount

def verifProg(prog, key):
    '''Recherche d'un programme'''
    # print (str(listeAv)+" test")#ok
    folder = key
    progFind = ""
    progFindPath = ""
    compteur = -1
    
    #recherche
    progFind, progFindPath, compteur = findThat(prog, key)

    if compteur == -1:
        print("\nCe programme a besoin de "+str(prog)+" pour fonctionner !")

    return progFind, progFindPath
 
def findThat(prog, folder):
    ''' Recherche récursive'''
    progFind = ""
    progFindPath = ""
    compteur = -1
    # print("Recherche dans => "+str(folder)) 
    folder = folder+"/"
    for root, dirnames, filenames in os.walk(folder):
        # print ("dirnames : ",str(dirnames))
        # print ("filenames : ",str(filenames))
        for filename in filenames:
            # print("recherche de "+str(prog)+" dans "+str(folder))#"ok
            # print("filename :",filename)#ok
            if prog == filename:
                progFindPath = os.path.abspath(root)
                # print("path prog :"+str(progFindPath)) #test ok
                progFind = os.path.abspath(root)+"/"+prog
                # root+"/"+prog
                # print (progFind+" trouvé !\n") #ok
                compteur = 1  
                return progFind, progFindPath, compteur
    return progFind, progFindPath, compteur

def amovible(key):
    '''Test si le périphérique passé en paramètre est amovible'''
    return win32file.GetDriveType(key) == win32file.DRIVE_REMOVABLE