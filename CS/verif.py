#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - verif.py
#@Alexandre Buissé - 2019

#Standard imports
import win32file

#Project modules imports
import ask_dismount

def verifProg(prog, key):
    '''
    **FR**
    Recherche d'un programme
    **EN**
    Search an app
    '''
    # print (str(listeAv)+" test")#ok
    folder = key
    progFind = ""
    progFindPath = ""
    compteur = -1
    
    #recherche
    progFind, progFindPath, compteur = findThat(prog, key)

    if compteur == -1:
        print("\nThis app needs " + str(prog) + " !")

    return progFind, progFindPath
 
def findThat(prog, folder):
    ''' 
    **FR**
    Recherche récursive
    **EN**
    Recursive search
    '''
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
    '''
    **FR**
    Test si le périphérique passé en paramètre est amovible
    **EN**
    Test if device in parameter is removable
    '''
    return win32file.GetDriveType(key) == win32file.DRIVE_REMOVABLE