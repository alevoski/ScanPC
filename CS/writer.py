#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - writer.py
#@Alexandre Buissé - 2019

#Standard imports
from sys import stdout
from time import sleep
import os
import shutil

def prepaLogScan(logtowrite, elem):
    elem0 = '*'*len(elem)
    writeLog(logtowrite, elem0 + '\n' +elem + '\n' + elem0 + '\n')

def copyFile(file, newFile):
    '''
    Effectue la copie d'un fichier passé en paramètre
    '''
    file1 = open(file, 'r', errors='ignore')#, encoding = 'cp1252')
    file2 = open(newFile, 'w', errors='ignore')#, encoding = 'cp1252')
    shutil.copyfileobj(file1, file2)
    file1.close()
    file2.close()
    return file2 

def write(txt, timeout=100, eol=True):
    ''' Forcer l'affichage (car problème : affiche tout d'un coup à la fin (buffer ?)) '''
    try:
        timeout = timeout / 400000  # ne pas faire le calcul à chaque itération !
        for c in txt:
            print(c, end='')
            stdout.flush()       # force l'affichage du buffer
            if not c in ' \n\t': # vérifie que c n'est pas dans la chaine " \n\t"
                sleep(timeout)
        if eol:
            print()
    except OSError:
        pass

def writeLog(logFile, element):
    '''Enregistrement des logs'''
    #Créer un répertoire où seront repertorié tous les logs
    # print('logFile sent to writeLog : ', logFile)
    dirToCreate = os.path.dirname(os.path.abspath(logFile))
    # print (dirToCreate)
    if not os.path.exists(dirToCreate):
        os.makedirs(dirToCreate)      
    with open(logFile, mode='a', encoding='utf-8') as f:
        f.write(element)

def writeFinalLog(concatenateBasesFiles, listFichier):
    '''Concaténation des logs''' 
    fichier_final = open(concatenateBasesFiles, "w", errors='ignore', encoding='utf-8')
    
    for i in listFichier:
        # print(i)
        while os.path.isfile(i):
            # print ("testFile : "+str(i)) #ok
            shutil.copyfileobj(open(i, 'r', errors='ignore', encoding='utf-8'), fichier_final)
            try:
                os.remove(i) #suppression des fichiers concaténés
                # print(i, 'deleted')
            except Exception:#OSError:
                # print('oups')
                continue
    fichier_final.close()