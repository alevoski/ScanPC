#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - writer.py
#@Alexandre Buissé - 2019

#Standard imports
from sys import stdout
from time import sleep
import os

def prepaLogScan(logtowrite, elem):
    '''
    **FR**
    Ecrire un elément dans un fichier passés en paramètre
    **EN**
    Write elem in a file pass in parameter
    '''
    elem0 = '*'*len(elem)
    writeLog(logtowrite, elem0 + '\n' +elem + '\n' + elem0 + '\n')

def write(txt, timeout=100, eol=True):
    '''
    **FR**
    Forcer un affichage harmonieux (car sinon tout est affiché trop vite sans)
    **EN**
    Force a smooth display (because everything is printed to quickly without it)
    '''
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
    '''
    **FR**
    Enregistrement des logs
    **EN**
    Write logs
    '''
    #Créer un répertoire où seront repertorié tous les logs
    # print('logFile sent to writeLog : ', logFile)
    dirToCreate = os.path.dirname(os.path.abspath(logFile))
    # print (dirToCreate)
    if not os.path.exists(dirToCreate):
        os.makedirs(dirToCreate)      
    with open(logFile, mode='a', encoding='utf-8') as f:
        f.write(element)