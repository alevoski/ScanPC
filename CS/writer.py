#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - writer.py
#@Alexandre Buissé - 2019

#Standard imports
from sys import stdout
from time import sleep
import os
import csv
import shutil
from cgi import escape

def copyFile(source, destination):
    '''
    Copy a file (source) to a destination
    '''
    fileName = os.path.basename(source)
    try:
        shutil.copy2(source, destination)
        if os.path.isfile(destination):
            msg = "\n" + fileName + " has been copied !\n"
            write(msg)
        else:
            raise IOError
    except (IOError, PermissionError) as e:
        msg = "\nUnable to copy " + fileName + " ! :\nPermission denied !"
        write(msg)

def prepaLogScan(logtowrite, elem):
    '''
    **FR**
    Ecrire un elément dans un fichier passés en paramètre
    **EN**
    Write elem in a file pass in parameter
    '''
    elem0 = '*'*len(elem)
    writeLog(logtowrite, elem0 + '<br>\n' +elem + '<br>\n' + elem0 + '<br>\n<br>\n')

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

def writeCSV(logFile, fieldnames, dictElement):
    '''
    **FR**
    Enregistrement des logs au format CSV
    **EN**
    Write logs in CSV format
    '''
    with open(logFile, mode='w', newline='', encoding='utf-8-sig') as csvfile:
        writerCSV = csv.DictWriter(csvfile, delimiter=';', fieldnames=fieldnames)
        writerCSV.writeheader()
        for elems, values in sorted(dictElement.items()):
            writerCSV.writerow(values)

def _row2tr(row, attr=None):
    '''
    **FR**
    Transforme les lignes (issues d'un fichier CSV) en lignes de tableau HTML
    **EN**
    Transform rows (from a CSV file) in HTML table rows
    '''
    cols = row.split(';')
    return ('<TR>'
            + ''.join('<TD>%s</TD>' % data for data in cols)
            + '</TR>')
            
def _rowheader(row, attr=None):
    '''
    **FR**
    Transforme les lignes (issues d'un fichier CSV) en lignes d'entête de tableau HTML
    **EN**
    Transform rows (from a CSV file) in HTML table rows header
    '''
    cols = escape(row).split(';')
    return ('<TR>'
            + ''.join('<TH>%s</TH>' % data for data in cols)
            + '</TR>')

def csv2html(csvFile, summary):
    '''
    **FR**
    Transforme un fichier CSV en tableau HTML
    **EN**
    Transform a CSV file in an HTML table
    '''
    htmltxt = '<TABLE summary="' + summary + '">\n'
    with open(csvFile, mode='r', encoding='utf-8-sig') as f:
        csvfile = f.read()
    for rownum, row in enumerate(csvfile.split('\n')):
        if row != '':
            #Prepare header
            if rownum == 0:
                # print(row)
                htmlrow = _rowheader(row)
            else:
                htmlrow = _row2tr(row)
            htmlrow = '  <TBODY>%s</TBODY>\n' % htmlrow
            htmltxt += htmlrow
    htmltxt += '</TABLE>\n'
    return htmltxt