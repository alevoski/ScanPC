#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - scans.py
#@Alexandre Buissé - 2019

#Standard imports
import os
import re
from datetime import datetime
import platform
import psutil
import win32com.client
from pywintypes import com_error
import winreg
import ctypes
import time

#Project modules imports
import writer

computername = os.environ['COMPUTERNAME']
userFlagsDict = {'SCRIPT':1,
                'ACCOUNTDISABLE':2,
                'HOMEDIR_REQUIRED':8,
                'LOCKOUT':16,
                'PASSWD_NOTREQD':32,
                'PASSWD_CANT_CHANGE':64,
                'ENCRYPTED_TEXT_PWD_ALLOWED':128,
                'TEMP_DUPLICATE_ACCOUNT':256,
                'NORMAL_ACCOUNT':512,
                'INTERDOMAIN_TRUST_ACCOUNT':2048,
                'WORKSTATION_TRUST_ACCOUNT':4096,
                'SERVER_TRUST_ACCOUNT':8192,
                'DONT_EXPIRE_PASSWORD':65536,
                'MNS_LOGON_ACCOUNT':131072,
                'SMARTCARD_REQUIRED':262144,
                'TRUSTED_FOR_DELEGATION':524288,
                'NOT_DELEGATED':1048576,
                'USE_DES_KEY_ONLY':2097152,
                'DONT_REQ_PREAUTH':4194304,
                'PASSWORD_EXPIRED':8388608,
                'TRUSTED_TO_AUTH_FOR_DELEGATION':16777216,
                'PARTIAL_SECRETS_ACCOUNT':67108864}
                
productStateDict = {'262144':{'defstatus':'Up to date','rtstatus':'Disabled'},
                    '266240':{'defstatus':'Up to date','rtstatus':'Enabled'},
                    '262160':{'defstatus':'Outdated','rtstatus':'Disabled'},
                    '266256':{'defstatus':'Outdated','rtstatus':'Enabled'},
                    # '270336'
                    # '327680'
                    # '327696'
                    '331776':{'defstatus':'Up to date','rtstatus':'Enabled'},
                    # '335872'
                    # '344064'
                    '393216':{'defstatus':'Up to date','rtstatus':'Disabled'},
                    '393232':{'defstatus':'Outdated','rtstatus':'Disabled'},
                    '393488':{'defstatus':'Outdated','rtstatus':'Disabled'},
                    '397312':{'defstatus':'Up to date','rtstatus':'Enabled'},
                    '397328':{'defstatus':'Outdated','rtstatus':'Enabled'},
                    '393472':{'defstatus':'Up to date','rtstatus':'Disabled'},
                    '397584':{'defstatus':'Outdated','$rtstatus':'Enabled'},
                    '397568':{'defstatus':'Up to date','rtstatus':'Enabled'},
                    # '458768'
                    # '458752'
                    '462864':{'defstatus':'Outdated','rtstatus':'Enabled'},
                    '462848':{'defstatus':'Up to date','rtstatus':'Enabled'}}

def detectOS():
    '''
    **FR**
    Retourne l'OS utilisé
    **EN**
    Return OS used
    '''
    TP = str(platform.platform().lower())
    # print(TP)
    if 'xp' in TP or '2003' in TP:
        rep = 'xp'
    if '7' in TP or 'vista' in TP or '2008' in TP or '2012' in TP:
        rep = '7'
    if '10' in TP or '2016' in TP:
        rep = '10'
    return rep

def calcFlag(userFlagsDict, flag):
    '''
    https://support.microsoft.com/en-sg/help/305144/how-to-use-useraccountcontrol-to-manipulate-user-account-properties
    **FR**
    Calcul l'id de la propriété d'un compte utilisateur avec le dictionnaire de propriétés 
    et le flag du compte utilisateur en paramètre
    Retourne :
        - La différence la plus petite trouvé entre le flag et la valeur d'une des propriétés
        Cette différence sera réutilisée pour trouver d'autres valeurs de propriété par la même opération
        - L'id de la propriété du compte trouvée
    **EN**
    Get propertie id of an user account with the flags dict and the userflag in parameters
    Return :
        - The lowest difference found between the flag and one of the propertie value
        - The found account propertie id
    '''
    flagTosave = 0
    min = 0
    for k, v in userFlagsDict.items():
        if v == flag:
            min = flag - v
            flagTosave = v
            break
        if v < flag and v > flagTosave:
            min = flag - v
            flagTosave = v
    return min, flagTosave

def userInfo(logFilePath):
    '''
    ~ net user <username> /domain    &    net user <username>
    **FR**
    Obtient la liste des utilisateurs AD et locaux de l'ordinateur avec la base de registre et WinNT
    Retourne le log généré
    **EN**
    Get AD and local users list of the computer with regedit and WinNT
    Return generated log
    '''
    writer.write('Getting computer users')
    hive = winreg.HKEY_LOCAL_MACHINE
    keyUsr = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" # AD users
    registry_key = None
    usrLstTEMP = []
    usrLstTEMP2 = []
    profilPath = "ProfileImagePath"

    # 1 - Liste tous les profils
    try:
        registry_key = winreg.OpenKey(hive, keyUsr, 0, winreg.KEY_READ)
        # print("ok")
    except FileNotFoundError:
        # print("oups")
        pass
    # print(registry_key)
    if registry_key != None:
        i = 0
        while True:
            try:
                usrLstTEMP.append(winreg.EnumKey(registry_key, i)) #ok pour 1 clé
                i+=1
                # print("value :", str(registry_key))
            except Exception:
                # print(Exception)
                winreg.CloseKey(registry_key)
                break

    # 2 - Tri de la liste pour ne garder que les SID commençant par "1-5-21" et récupérer le nom des utilisateurs
    for item in usrLstTEMP:
        if "1-5-21" in item:
            # print(item)
            key = keyUsr + '\\' + item
            # print(key)
            try:
                registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ)
                value, regtype = winreg.QueryValueEx(registry_key, profilPath)
                winreg.CloseKey(registry_key)
                # print(value)
                valueMOD = re.findall(r'[^\\]+\s*$', value)
                # print(type(valueMOD[0]))
                usrLstTEMP2.append(valueMOD[0])
            except FileNotFoundError:
                # print("oups")
                pass

    domain = os.environ['userdomain']
    # 3 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Getting details about the ' + str(len(usrLstTEMP2)) + ' user(s) of computer "' + computername + '"</h2>'
    elem2 = "<br>Domain : " + str(domain) + '\n'
    writer.prepaLogScan(logFile, elem)
    writer.writeLog(logFile, elem2)

    # 4 - Détails sur les utilisateurs
    userDict = {}
    for user in usrLstTEMP2:
        try:
            objOU = win32com.client.GetObject("WinNT://" + domain + "/" + user + ",user")
            #User account
            fullname = str(objOU.Get('fullname'))
            # print(fullname)
            try:
                expirationDate = str(objOU.Get('AccountExpirationDate'))
            except Exception as e:
                expirationDate = 'N/A'
            profile = str(objOU.Get('Profile'))
            loginScript = str(objOU.Get('LoginScript'))
            try:
                lastLogin = str(objOU.Get('lastlogin'))
            except com_error as e:
                lastLogin = 'N/A'
            primaryGroupID = str(objOU.Get('PrimaryGroupID'))
            autoUnlockInterval = str(objOU.Get('AutoUnlockInterval'))
            lockoutObservationInterval = str(objOU.Get('LockoutObservationInterval'))
            homedir = str(objOU.Get('HomeDirectory'))
            homedirDrive = str(objOU.Get('HomeDirDrive'))
            
            #Password
            pwdAge = str(round(objOU.Get('PasswordAge')/3600/24))
            pwdMinAge = str(round(objOU.Get('MinPasswordAge')/3600/24))
            pwdMaxAge = str(round(objOU.Get('MaxPasswordAge')/3600/24))
            pwdExpired = str(objOU.Get('PasswordExpired'))
            pwdMaxBadpwdAllowed = str(objOU.Get('MaxBadPasswordsAllowed'))
            pwdMinLength = str(objOU.Get('MinPasswordLength'))
            pwdHistoryLength = str(objOU.Get('PasswordHistoryLength'))

            #Groups
            groupsList = []
            for grp in objOU.Groups():
                groupsList.append(grp.Name + '<br>')

            #Flag
            flagFinalList = []
            flag = objOU.Get('UserFlags')
            #Get flags with user flag (user flag = sum(flags))
            flagList = []
            min, flagTosave = calcFlag(userFlagsDict, flag)
            flagList.append(flagTosave)
            while min != 0:
                min, flagTosave = calcFlag(userFlagsDict, min)
                flagList.append(flagTosave)
            #Verify if flag = sum(flagList)
            if flag == sum(flagList):
                flagFinalList.append(str(flag) + ' => ' + str(flagList) + '<br>')
                #Get flags description with flag nums saved
                flagFinalList.append('Flag properties :')
                for k, v in userFlagsDict.items():
                    if v in flagList:
                        flagFinalList.append('<br>' + str(v) + ' : ' + str(k))
            userDict[user] = {'User':user, 'Fullname':fullname, 'Expiration_Date':expirationDate,
                            'Profile':profile, 'Login_Script':loginScript, 'Last_Login':lastLogin, 
                            'Primary_Group_ID':primaryGroupID, 'Auto_Unlock_Interval (secs)': autoUnlockInterval, 
                            'Lockout_Observation_Interval (secs)':lockoutObservationInterval,
                            'HomeDir':homedir, 'HomeDirDrive':homedirDrive, 
                            'pwdAge (days)':pwdAge, 'pwdMinAge (days)':pwdMinAge, 'pwdMaxAge (days)':pwdMaxAge, 
                            'pwdExpired':pwdExpired, 'pwdMaxBadpwdAllowed':pwdMaxBadpwdAllowed,
                            'pwdMinLength':pwdMinLength, 'pwdHistoryLength':pwdHistoryLength,
                            'Groups':(''.join(groupsList)), 'Flag':(''.join(flagFinalList))}

        except Exception as e:
            pass
        i+=1

    if userDict != '':
        # Ecriture du fichier CSV
        header = ['User', 'Fullname', 'Expiration_Date', 'Profile', 'Login_Script', 'Last_Login',
                'Primary_Group_ID', 'Auto_Unlock_Interval (secs)', 'Lockout_Observation_Interval (secs)',
                'HomeDir', 'HomeDirDrive','pwdAge (days)', 'pwdMinAge (days)', 'pwdMaxAge (days)', 
                'pwdExpired', 'pwdMaxBadpwdAllowed', 'pwdMinLength', 'pwdHistoryLength', 'Groups', 'Flag']
        csvFile = logFilePath + "users.csv"
        writer.writeCSV(csvFile, header, userDict)
        # Transformation du CSV en HTML
        htmltxt = writer.csv2html(csvFile, 'Users details')
        writer.writeLog(logFile, htmltxt)

    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, '\n</div>\n')

    return logFile

def sharedFolders(logFilePath):
    '''
    ~ net share
    **FR**
    Liste les dossiers partagés
    Retourne le log généré
    **EN**
    List shared folders
    Return generated log
    '''
    writer.write('Getting shared folders')
    # 1 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Shared folders of computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des dossiers partagés de l'ordinateur
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_Share")
    sfDict = {}
    for objItem in colItems:
        sfDict[str(objItem.Name)] = {'Name':str(objItem.Name), 'Path':str(objItem.Path),
                                    'Caption':str(objItem.Caption), 'Description':str(objItem.Description)}

    # 3 - Ecriture du fichier CSV
    header = ['Name', 'Path', 'Caption', 'Description']
    csvFile = logFilePath + "sharedFolders.csv"
    writer.writeCSV(csvFile, header, sfDict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Shared Folders')
    writer.writeLog(logFile, htmltxt)

    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, '\n</div>\n')

    return logFile

def hotFixesInfo(logFilePath):
    '''
    https://www.activexperts.com/admin/scripts/wmi/python/0417/
    ~ wmic qfe get HotfixID,InstalledOn | more
    **FR**
    Liste les patchs de sécurité Windows installés
    Retourne le log généré
    **EN**
    List Windows security updates installed
    Return generated log
    '''
    writer.write('Getting Windows security updates')
    # 1 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Windows security updates of computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des correctifs de l'ordinateur
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_QuickFixEngineering WHERE Description='Security Update'")
    hotfixDict = {}
    for objItem in colItems:
        try:
            hotfixDict[objItem.HotFixID] = {'HotFixID':objItem.HotFixID, 
                        'InstalledOn':str(datetime.strptime(str(objItem.InstalledOn), "%m/%d/%Y").strftime("%Y-%m-%d"))}
        except ValueError:
            hotfixDict[objItem.HotFixID] = {'HotFixID':objItem.HotFixID, 
                        'InstalledOn':str(objItem.InstalledOn)}

    # 3 - Ecriture du fichier CSV
    header = ['HotFixID', 'InstalledOn']
    csvFile = logFilePath + "hotfixes.csv"
    writer.writeCSV(csvFile, header, hotfixDict)
    
    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Hotfixes')

    # 5 - Ecriture fin de log
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, '\n</div>\n')

    return hotfixDict

def systemInfo(logFilePath):
    '''
    ~ systeminfo | find /V /I "hotfix" | find /V "KB"
    ~ wmic logicaldisk get volumename, description, FileSystem, Caption, ProviderName
    **FR**
    Scan le système (os, bios, cpu, ram, cartes réseaux, disques durs, etc)
    Retourne le log généré
    **EN**
    Scan system (os, bios, cpu, ram, network interfaces, drives, etc)
    Return generated log
    '''
    writer.write('Getting system informations')
    delimiter = '*' * 40

    # 1 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>System information of computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Get information from regedit
    hive = winreg.HKEY_LOCAL_MACHINE
    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion" #W7-64/32
    registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) #W7-64/32

    # 2' - Specific to W10 (ok for 64 bits)
    if detectOS() == '10':
        writer.write('Perform operations for W10+ OS')
        key = r"SYSTEM\Setup" #\Source OS" #W10
        registry_keys = winreg.OpenKey(hive, key, 0, winreg.KEY_READ)
        i = 0
        while 1:
            try:
                name = winreg.EnumKey(registry_keys, i)
                # print(name)
                if 'Source OS' in name:
                    # print(name)
                    key = key + '\\' + name
                    registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ) #W10-64
                    break
            except Exception as e:
                # print(e)
                break
                pass
            i+=1
    
    # Continuing 
    dictToSave = {'InstallDate':'N/A', 'SystemRoot':'N/A', 'ProductId':'N/A',
    'ProductName':'N/A', 'RegisteredOrganization':'N/A', 'RegisteredOwner':'N/A'}
    infoToGet = ['InstallDate', 'SystemRoot', 'ProductId',
    'ProductName', 'RegisteredOrganization', 'RegisteredOwner']
    i = 0
    while 1:
        try:
            name, value, type = winreg.EnumValue(registry_key, i)
            if name in infoToGet:
                # print(repr(name), value, type)
                dictToSave[name] = value
        except Exception as e:
            # print('error : ' + name, type)
            pass
            break
        i+=1
    # print(dictToSave)

    # OS
    writer.writeLog(logFile, 'Hostname : ' + platform.node() + '<br>\n')
    # print('OS : ' + platform.system() + ' ' + platform.version())
    writer.writeLog(logFile, 'OS : ' + dictToSave.get('ProductName') + '<br>\n')
    writer.writeLog(logFile, 'OS type : ' + platform.machine() + '<br>\n')
    writer.writeLog(logFile, 'Product Id : ' + dictToSave.get('ProductId') + '<br>\n')
    writer.writeLog(logFile, 'Install Date: ' + str(datetime.fromtimestamp(dictToSave.get('InstallDate'))) + '<br>\n')
    writer.writeLog(logFile, 'System Root: ' + dictToSave.get('SystemRoot') + '<br>\n')
    
    # Language
    bufSize = 32
    buf = ctypes.create_unicode_buffer(bufSize)
    dwFlags = 0
    lcid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    ctypes.windll.kernel32.LCIDToLocaleName(lcid, buf, bufSize, dwFlags)
    writer.writeLog(logFile, 'Regional and language options : ' + buf.value + '<br>\n')
    
    # Time zone
    writer.writeLog(logFile, 'Time zone : ' + time.tzname[0] + '<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    # Owner
    writer.writeLog(logFile, 'Registered Organization : ' + dictToSave.get('RegisteredOrganization') + '<br>\n')
    writer.writeLog(logFile, 'Registered Owner : ' + dictToSave.get('RegisteredOwner') + '<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    # Computer model/brand
    strComputer = "."
    objWMIService = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    objSWbemServices = objWMIService.ConnectServer(strComputer, 'root\cimv2')
    colItems = objSWbemServices.ExecQuery('SELECT * FROM Win32_ComputerSystem')
    for objItem in colItems:
        writer.writeLog(logFile, 'Manufacturer : ' + objItem.Manufacturer + '<br>\n')
        try:
            writer.writeLog(logFile, 'SystemFamily : ' + objItem.SystemFamily + '<br>\n')
        except AttributeError as e:
            writer.writeLog(logFile, 'SystemFamily : N/A<br>\n')
        writer.writeLog(logFile, 'Model : ' + objItem.Model + '<br>\n')
        writer.writeLog(logFile, 'Type : ' + objItem.SystemType + '<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')

    # Processor
    colItems = objSWbemServices.ExecQuery('Select * from Win32_Processor')
    for objItem in colItems:
        writer.writeLog(logFile, 'Processor : ' + objItem.Name + '<br>\n')
        writer.writeLog(logFile, 'Core : ' + str(objItem.NumberOfCores) + '<br>\n')
        writer.writeLog(logFile, 'Logical core : ' + str(objItem.NumberOfLogicalProcessors) + '<br>\n')
        try:
            writer.writeLog(logFile, 'Virtualization enabled : ' + str(objItem.VirtualizationFirmwareEnabled) + '<br>\n')
        except AttributeError as e:
            writer.writeLog(logFile, 'Virtualization enabled : N/A<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    # Bios
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BIOS")
    for objItem in colItems:
        writer.writeLog(logFile, "BIOS Version : " + str(objItem.BIOSVersion) + '<br>\n')
        writer.writeLog(logFile, "Release Date : " + str(datetime.strptime(str(objItem.ReleaseDate[:8]), "%Y%m%d").strftime("%Y-%m-%d")) + '<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')

    # Memory
    mem = psutil.virtual_memory()
    writer.writeLog(logFile, 'Total memory : ' + str(mem.total >> 20) + ' Mo<br>\n')
    writer.writeLog(logFile, 'Available memory : ' + str(mem.available >> 20) + ' Mo<br>\n')
    writer.writeLog(logFile, 'Used memory : ' + str(mem.used >> 20) + ' Mo (' + str(mem.percent) + ' %)<br>\n')
    writer.writeLog(logFile, 'Free memory : ' + str(mem.free >> 20) + ' Mo<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    # Domaine
    writer.writeLog(logFile, "Domain : " + os.environ['userdomain'] + '<br>\n')
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    #Network interfaces
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("select * from Win32_NetworkAdapterConfiguration")
    networkDict = {}
    for objItem in colItems:
        if objItem.MACAddress != None:
            netCaption = re.findall("] (.*)", objItem.Caption)[0]
            # print(netCaption, objItem.IPAddress)
            try:
                dnsSrv = '<br>'.join(objItem.DNSServerSearchOrder)
            except TypeError:
                dnsSrv = objItem.DNSServerSearchOrder
            try:
                ipv4 = objItem.IPAddress[0]
            except IndexError:
                ipv4 = ''
            except TypeError:
                ipv4 = ''
            try:
                ipv6 = objItem.IPAddress[1]
            except IndexError:
                ipv6 = ''
            except TypeError:
                ipv6 = ''
            try:
                gw = ''.join(objItem.DefaultIPGateway)
            except TypeError:
                gw = objItem.DefaultIPGateway
            try:
                dhcpSrv = ''.join(objItem.DHCPServer)
            except TypeError:
                dhcpSrv = objItem.DHCPServer
            networkDict[netCaption] = {'Name':netCaption,
                    'MAC':objItem.MACAddress,
                    'Enabled':objItem.IPEnabled,
                    'IpV4':ipv4, 
                    'IpV6':ipv6,
                    'Default Gateway':gw,
                    'DHCP Enabled':objItem.DHCPEnabled,
                    'DHCP Server':dhcpSrv,
                    'WINS Primary Server':objItem.WINSPrimaryServer,
                    'WINS Secondary Server':objItem.WINSSecondaryServer,
                    'DNS Domain':objItem.DNSDomain,
                    'DNS Servers':dnsSrv}
    writer.writeLog(logFile, str(len(networkDict)) + ' network interfaces found :<br>\n')
    # Ecriture du fichier CSV 
    header = ['Name', 'MAC', 'Enabled', 'IpV4', 'IpV6', 'Default Gateway',
            'DHCP Enabled', 'DHCP Server', 'WINS Primary Server',
            'WINS Secondary Server', 'DNS Domain', 'DNS Servers']
    csvFile = logFilePath + "networkInterfaces.csv"
    writer.writeCSV(csvFile, header, networkDict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Network Interfaces')
    # Ecriture sur le log
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, delimiter + '<br>\n')
    
    # Drives
    drives = psutil.disk_partitions()
    drivesDict = {}
    for drive in drives:
        drivePath = re.findall("device='(.*?)'", str(drive))[0]
        if drive.opts != 'cdrom':
            try:
                diskUsage = psutil.disk_usage(drivePath)
                drivesDict[str(drivePath)] = {'Path':str(drivePath), 
                                        'Total_Space (Go)':str(diskUsage.total >> 30),
                                        'Free_Space (Go)':str(diskUsage.free >> 30),
                                        'Used_space (Go)':str(diskUsage.used >> 30),
                                        'Used_space (%)':str(diskUsage.percent)}
            except Exception:
                pass
    # Ecriture du fichier CSV 
    header = ['Path', 'Total_Space (Go)', 'Free_Space (Go)', 'Used_space (Go)', 'Used_space (%)']
    csvFile = logFilePath + "drives.csv"
    writer.writeCSV(csvFile, header, drivesDict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Drives')
    # Ecriture sur le log
    writer.writeLog(logFile, str(len(drivesDict)) + ' drive(s) found :<br>\n')
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, delimiter + '<br>\n')

    # 3 - Ecriture fin de log
    writer.writeLog(logFile, '\n</div>\n')

    return logFile
    
def securityProductState(productStateDict, productState):
    '''
    **FR**
    Retourne la signification du productState passé en paramètre
    **EN**
    Return productState meaning pasted in parameter
    '''
    rtstatus = 'Unknown'
    defstatus = 'Unknown'
    for keys, values in productStateDict.items():
        if productState == keys:
            rtstatus = values['rtstatus']
            defstatus = values['defstatus']
            break
    return rtstatus, defstatus
    
def securityProductInfo(logFilePath):
    '''
    http://neophob.com/2010/03/wmi-query-windows-securitycenter2/
    **FR**
    Scan le système pour obtenir la liste et le status des produits de sécurité (antivirus, parefeu, antispyware)
    Retourne le dictionnaire des produits de sécurité
    **EN**
    Scan to get security product in the system with their status (antivirus, firewall, antispyware)
    Return security products dict
    '''
    writer.write('Getting security products')
    # 1 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Security products on computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtention des informations
    securityProductDict = {}
    i = 1
    # Antivirus
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\SecurityCenter2")
    colItems = objSWbemServices.ExecQuery("select * from AntivirusProduct")
    for objItem in colItems:
        productName = objItem.displayName
        instanceGuid = objItem.instanceGuid
        pathProduct = objItem.pathToSignedProductExe
        pathReporting = objItem.pathToSignedReportingExe
        productState = str(objItem.productState)
        # ProductState conversion
        rtstatus, defstatus = securityProductState(productStateDict, productState)
        i+=1
        # Put in dict
        securityProductDict[i] = {'Name':productName, 'GUID':instanceGuid,
                                            'pathProduct':pathProduct, 'pathReporting':pathReporting,
                                            'productState':productState,
                                            'rtstatus':rtstatus, 'defstatus':defstatus,
                                            'Type':'Antivirus'}

    # Firewalls
    # Windows
    XPFW = win32com.client.gencache.EnsureDispatch('HNetCfg.FwMgr', 0)
    XPFW_policy = XPFW.LocalPolicy.CurrentProfile
    writer.writeLog(logFile, 'Windows firewall enabled : ' + str(XPFW_policy.FirewallEnabled) + '<br>\n')
    # Third party firewall
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\SecurityCenter2")
    colItems = objSWbemServices.ExecQuery("select * from FireWallProduct")
    for objItem in colItems:
        productName = objItem.displayName
        instanceGuid = objItem.instanceGuid
        pathProduct = objItem.pathToSignedProductExe
        pathReporting = objItem.pathToSignedReportingExe
        productState = str(objItem.productState)
        # ProductState conversion
        rtstatus, defstatus = securityProductState(productStateDict, productState)
        i+=1
        # Put in dict
        securityProductDict[i] = {'Name':productName, 'GUID':instanceGuid,
                                            'pathProduct':pathProduct, 'pathReporting':pathReporting,
                                            'productState':productState,
                                            'rtstatus':rtstatus, 'defstatus':defstatus,
                                            'Type':'Firewall'}

    # AntiSpyware
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\SecurityCenter2")
    colItems = objSWbemServices.ExecQuery("select * from AntiSpywareProduct")
    for objItem in colItems:
        productName = objItem.displayName
        instanceGuid = objItem.instanceGuid
        pathProduct = objItem.pathToSignedProductExe
        pathReporting = objItem.pathToSignedReportingExe
        productState = str(objItem.productState)
        # ProductState conversion
        rtstatus, defstatus = securityProductState(productStateDict, productState)
        i+=1
        # Put in dict
        securityProductDict[i] = {'Name':productName, 'GUID':instanceGuid,
                                            'pathProduct':pathProduct, 'pathReporting':pathReporting,
                                            'productState':productState,
                                            'rtstatus':rtstatus, 'defstatus':defstatus,
                                            'Type':'Antispyware'}

    # 3 - Ecrire du fichier CSV
    header = ['Type', 'Name', 'GUID', 'pathProduct', 'pathReporting', 'productState', 'rtstatus', 'defstatus']
    csvFile = logFilePath + "security_products.csv"
    writer.writeCSV(csvFile, header, securityProductDict)
    
    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Security Products')
    htmltxt = htmltxt.replace('<TD>Outdated</TD>', '<TD id="ko">Outdated</TD>')
    htmltxt = htmltxt.replace('<TD>Up to date</TD>', '<TD id="ok">Up to date</TD>')
    htmltxt = htmltxt.replace('<TD>Disabled</TD>', '<TD id="ko">Disabled</TD>')
    htmltxt = htmltxt.replace('<TD>Enabled</TD>', '<TD id="ok">Enabled</TD>')

    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, str(len(securityProductDict)) + ' security products :\n')
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, '\n</div>\n')
    
    return securityProductDict

def processInfo(logFilePath):
    '''
    ~ tasklist /SVC
    ***FR**
    Liste les processus démarrés sur l'ordinateur
    Retourne la liste des processus
    **EN**
    List computer running processes
    Retourne the processes list
    '''
    writer.write('Getting running processes')
    # 1 - Ecriture début de log
    logFile = logFilePath + 'final.html'
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Running processes on computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des processus démarrés
    # i = 1
    procDict = {}
    for proc in psutil.process_iter():
        try:
            procDict[str(proc.name())] = {'Name':str(proc.name()), 'PID':str(proc.pid)}
            # i+=1
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
            
    # 3 - Ecrire du fichier CSV
    header = ['Name', 'PID']
    csvFile = logFilePath + 'processes.csv'
    writer.writeCSV(csvFile, header, procDict)
    
    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Running processes')

    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, str(len(procDict)) + ' running processes :\n')
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, '\n</div>\n')

    return procDict

def servicesInfo(logFilePath):
    '''
    ~ wmic service where started=true get name, startname
    **FR**
    Liste les services démarrés sur l'ordinateur
    Retourne la liste des noms des services démarrés
    **EN**
    List computer started services
    Return the running services names list
    '''
    writer.write('Getting running services')
    # 1 - Ecriture début de log   
    logFile = logFilePath + 'final.html'
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Running services on computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des services démarrés
    servicesList = psutil.win_service_iter()
    servicesDictRunning = {}
    for srv in servicesList:
        srvname = re.findall("name='(.*)',", str(srv))[0]
        serviceTemp = psutil.win_service_get(srvname)
        service = serviceTemp.as_dict()
        # Keep running services
        if service['status'] == 'running':
            servicesDictRunning[service['name']] = {'Name':service['name'], 'PID':service['pid'],
                                'Display_Name':service['display_name'], 'Start_Type':service['start_type'],
                                'Username':service['username'], 'Binpath':service['binpath']}

    # 3 - Ecriture du fichier CSV
    header = ['Name', 'Display_Name', 'PID', 'Start_Type', 'Username', 'Binpath']
    csvFile = logFilePath + "services.csv"
    writer.writeCSV(csvFile, header, servicesDictRunning)
    
    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Running services')

    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, str(len(servicesDictRunning)) + ' running services :<br>\n')
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, '\n</div>\n')

    return servicesDictRunning

def portsInfo(logFilePath):
    '''
    ~ netstat -a
    **FR**
    Liste les communications réseaux de l'ordinateur
    Retourne la liste des connexions réseaux
    **EN**
    List computer network connections
    Return the network connections list
    '''
    writer.write('Getting network connections')
    # 1 - Ecriture début de log
    logFile = logFilePath + "final.html"
    writer.writeLog(logFile, '<div><br>\n')
    elem = '<h2>Informations about network connections of computer "' + computername + '"</h2>'
    writer.prepaLogScan(logFile, elem)

    # 2 - obtenir la liste des connexions actives
    portsList = psutil.net_connections()
    portsDict = {}
    for ports in portsList:
        if ports.status == 'ESTABLISHED' or ports.status == 'CLOSE_WAIT':
            portsDict[str(ports.laddr)] = {'Local_addr':str(ports.laddr[0]), 'Local_port':str(ports.laddr[1]),
                                        'Remote_addr':str(ports.raddr[0]), 'Remote_port':str(ports.raddr[1]),
                                        'Status':str(ports.status), 'PID':str(ports.pid)}

        if ports.status == 'LISTEN':
            portsDict[str(ports.laddr)] = {'Local_addr':str(ports.laddr[0]), 'Local_port':str(ports.laddr[1]),
                                        'Remote_addr':'', 'Remote_port':'',
                                        'Status':str(ports.status), 'PID':str(ports.pid)}

    # 3 - Ecriture du fichier CSV
    header = ['Local_addr', 'Local_port', 'Remote_addr', 'Remote_port', 'Status', 'PID']
    csvFile = logFilePath + "ports.csv"
    writer.writeCSV(csvFile, header, portsDict)
    
    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csvFile, 'Network connections')
    
    # 5 - Ecriture de la fin du log
    writer.writeLog(logFile, str(len(portsDict)) + ' network connections :<br>\n')
    writer.writeLog(logFile, htmltxt)
    writer.writeLog(logFile, '\n</div>\n')

    return portsDict
  
if __name__ == '__main__':
    # softwareList = softwareInit()
    # softwareDict = searchSoftware(softwareList)
    # userInfo(r'E:\scanPC_dev_encours\TESTS\scans/')
    # userInfo(r'C:\STOCKAGE\logScanPC\2019\05\7/')
    # print(locale.getpreferredencoding())
    # usrfile = userInfo(r"E:\scanPC_dev_encours\TESTS\scans/")
    # pwdfile = pwdPolicy(r"E:\scanPC_dev_encours\TESTS\scans/")
    # sFfile = sharedFolders(r"C:\STOCKAGE\logScanPC\2019\05\7/")
    # hotFixesfile = hotFixesInfo(r"C:\STOCKAGE\logScanPC\2019\05\7/")
    # systelInfofile = systemInfo(r"C:\scanPC_dev_encours\TESTS\scans/")
    # systelInfofile = systemInfo(r'C:\STOCKAGE\logScanPC\2019\05\7/')
    securityProduct = securityProductInfo(r'C:\STOCKAGE\logScanPC\2019\05\7/')
    # processfile = processInfo(r"C:\STOCKAGE\logScanPC\2019\05\7/")
    # servicesfile = servicesInfo(r"C:\STOCKAGE\logScanPC\2019\05\7/")
    # portfile = portsInfo(r"C:\STOCKAGE\logScanPC\2019\05\7/")
    
# mcAfeeLog = str(logFilePath)+"mcAfeeLog.txt"