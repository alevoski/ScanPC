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
        if v < flag:
            min = flag - v
            flagTosave = v
    return min, flagTosave

def userInfo(logFile):
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

    # 3 - Ecriture début de log
    # logAllUsersTEMP = logFilePath + "users_TEMP.txt"
    elem = "*******Getting details about the "+str(len(usrLstTEMP2)) + " user(s) of computer ''" + computername + "''*******"
    writer.prepaLogScan(logFile, elem)

    # 4 - Détails sur les utilisateurs
    domain = os.environ['userdomain']
    # print(str(len(usrLstTEMP2)) + ' user(s) found on computer ' + computername)
    writer.writeLog(logFile, str(len(usrLstTEMP2)) + ' user(s) found on computer ' + computername + '\n')
    writer.writeLog(logFile, 'Domain : ' + domain + '\n')
    i = 1
    for user in usrLstTEMP2:
        writer.writeLog(logFile, '[' + str(i) + '] : ' + user + '\n')
        try:
            objOU = win32com.client.GetObject("WinNT://" + domain + "/" + user + ",user")
            #User account
            writer.writeLog(logFile, '****User account****\n')
            writer.writeLog(logFile, 'Full name : ' + str(objOU.Get('fullname')) + '\n')
            # print('Description : ' + str(objOU.Get('Description')))
            # print('Account Disabled ? : ' + str(objOU.Get('AccountDisabled')))
            # print('Account active ? : ' + str(objOU.Get('active')))
            try:
                writer.writeLog(logFile, 'Expiration date : ' + str(objOU.Get('AccountExpirationDate')) + '\n')
            except Exception as e:
                writer.writeLog(logFile, 'Expiration date : N/A\n')
            # print('Password Last Changed : ' + str(objOU.Get('PasswordLastChanged')))
            # print('Account Locked ? : ' + str(objOU.Get('IsAccountLocked')))
            writer.writeLog(logFile, 'Profile : ' + str(objOU.Get('Profile')) + '\n')
            writer.writeLog(logFile, 'Login script : ' + str(objOU.Get('LoginScript')) + '\n')
            writer.writeLog(logFile, 'Last login : ' + str(objOU.Get('lastlogin')) + '\n')
            # print('ObjectSID : ' + str(objOU.Get('ObjectSID')))
            # print(objOU.Get('Parameters').encode('iso-8859-1', 'ignore'))
            writer.writeLog(logFile, 'Primary group ID : ' + str(objOU.Get('PrimaryGroupID')) + '\n')
            writer.writeLog(logFile, 'Auto unlock interval : ' + str(objOU.Get('AutoUnlockInterval')) + ' secs\n')
            writer.writeLog(logFile, 'Lockout observation interval : ' + str(objOU.Get('LockoutObservationInterval')) + ' secs\n')
            # print('LoginHours : ' + str(objOU.Get('LoginHours')))
            writer.writeLog(logFile, 'Home directory : ' + str(objOU.Get('HomeDirectory')) + '\n')
            writer.writeLog(logFile, 'Home dir drive : ' + str(objOU.Get('HomeDirDrive')) + '\n')
            
            #Password
            writer.writeLog(logFile, '****Password****\n')
            # print('UserCannotChangePassword : ' + str(objOU.Get('UserCannotChangePassword')))
            writer.writeLog(logFile, 'Age : ' + str(round(objOU.Get('PasswordAge')/3600/24)) + ' days\n')
            # print('Password last change date : ' + str())
            # print('Password Age : ' + str(objOU.Get('PasswordAgeDate')))
            writer.writeLog(logFile, 'Min age : ' + str(round(objOU.Get('MinPasswordAge')/3600/24)) + ' days\n')
            writer.writeLog(logFile, 'Max age : ' + str(round(objOU.Get('MaxPasswordAge')/3600/24)) + ' days\n')
            # print('Password expiration date : ' + str(objOU.Get('PasswordExpirationDate')))
            writer.writeLog(logFile, 'Expired ? : ' + str(objOU.Get('PasswordExpired')) + '\n')
            writer.writeLog(logFile, 'Max bad passwords allowed : ' + str(objOU.Get('MaxBadPasswordsAllowed')) + '\n')
            writer.writeLog(logFile, 'Min length : ' + str(objOU.Get('MinPasswordLength')) + '\n')
            writer.writeLog(logFile, 'History length : ' + str(objOU.Get('PasswordHistoryLength')) + '\n')
            
            #Groups
            writer.writeLog(logFile,'****Groups****\n')
            # print('NAME      |       DESCRIPTION')
            for grp in objOU.Groups():
                writer.writeLog(logFile, grp.Name + '\n') # + '      |       ' + grp.Description)
            
            #Flag
            writer.writeLog(logFile, '****Flag****\n')
            flag = objOU.Get('UserFlags')
            writer.writeLog(logFile, 'User flag : ' + str(flag) + '\n')
            #Get flags with user flag (user flag = sum(flags))
            flagList = []
            min, flagTosave = calcFlag(userFlagsDict, flag)
            flagList.append(flagTosave)
            while min != 0:
                min, flagTosave = calcFlag(userFlagsDict, min)
                flagList.append(flagTosave)
            
            #Verify if flag = sum(flagList)
            if flag == sum(flagList):
                writer.writeLog(logFile, str(flag) + ' = ' + str(flagList) + '\n')
                
                #Get flags description with flag nums saved
                writer.writeLog(logFile, 'This flag means the user account has this properties :\n')
                for k, v in userFlagsDict.items():
                    if v in flagList:
                        writer.writeLog(logFile, str(v) + ' : ' + str(k) + '\n')
            writer.writeLog(logFile, '*' * 20 + '\n')
            
        except Exception as e:
            # print(e)
            pass
        i+=1
        
    # 5 - Ecriture de la fin du log
    elem = "----------------------- Users listing ended ------------------------"
    writer.prepaLogScan(logFile, elem)
    
    return logFile

def sharedFolders(logFile):
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
    # logSFTEMP = str(logFilePath)+"SF_TEMP.txt"
    elem = "**** Shared folders of computer ''" + computername + "''****"
    writer.prepaLogScan(logFile, elem)
    
    # 2 - Obtenir la liste des dossiers partagés de l'ordinateur
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_Share")
    i = 1
    for objItem in colItems:
        writer.writeLog(logFile, '[' + str(i) + ']Name : ' + str(objItem.Name) + '\n')
        writer.writeLog(logFile, 'Path : ' + str(objItem.Path) + '\n')
        writer.writeLog(logFile, 'Caption : ' + str(objItem.Caption) + '\n')
        writer.writeLog(logFile, 'Description : ' + str(objItem.Description) + '\n')
        i+=1
        
    # 3 - Ecriture de la fin du log
    elem = "------------------- Shared folders listing ended -------------------"
    writer.prepaLogScan(logFile, elem)
    
    return logFile

def hotFixesInfo(logFile):
    '''
    https://www.activexperts.com/admin/scripts/wmi/python/0417/
    ~ wmic qfe get HotfixID,InstalledOn | more
    **FR**
    Liste les patchs Windows installés
    Retourne le log généré
    **EN**
    List Windows updates installed
    Return generated log
    '''
    writer.write('Getting Windows updates')
    # 1 - Ecriture début de log
    # logHotFix = str(logFilePath) + "TEMP_hotFixes.txt"
    elem = "**** Windows updates of computer ''" + computername + "''****"
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des correctifs de l'ordinateur
    strComputer = "."
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_QuickFixEngineering")
    writer.writeLog(logFile, 'HotFixID' + '   ' + 'InstalledOn' + '\n')
    for objItem in colItems:
        # print(objItem.HotFixID + ' ' + str(datetime.strptime(str(objItem.InstalledOn), "%m/%d/%Y").strftime("%Y-%m-%d")))
        writer.writeLog(logFile, objItem.HotFixID + '  ' + str(datetime.strptime(str(objItem.InstalledOn), "%m/%d/%Y").strftime("%Y-%m-%d")) + '\n')

    # 3 - Ecriture fin de log
    elem = "------------------- Windows updates listing ended -------------------"
    writer.prepaLogScan(logFile, elem)

    return logFile

def systemInfo(logFile):
    '''
    ~ systeminfo | find /V /I "hotfix" | find /V "KB"
    ~ wmic logicaldisk get volumename, description, FileSystem, Caption, ProviderName
    ~ netsh advfirewall show all state        ||          netsh firewall show state
    **FR**
    Scan le système (os, bios, cpu, ram, cartes réseaux, disques durs, parefeu, etc)
    Retourne le log généré
    **EN**
    Scan system (os, bios, cpu, ram, network interfaces, drives, firewall, etc)
    Return generated log
    '''
    writer.write('Getting system informations')
    # logFile = str(logFilePath) + "SysInfo_TEMP.txt"
    delimiter = '*' * 40

    # 1 - Ecriture début de log
    elem = "*************************** System information of computer ''" + computername + "''***************************"
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
    writer.writeLog(logFile, 'Hostname : ' + platform.node() + '\n')
    # print('OS : ' + platform.system() + ' ' + platform.version())
    writer.writeLog(logFile, 'OS : ' + dictToSave.get('ProductName') + '\n')
    writer.writeLog(logFile, 'OS type : ' + platform.machine() + '\n')
    writer.writeLog(logFile, 'Product Id : ' + dictToSave.get('ProductId') + '\n')
    writer.writeLog(logFile, 'Install Date: ' + str(datetime.fromtimestamp(dictToSave.get('InstallDate'))) + '\n')
    writer.writeLog(logFile, 'System Root: ' + dictToSave.get('SystemRoot') + '\n')
    
    # Language
    bufSize = 32
    buf = ctypes.create_unicode_buffer(bufSize)
    dwFlags = 0
    lcid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    ctypes.windll.kernel32.LCIDToLocaleName(lcid, buf, bufSize, dwFlags)
    writer.writeLog(logFile, 'Regional and language options : ' + buf.value + '\n')
    
    # Time zone
    writer.writeLog(logFile, 'Time zone : ' + time.tzname[0] + '\n')
    writer.writeLog(logFile, delimiter + '\n')
    
    # Owner
    writer.writeLog(logFile, 'Registered Organization : ' + dictToSave.get('RegisteredOrganization') + '\n')
    writer.writeLog(logFile, 'Registered Owner : ' + dictToSave.get('RegisteredOwner') + '\n')
    writer.writeLog(logFile, delimiter + '\n')
    
    # Computer model/brand
    strComputer = "."
    objWMIService = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    objSWbemServices = objWMIService.ConnectServer(strComputer, 'root\cimv2')
    colItems = objSWbemServices.ExecQuery('SELECT * FROM Win32_ComputerSystem')
    for objItem in colItems:
        writer.writeLog(logFile, 'Manufacturer : ' + objItem.Manufacturer + '\n')
        try:
            writer.writeLog(logFile, 'SystemFamily : ' + objItem.SystemFamily + '\n')
        except AttributeError as e:
            writer.writeLog(logFile, 'SystemFamily : N/A\n')
        writer.writeLog(logFile, 'Model : ' + objItem.Model + '\n')
        writer.writeLog(logFile, 'Type : ' + objItem.SystemType + '\n')
    writer.writeLog(logFile, delimiter + '\n')

    # Processor
    colItems = objSWbemServices.ExecQuery('Select * from Win32_Processor')
    for objItem in colItems:
        writer.writeLog(logFile, 'Processor : ' + objItem.Name + '\n')
        writer.writeLog(logFile, 'Core : ' + str(objItem.NumberOfCores) + '\n')
        writer.writeLog(logFile, 'Logical core : ' + str(objItem.NumberOfLogicalProcessors) + '\n')
        try:
            writer.writeLog(logFile, 'Virtualization enabled : ' + str(objItem.VirtualizationFirmwareEnabled) + '\n')
        except AttributeError as e:
            writer.writeLog(logFile, 'Virtualization enabled : N/A\n')
    writer.writeLog(logFile, delimiter + '\n')
    
    # Bios
    colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BIOS")
    for objItem in colItems:
        writer.writeLog(logFile, "BIOS Version : " + str(objItem.BIOSVersion) + '\n')
        writer.writeLog(logFile, "Release Date : " + str(datetime.strptime(str(objItem.ReleaseDate[:8]), "%Y%m%d").strftime("%Y-%m-%d")) + '\n')
    writer.writeLog(logFile, delimiter + '\n')

    # Memory
    mem = psutil.virtual_memory()
    writer.writeLog(logFile, 'Total memory : ' + str(mem.total >> 20) + ' Mo\n')
    writer.writeLog(logFile, 'Available memory : ' + str(mem.available >> 20) + ' Mo\n')
    writer.writeLog(logFile, 'Used memory : ' + str(mem.used >> 20) + ' Mo (' + str(mem.percent) + ' %)\n')
    writer.writeLog(logFile, 'Free memory : ' + str(mem.free >> 20) + ' Mo\n')
    writer.writeLog(logFile, delimiter + '\n')
    
    # Domaine
    writer.writeLog(logFile, "Domain : " + os.environ['userdomain'] + '\n')
    writer.writeLog(logFile, delimiter + '\n')
    
    #Network interfaces
    network_interface = psutil.net_if_addrs()
    # print(str(len(network_interface)) + ' network interface(s) found\n')
    writer.writeLog(logFile, str(len(network_interface)) + ' network interfaces found\n')
    i = 1
    for k, v in network_interface.items():
        # print('[' + str(i) + '] : ' + k)
        writer.writeLog(logFile, '[' + str(i) + '] : ' + k + '\n')
        for items in v:
            regex = re.findall("address='(.*?)'", str(items))[0]
            # print(regex)
            writer.writeLog(logFile, regex + '\n')
        # print('\n')
        writer.writeLog(logFile, '\n')
        i+=1
    writer.writeLog(logFile, delimiter + '\n')
    
    # Drives
    drives = psutil.disk_partitions()
    # print(str(len(drives)) + ' drive(s) found\n')
    writer.writeLog(logFile, str(len(drives)) + ' drive(s) found\n')
    # print(drives)
    i = 1
    for drive in drives:
        drivePath = re.findall("device='(.*?)'", str(drive))[0]
        # print('[' + str(i) + '] : ' + drivePath)
        writer.writeLog(logFile, '[' + str(i) + '] : ' + drivePath + '\n')
        # print(str(psutil.disk_usage(drivePath)) + '\n')
        # print('opts', drive.opts)
        # print('fstype', drive.fstype)
        if drive.opts != 'cdrom':
            try:
                diskUsage = psutil.disk_usage(drivePath)
                writer.writeLog(logFile, 'Total space : ' + str(diskUsage.total >> 30) + ' Go\n')
                writer.writeLog(logFile, 'Free space : ' + str(diskUsage.free >> 30) + ' Go\n')
                writer.writeLog(logFile, 'Used space : ' + str(diskUsage.used >> 30) + ' Go (' + str(diskUsage.percent) + ' %)\n')
            except Exception:
                pass
        i+=1
    writer.writeLog(logFile, delimiter + '\n')
    
    # Firewalls
    XPFW = win32com.client.gencache.EnsureDispatch('HNetCfg.FwMgr', 0)
    XPFW_policy = XPFW.LocalPolicy.CurrentProfile
    writer.writeLog(logFile, 'Firewall enabled : ' + str(XPFW_policy.FirewallEnabled) + '\n')

    # 3 - Ecriture fin de log
    elem = "------------------- System informations listing ended -------------------"
    writer.prepaLogScan(logFile, elem)

    return logFile

def processInfo(logFile):
    '''
    ~ tasklist /SVC
    ***FR**
    Liste les processus démarrés sur l'ordinateur
    Retourne le log généré
    **EN**
    List computer running processes
    Retourne generated log
    '''
    writer.write('Getting running processes')
    # 1 - Ecriture début de log
    # logFile = str(logFilePath)+"TEMP_proc.txt"
    elem = "***Running processes on computer ''" + computername + "''***"
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des processus démarrés
    for proc in psutil.process_iter():
        try:
            writer.writeLog(logFile, str(proc.name()) + ' ::: ' + str(proc.pid) + '\n')
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    # 3 - Ecriture de la fin du log
    elem = "------------------- Running processes listing finished -------------------"
    writer.prepaLogScan(logFile, elem)

    return logFile

def servicesInfo(logFile):
    '''
    ~ wmic service where started=true get name, startname
    **FR**
    Liste les services démarrés sur l'ordinateur
    Retourne le log généré
    **EN**
    List computer started services
    Return generated log
    '''
    writer.write('Getting running services')
    # 1 - Ecriture début de log   
    # logFile = logFilePath + "TEMP_services.txt"
    elem = "***Running services on computer ''" + computername + "''***"
    writer.prepaLogScan(logFile, elem)

    # 2 - Obtenir la liste des services démarrés
    servicesList = psutil.win_service_iter()
    servicesListRunning = []
    i = 1
    for srv in servicesList:
        srvname = re.findall("name='(.*)',", str(srv))[0]
        serviceTemp = psutil.win_service_get(srvname)
        service = serviceTemp.as_dict()
        # Remove description
        del service['description']
        # Keep running services
        if service['status'] == 'running':
            # print('[' + str(i) + ']')
            # print(service)
            servicesListRunning.append(str(service['name']))
            writer.writeLog(logFile, '[' + str(i) + ']' + '\n')
            for k, v in service.items():
                writer.writeLog(logFile, str(k) + ':' + str(v) + '\n')
            i+=1

    # 3 - Ecriture de la fin du log
    elem = "------------------- Running services listing finished -------------------"
    writer.prepaLogScan(logFile, elem)

    return servicesListRunning

def portsInfo(logFile):
    '''
    ~ netstat -a
    **FR**
    Liste les communications réseaux de l'ordinateur
    Retourne le log généré
    **EN**
    List computer network connections
    Return generated log
    '''
    writer.write('Getting network connections')
    # 1 - Ecriture début de log
    # logFile = str(logFilePath) + "TEMP_ports.txt"
    elem = "***Informations about network connections of computer ''" + computername + "''***"
    writer.prepaLogScan(logFile, elem)

    # 2 - obtenir la liste des connexions actives
    portsList = psutil.net_connections()
    i = 1
    for ports in portsList:
        # if ports.status == 'ESTABLISHED' or ports.status == 'LISTEN':
        # print('[' + str(i) + ']' + ' laddr:' + str(ports.laddr) 
            # + '; raddr:' + str(ports.raddr) 
            # + '; status:' + str(ports.status) 
            # + '; pid:' + str(ports.pid))
        writer.writeLog(logFile, '[' + str(i) + ']')
        writer.writeLog(logFile, ' laddr:' + str(ports.laddr) 
                        + '; raddr:' + str(ports.raddr) 
                        + '; status:' + str(ports.status) 
                        + '; pid:' + str(ports.pid) + '\n')
        i+=1

    # 3 - Ecriture de la fin du log
    elem = "------------------- Computer network connections listing finished -------------------"
    writer.prepaLogScan(logFile, elem)

    return logFile  
  
if __name__ == '__main__':
    # softwareList = softwareInit()
    # softwareDict = searchSoftware(softwareList)
    # userInfo(r'E:\scanPC_dev_encours\TESTS\scans/')
    # userInfo(r'C:\STOCKAGE\logScanPC\2019\4\7/')
    # print(locale.getpreferredencoding())
    # usrfile = userInfo(r"E:\scanPC_dev_encours\TESTS\scans/")
    # pwdfile = pwdPolicy(r"E:\scanPC_dev_encours\TESTS\scans/")
    # sFfile = sharedFolders(r"C:\STOCKAGE\logScanPC\2019\4\7/")
    # hotFixesfile = hotFixesInfo(r"C:\STOCKAGE\logScanPC\2019\04\7/testhotfixes.txt")
    # systelInfofile = systemInfo(r"C:\scanPC_dev_encours\TESTS\scans/")
    # systelInfofile = systemInfo(r'C:\STOCKAGE\logScanPC\2019\04/')
    # processfile = processInfo(r"C:\scanPC_dev_encours\TESTS\scans/")
    # servicesfile = servicesInfo(r"C:\STOCKAGE\logScanPC\2019\04\7/testservices.txt")
    portfile = portsInfo(r"C:\STOCKAGE\logScanPC\2019\04\7/testports.txt")
    
# mcAfeeLog = str(logFilePath)+"mcAfeeLog.txt"