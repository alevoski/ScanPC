#!/usr/bin/python3
# -*- coding: utf-8 -*-
#ScanPC - scans.py
#@Alexandre Buissé - 2019/2020

'''
Scanning module : performs different scans on the computer
'''

#Standard imports
import os
import re
from datetime import datetime
import platform
import ctypes
import time
import winreg

#Third party imports
import psutil
from pywintypes import com_error
import win32com.client

#Project modules imports
import writer

COMPUTERNAME = os.environ['COMPUTERNAME']
USER_FLAGS_DICT = {'SCRIPT':1,
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
PRODUCT_STATE_DICT = {'262144':{'defstatus':'Up to date', 'rtstatus':'Disabled'},
                      '266240':{'defstatus':'Up to date', 'rtstatus':'Enabled'},
                      '262160':{'defstatus':'Outdated', 'rtstatus':'Disabled'},
                      '266256':{'defstatus':'Outdated', 'rtstatus':'Enabled'},
                      # '270336'
                      # '327680'
                      # '327696'
                      '331776':{'defstatus':'Up to date', 'rtstatus':'Enabled'},
                      # '335872'
                      # '344064'
                      '393216':{'defstatus':'Up to date', 'rtstatus':'Disabled'},
                      '393232':{'defstatus':'Outdated', 'rtstatus':'Disabled'},
                      '393488':{'defstatus':'Outdated', 'rtstatus':'Disabled'},
                      '397312':{'defstatus':'Up to date', 'rtstatus':'Enabled'},
                      '397328':{'defstatus':'Outdated', 'rtstatus':'Enabled'},
                      '393472':{'defstatus':'Up to date', 'rtstatus':'Disabled'},
                      '397584':{'defstatus':'Outdated', 'rtstatus':'Enabled'},
                      '397568':{'defstatus':'Up to date', 'rtstatus':'Enabled'},
                      # '458768'
                      # '458752'
                      '462864':{'defstatus':'Outdated', 'rtstatus':'Enabled'},
                      '462848':{'defstatus':'Up to date', 'rtstatus':'Enabled'}}

def detect_os():
    '''
    **FR**
    Retourne l'OS utilisé
    **EN**
    Return used OS
    '''
    computer_platform = str(platform.platform().lower())
    # print(computer_platform)
    if 'xp' in computer_platform or '2003' in computer_platform:
        rep = 'xp'
    if '7' in computer_platform or 'vista' in computer_platform or '2008' in computer_platform or '2012' in computer_platform:
        rep = '7'
    if '10' in computer_platform or '2016' in computer_platform:
        rep = '10'
    return rep

def calc_flag(flag):
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
    flag_to_save = 0
    mini = 0
    for _, value in USER_FLAGS_DICT.items():
        if value == flag:
            mini = flag - value
            flag_to_save = value
            break
        if value < flag and value > flag_to_save:
            mini = flag - value
            flag_to_save = value
    return mini, flag_to_save

def user_info(log_file_path):
    '''
    ~ net user <username> /domain    &    net user <username>
    **FR**
    Obtient la liste des utilisateurs AD et locaux de l'ordinateur avec la base de registre et WinNT
    Retourne le log généré
    **EN**
    Get AD and local users list of the computer with regedit and WinNT
    Return generated log
    '''
    writer.writer('Getting computer users')
    hive = winreg.HKEY_LOCAL_MACHINE
    key_user = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" # AD users
    registry_key = None
    temp_user_list = []
    temp2_user_list = []
    profil_path = "ProfileImagePath"

    # 1 - Liste tous les profils
    try:
        registry_key = winreg.OpenKey(hive, key_user, 0, winreg.KEY_READ)
        # print("ok")
    except FileNotFoundError:
        # print("oups")
        pass
    # print(registry_key)
    if registry_key is not None:
        i = 0
        while True:
            try:
                temp_user_list.append(winreg.EnumKey(registry_key, i)) # ok pour 1 clé
                i += 1
                # print("value :", str(registry_key))
            except Exception:
                # print(Exception)
                winreg.CloseKey(registry_key)
                break

    # 2 - Tri de la liste pour ne garder que les SID commençant par "1-5-21" et récupérer le nom des utilisateurs
    for item in temp_user_list:
        if "1-5-21" in item:
            # print(item)
            key = key_user + '\\' + item
            # print(key)
            try:
                registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ)
                value, _ = winreg.QueryValueEx(registry_key, profil_path)
                winreg.CloseKey(registry_key)
                # print(value)
                value_mod = re.findall(r'[^\\]+\s*$', value)
                # print(type(value_mod[0]))
                temp2_user_list.append(value_mod[0])
            except FileNotFoundError:
                # print("oups")
                pass

    domain = os.environ['userdomain']
    # 3 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Getting details about the ' + str(len(temp2_user_list)) + ' user(s) of computer "' + COMPUTERNAME + '"</h2>'
    elem2 = "<br>Domain : " + str(domain) + '\n'
    writer.prepa_log_scan(log_file, elem)
    writer.writelog(log_file, elem2)

    # 4 - Détails sur les utilisateurs
    user_dict = {}
    for user in temp2_user_list:
        try:
            obj_ou = win32com.client.GetObject("WinNT://" + domain + "/" + user + ",user")
            #User account
            fullname = str(obj_ou.Get('fullname'))
            # print(fullname)
            try:
                expiration_date = str(obj_ou.Get('AccountExpirationDate'))
            except Exception:
                expiration_date = 'N/A'
            profile = str(obj_ou.Get('Profile'))
            login_script = str(obj_ou.Get('LoginScript'))
            try:
                last_login = str(obj_ou.Get('lastlogin'))
            except com_error:
                last_login = 'N/A'
            primary_gid = str(obj_ou.Get('PrimaryGroupID'))
            auto_unlock_interval = str(obj_ou.Get('AutoUnlockInterval'))
            lockout_observation_interval = str(obj_ou.Get('LockoutObservationInterval'))
            homedir = str(obj_ou.Get('HomeDirectory'))
            homedir_drive = str(obj_ou.Get('HomeDirDrive'))

            #Password
            pwd_age = str(round(obj_ou.Get('PasswordAge')/3600/24))
            pwd_min_age = str(round(obj_ou.Get('MinPasswordAge')/3600/24))
            pwd_max_age = str(round(obj_ou.Get('MaxPasswordAge')/3600/24))
            pwd_expired = str(obj_ou.Get('PasswordExpired'))
            pwd_max_bad_pwd_allowed = str(obj_ou.Get('MaxBadPasswordsAllowed'))
            pwd_min_length = str(obj_ou.Get('MinPasswordLength'))
            pwd_history_length = str(obj_ou.Get('PasswordHistoryLength'))

            #Groups
            groups_list = []
            for grp in obj_ou.Groups():
                groups_list.append(grp.Name + '<br>')

            #Flag
            flag_final_list = []
            flag = obj_ou.Get('UserFlags')
            #Get flags with user flag (user flag = sum(flags))
            flag_list = []
            mini, flag_to_save = calc_flag(flag)
            flag_list.append(flag_to_save)
            while mini != 0:
                mini, flag_to_save = calc_flag(mini)
                flag_list.append(flag_to_save)
            #Verify if flag = sum(flag_list)
            if flag == sum(flag_list):
                flag_final_list.append(str(flag) + ' => ' + str(flag_list) + '<br>')
                #Get flags description with flag nums saved
                flag_final_list.append('Flag properties :')
                for k, value in USER_FLAGS_DICT.items():
                    if value in flag_list:
                        flag_final_list.append('<br>' + str(value) + ' : ' + str(k))
            user_dict[user] = {'User':user, 'Fullname':fullname, 'Expiration_Date':expiration_date,
                               'Profile':profile, 'Login_Script':login_script, 'Last_Login':last_login,
                               'Primary_Group_ID':primary_gid, 'Auto_Unlock_Interval (secs)': auto_unlock_interval,
                               'Lockout_Observation_Interval (secs)':lockout_observation_interval,
                               'HomeDir':homedir, 'HomeDirDrive':homedir_drive,
                               'pwd_age (days)':pwd_age, 'pwd_min_age (days)':pwd_min_age, 'pwd_max_age (days)':pwd_max_age,
                               'pwd_expired':pwd_expired, 'pwd_max_bad_pwd_allowed':pwd_max_bad_pwd_allowed,
                               'pwd_min_length':pwd_min_length, 'pwd_history_length':pwd_history_length,
                               'Groups':(''.join(groups_list)), 'Flag':(''.join(flag_final_list))}

        except Exception:
            pass
        i += 1

    if user_dict != '':
        # Ecriture du fichier CSV
        header = ['User', 'Fullname', 'Expiration_Date', 'Profile', 'Login_Script', 'Last_Login',
                  'Primary_Group_ID', 'Auto_Unlock_Interval (secs)', 'Lockout_Observation_Interval (secs)',
                  'HomeDir', 'HomeDirDrive', 'pwd_age (days)', 'pwd_min_age (days)', 'pwd_max_age (days)',
                  'pwd_expired', 'pwd_max_bad_pwd_allowed', 'pwd_min_length', 'pwd_history_length', 'Groups', 'Flag']
        csv_file = log_file_path + "users.csv"
        writer.write_csv(csv_file, header, user_dict)
        # Transformation du CSV en HTML
        htmltxt = writer.csv2html(csv_file, 'Users details')
        writer.writelog(log_file, htmltxt)

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, '\n</div>\n')

    return log_file

def shared_folders_info(log_file_path):
    '''
    ~ net share
    **FR**
    Liste les dossiers partagés
    Retourne le log généré
    **EN**
    List shared folders
    Return generated log
    '''
    writer.writer('Getting shared folders')
    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Shared folders of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtenir la liste des dossiers partagés de l'ordinateur
    str_computer = '.'
    obj_wmi_service = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    obj_sw_bem_services = obj_wmi_service.ConnectServer(str_computer, 'root\\cimv2')
    col_items = obj_sw_bem_services.ExecQuery('SELECT * FROM Win32_Share')
    shared_folders_dict = {}
    for obj_item in col_items:
        shared_folders_dict[str(obj_item.Name)] = {'Name':str(obj_item.Name), 'Path':str(obj_item.Path),
                                                   'Caption':str(obj_item.Caption), 'Description':str(obj_item.Description)}

    # 3 - Ecriture du fichier CSV
    header = ['Name', 'Path', 'Caption', 'Description']
    csv_file = log_file_path + "shared_folders_info.csv"
    writer.write_csv(csv_file, header, shared_folders_dict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Shared Folders')
    writer.writelog(log_file, htmltxt)

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, '\n</div>\n')

    return log_file

def hotfixes_info(log_file_path):
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
    writer.writer('Getting Windows security updates')
    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Windows security updates of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtenir la liste des correctifs de l'ordinateur
    str_computer = '.'
    obj_wmi_service = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    obj_sw_bem_services = obj_wmi_service.ConnectServer(str_computer, 'root\\cimv2')
    col_items = obj_sw_bem_services.ExecQuery("SELECT * FROM Win32_QuickFixEngineering WHERE Description='Security Update'")
    hotfix_dict = {}
    for obj_item in col_items:
        try:
            hotfix_dict[obj_item.HotFixID] = {'HotFixID':obj_item.HotFixID,
                                              'InstalledOn':str(datetime.strptime(str(obj_item.InstalledOn), "%m/%d/%Y").strftime("%Y-%m-%d"))}
        except ValueError:
            hotfix_dict[obj_item.HotFixID] = {'HotFixID':obj_item.HotFixID,
                                              'InstalledOn':str(obj_item.InstalledOn)}

    # 3 - Ecriture du fichier CSV
    header = ['HotFixID', 'InstalledOn']
    csv_file = log_file_path + "hotfixes.csv"
    writer.write_csv(csv_file, header, hotfix_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Hotfixes')

    # 5 - Ecriture fin de log
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')

    return hotfix_dict

def system_info(log_file_path):
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
    writer.writer('Getting system informations')
    delimiter = '*' * 40

    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>System information of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Get information from regedit
    hive = winreg.HKEY_LOCAL_MACHINE
    key = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion" # W7-64/32
    registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) # W7-64/32

    # 2' - Specific to W10 (ok for 64 bits)
    if detect_os() == '10':
        writer.writer('Perform operations for W10+ OS')
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
            except Exception:
                break
            i += 1

    # Continuing
    tosave_dict = {'InstallDate':'N/A', 'SystemRoot':'N/A', 'ProductId':'N/A',
                   'ProductName':'N/A', 'RegisteredOrganization':'N/A', 'RegisteredOwner':'N/A'}
    to_get_list = ['InstallDate', 'SystemRoot', 'ProductId',
                   'ProductName', 'RegisteredOrganization', 'RegisteredOwner']
    i = 0
    while 1:
        try:
            name, value, _ = winreg.EnumValue(registry_key, i)
            if name in to_get_list:
                # print(repr(name), value)
                tosave_dict[name] = value
        except Exception:
            # print('error : ' + name)
            break
        i += 1
    # print(tosave_dict)

    # OS
    writer.writelog(log_file, 'Hostname : ' + platform.node() + '<br>\n')
    # print('OS : ' + platform.system() + ' ' + platform.version())
    writer.writelog(log_file, 'OS : ' + tosave_dict.get('ProductName') + '<br>\n')
    writer.writelog(log_file, 'OS type : ' + platform.machine() + '<br>\n')
    writer.writelog(log_file, 'Product Id : ' + tosave_dict.get('ProductId') + '<br>\n')
    writer.writelog(log_file, 'Install Date: ' + str(datetime.fromtimestamp(tosave_dict.get('InstallDate'))) + '<br>\n')
    writer.writelog(log_file, 'System Root: ' + tosave_dict.get('SystemRoot') + '<br>\n')

    # Language (Windows Vista or further)
    if detect_os() != 'xp':
        buf_size = 32
        buf = ctypes.create_unicode_buffer(buf_size)
        dw_flags = 0
        lcid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
        ctypes.windll.kernel32.LCIDToLocaleName(lcid, buf, buf_size, dw_flags)
        writer.writelog(log_file, 'Regional and language options : ' + buf.value + '<br>\n')

    # Time zone
    writer.writelog(log_file, 'Time zone : ' + time.tzname[0] + '<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Owner
    writer.writelog(log_file, 'Registered Organization : ' + tosave_dict.get('RegisteredOrganization') + '<br>\n')
    writer.writelog(log_file, 'Registered Owner : ' + tosave_dict.get('RegisteredOwner') + '<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Computer model/brand
    str_computer = '.'
    obj_wmi_service = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    obj_sw_bem_services = obj_wmi_service.ConnectServer(str_computer, 'root\\cimv2')
    col_items = obj_sw_bem_services.ExecQuery('SELECT * FROM Win32_ComputerSystem')
    for obj_item in col_items:
        writer.writelog(log_file, 'Manufacturer : ' + obj_item.Manufacturer + '<br>\n')
        try:
            writer.writelog(log_file, 'SystemFamily : ' + obj_item.SystemFamily + '<br>\n')
        except AttributeError:
            writer.writelog(log_file, 'SystemFamily : N/A<br>\n')
        writer.writelog(log_file, 'Model : ' + obj_item.Model + '<br>\n')
        writer.writelog(log_file, 'Type : ' + obj_item.SystemType + '<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Processor
    col_items = obj_sw_bem_services.ExecQuery('Select * from Win32_Processor')
    for obj_item in col_items:
        writer.writelog(log_file, 'Processor : ' + obj_item.Name + '<br>\n')
        writer.writelog(log_file, 'Core : ' + str(obj_item.NumberOfCores) + '<br>\n')
        writer.writelog(log_file, 'Logical core : ' + str(obj_item.NumberOfLogicalProcessors) + '<br>\n')
        try:
            writer.writelog(log_file, 'Virtualization enabled : ' + str(obj_item.VirtualizationFirmwareEnabled) + '<br>\n')
        except AttributeError:
            writer.writelog(log_file, 'Virtualization enabled : N/A<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Bios
    col_items = obj_sw_bem_services.ExecQuery("SELECT * FROM Win32_BIOS")
    for obj_item in col_items:
        writer.writelog(log_file, "BIOS Version : " + str(obj_item.BIOSVersion) + '<br>\n')
        writer.writelog(log_file, "Release Date : " + str(datetime.strptime(str(obj_item.ReleaseDate[:8]), "%Y%m%d").strftime("%Y-%m-%d")) + '<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Memory
    mem = psutil.virtual_memory()
    writer.writelog(log_file, 'Total memory : ' + str(mem.total >> 20) + ' Mo<br>\n')
    writer.writelog(log_file, 'Available memory : ' + str(mem.available >> 20) + ' Mo<br>\n')
    writer.writelog(log_file, 'Used memory : ' + str(mem.used >> 20) + ' Mo (' + str(mem.percent) + ' %)<br>\n')
    writer.writelog(log_file, 'Free memory : ' + str(mem.free >> 20) + ' Mo<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Domaine
    writer.writelog(log_file, 'Domain : ' + os.environ['userdomain'] + '<br>\n')
    writer.writelog(log_file, delimiter + '<br>\n')

    # Network interfaces
    str_computer = '.'
    obj_wmi_service = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    obj_sw_bem_services = obj_wmi_service.ConnectServer(str_computer, 'root\\cimv2')
    col_items = obj_sw_bem_services.ExecQuery('select * from Win32_NetworkAdapterConfiguration')
    network_dict = {}
    for obj_item in col_items:
        if obj_item.MACAddress is not None:
            net_caption = re.findall("] (.*)", obj_item.Caption)[0]
            # print(net_caption, obj_item.IPAddress)
            try:
                dns_srv = '<br>'.join(obj_item.DNSServerSearchOrder)
            except TypeError:
                dns_srv = obj_item.DNSServerSearchOrder
            try:
                ipv4 = obj_item.IPAddress[0]
            except IndexError:
                ipv4 = ''
            except TypeError:
                ipv4 = ''
            try:
                ipv6 = obj_item.IPAddress[1]
            except IndexError:
                ipv6 = ''
            except TypeError:
                ipv6 = ''
            try:
                dgw = ''.join(obj_item.DefaultIPGateway)
            except TypeError:
                dgw = obj_item.DefaultIPGateway
            try:
                dhcp_srv = ''.join(obj_item.DHCPServer)
            except TypeError:
                dhcp_srv = obj_item.DHCPServer
            network_dict[net_caption] = {'Name':net_caption,
                                         'MAC':obj_item.MACAddress,
                                         'Enabled':obj_item.IPEnabled,
                                         'IpV4':ipv4,
                                         'IpV6':ipv6,
                                         'Default Gateway':dgw,
                                         'DHCP Enabled':obj_item.DHCPEnabled,
                                         'DHCP Server':dhcp_srv,
                                         'WINS Primary Server':obj_item.WINSPrimaryServer,
                                         'WINS Secondary Server':obj_item.WINSSecondaryServer,
                                         'DNS Domain':obj_item.DNSDomain,
                                         'DNS Servers':dns_srv}
    writer.writelog(log_file, str(len(network_dict)) + ' network interfaces found :<br>\n')
    # Ecriture du fichier CSV
    header = ['Name', 'MAC', 'Enabled', 'IpV4', 'IpV6', 'Default Gateway',
              'DHCP Enabled', 'DHCP Server', 'WINS Primary Server',
              'WINS Secondary Server', 'DNS Domain', 'DNS Servers']
    csv_file = log_file_path + "networkInterfaces.csv"
    writer.write_csv(csv_file, header, network_dict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Network Interfaces')
    # Ecriture sur le log
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, delimiter + '<br>\n')

    # Drives
    drives = psutil.disk_partitions()
    drives_dict = {}
    for drive in drives:
        drive_path = re.findall("device='(.*?)'", str(drive))[0]
        if drive.opts != 'cdrom':
            try:
                disk_usage = psutil.disk_usage(drive_path)
                drives_dict[str(drive_path)] = {'Path':str(drive_path),
                                                'Total_Space (Go)':str(disk_usage.total >> 30),
                                                'Free_Space (Go)':str(disk_usage.free >> 30),
                                                'Used_space (Go)':str(disk_usage.used >> 30),
                                                'Used_space (%)':str(disk_usage.percent)}
            except Exception:
                pass
    # Ecriture du fichier CSV
    header = ['Path', 'Total_Space (Go)', 'Free_Space (Go)', 'Used_space (Go)', 'Used_space (%)']
    csv_file = log_file_path + "drives.csv"
    writer.write_csv(csv_file, header, drives_dict)
    # Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Drives')
    # Ecriture sur le log
    writer.writelog(log_file, str(len(drives_dict)) + ' drive(s) found :<br>\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, delimiter + '<br>\n')

    # 3 - Ecriture fin de log
    writer.writelog(log_file, '\n</div>\n')

    return log_file

def security_product_state(product_state):
    '''
    **FR**
    Retourne la signification du product_state passé en paramètre
    **EN**
    Return product_state meaning pasted in parameter
    '''
    rtstatus = 'Unknown'
    defstatus = 'Unknown'
    for keys, values in PRODUCT_STATE_DICT.items():
        if product_state == keys:
            rtstatus = values['rtstatus']
            defstatus = values['defstatus']
            break
    return rtstatus, defstatus

def security_product_get(i, security_product_dict, col_items, product_type):
    '''
    **FR**
    Met à jour et retourne security_product_dict du col_items et du product_type en paramètre
    **EN**
    Update and return security_product_dict of the col_items and product_type in parameter
    '''
    for obj_item in col_items:
        product_name = obj_item.displayName
        instance_guid = obj_item.instanceGuid
        path_product = obj_item.pathToSignedProductExe
        path_reporting = obj_item.pathToSignedReportingExe
        product_state = str(obj_item.productState)
        # product_state conversion
        rtstatus, defstatus = security_product_state(product_state)
        i += 1
        # Put in dict
        security_product_dict[i] = {'Name':product_name, 'GUID':instance_guid,
                                    'pathProduct':path_product, 'pathReporting':path_reporting,
                                    'productState':product_state,
                                    'rtstatus':rtstatus, 'defstatus':defstatus,
                                    'Type':product_type}
    return i, security_product_dict

def security_product_info(log_file_path):
    '''
    http://neophob.com/2010/03/wmi-query-windows-securitycenter2/
    **FR**
    Scan le système pour obtenir la liste et le status des produits de sécurité (antivirus, parefeu, antispyware)
    Retourne le dictionnaire des produits de sécurité
    **EN**
    Scan to get security product in the system with their status (antivirus, firewall, antispyware)
    Return security products dict
    '''
    # Determinate what namespace to call
    sec_name_space = 'SecurityCenter'
    if detect_os() != 'xp':
        sec_name_space = 'SecurityCenter2'
    str_computer = '.'
    obj_wmi_service = win32com.client.Dispatch('WbemScripting.SWbemLocator')
    obj_sw_bem_services = obj_wmi_service.ConnectServer(str_computer, 'root\\' + sec_name_space)

    writer.writer('Getting security products')
    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Security products on computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtention des informations
    security_product_dict = {}
    i = 1
    # Antivirus
    col_items = obj_sw_bem_services.ExecQuery("select * from AntivirusProduct")
    i, security_product_dict = security_product_get(i, security_product_dict, col_items, 'Antivirus')

    # Firewalls
    # Windows
    xpfw = win32com.client.gencache.EnsureDispatch('HNetCfg.FwMgr', 0)
    xpfw_policy = xpfw.LocalPolicy.CurrentProfile
    writer.writelog(log_file, 'Windows firewall enabled : ' + str(xpfw_policy.FirewallEnabled) + '<br>\n')
    # Third party firewall
    col_items = obj_sw_bem_services.ExecQuery("select * from FireWallProduct")
    i, security_product_dict = security_product_get(i, security_product_dict, col_items, 'Firewall')

    if detect_os() != 'xp':
        # AntiSpyware
        col_items = obj_sw_bem_services.ExecQuery("select * from AntiSpywareProduct")
        i, security_product_dict = security_product_get(i, security_product_dict, col_items, 'Antispyware')

    # 3 - Ecrire du fichier CSV
    header = ['Type', 'Name', 'GUID', 'pathProduct', 'pathReporting', 'productState', 'rtstatus', 'defstatus']
    csv_file = log_file_path + "security_products.csv"
    writer.write_csv(csv_file, header, security_product_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Security Products')
    htmltxt = htmltxt.replace('<TD>Outdated</TD>', '<TD id="ko">Outdated</TD>')
    htmltxt = htmltxt.replace('<TD>Up to date</TD>', '<TD id="ok">Up to date</TD>')
    htmltxt = htmltxt.replace('<TD>Disabled</TD>', '<TD id="ko">Disabled</TD>')
    htmltxt = htmltxt.replace('<TD>Enabled</TD>', '<TD id="ok">Enabled</TD>')

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, str(len(security_product_dict)) + ' security products :\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')

    return security_product_dict

def process_info(log_file_path):
    '''
    ~ tasklist /SVC
    ***FR**
    Liste les processus démarrés sur l'ordinateur
    Retourne le dictionnaire des processus
    **EN**
    List computer running processes
    Retourne the processes dict
    '''
    writer.writer('Getting running processes')
    # 1 - Ecriture début de log
    log_file = log_file_path + 'FINAL.html'
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Running processes on computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtenir la liste des processus démarrés
    proc_dict = {}
    for proc in psutil.process_iter():
        try:
            proc_timestamp = proc.create_time()
            proc_start_date = datetime.fromtimestamp(proc_timestamp)
            proc_dict[str(proc.pid)] = {'Name':str(proc.name()),
                                        'PID':str(proc.pid),
                                        'PPID': str(proc.ppid()),
                                        'Start': str(proc_start_date)}
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    # 3 - Ecrire du fichier CSV
    header = ['Name', 'PID', 'PPID', 'Start']
    csv_file = log_file_path + 'processes.csv'
    writer.write_csv(csv_file, header, proc_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Running processes')

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, str(len(proc_dict)) + ' running processes :\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')

    return proc_dict

def services_info(log_file_path):
    '''
    ~ wmic service where started=true get name, displayname, processid, startmode, startname, pathname
    **FR**
    Liste les services démarrés sur l'ordinateur
    Retourne le dictionnaire des noms des services démarrés
    **EN**
    List computer started services
    Return the running services names dict
    '''
    writer.writer('Getting running services')
    # 1 - Ecriture début de log
    log_file = log_file_path + 'FINAL.html'
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Running services on computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - Obtenir la liste des services démarrés
    if detect_os() != 'xp' or psutil.__version__ != '3.4.2':
        services_list = psutil.win_service_iter()
        services_running_dict = {}
        for srv in services_list:
            srvname = re.findall("name='(.*)',", str(srv))[0]
            service_temp = psutil.win_service_get(srvname)
            service = service_temp.as_dict()
            # Keep running services
            if service['status'] == 'running':
                services_running_dict[service['name']] = {'Name':service['name'], 'PID':service['pid'],
                                                          'Display_Name':service['display_name'], 'Start_Type':service['start_type'],
                                                          'Username':service['username'], 'Binpath':service['binpath']}
    else:
        import wmi
        wmi_instance = wmi.WMI()
        services_list = wmi_instance.Win32_Service()
        services_running_dict = {}
        for srv in services_list:
            # Keep running services
            if srv.State == 'Running':
                services_running_dict[srv.name] = {'Name':srv.name, 'PID':srv.processid,
                                                   'Display_Name':srv.displayname, 'Start_Type':srv.startmode,
                                                   'Username':srv.startname, 'Binpath':srv.pathname}

    # 3 - Ecriture du fichier CSV
    header = ['Name', 'Display_Name', 'PID', 'Start_Type', 'Username', 'Binpath']
    csv_file = log_file_path + "services.csv"
    writer.write_csv(csv_file, header, services_running_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Running services')

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, str(len(services_running_dict)) + ' running services :<br>\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')

    return services_running_dict

def ports_info(log_file_path):
    '''
    ~ netstat -a
    **FR**
    Liste les communications réseaux de l'ordinateur
    Retourne le dictionnaire des connexions réseaux
    **EN**
    List computer network connections
    Return the network connections dict
    '''
    writer.writer('Getting network connections')
    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Informations about network connections of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - obtenir la liste des connexions actives
    ports_list = psutil.net_connections()
    ports_dict = {}
    for ports in ports_list:
        if ports.status == 'ESTABLISHED' or ports.status == 'CLOSE_WAIT':
            ports_dict[str(ports.laddr)] = {'Local_addr':str(ports.laddr[0]), 'Local_port':str(ports.laddr[1]),
                                            'Remote_addr':str(ports.raddr[0]), 'Remote_port':str(ports.raddr[1]),
                                            'Status':str(ports.status), 'PID':str(ports.pid)}

        if ports.status == 'LISTEN':
            ports_dict[str(ports.laddr)] = {'Local_addr':str(ports.laddr[0]), 'Local_port':str(ports.laddr[1]),
                                            'Remote_addr':'', 'Remote_port':'',
                                            'Status':str(ports.status), 'PID':str(ports.pid)}

    # 3 - Ecriture du fichier CSV
    header = ['Local_addr', 'Local_port', 'Remote_addr', 'Remote_port', 'Status', 'PID']
    csv_file = log_file_path + "ports.csv"
    writer.write_csv(csv_file, header, ports_dict)

    # 4 - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Network connections')

    # 5 - Ecriture de la fin du log
    writer.writelog(log_file, str(len(ports_dict)) + ' network connections :<br>\n')
    writer.writelog(log_file, htmltxt)
    writer.writelog(log_file, '\n</div>\n')

    return ports_dict

def get_files(directory, files_list):
    '''
    **FR**
    Obtenir les fichiers du répertoire en paramètre, retourne files_list
    **EN**
    Get files of the directory in parameter, return files_list
    '''
    if float(platform.python_version()[:3]) >= 3.5:
        for elems in os.scandir(directory):
            try:
                if elems.is_dir(follow_symlinks=False):
                    get_files(elems.path, files_list)
                else:
                    files_list.append(elems.path)
            except PermissionError:
                pass
    else: # Python 3.4.X or before => Windows XP
        for root_dir, _, files in os.walk(directory):
            for filename in files:
                files_list.append(os.path.join(root_dir, filename))
    return files_list

def search_reg(hkey, key):
    '''
    **FR**
    Chercher dans regedit les informations contenus dans les sous clés de registre de key en paramètre,
    Retourne tosave_dict
    **EN**
    Search in regedit informations in subkey of hkey in parameter,
    Return tosave_dict
    '''
    hive = hkey
    try:
        registry_key = winreg.OpenKey(hive, key, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) # W7-64/32
    except FileNotFoundError:
        return 0

    tosave_dict = {}

    i = 0
    while 1:
        try:
            reg_name, reg_value, reg_type = winreg.EnumValue(registry_key, i)
            if reg_type == 3: # binary
                try:
                    reg_value = reg_value[::2][:reg_value[::2].find(b'\x00')].decode()
                except UnicodeDecodeError:
                    pass
            tosave_dict[reg_name] = str(reg_value).replace(';', ',')
        except OSError:
            break
        i += 1
    return tosave_dict

def persistence_info(log_file_path):
    '''
    cf https://www.fuzzysecurity.com/tutorials/19.html
    **FR**
    Chercher les élements persistants dans :
        - les répertoires de démarrage
        - Regedit
    **EN**
    Search for persistent objects in :
        - Startup directories
        - Regedit
    '''
    writer.writer('Search for persistent objects')
    # 1 - Ecriture début de log
    log_file = log_file_path + "FINAL.html"
    writer.writelog(log_file, '<div><br>\n')
    elem = '<h2>Informations about persistent objects of computer "' + COMPUTERNAME + '"</h2>'
    writer.prepa_log_scan(log_file, elem)

    # 2 - obtenir les objets persistants
    # 2.1a - startup dirs
    drive = os.environ['SYSTEMDRIVE']
    user = os.environ['USERNAME']
    dir_list = [drive + '/Users/' + user + '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup',
                drive + '/ProgramData/Microsoft/Windows/Start Menu/Programs/Startup',
                drive + '/Documents and Settings/All Users/Start Menu/Programs/Startup']
    files_list = []
    for dirs in dir_list:
        if os.path.isdir(dirs):
            files_list = get_files(dirs, files_list)

    # 2.1b - Ecriture dans le log
    writer.writelog(log_file, 'Persistent objects in startup dirs :<br>\n')
    for elem in files_list:
        writer.writelog(log_file, elem + '<br>\n')

    # 2.2a - regedit
    hive_list = {'HKEY_LOCAL_MACHINE': winreg.HKEY_LOCAL_MACHINE, 'HKEY_CURRENT_USER':winreg.HKEY_CURRENT_USER}
    reg_list = [r'Software\Microsoft\Windows\CurrentVersion\Run',
                r'Software\Microsoft\Windows\CurrentVersion\RunOnce',
                r'Software\Microsoft\Windows\CurrentVersion\RunServices',
                r'Software\Microsoft\Windows\CurrentVersion\RunServicesOnce',
                r'Software\Microsoft\Windows NT\CurrentVersion\Winlogon']

    per_reg_dict = {}
    for hkeyname, hkeyvalue in hive_list.items():
        for key in reg_list:
            fullhkeyreg = hkeyname + '\\' + key
            tosave_dict = search_reg(hkeyvalue, key)
            if tosave_dict:
                for name, value in tosave_dict.items():
                    per_reg_dict[fullhkeyreg + '_' + name] = {'regkey':fullhkeyreg, 'name':name, 'value':value}

    # 2.2b - Ecriture du fichier CSV
    header = ['regkey', 'name', 'value']
    csv_file = log_file_path + "persistants_reg.csv"
    writer.write_csv(csv_file, header, per_reg_dict)

    # 2.2c - Transformation du CSV en HTML
    htmltxt = writer.csv2html(csv_file, 'Persistent reg objects')

    # 2.2d - Ecriture de la fin du log
    writer.writelog(log_file, '<br>\nPersistent reg objects :<br>\n')
    writer.writelog(log_file, htmltxt)

    # 3 - Ecriture de la fin du log
    writer.writelog(log_file, '\n</div>\n')

    return log_file
