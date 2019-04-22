# ScanPC
[![License](https://img.shields.io/badge/licence-AGPLv3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0.en.html)
![Platform](https://img.shields.io/badge/platform-win--32%20%7C%20win--64-lightgrey.svg)
[![Language](https://img.shields.io/badge/language-python3-orange.svg)](https://www.python.org/)

ScanPC is an audit tool for Windows.  
It scans your Windows computer to gather informations like :  
- the user accounts list
- the password policy
- the share folders list
- the hardware configuration
- the OS version
- the network interfaces
- the Windows updates (KB) list
- the softwares installed
- the firewall state
- the processes list
- the services list
- the antivirus state  

Started since October 2016  
Python 3.6 32 bit is used for all new developments.  
Project ongoing.  

## Compatibility
This software has been succesfully tested on the following Microsoft Windows systems (32 and 64 bits):  
- 7
- 10

## Getting Started
**Download the project on your computer.**
```
git clone https://github.com/alevoski/ScanPC.git
```

**Install the required pip modules**
```
pip install -r requirements.txt
```

**Optional tool**
RemoveDrive.exe from <https://www.uwe-sieber.de/drivetools_e.html> website is used to dismount device used to scan the computers.  


## HOW TO USE ?
**Method 1 - With Python and all dependencies**  
```
cd CS/
python main.py
```

**Method 2 - Compile the code into an executable file**  
You will need pyinstaller module 
```
pip install pyinstaller
c:\<yourpythonpath>\Scripts\pyinstaller.exe --onefile --icon=logoScanPC.ico main.py
```
A folder named "dist" will be created with a scanPC executable file.  
Note : you need to compile with a 32 bit Python version to perform scan on 32 bit Windows OS  
  
Put the code or the exe on a USB key and go scanning some Windows computers.  

***Note : To limit viruses spreading, you should always analyze your scanning devices between two scans !***
[Decontamine_Linux can help you !](https://github.com/alevoski/decontamine_Linux)

#### Demo  
![](DOCS/DEMO/scanpc_demo.gif)  

## Author
Alexandre Buissé

## License
ScanPC. Audit tool for Windows.  
Copyright (C) 2019 Alexandre Buissé alevoski@pm.me

This program is free software: you can redistribute it and/or modify  
it under the terms of the GNU Affero General Public License as published  
by the Free Software Foundation, either version 3 of the License, or  
(at your option) any later version.  

This program is distributed in the hope that it will be useful,  
but WITHOUT ANY WARRANTY; without even the implied warranty of  
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the  
GNU Affero General Public License for more details.  

You should have received a copy of the GNU Affero General Public License  
along with this program.  If not, see <https://www.gnu.org/licenses/>.


