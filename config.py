#!/usr/bin/python3

#server details
username=''
password=''

#database connection details
dbuserid = '' #Database Userid
dbpwd = '' #Database Password
connectionstring = ''    #DatabaseServer:DatabasePort/SID

#python 3
#install paramiko, cx_Oracle, openpyxl, pillow

#For the following error: "DPI-1047: Cannot locate a 64-bit Oracle Client library: The specified module could not be found"
#Download 64-bit version of oracle instantClient (Basic Package) from: https://www.oracle.com/database/technologies/instant-client/winx64-64-downloads.html
#Extract the files in the instantclient_directory.zip file to the python directory
#Copy the path of the extracted folder and paste it in line 6 of the dsr.py code

