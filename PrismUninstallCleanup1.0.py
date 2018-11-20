#! python3
# Prism Uninstall cleanup
# Version 1.0
# By Zach Cutberth

# Cleans up remaining Prism files, folders, services, registry entries, and RPS/DRS database users.

# Imports
import sys
import shutil
import subprocess
import cx_Oracle
import config

# Check to make sure they want to run Prism clean up.
runCleanUp = input('Type \'Yes\' to run Prism Clean Up: ').lower()
if runCleanUp == 'yes':
    print('\nBeginning Prism Clean Up...')
else:
    sys.exit()

# Stop Prism Services

print('\n**************************************************\n')
print('Stopping Prism Services...\n')

subprocess.run(["net", "stop", "Apache"])
subprocess.run(["net", "stop", "PrismMQService"])
subprocess.run(["net", "stop", "RabbitMQ"])
subprocess.run(["net", "stop", "PrismBackOfficeService"])
subprocess.run(["net", "stop", "PrismCommonService"])
subprocess.run(["net", "stop", "PrismLicSvr"])
subprocess.run(["net", "stop", "PrismResiliencyService"])
subprocess.run(["net", "stop", "PrismV9Service"])

print('\nPrism Services Stopped.\n')

# Remove Prism Services

print('**************************************************\n')
print('Removing Prism Services...\n')

subprocess.run(["sc", "delete", "Apache"])
subprocess.run(["sc", "delete", "RabbitMQ"])
subprocess.run(["sc", "delete", "PrismBackOfficeService"])
subprocess.run(["sc", "delete", "PrismCommonService"])
subprocess.run(["sc", "delete", "PrismLicSvr"])
subprocess.run(["sc", "delete", "PrismMQService"])
subprocess.run(["sc", "delete", "PrismResiliencyService"])
subprocess.run(["sc", "delete", "PrismV9Service"])

print('\nPrism Services Removed.\n')

# Remove Files and Folders

print('**************************************************\n')
print('Deleting Prism Files and Folders...\n')

path = 'c:\ProgramData\RetailPro'
subprocess.run(['rd', '/S', '/Q', path], shell=True)

path = 'c:\Progra~2\Apache Software Foundation'
subprocess.run(['rd', '/S', '/Q', path], shell=True)

path = 'c:\Progra~2\erl8.1'
subprocess.run(['rd', '/S', '/Q', path], shell=True)

path = 'c:\Progra~2\RabbitMQ Server'
subprocess.run(['rd', '/S', '/Q', path], shell=True)

path = 'c:\Progra~2\RetailPro'
subprocess.run(['rd', '/S', '/Q', path], shell=True)

print('\nPrism Files and Folders Deleted.\n')

# Remove Registry Entries.

print('**************************************************\n')
print('Deleting Prism Registry Keys...\n')

subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\Prism", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Apache Software Foundation", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Ericsson", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\Document Designer", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismBackOfficeService", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismCommonService", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismLicSvr", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismMQService", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismResiliencyService", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Retail Pro\PrismV9Service", "/f"])
subprocess.run(["reg", "delete", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\VMware, Inc.", "/f"])

print('\nPrism Registry Keys Deleted.\n')

# Remove RPS and DRS database users.

print('**************************************************\n')
print('Dropping RPS/DRS Oracle Users...\n')

connstr = config.connstring
dbconnection = cx_Oracle.connect(connstr)
cursor = dbconnection.cursor()
try:
    cursor.execute('drop user rps cascade')
except:
    print('RPS user does not exist.')
try:
    cursor.execute('drop user drs cascade')
except:
    print('DRS user does not exist.')
cursor.close()
dbconnection.close()

print('\nRPS/DRS Oracle Users Dropped.\n')
      
print('**************************************************\n')
print('Prism Clean Up Completed.')

input('\nPress Any Key To Exit...')
