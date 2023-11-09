import os

try:
    os.chdir('\\\\192.168.11.10\\PM_secretC')
except PermissionError:
    print('Unable to connect to server, please check the directory again.')
else:
    print('Connected to server successfully!')