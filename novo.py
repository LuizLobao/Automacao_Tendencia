import os, time
from datetime import date, datetime



arquivo = r'c:/audio.log'


login = os.getlogin()
print(login)


usuario = os.getenv('username')
print(usuario)


print("last modified: %s" % time.ctime(os.path.getmtime(arquivo)))
print("created: %s" % time.ctime(os.path.getctime(arquivo)))