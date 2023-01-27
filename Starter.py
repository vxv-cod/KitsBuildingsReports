import os
import sys
from rich import print
import TZnaII
import ExplicationGP
import traceback
from datetime import datetime


def ER():
    errortext = traceback.format_exc()
    current_datetime = str(datetime.now())
    current_datetime = current_datetime.replace('.', '_').replace(' ', '_').replace(':', '-')
    namelog  = f"{os.getcwd()}\\logs\\{current_datetime}.txt"
    with open(namelog, "w") as f:
        text = f"""{Id} {TypeReport} {NameFaileDoc}\n{errortext}"""
        print(text)
        f.write(text)

try:
    TypeReport = sys.argv[2]
except:
    pass    

'''# создать пустой каталог (папку)'''
try:
    os.mkdir("logs")
except:
    pass    

if TypeReport == 'TZnaII':
    try:
        Id, TypeReport, NameFaileDoc = sys.argv[1:]
        sys.exit(TZnaII.GO(Id, NameFaileDoc))
    except:
        ER()

if TypeReport == 'ExpGP':
    try:
        Id, TypeReport, NameFaileDoc, KO = sys.argv[1:]
        sys.exit(ExplicationGP.GO(Id, NameFaileDoc, KO))
    except:
        ER()


