import os
import sys
from rich import print
import Kits_buildings_reports
import traceback
from datetime import datetime

try:
    Id, TypeReport, NameFaileDoc = sys.argv[1:]
except:
    Id, TypeReport, NameFaileDoc = '', '', ''
    print("'# входные параметры не корректны'")
    sys.exit(1)

# Id, TypeReport, NameFaileDoc = '2286', 'TZnaII', 'TZII.docx'


'''# создать пустой каталог (папку)'''
try:
    os.mkdir("logs")
except:
    pass    

if TypeReport == 'TZnaII':
    try:
        sys.exit(Kits_buildings_reports.GO(Id, TypeReport, NameFaileDoc))
    except:
        errortext = traceback.format_exc()
        
        current_datetime = str(datetime.now())
        current_datetime = current_datetime.replace('.', '_').replace(' ', '_').replace(':', '-')
        namelog  = f"{os.getcwd()}\\logs\\{current_datetime}.txt"

        with open(namelog, "w") as f:
            text = f"""{Id} {TypeReport} {NameFaileDoc}\n{errortext}"""
            print(text)
            f.write(text)

