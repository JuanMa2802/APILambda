from distutils.command.install_egg_info import safe_name
import json

salida = [{'obra':'436','datos':['hola','josa']}]
obras = ['209','211','435','436','437','438','439','441','110','411','426','443']
proveedorInfo = [{'obra':'436','tercero':'pepito'},{'obra':'438','tercero':'pepito'},{'obra':'439','tercero':'pepito'}]

for obra in obras:
    safe_name = [sa['datos'] for sa in salida if obra in sa['obra']]
    for row in salida:
        obrae = set([proy['tercero'] for proy in proveedorInfo if obra in proy['obra']])
        print(obrae)
