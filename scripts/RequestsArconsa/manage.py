import sys
from Extras.Colorama import bcolors

if sys.argv[1].lower() == "arconsa":
    try:
        print(bcolors.OK + "Inicializando el archivo RequestsArconsa" + bcolors.RESET)
        exec(open("Arconsa/RequestsArconsa.py").read())
    except Exception as e:
        print(e)
        print(bcolors.OK + "Inicializando el archivo RequestsArconsa" + bcolors.RESET)
        exec(open("Arconsa/RequestsArconsa.py").read())
else:
    print(bcolors.FAIL + "El argumento dado no tiene ningúna función definida" + bcolors.RESET)
