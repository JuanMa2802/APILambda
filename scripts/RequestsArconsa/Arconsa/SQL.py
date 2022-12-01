# import pyodbc 
# # Some other example server values are
# # server = 'myserver,port' # to specify an alternate port
# server = 'datamart.sincoerp.com:4263'
# database = 'SincoArconsaDW'
# username = 'SincoArconsaDW'
# password = 'SincoArconsa.382.1866'
# # ENCRYPT defaults to yes starting in ODBC Driver 18. It's good to always specify ENCRYPT=yes on the client side to avoid MITM attacks.
# cnxn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER='+server+';DATABASE='+database+';ENCRYPT=yes;UID='+username+';PWD='+ password)
# cursor = cnxn.cursor()

# import mysql.connector
# miConexion = mysql.connector.connect( host='datamart.sincoerp.com:4263', user= 'SincoArconsaDW', passwd='SincoArconsa.382.1866', db='SincoArconsaDW')
# cur = miConexion.cursor()
# # cur.execute( "SELECT nombre, apellido FROM usuarios" )
# # for nombre, apellido in cur.fetchall():
# #     print(nombre, apellido)
# miConexion.close()