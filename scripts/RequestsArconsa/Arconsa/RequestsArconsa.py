# Imports scrapping
import requests
import json
import pandas as pd
import subprocess
from bs4 import BeautifulSoup
from time import time
from Extras.proveedores import dataframeSP
from Extras.GetDays import getDays
import time as tiempo
import sys
import holidays_co
import warnings
import datetime
from Arconsa.mail import mail
from os import remove

warnings.filterwarnings("ignore")
fechaActual = datetime.date.today()
diaActualSemana = datetime.date.today().strftime('%A')
diaFestivoFT = holidays_co.is_holiday_date(fechaActual)

# prueba = False
prueba = True
print('test:',prueba)

# Iniciar Sesión 
def login(user,enc_pwd,auth_url):
    s=requests.Session()
    s.headers['User-Agent']="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.134 Safari/537.36 OPR/89.0.4447.91"
    get = s.get('https://www1.sincoerp.com/SincoArconsa/V3/Marco/Login.aspx')
    payload={'ClaveUsuario': enc_pwd.replace('\n', ''), 'NomUsuario': user}
    res =s.post("https://www1.sincoerp.com/SincoArconsa/V3/API/Auth/Usuario", json = payload)
    contenido=json.loads(res.content)
    token=contenido.get('access_token')
    s.headers['Authorization'] = contenido.get('token_type')+' '+token
    s.headers['Referer'] = 'https://www1.sincoerp.com/SincoArconsa/V3/Marco/Seleccion.aspx'
    data = contenido['data']
    s.post('https://www1.sincoerp.com/SincoArconsa/V3/Marco/Session', data=data)
    get = s.get('https://www1.sincoerp.com/SincoArconsa/V3/API/Auth/Sesion/Iniciar/1/Empresa/1/Sucursal/250')
    contenido=json.loads(get.content)
    data = {'asp':contenido['asp'],'token':contenido['token']['token_type']+" "+contenido['token']['access_token']}
    s.headers['Authorization'] = contenido['token']['token_type']+" "+contenido['token']['access_token']
    data = json.dumps(data)
    headers={'Content-Type':'application/json'}
    s.put('https://www1.sincoerp.com/SincoArconsa/V3/Marco/Session',data=data, headers=headers)
    return s

# Cargar excel de los insumos
InsumosAgr = pd.read_excel(r"C:\Lambda Analytics\RequestsArconsa\Extras\Insumos.xlsx")
# Cargar excel de proveedores sugeridos
proveedores = dataframeSP('Proveedores_Orden-Compra')
# Ecripta la contraseña como lo hace SINCO
password = str(subprocess.run(["node", "Arconsa/funciones.js", 'password','Lambda2022*'], capture_output=True).stdout.decode())
# Manda los datos parar realizar el logueo
session = login('sandra.r',password,'https://www1.sincoerp.com/SincoArconsa/V3/API/Auth/Usuario')
session.get('https://www1.sincoerp.com/SincoArconsa/V3/Marco/')
# Paso los datos precisos desde el netword del navegador
session.headers['Referer']="https://www1.sincoerp.com/SincoArconsa/V3/ADPRO/Views/Almacen/Pedidos/ComprarPedidosProveedorSugerido.html"

try:
    Proyectos = list()
    Proyectos.append(sys.argv[2])
except:
    Proyectos = getDays(proveedores)

# Define las variables que se van a usar  
Insumos, UMs , Cantidads, Inventarios, FechaPedidos, FechaRequs, Adjuntos, Comentarios, Proveedors, Proys, orden, ordenesA,insumos, newSupplier, OC, OCG, datosSalida, datosProveedores = [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]]

# Se analizan las reglas de ARCONSA
def reglas(idInsumo, proveedorInfo, proy, row):
    producto = row.get('ObjProductos').get('ProDesc').split(' - ')[1]
    cantidad = float(row.get('Objpedidossucursal').get('PdSCant'))
    index = InsumosAgr.index[InsumosAgr['Codigo Insumo']==int(idInsumo)]
    insumo = None
    Aceros = False
    acepta = True
    tercero = ""
    Pindex = proveedores.index[proveedores['# Obra']==proy].to_numpy()
    proveedor = ""
    for ind in Pindex:
        proveedor = proveedores[ind:ind+1]
    for ind in index:
            insumo = InsumosAgr[ind:ind+1]
    # Regla del acero
    if (insumo['Agrupacion'] == int(1)).values[0]:
        Prov = proveedor['Acero'].values[0]
        for data in proveedorInfo:
            if data.get('objtercero')['TerNombre'].upper().find(Prov.upper()) != -1:
                    tercero = data.get('objtercero')
        Aceros = True
        if tercero == "":
            tercero = proveedorInfo[0].get('objtercero')
    
    # Regla del cemento
    elif (insumo['Agrupacion'] == int(11)).values[0]:
        Prov = proveedor['Cemento'].values[0]
        for data in proveedorInfo:
            if data.get('objtercero')['TerNombre'].upper().find(Prov.upper()) != -1:
                    tercero = data.get('objtercero')
                        
        if tercero == "":
            tercero = proveedorInfo[0].get('objtercero')
    # Regla del ACPM
    elif (insumo['Agrupacion'] == int(10)).values[0]:
        Prov = proveedor['ACPM'].values[0]
            # Cuando sea riovivo o accanto el proveedor va a ser SERVILLANOGRANDE
        if cantidad >= int(proveedor['Cantidad'].values[0]):
            acepta = True
            if proy == 342 or proy == 326 or proy == 265:
                for data in proveedorInfo:
                    if data.get('objtercero')['TerNombre'].upper().find(Prov.upper()) != -1:
                        # Si el insumo es de ACPM con transporte le pone el servillanogrande de tranporte ACPM
                        if str(producto) == 'Transporte ACPM':
                            tercero = data.get('objtercero')
                        # Si el insumo es de ACPM y no tiene transporte le pone el servillanogrande de ACPM
                        elif str(producto) == 'ACPM':
                            tercero = data.get('objtercero')
                        # si no le pone el servillanogrande normal
                        else:
                            tercero = data.get('objtercero')
            # Para todos los demás insumos de ACPM le pone acevedo hermanos y cia ltda
            else:
                print('Acevedo')
                for data in proveedorInfo:
                    if int(data.get('objtercero').get('TerID'))==470:
                        tercero = data.get('objtercero')    
            if tercero == "":
                tercero = proveedorInfo[0].get('objtercero')
        else:
            acepta = False
        
    # Regla basica poner el primer proveedor
    else:
        tercero = proveedorInfo[0].get('objtercero')
    return tercero, Aceros, acepta

# Obtiene todos los datos de los insumos y actualiza el proveedor
def scrap():
    for proy in Proyectos:
        url="https://www1.sincoerp.com/sincoarconsa/V3/ADPRO/api/ComprarPedidos/ComprarPedSugerido/proyecto/%s/FchIniped/-1/FchFinped/-1/FchInireq/-1/FchFinreq/-1/Usuario/-1/Producto/-1/urgentes/-1" %str(proy)
        get = session.get(url, timeout=30)
        print(str(proy))
        salida=json.loads(get.content)
        datosSalida.append({'obra':proy,'datos':salida})
        if len(salida)>0:
            for row in salida:
                producto = int(row.get('ObjProductos').get('ProDesc').split()[0])
                get = session.get("https://www1.sincoerp.com/sincoarconsa/V3/ADPRO/api/ComprarPedidos/ComprarPedSugerido/TraeProvedorSugerido/Producto/"+str(producto)+"/proyecto/"+proy)
                proveedorInfo, tercero, idpedido =["","",""]
                try:
                    proveedorInfo = json.loads(get.content.decode())
                    datosProveedores.append({'obra':proy,'producto':producto,'datos':proveedorInfo})
                    idpedido = row.get('Objpedidossucursal')['PdsID']
                except Exception as e:
                    pass
                if proveedorInfo!=[]:
                    try:
                        idInsumo = row.get('Objpedidossucursal').get('PdSProd')
                        observacion = row.get('Objpedidossucursal').get('PdSComentarios')
                        tercero,Aceros,acepta = reglas(idInsumo, proveedorInfo, proy, row)
                        data = {"idpedidos":idpedido,"obra":proy,"prod":producto,"tercero":tercero['TerID']}
                        url = "https://www1.sincoerp.com/sincoarconsa/V3/ADPRO/api/ComprarPedidos/CambiarProvSugerido"
                        post = session.post(url, data=data)
                        _response = json.loads(post.content.decode())
                        if _response['ObjMensajes']['codigo']==1:
                            if int(tercero['TerID']) == int(_response['tercero']):
                                newSupplier.append({'idpedidos':_response['idpedidos'], 'terceroID':str(_response['tercero'])+"/"+str(tercero['TerID']), 'producto':str(_response['prod']), 'tercero':str(tercero['TerNombre']),'obra':proy,'Observacion':observacion, 'Aceros':Aceros})
                                if acepta == False:
                                    newSupplier.pop()
                                    
                    except Exception as e:
                        print(row.get('ObjProductos').get('ProDesc'), 'No se encontro el recurso', e)
                    # print("No se encontro el recurso")
                    
start_time = time()
scrap()
elapsed_time = time() - start_time
elapsed_time = round(float(elapsed_time))

# Realiza las ordenes 
def ordenes():
    obrasUnicas=list(set([el.get('obra') for el in newSupplier]))
    for obraUnica in obrasUnicas:
        provedoresUnicos=list(set([el.get('tercero') for el in newSupplier if el.get('obra') == obraUnica]))
        for proveedor in provedoresUnicos:
            observacionesAceros = set([el.get('Observacion').replace(' ', '').lower() for el in newSupplier if el.get('tercero') == proveedor and el.get('obra') == obraUnica and el.get('Aceros') == True])
            idpedidos = []
            auth = str(session.headers['Authorization'])
            Insumo = []
            if observacionesAceros:
                for observacion in observacionesAceros:
                    print(observacion)
                    idpedidos = json.dumps([el.get('idpedidos') for el in newSupplier if el.get('tercero') == proveedor and el.get('obra') == obraUnica and el.get('Observacion').replace(' ','').lower() == observacion])
                    obra = str(obraUnica)
                    if prueba == False:
                        Insumo = [el.get('producto') for el in newSupplier if el.get('tercero') == proveedor and el.get('obra') == obraUnica]
                        response = subprocess.run(["node", r"C:\Lambda Analytics\RequestsArconsa\Arconsa\js.js", idpedidos, obra, auth], capture_output=True)
                        dict = json.loads(response.stdout.decode())
                        print(idpedidos, obra, Insumo, dict['mensaje'])
                        OCG.append({'Obra':obra,'Pedidos':json.loads(idpedidos),'OC':dict['mensaje']})
                    else:
                        print(idpedidos, obra, Insumo, 'mensaje')
                        OCG.append({'Obra':obra,'Pedidos':json.loads(idpedidos),'OC':'mensaje'})   
                    tiempo.sleep(5)
                    ordenesA.append(True)
            else:
                idpedidos = json.dumps([el.get('idpedidos') for el in newSupplier if el.get('tercero') == proveedor and el.get('obra') == obraUnica])
                obra = str(obraUnica)
                if prueba == False:
                    Insumo = [el.get('producto') for el in newSupplier if el.get('tercero') == proveedor and el.get('obra') == obraUnica]
                    response = subprocess.run(["node", "Arconsa/js.js", idpedidos, obra, auth], capture_output=True)
                    dict = json.loads(response.stdout.decode())
                    print(idpedidos, obra, Insumo, dict['mensaje'])
                    OCG.append({'Obra':obra,'Pedidos':json.loads(idpedidos),'OC':dict['mensaje']})
                else:
                    print(idpedidos, obra, Insumo, 'mensaje')
                    OCG.append({'Obra':obra,'Pedidos':json.loads(idpedidos),'OC':'mensaje'})   
                tiempo.sleep(5)
                ordenesA.append(True)
                
# Genera el reporte
def reporte():
    start_time = time()
    name_excel = ordenes()
    elapsed_timeO = time() - start_time
    elapsed_timeO = round(float(elapsed_timeO))
    print("Proceso finalizado correctamente, se procesaron: "+str(len(ordenesA))+" ordenes ("+diaActualSemana+")")
    print('Generando Reporte... Tiempo estimado: %s segundos.'%str(elapsed_time))
    for proy in Proyectos:
        salida = [sa.get('datos') for sa in datosSalida if proy in sa['obra']][0]
        if len(salida)>0:
            print(str(proy))
            for row in salida:
                producto = int(row.get('ObjProductos').get('ProDesc').split()[0])
                try:
                    idpedido = row.get('Objpedidossucursal')['PdsID']
                    OC.append([id.get('OC') for id in [idP for idP in [si for si in OCG if proy in si.get('Obra')] if str(idpedido) in idP.get('Pedidos')] if str(idpedido) in id.get('Pedidos')][0])
                    for supplier in newSupplier:
                        if int(supplier['idpedidos']) == int(idpedido):
                            print(supplier['tercero'], supplier['producto'])
                            Proveedors.append(supplier['tercero'])
                    Insumos.append(row.get('ObjProductos').get('ProDesc'))
                    UMs.append(row.get('ObjProductos').get('ProUnidadCont'))
                    Cantidads.append(row.get('Objpedidossucursal').get('PdSCant'))
                    Inventarios.append(row.get('CantidadInventario'))
                    FechaPedidos.append(row.get('Objpedidossucursal').get('PdSFechaPed'))
                    FechaRequs.append(row.get('Objpedidossucursal').get('PdSFechaReq'))
                    Adjuntos.append(row.get('Objpedidossucursal').get('RutaAdjuntos'))
                    Comentarios.append(row.get('Objpedidossucursal').get('PdSComentarios'))
                    Proys.append(proy)
                except Exception as e:
                    print("No se implemento en una orden el producto: "+str(producto))
    # Genera el dataframe y el excel
    base=pd.DataFrame()     
    base['Insumos']=Insumos
    base['UMs']=UMs
    base['Cantidads']=Cantidads
    base['Inventarios']=Inventarios
    base['FechaPedidos']=FechaPedidos
    base['FechaRequs']=FechaRequs
    base['Adjuntos']=Adjuntos
    base['Comentarios']=Comentarios
    base['Proveedors']=Proveedors
    base['Proyecto']=Proys
    base['OC']=OC
    print(base)
    if len(OC) > 0:
        if prueba == False:
            try:
                remove('Reportes/pedidosOrden.xlsx')
            except:
                pass
            base.to_excel('Reportes/pedidosOrden.xlsx',header=True, index=False)
            # Envio Correo
            mail('jhormangallegogallego@gmail.com')
            mail('karen.giraldo@lambdaanalytics.co')
            mail('sandra.rodriguez@arconsa.com.co')
            return 'Reportes/pedidosOrden.xlsx'

start_time = time()
name_excel = reporte()
elapsed_timeO = time() - start_time
elapsed_timeO = round(float(elapsed_timeO))

if name_excel != None:
    print(name_excel+' generado correctamente En: %s segundos'%str(elapsed_timeO))    
