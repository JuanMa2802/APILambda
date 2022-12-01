email = 'automations@arconsa.com.co'
password = 'Lambda63074'

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import pandas as pd
import re

site_url = "https://arquitecturayconstrucciones.sharepoint.com/sites/OrdenesdeCompra"
ctx = ClientContext(site_url).with_credentials(UserCredential(email, password))
def dataframeSP(lista):
    sp_list = lista    
    sp_lists = ctx.web.lists    
    s_list = sp_lists.get_by_title(sp_list)
    l_items = s_list.get_items()
    ctx.load(l_items)
    ctx.execute_query()
    columnas=[]
    for columna in l_items[0].properties.items():
        if columna[0]=='Title':
            columnas.append('# Obra')
        elif columna[0]=='field_1':
            columnas.append('Obra')
        elif columna[0]=='field_2':
            columnas.append('Dias')
        elif columna[0]=='field_3':
            columnas.append('email')
        elif columna[0]=='field_4':
            columnas.append('Nombre Almacenista')
        elif columna[0]=='field_5':
            columnas.append('Cemento')
        elif columna[0]=='field_6':
            columnas.append('Acero')
        elif columna[0]=='field_7':
            columnas.append('Cantidad')
        elif columna[0]=='field_8':
            columnas.append('ACPM')
    valores=list()
    for item in l_items:
        fila = list(item.properties.items())
        data = []
        for columna in fila:
            if columna[0]=='Title':
                data.append(columna[1])
            elif columna[0]=='field_1':
                data.append(columna[1])
            elif columna[0]=='field_2':
                data.append(columna[1])
            elif columna[0]=='field_3':
                data.append(columna[1])
            elif columna[0]=='field_4':
                data.append(columna[1])
            elif columna[0]=='field_5':
                data.append(columna[1])
            elif columna[0]=='field_6':
                data.append(columna[1])
            elif columna[0]=='field_7':
                if columna[1] != None:
                    data.append(columna[1])
                else: 
                    data.append(55)
            elif columna[0]=='field_8':
                data.append(columna[1])
        valores.append(data)
    resultado=pd.DataFrame(valores,columns=columnas)
    return resultado

