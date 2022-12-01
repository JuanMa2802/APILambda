import pandas as pd
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import requests
import json
import re
from datetime import datetime
from dateutil import relativedelta
from datetime import date
from pytz import utc
from pprint import pprint
import sys

querystring = {"hapikey":"5a1e6e50-46c5-47f8-aa2d-f35be24978a1"}

def FormatNamePer(name1,name2):
      
    if(name1=='None'):
        
        name1 = ""
        
    else:
        pLetra =  name1[0].upper()
        
        rLetras = name1[1:].lower()
        
        name1 = pLetra + rLetras
        
    if(name2=='None'):
        
        name2 = ""
        
    else:
        
        pLetra =  name2[0].upper()
        
        rLetras = name2[1:].lower()
        
        name2 = pLetra + rLetras

    response = str(name1 +  ' ' + name2+ ' ')
    
    return response

def FormatSurnamePer(surname1,surname2):
    
    if(surname1=='None'):
        
        surname1 = ""
        
    else:
        
        pLetra =  surname1[0].upper()
        
        rLetras = surname1[1:].lower()
        
        surname1 = pLetra + rLetras
        
    if(surname2=='None'):
        
        surname2 = ""
        
    else:
        
        pLetra =  surname2[0].upper()
        
        rLetras = surname2[1:].lower()
        
        surname2 = pLetra + rLetras
        
    response = str(surname1 + ' ' + surname2)
    
    return response
   

def IsNone(value):
    
    if(value=='None'):
        
        value = ""
        
    return value

def ValuestoDictOp(value):
    
      
        value1 = str(value.get(0, 'Value'))
        
        value2 = str(value.get(1, 'Value'))
        
        value3 = str(value.get(2, 'Value'))
        
        value4 = str(value.get(3, 'Value'))
        
        value5 = str(value.get(4, 'Value'))
        
        value6 = str(value.get(5, 'Value'))
        
        
        if(value1=='Value' or value1 == None):
            
            value1 = ""
            
        else:
            
            value1 = value1+';'
            
        if(value2=='Value' or value2 == None):
            
            value2 = ""
            
        else: 
            
            value2 = value2+';'
            
        if(value3=='Value' or value3 == None):
            
            value3 = ""
            
        else:
            
            value3 = value3 + ';'
            
            
        if(value4=='Value' or value4 == None):
            
            value4 = ""
            
        else:
            
             value4 = value4 +';'
            
        if(value5=='Value' or value5 == None):
            
            value5 = ""
            
        else:
            
            value5 = value5 + ';'
            
        if(value6=='Value' or value6 == None):
            
            value6 = ""
            
        else:
            
            value6 = value6
            
        Response =f"{value1}{value2}{value3}{value4}{value5}{value6}"
        
        return Response


def ValuestoDictRg(value):
    
        value1 = str(value.get(0, 'Value'))
        
        value2 = str(value.get(1, 'Value'))
        
        value3 = str(value.get(2, 'Value'))
        
        value4 = str(value.get(3, 'Value'))
        
        
        if(value1=='Value' or value1 == None):
            
            value1 = ""
            
        else:
            
            value1 = value1 + ';'
            
            
        if(value2=='Value' or value2 == None):
            
            value2 = ""
            
        else:
            
            value2 =value2 + ';'
            
        if(value3=='Value' or value3 == None):
            
            value3 = ""
            
        else:
            
            value3 = value3 + ';'
            
        if(value4=='Value' or value4== None):
            
            value4 = ""
            
        else:
            
            value4 = value4 + ';'
            
        Response =f"{value1}{value2}{value3}{value4}"
        
        return Response


def ValuestoDictOb(value):
    
        value1 = str(value.get(0, 'Value'))
        
        value2 = str(value.get(1, 'Value'))
        
        value3 = str(value.get(2, 'Value'))
        
        value4 = str(value.get(3, 'Value'))
        
        value5 = str(value.get(4, 'Value'))
        
        if(value1=='Value' or value1 == None ):
            
            value1 = ""
            
        else:
            
            value1 = value1+';'
            
        if(value2=='Value' or value2 == None ):
            
            value2 = ""
            
        else: 
            
            value2 = value2+';'
            
        if(value3=='Value' or value3 == None ):
            
            value3 = ""
            
        else:
            
            value3 = value3 + ';'
            
        if(value4=='Value' or value4 == None):
            
            value4 = ""
            
        else:
            
             value4 = value4 +';'
            
        if(value5=='Value' or value5 == None):
            
            value5 = ""
            
        else:
            
            value5 = value5 + ';'
            
        Response =f"{value1}{value2}{value3}{value4}{value5}"
        
        return Response
    
    
def FormateDateYear(fh):
    
    if(fh=='None'):
        
        year = ""
        
    else:
        
        year = datetime.fromisoformat(fh[:-1]).strftime('%Y')
        
        year =int(year)
   
    return year


def FormateDateMonth(fh):
    
    if(fh=='None'):
        
        month = ""
        
    else:
        
        month = datetime.fromisoformat(fh[:-1]).strftime('%m')
        
        month = int(month)
        
    return month


def FormateDateDay(fh):
    
    if(fh=='None'):
        
        day = ""
        
    else:
        
        day = datetime.fromisoformat(fh[:-1]).strftime('%d')
        
        day = int(day)
        
    return day


def UTC(yy,mm,dd):
    
        utcdt = utc.localize(
               datetime(
                   year=yy,
                   month=mm,
                   day=dd
                   )
           )
        
        return utcdt

def ResponseUtc(yy,mm,dd):
    
    if (yy == "" and mm=="" and dd == ""):
        
        value =""
        
    else:
        
        value = UTC(yy, mm, dd)
        
    return value


site_url ='https://arquitecturayconstrucciones.sharepoint.com/sites/Automatizaciones/Proveedores'

ctx = ClientContext(site_url).with_credentials(UserCredential("automations@arconsa.com.co", "Lambda63074"))

# idRegister= int(77)
idRegister= int(sys.argv[1])

def dataframeSP(lista):
    sp_list = lista
    sp_lists = ctx.web.lists
    s_list = sp_lists.get_by_title(sp_list)
    l_items = s_list.get_items()
    ctx.load(l_items)
    ctx.execute_query()
    columnas=list(pd.DataFrame.from_dict(l_items[0].properties.items()).iloc[:,0])
    valores=list()
    for item in l_items:
        data=list(pd.DataFrame.from_dict(item.properties.items()).iloc[:,1])
        valores.append(data)
    resultado=pd.DataFrame(valores,columns=columnas)
    return resultado


contact = dataframeSP("Proveedores")

Accionistas = dataframeSP("Accionista")

Miembros = dataframeSP("Miembros")

contactPrinc = contact[contact.ID == idRegister]

contactColumns = pd.Series(list(contactPrinc.columns))


persona = str(contactPrinc.iat[0,17])

tipoSolicitud = str(contactPrinc.iat[0,15])

tipoPersona = str(contactPrinc.iat[0,19])

email = str(contactPrinc.iat[0,30])

primerNombre = str(contactPrinc.iat[0,20])

segundoNombre = str(contactPrinc.iat[0,21])

primerApellido = str(contactPrinc.iat[0,22])

segundoApellido= str(contactPrinc.iat[0,23])

tipoDocumentPer = str(contactPrinc.iat[0,24])

documentNIT  = str(contactPrinc.iat[0,36])

lugarNacimientoPer = str(contactPrinc.iat[0,25])

lugarExpedicionPerNat = str(contactPrinc.iat[0,28])

ocupacionOficioPerNat = str(contactPrinc.iat[0,29])

celularPerNat = str(contactPrinc.iat[0,32])

antiguedadEmpresa = str(contactPrinc.iat[0,31])

direccion = str(contactPrinc.iat[0,34])

telefonofijo = str(contactPrinc.iat[0,33])

pais = str(contactPrinc.iat[0,156])

Ciudad  = str(contactPrinc.iat[0,155])

fecha = str(contactPrinc.iat[0,16])

fechaNacimientoPer = str(contactPrinc.iat[0,26])

fechaExpedicionPerNat = str(contactPrinc.iat[0,27])

razonSocial = str(contactPrinc.iat[0,35])

tieneAlgunaCertificacion  = str(contactPrinc.iat[0,42])

estaObligadoATenerUnsistemaDePre  = str(contactPrinc.iat[0,39])

estaObligadoAUnaLicenciaOPermiso = str(contactPrinc.iat[0,39])

representacion = str(contactPrinc.iat[0,30])

primerNombreRf = str(contactPrinc.iat[0,158])

segundoNombreRf = str(contactPrinc.iat[0,159])

primerApellidoRf = str(contactPrinc.iat[0,160])

segundoApellidoRf = str(contactPrinc.iat[0,161])

correoRf = str(contactPrinc.iat[0,162])

numeroRf = str(contactPrinc.iat[0,163])

regimenTributario = contactPrinc.iat[0,101]

obligaciones = contactPrinc.iat[0,102]

if obligaciones == None:
    
    obligaciones = ""
    
    ResponseObligaciones = obligaciones
    
else:
    
    ResponseObligaciones = ValuestoDictOb(obligaciones)
    


sector = str(contactPrinc.iat[0,127])

CodigoCIIU = str(contactPrinc.iat[0,100])

valueCodigoCIIU = str(CodigoCIIU[0:4])

tipoCuentaBancaria = str(contactPrinc.iat[0,103])

bancoSucursal = str(contactPrinc.iat[0,141])

numeroCuentaBancaria = str(contactPrinc.iat[0,104])

emailNotificacionPagoCuentaBanca = str(contactPrinc.iat[0,106])

primerNombreContactoPago  = str(contactPrinc.iat[0,169])

segundoNombreContactoPago = str(contactPrinc.iat[0,170])

primerApellidoContactoPago = str(contactPrinc.iat[0,171])

segundoApellidoContactoPago = str(contactPrinc.iat[0,172])

numeroDocContactoPago = str(contactPrinc.iat[0,173])

tipoDocumContactoPago = str(contactPrinc.iat[0,180])

primerNombreContactoRetencion = str(contactPrinc.iat[0,174])

segundoNombreContactoRetencion = str(contactPrinc.iat[0,175])

primerApellidoContactoRetencion = str(contactPrinc.iat[0,176])

segundoApellidoContactoRetencion = str(contactPrinc.iat[0,177])

numeroDocContactoRetencion = str(contactPrinc.iat[0,178])

emailEnvioCertificadoRetencionCu = str(contactPrinc.iat[0,107])

tTipoDocumContactoRetencion = str(contactPrinc.iat[0,179])

administraActivosVirtualesCripto = str(contactPrinc.iat[0,113])

cualOperacioInternacional = str(contactPrinc.iat[0,114])

enqPaises = str(contactPrinc.iat[0,115])

tipoOperacionesInternacionales = contactPrinc.iat[0,116]

if tipoOperacionesInternacionales == None:
    
    tipoOperacionesInternacionales = ""
    
    responseOperaciones = tipoOperacionesInternacionales
    
else:
    
    responseOperaciones = ValuestoDictOp(tipoOperacionesInternacionales)
    

nombreEntidadPublica = str(contactPrinc.iat[0,151])

poseeProductosEnMonedaExtranjera = str(contactPrinc.iat[0,117])

poseCuentasEnMonedaExtranjera = str(contactPrinc.iat[0,118])

ciudadOperacionInternacional = str(contactPrinc.iat[0,119])

montoEntidadInternacional = str(contactPrinc.iat[0,120])

productoInternacional = str(contactPrinc.iat[0,121])

numeroCuentaInternacional = str(contactPrinc.iat[0,122])

transacionMonedaExtrangera = str(contactPrinc.iat[0,142])

primerNombreSuplenteRL = str(contactPrinc.iat[0,77])

segundoNombreSuplenteRL = str(contactPrinc.iat[0,78])

primerApellidoSuplenteRL = str(contactPrinc.iat[0,79])

segundoApellidoSuplenteRL = str(contactPrinc.iat[0,80])

emailSuplenteRL = str(contactPrinc.iat[0,81])

numeroDocumentoSuplenteRL = str(contactPrinc.iat[0,82])

TipoDocumentoSuplenteRL = str(contactPrinc.iat[0,83])

ingresoMensual = str(contactPrinc.iat[0,108])

conceptos = str(contactPrinc.iat[0,109])

egresosMensuales = str(contactPrinc.iat[0,110])

activos = str(contactPrinc.iat[0,111])

pasivos = str(contactPrinc.iat[0,112])

patrimonio = str(contactPrinc.iat[0,140])

primerNombreRL = str(contactPrinc.iat[0,70])

segundoNombreRL = str(contactPrinc.iat[0,71])

primerApellidoRL = str(contactPrinc.iat[0,72])

segundoApellidoRL = str(contactPrinc.iat[0,73])

numeroDocumentoRL = str(contactPrinc.iat[0,74])

tipoDocumentoRL = str(contactPrinc.iat[0,75])

emailRL = str(contactPrinc.iat[0,76])

primerNombreResponsableSGSGT = str(contactPrinc.iat[0,56])

SegundoNombreResponsableSGSGT = str(contactPrinc.iat[0,57])

primerApellidoResponsableSGSGT = str(contactPrinc.iat[0,58])

segundoApellidoResponsableSGSGT = str(contactPrinc.iat[0,59])

cargoResponsableSGSGT= str(contactPrinc.iat[0,60])

emailResponsableSGSGT = str(contactPrinc.iat[0,61])

telefonoResponsableSGSGT = str(contactPrinc.iat[0,62])

primerNombreQuienLiquidaContrato = str(contactPrinc.iat[0,63])

segundoNombreQuienLiquidaContrato = str(contactPrinc.iat[0,64])

primerApellidoQuienLiquidaContrato = str(contactPrinc.iat[0,65])

segundoApellidoQuienLiquidaContrato = str(contactPrinc.iat[0,66])

cargoQuienLiquidaContrato = str(contactPrinc.iat[0,67])

correoQuienLiquidaContrato = str(contactPrinc.iat[0,68])

telefonoQuienLiquidaContrato = str(contactPrinc.iat[0,68])

funcionesPublicasDestacadasRl = str(contactPrinc.iat[0,165])

EjerceFuncionesDirectivasRL = str(contactPrinc.iat[0,166])

FuncionesPublicasProminentesyDesRL = str(contactPrinc.iat[0,167])

ManejaoAdministraRecursosPublicoRL = str(contactPrinc.iat[0,168])

funcionesPublicasDestacadas  = str(contactPrinc.iat[0,92])

ejerceFuncionesDirectivasEnUnaOr = str(contactPrinc.iat[0,93])

funcionesPublicasProminentesyDes = str(contactPrinc.iat[0,94])

manejaoAdministraRecursosPublico = str(contactPrinc.iat[0,95])



funcionesPublicasDestacadasRl = IsNone(funcionesPublicasDestacadasRl)

EjerceFuncionesDirectivasRL = IsNone(EjerceFuncionesDirectivasRL)

FuncionesPublicasProminentesyDesRL = IsNone(FuncionesPublicasProminentesyDesRL)

ManejaoAdministraRecursosPublicoRL = IsNone(ManejaoAdministraRecursosPublicoRL)

funcionesPublicasDestacadas  = IsNone(funcionesPublicasDestacadas)

ejerceFuncionesDirectivasEnUnaOr = IsNone(ejerceFuncionesDirectivasEnUnaOr)

funcionesPublicasProminentesyDes = IsNone(funcionesPublicasProminentesyDes)

manejaoAdministraRecursosPublico = IsNone(manejaoAdministraRecursosPublico)

nombresQuienLiquidaContrato = FormatNamePer(primerNombreQuienLiquidaContrato,segundoNombreQuienLiquidaContrato)

apellidosQuienLiquidaContrato = FormatSurnamePer(primerApellidoQuienLiquidaContrato,segundoApellidoQuienLiquidaContrato)

nombresSuplenteRl = FormatNamePer(primerNombreSuplenteRL,segundoNombreSuplenteRL)

apellidosSuplenteRl = FormatSurnamePer(primerApellidoSuplenteRL,segundoApellidoSuplenteRL)

tipoPersona = IsNone(tipoPersona)

email = IsNone(email)

tipoDocumentPer = IsNone(tipoDocumentPer)

documentNIT  = IsNone(documentNIT)

lugarNacimientoPer = IsNone(lugarNacimientoPer)

lugarExpedicionPerNat = IsNone(lugarExpedicionPerNat)

ocupacionOficioPerNat = IsNone(ocupacionOficioPerNat)

celularPerNat = IsNone(celularPerNat)

antiguedadEmpresa = IsNone(antiguedadEmpresa)

direccion = IsNone(direccion)

telefonofijo = IsNone(telefonofijo)

pais = IsNone(pais)

Ciudad  = IsNone(Ciudad)

razonSocial = IsNone(razonSocial)

representacion = IsNone(representacion)

correoRf = IsNone(correoRf)

numeroRf = IsNone(numeroRf)

regimenTributario =  IsNone(regimenTributario)


# CodigoCIIU =  IsNone(CodigoCIIU)

sector =  IsNone(sector)

administraActivosVirtualesCripto = IsNone(administraActivosVirtualesCripto)

cualOperacioInternacional = IsNone(cualOperacioInternacional)

enqPaises = IsNone(enqPaises)

# tipoOperacionesInternacionales = IsNone(tipoOperacionesInternacionales)

# responseOperaciones = IsNone(responseOperaciones)

nombreEntidadPublica = IsNone(nombreEntidadPublica)

patrimonio = IsNone(patrimonio)

poseeProductosEnMonedaExtranjera = IsNone(poseeProductosEnMonedaExtranjera)

poseCuentasEnMonedaExtranjera = IsNone(poseCuentasEnMonedaExtranjera)

ciudadOperacionInternacional = IsNone(ciudadOperacionInternacional)

montoEntidadInternacional = IsNone(montoEntidadInternacional)

productoInternacional = IsNone(productoInternacional)

numeroCuentaInternacional =IsNone(numeroCuentaInternacional)

tipoCuentaBancaria =  IsNone(tipoCuentaBancaria)

emailNotificacionPagoCuentaBanca =  IsNone(emailNotificacionPagoCuentaBanca)

emailEnvioCertificadoRetencionCu =  IsNone(emailEnvioCertificadoRetencionCu)

bancoSucursal =  IsNone(bancoSucursal)

numeroDocContactoPago =  IsNone(numeroDocContactoPago)

numeroDocContactoRetencion =  IsNone(numeroDocContactoRetencion)

ingresoMensual = IsNone(ingresoMensual)

conceptos = IsNone(conceptos)

egresosMensuales = IsNone(egresosMensuales)

activos = IsNone(activos)

pasivos = IsNone(pasivos)

patrimonio = IsNone(patrimonio)

numeroDocumentoRL = IsNone(numeroDocumentoRL)

tipoDocumentoRL = IsNone(tipoDocumentoRL)

emailRL = IsNone(emailRL)

emailSuplenteRL = IsNone(emailSuplenteRL)

numeroDocumentoSuplenteRL = IsNone(numeroDocumentoSuplenteRL)

TipoDocumentoSuplenteRL = IsNone(TipoDocumentoSuplenteRL)

cargoResponsableSGSGT = IsNone(cargoResponsableSGSGT)

emailResponsableSGSGT = IsNone(emailResponsableSGSGT)

telefonoResponsableSGSGT = IsNone(telefonoResponsableSGSGT)

cargoQuienLiquidaContrato  = IsNone(cargoQuienLiquidaContrato)

correoQuienLiquidaContrato  = IsNone(correoQuienLiquidaContrato)

telefonoQuienLiquidaContrato  = IsNone(telefonoQuienLiquidaContrato)


year = FormateDateYear(fecha)

month = FormateDateMonth(fecha)

day = FormateDateDay(fecha)


yearNacimientoPer = FormateDateYear(fechaNacimientoPer)

monthNacimientoPer = FormateDateMonth(fechaNacimientoPer)

dayNacimientoPer = FormateDateDay(fechaNacimientoPer)


yearExpedicionPerNat = FormateDateYear(fechaExpedicionPerNat)

monthExpedicionPerNat = FormateDateMonth(fechaExpedicionPerNat)

dayExpedicionPerNat = FormateDateDay(fechaExpedicionPerNat)


Fecha = ResponseUtc(year, month, day)

date_of_birth = ResponseUtc(yearNacimientoPer, monthNacimientoPer, dayNacimientoPer)

fecha_de_expedicion = ResponseUtc(yearExpedicionPerNat, monthExpedicionPerNat, dayExpedicionPerNat)

def fecha (Formato):
    
    if(Formato == ''):
        
        d =''
    else:
        d=int(Formato.timestamp()*1000)
    
    return d

TipoDocumentoSuplenteRL = IsNone(TipoDocumentoSuplenteRL)

tipoDocumContactoPago = IsNone(tipoDocumContactoPago)

valueCodigoCIIU = IsNone(valueCodigoCIIU)

tipoDocumContactoRetencion = IsNone(tTipoDocumContactoRetencion)

# TipoDocumentoSuplenteRL = ValueTipoDocument(TipoDocumentoSuplenteRL)

# tipoDocumContactoPago = ValueTipoDocument(tipoDocumContactoPago)

# tipoDocumContactoRetencion = ValueTipoDocument(tipoDocumContactoRetencion)

ResponseFecha = fecha(Fecha)

Responsedate_of_birth = fecha(date_of_birth)

Responsefecha_de_expedicion = fecha(fecha_de_expedicion)


nombresPersona = FormatNamePer(primerNombre,segundoNombre)

apellidosPersona = FormatSurnamePer(primerApellido,segundoApellido)


nombresRf =  FormatNamePer(primerNombreRf,segundoNombreRf)

apellidosRf = FormatSurnamePer(primerApellidoRf,segundoApellidoRf)


nombresContactoRetencion = FormatNamePer(primerNombreContactoRetencion,segundoNombreContactoRetencion)

apellidosContactoRetencion = FormatSurnamePer(primerApellidoContactoRetencion,segundoApellidoContactoRetencion)


nombresSuplenteRl = FormatNamePer(primerNombreSuplenteRL,segundoNombreSuplenteRL)

apellidosSuplenteRl = FormatSurnamePer(primerApellidoSuplenteRL,segundoApellidoSuplenteRL)


nombresRL = FormatNamePer(primerNombreRL,segundoNombreRL)

apellidosRL = FormatSurnamePer(primerApellidoRL,segundoApellidoRL)


nombresResponsableSGSGT = FormatNamePer(primerNombreResponsableSGSGT,SegundoNombreResponsableSGSGT)

apellidosResponsableSGSGT = FormatSurnamePer(primerApellidoResponsableSGSGT,segundoApellidoResponsableSGSGT)



nombresContactoRetencion = FormatNamePer(primerNombreContactoRetencion,segundoNombreContactoRetencion)

apellidosContactoRetencion = FormatSurnamePer(primerApellidoContactoRetencion,segundoApellidoContactoRetencion)


nombresContactoPago =  FormatNamePer(primerNombreContactoPago,segundoNombreContactoPago)

apellidosContactoPago = FormatSurnamePer(primerApellidoContactoPago,segundoApellidoContactoPago)


nombresContactoPago =  FormatNamePer(primerNombreContactoPago,segundoNombreContactoPago)

apellidosContactoPago = FormatSurnamePer(primerApellidoContactoPago,segundoApellidoContactoPago)



responseRegimen = ValuestoDictRg(regimenTributario)


ContactosId = []

# try:
   
if email != '' and persona == 'Persona natural':
    
    if tipoPersona == 'Contratista/Proveedor' or tipoPersona == 'Contratista' or tipoPersona == 'Proveedor':
        
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{email}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
    
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
            {
                
              "property": "email",
              
              "value": email
              
            },
            {
                
              "property": "persona",
              
              "value": tipoPersona 
            },
            
            {
                
              "property": "clase_de_persona",
              
              "value": persona
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresPersona
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosPersona
              
            },
            
            {
                
              "property": "tipo_de_documento",
              
              "value": tipoDocumentPer
              
            },
            {
                
            "property": "numero_del_documento",
            
            "value": documentNIT
            
            }, 
            
            {
                
              "property": "lugar_de_nacimiento",
              
              "value": lugarNacimientoPer
              
            },
            
            {
                
                "property": "lugar_de_expedicion",
                
                "value": lugarExpedicionPerNat
                
              },
            
              {
                  
                "property": "date_of_birth",
                
                "value": Responsedate_of_birth
              },
              
              {
                  
                "property": "fecha_de_expedicion",
                
                "value": Responsefecha_de_expedicion
                
              },
              
              {
                  
                "property": "profesion_ocupacion_oficio",
                
                "value": ocupacionOficioPerNat
                
              },
              # {
              #   "property": "jobtitle",
              #   "value": cargo
              # },
              {
                  
                "property": "mobilephone",
                
                "value": celularPerNat
                
              },
              
              {
                  
                "property": "ciudad",
                
                "value": Ciudad
                
              },
              
              {
                  
                "property": "address",
                
                "value": direccion
                
              },
              
            {
                
              "property": "pa_s",
              
              "value": pais
              
            },
            
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
              "property": "tipo_de_suscripcion",
              
              "value": tipoSolicitud
              
            },
            
            {
                
              "property": "obligaciones",#seleccion multiple
              
              "value": ResponseObligaciones
              
            },
            
            {
                
              "property": "ciiu__actividad_economica_segun_rut_",
              
              "value": valueCodigoCIIU
              
            },
            
            {
                
              "property": "sector",
              
              "value": sector
              
            },
            
            {
                
              "property": "tipo_de_cuenta",
              
              "value": tipoCuentaBancaria
              
            },
            
            {
                
              "property": "numero_de_cuenta",
              
              "value": numeroCuentaBancaria
              
            },
            
            {    
                
              "property": "banco_sucursal",
              
              "value": bancoSucursal
              
            },
            
            {
                
              "property": "correo_electronico_de_notificacion_de_pagos",
              
              "value": emailNotificacionPagoCuentaBanca
              
            },
            
            {
                
              "property": "correo_electronico_para_el_envio_de_certificados_de_retencion",
              
              "value": emailEnvioCertificadoRetencionCu
              
            },
            
            {
                
              "property": "administra_activos_virtuales_criptoactivos_",
              
              "value": administraActivosVirtualesCripto
              
            },
            
            {
            
              "property": "cuales_virtuales_criptoactivos_",
              
              "value": cualOperacioInternacional
              
            },
            
            {
                
              "property": "en_que_paises_realiza_transaccion_en_moneda_extranjera_",
              
              "value": enqPaises
              
            },
            
            {
                
              "property": "tipo_de_operaciones",##multiple
              
              "value": responseOperaciones
              
            },
            
            {
                
              "property": "nombre_de_la_entidad_publica_u_organizacion_internacional",
              
              "value": nombreEntidadPublica
              
            },
            

            {
                
             "property": "Patrimonio",
             
             "value": patrimonio
             
           },
            
           {
               
             "property": "realiza_transaccion_en_moneda_extranjera_",
             
             "value": transacionMonedaExtrangera
             
           },
           
           {
               
             "property": "posee_productos_en_moneda_extranjera_",
             
             "value":poseeProductosEnMonedaExtranjera
             
           },
           
           {
               
             "property": "posee_cuentas_en_moneda_extranjera_",
             
             "value": poseCuentasEnMonedaExtranjera
             
           },
           
           {
               
             "property": "ciudad_operaciones_internacionales",
             
             "value": ciudadOperacionInternacional
             
           },
           
           {
               
             "property": "producto",
             
             "value": productoInternacional
             
           },
           
           {
               
             "property": "numero_de_cuenta_operaciones_internacionales",
             
             "value": numeroCuentaInternacional
             
           },
           
           {
               
             "property": "monto_operaciones_internacionales",
             
             "value": montoEntidadInternacional
             
           },
           
           {
               
             "property": "regimen",
             
             "value": responseRegimen
             
           },
           
           {
               
             "property": "grupos_de_interes_1",
             
             "value": tipoPersona
             
           },
           
          ]
          
        },default=str)
        
        ContactoPrincipal = requests.post(data=data, url=url, headers=headers)
        
        ContactosId.append(ContactoPrincipal.json()["vid"])
        
        print(ContactoPrincipal.json()['vid'])
        
        print('Contacto persona natural  asociada  fue creado con éxito')
        
    elif tipoPersona == 'Accionista' or tipoPersona == 'Colaborador' or tipoPersona == 'Socio' or tipoPersona == 'Vendedor/Fideicomitente/Beneficiario de Área' or tipoPersona == 'Otro':
        
        
        #CONTACTO PERSONA NATURAL NO ASOCIADA
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{email}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
    
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
              
            {
                
              "property": "email",
              
              "value": email
              
            },
            
            # {
                
            #   "property": "persona",
              
            #   "value": tipoPersona 
              
            # },
            
            {
                
              "property": "clase_de_persona",
              
              "value": persona
              
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresPersona
              
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosPersona
              
            },
            
            {
                
              "property": "tipo_de_documento",
              
              "value": tipoDocumentPer
              
            },
            
            {
                
            "property": "numero_del_documento",
            
            "value": documentNIT
            
            }, 
            
            {
                
              "property": "lugar_de_nacimiento",
              
              "value": lugarNacimientoPer
              
            },
            
            {
                
               "property": "lugar_de_expedicion",
               
               "value": lugarExpedicionPerNat
               
             },
            
             {
                 
               "property": "date_of_birth",
               
               "value": Responsedate_of_birth
               
             },
             
             {
                 
               "property": "fecha_de_expedicion",
               
               "value": Responsefecha_de_expedicion
               
             },
             
             {
                 
               "property": "profesion_ocupacion_oficio",
               
               "value": ocupacionOficioPerNat
               
             },
             
              {
                  
                "property": "jobtitle",
                
                "value": "PERSONA NATURAL"
                
              },
              
              {
                  
                "property": "mobilephone",
                
                "value": celularPerNat
                
              },
              
              {
                  
                "property": "ciudad",
                
                "value": Ciudad
                
              },
              
              {
                  
                "property": "address",
                
                "value": direccion
                
              },
              
            {
                
              "property": "pa_s",
              
              "value": pais
              
            },
            
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
             "property": "tipo_de_suscripcion",
             
             "value": tipoSolicitud
             
           },
            
           {
            
             "property": "obligaciones",#seleccion multiple
             
             "value": ResponseObligaciones
             
           },
           
           {
               
             "property": "ciiu__actividad_economica_segun_rut_",
             
             "value": valueCodigoCIIU
             
           },
           
           {
               
             "property": "sector",
             
             "value": sector
             
           },
           
           {
               
             "property": "tipo_de_cuenta",
             
             "value": tipoCuentaBancaria
             
           },
           
           {
               
             "property": "numero_de_cuenta",
             
             "value": numeroCuentaBancaria
             
           },
           
           {    
               
             "property": "banco_sucursal",
             
             "value": bancoSucursal
           },
           
           {
               
             "property": "correo_electronico_de_notificacion_de_pagos",
             
             "value": emailNotificacionPagoCuentaBanca
             
           },
           
           {
               
             "property": "correo_electronico_para_el_envio_de_certificados_de_retencion",
             
             "value": emailEnvioCertificadoRetencionCu
             
           },
           
           {
               
             "property": "administra_activos_virtuales_criptoactivos_",
             
             "value": administraActivosVirtualesCripto
             
           },
           
           {
               
             "property": "cuales_virtuales_criptoactivos_",
             
             "value": cualOperacioInternacional
             
           },
           
           {
               
             "property": "en_que_paises_realiza_transaccion_en_moneda_extranjera_",
             
             "value": enqPaises
             
           },
           
           {
               
             "property": "tipo_de_operaciones",##multiple
             
             "value": responseOperaciones
             
           },
           
           {
               
              "property": "nombre_de_la_entidad_publica_u_organizacion_internacional",
              
              "value": nombreEntidadPublica
              
            },
           
         
           {
               
             "property": "posee_productos_en_moneda_extranjera_",
             
             "value":poseeProductosEnMonedaExtranjera
             
           },
           
           {
               
             "property": "posee_cuentas_en_moneda_extranjera_",
             
             "value": poseCuentasEnMonedaExtranjera
             
           },
           
           {
               
             "property": "ciudad_operaciones_internacionales",
             
             "value": ciudadOperacionInternacional
             
           },
           
           {
               
            "property": "producto",
            
            "value": productoInternacional
            
          },
           
          {
              
            "property": "numero_de_cuenta_operaciones_internacionales",
            
            "value": numeroCuentaInternacional
            
          },
          
          {
              
            "property": "monto_operaciones_internacionales",
            
            "value": montoEntidadInternacional
            
          },
          
          {
              
            "property": "regimen",
            
            "value": responseRegimen
            
          },
          
          {
              
            "property": "ingresos_mensuales",
            
            "value": ingresoMensual
            
          },
          
          {
              
              "property": "egresos_mensuales",
              
              "value": egresosMensuales
              
          },
          
          {
              
              "property": "activos",
              
              "value": activos
              
          },
          
          {
              
            "property": "Patrimonio",
            
            "value": patrimonio
            
          },
          
          {
              
             "property": "conceptos",
             
             "value": conceptos
             
           },
          
          {
              
              "property": "pasivos",
              
              "value":  pasivos
          },
          
          {
              
            "property": "realiza_transaccion_en_moneda_extranjera_",
            
            "value": transacionMonedaExtrangera
            
          },
          
          {
              
            "property": "grupos_de_interes_1",
            
            "value": tipoPersona
            
          }
          
          ]
        },default=str)
        
        ContactoPrincipal = requests.post(data=data, url=url, headers=headers)
        
        # ContactosId.append(ContactoPrincipal.json()["vid"])
        
        # print(ContactoPrincipal.json()['vid'])
        
        print('Contacto persona natural no asociada  fue creado con éxito')
        
else:
    
    if email != '' and persona == 'Persona jurídica':
        
        #este es el contacto comercial
        
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{email}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
        
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
              
            {
                
              "property": "email",
              
              "value": email
              
            },
            
            {
                
              "property": "persona",
              
              "value": tipoPersona
              
            },
            
            {
                
              "property": "clase_de_persona",
              
              "value": persona
              
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresPersona
              
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosPersona
              
            },
            
            {
                
              "property": "tipo_de_documento",
              
              "value": tipoDocumentPer
              
            },
            
            {
                
            "property": "numero_del_documento",
            
            "value": documentNIT
            
            }, 
            
            {
                
              "property": "jobtitle",
              
              "value": 'CONTACTO COMERCIAL'
            },
            
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
              "property": "tipo_de_suscripcion",
              
              "value": tipoSolicitud
              
            },
            
            {
                
              "property": "regimen",
              
              "value": responseRegimen
              
            },
            
            {
                
              "property": "grupos_de_interes_1",
              
              "value": tipoPersona
              
            }
            
          ]
          
        },default=str)
        
        ContactoPrincipal = requests.post(data=data, url=url, headers=headers)
        
        ContactosId.append(ContactoPrincipal.json()["vid"])
        
        print(ContactoPrincipal.text)
        
        print('se ha creado o actualizado el contacto comercial "la persona jurídica" ')
        


if emailNotificacionPagoCuentaBanca != '':
    
     url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailNotificacionPagoCuentaBanca}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
 
     headers = {}
 
     headers['Content-Type']= 'application/json'
 
     data=json.dumps({
         
       "properties": [
           
         {
             
           "property": "email",
           
           "value": emailNotificacionPagoCuentaBanca
           
         },
         
     
         {
             
           "property": "firstname",
           
           "value": nombresContactoPago
           
         },
         
         {
             
           "property": "lastname",
           
           "value": apellidosContactoPago
           
         },
         
          {
              
            "property": "tipo_de_documento",
            
            "value": tipoDocumContactoPago
            
          },
          
          {
              
          "property": "numero_del_documento",
          
          "value": numeroDocContactoPago
          
          }, 
          
          {
              
            "property": "jobtitle",
            
            "value": "CONTACTO DE PAGO"
            
           },
        
         {
             
           "property": "fecha",
           
           "value": ResponseFecha
           
          }
         
    
         # {
         #   "property": "tipo_de_suscripcion",
         #   "value": tipoSolicitud
         # }
       ]
       
     },default=str)
     
     ContactoPago = requests.post(data=data, url=url, headers=headers)
     
     ContactosId.append(ContactoPago.json()["vid"])
     # print(ContactoPago.json())
     
     print('Se ha creado el contacto de Pago con éxito ')
     
     print(ContactoPago.text)
     
     
if  emailEnvioCertificadoRetencionCu != '':
    
    
      url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailEnvioCertificadoRetencionCu}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
  
      headers = {}
  
      headers['Content-Type']= 'application/json'

      data=json.dumps({
          
        "properties": [
            
          {
              
            "property": "email",
            
            "value": emailEnvioCertificadoRetencionCu
            
          },
          
          {
              
            "property": "firstname",
            
            "value": nombresContactoRetencion
            
          },
          
          {
              
            "property": "lastname",
            
            "value": apellidosContactoRetencion
            
          },
          
           {
               
             "property": "tipo_de_documento",
             
             "value": tipoDocumContactoPago
             
           },
           
           {
               
           "property": "numero_del_documento",
           
           "value": numeroDocContactoRetencion
           
           }, 
           
           {
               
             "property": "jobtitle",
             
             "value": "CONTACTO DE CERTIFICADO DE RETENCIÓN"
             
            },
           
          {
              
            "property": "fecha",
            
            "value": ResponseFecha
            
           }
          
     
          # {
          #   "property": "tipo_de_suscripcion",
          #   "value": tipoSolicitud
          # }
        ]
      },default=str)
      
      ContactoRet = requests.post(data=data, url=url, headers=headers)
      
      ContactosId.append(ContactoRet.json()["vid"])
      
      print('Se ha creado el contacto de envio certificado de Retención ')
      
      print(ContactoPago.text)
      
if emailRL != '':
      
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailRL}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
    
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
              
            {
                
              "property": "email",
              
              "value": emailRL
              
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresRL
              
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosRL
              
            },
            
            {
                
              "property": "tipo_de_documento",
              
              "value": tipoDocumentoRL
              
            },
            
            {
                
            "property": "numero_del_documento",
            
            "value": numeroDocumentoRL
            
            }, 
            
            # {
            #   "property": "mobilephone",
            #   "value": 
            # },
               {
                   
                 "property": "jobtitle",
                 
                 "value": 'REPRESENTANTE LEGAL'
                 
               },
               
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
              "property": "tipo_de_suscripcion",
              
              "value": tipoSolicitud
              
            },
            
            {
                
              "property": "ejerce_funciones_directivas_en_una_organizacion_internacional_",
              
              "value": EjerceFuncionesDirectivasRL
              
            },
            
             {
                 
               "property": "desempena_funciones_publicas_destacadas___direccion_general__formulacion_de_politicas__adopcion_de_",
               
               "value": funcionesPublicasDestacadasRl
               
             },
             
            {
                
              "property": "desempena_funciones_publicas_prominentes_y_destacadas_en_otro_pais_",
              
              "value": FuncionesPublicasProminentesyDesRL
              
            },
            
            {
                
              "property": "maneja_o_administra_recursos_publicos_",
              
              "value": ManejaoAdministraRecursosPublicoRL
              
            }
            
          ]
              
        },default=str)
            
        ContactoRl = requests.post(data=data, url=url, headers=headers)
        
        ContactosId.append(ContactoRl.json()["vid"])
        
        print('se ha creado el contacto  del representante legal fue creado con éxito')
        
        print(ContactoRl.text)
        
if emailResponsableSGSGT != '':
        
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailResponsableSGSGT}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
    
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
              
            {
                
              "property": "email",
              
              "value": emailResponsableSGSGT
              
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresResponsableSGSGT
              
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosResponsableSGSGT
              
            },
            
            # {
            #   "property": "tipo_de_documento",
            #   "value": 
            # },
            # {
            # "property": "numero_del_documento",
            # "value": telefonoResponsableSGSGT
            # }, 
            {
                
              "property": "mobilephone",
              
              "value": telefonoResponsableSGSGT
              
            },
            
               {
                   
                 "property": "jobtitle",
                 
                 "value": 'RESPONSABLE SG-SGT '
                 
               },
               
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
              "property": "tipo_de_suscripcion",
              
              "value": tipoSolicitud
              
            }
            
          ]
              
        },default=str)
            
        ContactoSG = requests.post(data=data, url=url, headers=headers)
        
        ContactosId.append(ContactoSG.json()["vid"])
        
        print('se ha creado el contacto  Responsable SG-SGT fue creado con éxito')
        
        print(ContactoSG.json())
        

if emailSuplenteRL != '':
    
       url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailSuplenteRL}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
   
       headers = {}
   
       headers['Content-Type']= 'application/json'
   
       data=json.dumps({
           
         "properties": [
             
           {
               
             "property": "email",
             
             "value": emailSuplenteRL
             
           },
       
           {
               
             "property": "firstname",
             
             "value": nombresSuplenteRl
             
           },
           
           {
               
             "property": "lastname",
             
             "value": apellidosSuplenteRl
             
           },
           
           {
               
             "property": "tipo_de_documento",
             
             "value": TipoDocumentoSuplenteRL
             
           },
           
           {
               
           "property": "numero_del_documento",
           
           "value": numeroDocumentoSuplenteRL
           
           }, 
           
           # {
           #   "property": "mobilephone",
           #   "value": numeroRf
           # },
              {
                  
                "property": "jobtitle",
                
                "value": 'SUPLENTE REVISOR  FISCAL'
                
              },
              
           {
               
             "property": "fecha",
             
             "value": ResponseFecha
             
           },
           
           {
               
             "property": "tipo_de_suscripcion",
             
             "value": tipoSolicitud
             
           },
           
           {
               
             "property": "ejerce_funciones_directivas_en_una_organizacion_internacional_",
             
             "value": ejerceFuncionesDirectivasEnUnaOr
             
           },
           
            {
                
              "property": "desempena_funciones_publicas_destacadas___direccion_general__formulacion_de_politicas__adopcion_de_",
              
              "value": funcionesPublicasDestacadas
              
            },
            
           {
               
             "property": "desempena_funciones_publicas_prominentes_y_destacadas_en_otro_pais_",
             
             "value": funcionesPublicasProminentesyDes
             
           },
           
           {
               
             "property": "maneja_o_administra_recursos_publicos_",
             
             "value": manejaoAdministraRecursosPublico
             
           }
           
         ]
         
       },default=str)
       
       ContactoSuplente = requests.post(data=data, url=url, headers=headers)
       
       ContactosId.append(ContactoSuplente.json()["vid"])
       
       print('Se ha creado el contacto del suplente del representante legal ')
       
       print(ContactoSuplente.text)
       

if correoQuienLiquidaContrato != '':
    
       url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{correoQuienLiquidaContrato}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
   
       headers = {}
   
       headers['Content-Type']= 'application/json'
   
       data=json.dumps({
           
         "properties": [
             
           {
               
             "property": "email",
             
             "value": correoQuienLiquidaContrato
             
           },
           
           {
               
             "property": "firstname",
             
             "value": nombresQuienLiquidaContrato
             
           },
           
           # {
           #   "property": "lastname",
           #   "value": apellidosSuplenteRl
           # },
           # {
           #   "property": "tipo_de_documento",
           #   "value": TipoDocumentoSuplenteRL
           # },
           # {
           # "property": "numero_del_documento",
           # "value": numeroDocumentoSuplenteRL
           # }, 
            {
                
              "property": "mobilephone",
              
              "value": telefonoQuienLiquidaContrato
              
            },
            
              {
                  
                "property": "jobtitle",
                
                "value": cargoQuienLiquidaContrato
                
              },
              
           {
               
             "property": "fecha",
             
             "value": ResponseFecha
             
           },
           
           {
               
             "property": "tipo_de_suscripcion",
             
             "value": tipoSolicitud
             
           }
           
         ]
         
       },default=str)
       
       ContactoResLQ = requests.post(data=data, url=url, headers=headers)
       
       ContactosId.append(ContactoResLQ.json()["vid"])
       
       print('Se ha creado el contacto del responsable de liquidacion de contratos ')
       
       print(ContactoResLQ.text)
       

if correoRf != '':
    
        url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{correoRf}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
    
        headers = {}
    
        headers['Content-Type']= 'application/json'
    
        data=json.dumps({
            
          "properties": [
              
            {
                
              "property": "email",
              
              "value": correoRf
              
            },
            
            {
                
              "property": "firstname",
              
              "value": nombresRf
              
            },
            
            {
                
              "property": "lastname",
              
              "value": apellidosRf
              
            },
            
            # {
            #   "property": "tipo_de_documento",
            #   "value": tipoDocumentPer
            # },
            # {
            # "property": "numero_del_documento",
            # "value": documentNIT
            # }, 
            {
                
              "property": "mobilephone",
              
              "value": numeroRf
              
            },
            
               {
                   
                 "property": "jobtitle",
                 
                 "value": 'REVISOR FISCAL'
                 
               },
               
            {
                
              "property": "fecha",
              
              "value": ResponseFecha
              
            },
            
            {
                
              "property": "tipo_de_suscripcion",
              
              "value": tipoSolicitud
              
            }
            
          ]
        },default=str)
        
        ContactoRF = requests.post(data=data, url=url, headers=headers)
        
        ContactosId.append(ContactoRF.json()["vid"])
        
        print('se ha creado el contacto  del revisor fiscal con éxito')
        
        print(ContactoRF.text)
        

try:
    
    for i in range(len(Accionistas)):
        
        
            Accionistas = Accionistas[Accionistas.IdAccionista == razonSocial]
            
            # columnsAccionistas = pd.Series(list(Accionistas.columns))
            primerNombre = str(Accionistas.iat[i,15])
            
            segundoNombre = str(Accionistas.iat[i,16])
            
            primerApellido = str(Accionistas.iat[i,17])
            
            segundoApellido = str(Accionistas.iat[i,18])
            
            documento = str(Accionistas.iat[i,19])
            
            emailAC = str(Accionistas.iat[i,21])
            
            tipoDoc = str(Accionistas.iat[i,23])
            
            nombreEmpresa = str(Accionistas.iat[i,26])
            
            nombresAcc= FormatNamePer(primerNombre,segundoNombre)
            
            apellidosAcc = FormatSurnamePer(primerApellido,segundoApellido)
            
            nombresAcc =  IsNone(nombresAcc)
            
            apellidosAcc =  IsNone(apellidosAcc)
            
            documento =  IsNone(documento)
            
            tipoDoc =  IsNone(tipoDoc)
            
            emailAC =  IsNone(emailAC)
            
            
            if emailAC != '':
                
                url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailAC}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
            
                headers = {}
            
                headers['Content-Type']= 'application/json'
            
                data=json.dumps({
                    
                  "properties": [
                      
                      
                    #SGSGT contacto
                    {
                        
                      "property": "firstname",
                      
                      "value": nombresAcc
                      
                    },
                    
                    {
                        
                      "property": "lastname",
                      
                      "value": apellidosAcc
                      
                    },
                    
                    {
                        
                      "property": "company",
                      
                      "value": nombreEmpresa
                      
                    },
                    
                    {
                        
                      "property": "tipo_de_documento",
                      
                      "value": tipoDoc
                      
                    },
                    
                    {
                        
                    "property": "numero_del_documento",
                    
                    "value": documento
                    
                  },           
                    
                    {
                        
                      "property": "email",
                      
                      "value": emailAC
                      
                    },
                    
                    {
                        
                      "property": "jobtitle",
                      
                      "value": 'ACCIONISTA'
                      
                    },
                    
                  ]
                  
                },default=str)
                
                responseContacto2 = requests.post(data=data, url=url, headers=headers)
                
                print(responseContacto2.text)
                
                ContactosId.append(responseContacto2.json()["vid"])
                
                print(f'Accionista en posicion {i} registrado exitosamente')
                
                print(nombresAcc,apellidosAcc)
                
            else:
                
                print('No se registraron  Accionistas')
                

except Exception:
    
    print('No se encuantran Accionistas para registrar')

try:
    # CONTACTOS DE MIEMBROS
    Miembros = Miembros[Miembros.NombreEmpresa == razonSocial]
    
    for i in range(len(Miembros)):
        
        primerNombreM = str(Miembros.iat[i,15])
        
        segundoNombreM = str(Miembros.iat[i,16])
        
        primerApellidoM = str(Miembros.iat[i,17])
        
        segundoApellidoM = str(Miembros.iat[i,18])
        
        emailM = str(Miembros.iat[i,19])
        
        DocumentoM = str(Miembros.iat[i,20])
        
        tipoDocM = str(Miembros.iat[i,23])
        
        nombreEmpresaM = str(Miembros.iat[i,24])
        
        nombresM= FormatNamePer(primerNombreM,segundoNombreM)
        
        apellidosM = FormatSurnamePer(primerApellidoM,segundoApellidoM)
        
        
        responseApellidosM =  IsNone(apellidosM)
        
        responseNombresM =  IsNone(nombresM)
        
        nombreEmpresaM =  IsNone(nombreEmpresaM)
        
        tipoDocM =  IsNone(tipoDocM)
        
        DocumentoM =  IsNone(DocumentoM)
        
        emailM =  IsNone(emailM)
        
        
        if emailM != '':
       
                url =f"https://api.hubapi.com/contacts/v1/contact/createOrUpdate/email/{emailM}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
            
                headers = {}
            
                headers['Content-Type']= 'application/json'
            
                data=json.dumps({
                    
                  "properties": [
                      
                    {
                        
                      "property": "firstname",
                      
                      "value": responseNombresM
                      
                    },
                    
                    {
                        
                      "property": "lastname",
                      
                      "value": responseApellidosM
                      
                    },
                    
                    {
                        
                      "property": "company",
                      
                      "value": nombreEmpresaM
                      
                    },
                    
                    {
                        
                      "property": "tipo_de_documento",
                      
                      "value": tipoDocM
                      
                    },
                    
                    {
                        
                    "property": "numero_del_documento",
                    
                    "value": DocumentoM
                    
                  },           
                    
                    {
                        
                      "property": "email",
                      
                      "value": emailM
                      
                    },
                    
                    {
                        
                      "property": "jobtitle",
                      
                      "value": 'MIEMBRO'
                      
                    },
                    
                  ]
                  
                },default=str)
                
                responseContactMiembro = requests.post(data=data, url=url, headers=headers)
                
                print(responseContacto2.text)
                
                ContactosId.append(responseContactMiembro.json()["vid"])
                
                print(f'Miembro en posicion {i} registrado exitosamente')
                

        else:
            
            print('No se registraron Miembros')

except Exception:
    
    print('No se registraron Miembros')
    
        
# except:
#         print('No hay contacto registrado')
try:
    
    GET_Id = ContactoPrincipal.json()["vid"]
    
except:    
    
    print(ContactoPrincipal.json())
    
GetContact = requests.request("GET",f"https://api.hubapi.com/contacts/v1/contact/vid/{GET_Id}/profile?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1")
   
if 'associated-company' in GetContact.json():
    
   CompanyId = GetContact.json()['associated-company']['company-id']
   
else:
    
   CompanyId =''
  
   
CompanyId = str(CompanyId) 

print(type(CompanyId))

   
if email != '':
    
     Dominio = re.findall('@.*',email)[0].translate(str.maketrans('', '', '@'))
     
     if CompanyId == '' and persona == 'Persona jurídica':
              
           url = "https://api.hubapi.com/companies/v2/companies"
           
           payload = json.dumps(
               
           {
               
             "properties": [
                 
               {
                   #info bancaria
                 "name": "razon_social__p__juridica__o_nombres_y_apellidos_p__natural_",
                 
                 "value": razonSocial
                 
               },
               {    
                 "name": "name",
                 
                 "value": razonSocial
                 
               },
               
               {
                   
                 "name": "nit",
                 
                 "value": documentNIT
                 
               },
               
               {
                   
                 "name": "tipo_de_cuenta",
                 
                 "value": tipoCuentaBancaria
                 
               },
               
               {
                   
                 "name": "numero_de_cuenta",
                 
                 "value": numeroCuentaBancaria
                 
               },
               
               {    
                   
                 "name": "banco_sucursal",
                 
                 "value": bancoSucursal
                 
               },
               ##############info internacional a persona juridica
               {
                   
                 "name": "administra_activos_virtuales_criptoactivos_",
                 
                 "value": administraActivosVirtualesCripto
                 
               },
             #  {
             #    "property": "cuales_virtuales_criptoactivos_",
             #    "value": cualOperacioInternacional
             #  },
             #  {
             #    "property": "en_que_paises_realiza_transaccion_en_moneda_extranjera_",
             #    "value": enqPaises
             #  },
               {
                   
                 "name": "tipo_de_operaciones",##multiple
                 
                 "value": responseOperaciones
                 
               },
               
             #  {
             #    "property": "nombre_de_la_entidad_publica_u_organizacion_internacional",
             #    "value": nombreEntidadPublica
             #  },
             # {
             #   "property": "Patrimonio",
             #   "value": patrimonio
             # },
              {
                  
                "name": "realiza_transaccion_en_moneda_extranjera_",
                
                "value": transacionMonedaExtrangera
                
              },
              
              {
                  
                "name": "posee_productos_en_moneda_extranjera_",
                
                "value":poseeProductosEnMonedaExtranjera
                
              },
              
              {
                  
                "name": "posee_cuentas_en_moneda_extranjera_",
                
                "value": poseCuentasEnMonedaExtranjera
                
              },
              
              {
                  
                "name": "country",
                
                "value": pais
                
              },
              
              {
                  
                "name": "city",
                
                "value": Ciudad
                
              },
              
               {
                   
                 "name": "domain",
                 
                 "value": Dominio
                 
               },
               {
                   
                 "name": "address",
                 
                 "value": direccion
                 
               }
               
              
              # {
              #   "name": "ciudad_operaciones_internacionales",
              #   "value": Ciudad
              # },
              # {
              #   "name": "posee_productos_en_moneda_extranjera_",
              #   "value": productoInternacional
              # },
              # {
              #   "name": "numero_de_cuenta_operaciones_internacionales",
              #   "value": numeroCuentaInternacional
              # },
              # {
              #   "name": "recent_deal_amount",
              #   "value": montoEntidadInternacional
              # },
             #  {
             #    "property": "correo_electronico_de_notificacion_de_pagos",
             #    "value": emailNotificacionPagoCuentaBanca
             #  },
             #  {
             #    "property": "correo_electronico_para_el_envio_de_certificados_de_retencion",
             #    "value": emailEnvioCertificadoRetencionCu
             #  }
             ]
             
           }
           
           );
           
           headers = {
               
               'Content-Type': "application/json",
               
               }
           
           r = requests.request("POST", url, data=payload, headers=headers, params=querystring)
           
           r.json()
           
           companyId = r.json()["companyId"]
           
           print(type(companyId))
           
           print('Se ha creado la empresa de acuerdo a la persona jurídica')
           
           
           for i in ContactosId:
               
               url = f"https://api.hubapi.com/companies/v2/companies/{companyId}/contacts/{i}"
               
               asociar = requests.request("PUT", url=url, params=querystring)
               
               print(asociar)
               
     if CompanyId != '' and persona == 'Persona jurídica':
         
               CompanyId = int(CompanyId)
               
               url = f"https://api.hubapi.com/companies/v2/companies/{CompanyId}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
               
               payload = json.dumps(
                   
               {
                   
                 "properties": [
                     
                       {#info bancaria
                        
                         "name": "razon_social__p__juridica__o_nombres_y_apellidos_p__natural_",
                         
                         "value": razonSocial
                         
                       },
                       
                       {    
                         "name": "name",
                         
                         "value": razonSocial
                         
                       },
                       
                       {
                           
                         "name": "nit",
                         
                         "value": documentNIT
                         
                       },
                       
                       {
                           
                         "name": "tipo_de_cuenta",
                         
                         "value": tipoCuentaBancaria
                         
                       },
                       
                       {
                           
                         "name": "numero_de_cuenta",
                         
                         "value": numeroCuentaBancaria
                         
                       },
                       
                       {    
                           
                         "name": "banco_sucursal",
                         
                         "value": bancoSucursal
                         
                       },
                       ##############info internacional a persona juridica
                       {
                           
                         "name": "administra_activos_virtuales_criptoactivos_",
                         
                         "value": administraActivosVirtualesCripto
                         
                       },
                       
                       {
                           
                         "name": "country",
                         
                         "value": pais
                         
                       },
                       
                       {
                           
                         "name": "city",
                         
                         "value": Ciudad
                         
                       },
                       
                        {
                            
                          "name": "domain",
                          
                          "value": Dominio
                          
                        },
                        {
                            
                          "name": "address",
                          
                          "value": direccion
                          
                        }
                        
                       # {
                       #   "name": "realiza_transaccion_en_moneda_extranjera_",
                       #   "value": transacionMonedaExtrangera
                       # },
                       # {
                       #   "name": "posee_productos_en_moneda_extranjera_",
                       #   "value":poseeProductosEnMonedaExtranjera
                       # },
                       # {
                       #   "name": "posee_cuentas_en_moneda_extranjera_",
                       #   "value": poseCuentasEnMonedaExtranjera
                       # },
                       # {
                       #   "name": "ciudad_operaciones_internacionales",
                       #   "value": Ciudad
                       # },

                       # {
                       #   "name": "numero_de_cuenta_operaciones_internacionales",
                       #   "value": numeroCuentaInternacional
                       # },
                       # {
                       #   "name": "recent_deal_amount",
                       #   "value": montoEntidadInternacional
                       # },
                       # {
                       #   "name": "tipo_de_operaciones",##multiple
                       #   "value": responseOperaciones
                       # },
                       
                 ]
                 
               }
               
               );
               
               headers = {
                   
                   'Content-Type': "application/json",
                   
                   }
               
               r = requests.request("PUT", url, data=payload, headers=headers, params=querystring)
               
               r.json()
               
               print('actualizacion exitosa ')
               
               companyId = r.json()["companyId"]
               
               print(type(companyId))
               
               for i in ContactosId:
                   
                   url = f"https://api.hubapi.com/companies/v2/companies/{companyId}/contacts/{i}"
                   
                   asociar = requests.request("PUT", url=url, params=querystring)
                   
                   print(asociar)
                   
       
     if CompanyId == '' and persona == 'Persona natural':
         
         if tipoPersona == 'Contratista/Proveedor' or tipoPersona == 'Contratista' or tipoPersona == 'Proveedor':
           url = "https://api.hubapi.com/companies/v2/companies"
           
           querystring = {"hapikey":"5a1e6e50-46c5-47f8-aa2d-f35be24978a1"}
           
           payload = json.dumps(
               
           {
               
             "properties": [
                 
                   {#info bancaria
                    
                     "name": "razon_social__p__juridica__o_nombres_y_apellidos_p__natural_",
                     
                     "value": nombresPersona
                     
                   },
                   
                    {
                        
                      "name": "nit",
                      
                      "value": documentNIT
                      
                    },
                    
                   {
                       
                     "name": "tipo_de_cuenta",
                     
                     "value": tipoCuentaBancaria
                     
                   },
                   
                   {
                       
                     "name": "numero_de_cuenta",
                     
                     "value": numeroCuentaBancaria
                     
                   },
                   
                   {    
                       
                     "name": "banco_sucursal",
                     
                     "value": bancoSucursal
                     
                   },##############info internacional a persona juridica
                   
                   {
                       
                     "name": "administra_activos_virtuales_criptoactivos_",
                     
                     "value": administraActivosVirtualesCripto
                     
                   },
                   
                   {
                       
                     "name": "country",
                     
                     "value": pais
                     
                   },
                   
                   {
                       
                     "name": "city",
                     
                     "value": Ciudad
                     
                   },
                   
                    {
                        
                      "name": "domain",
                      
                      "value": Dominio
                      
                    }
                    
             ]
             
           }
           
           );
           
           headers = {
               
               'Content-Type': "application/json",
               
               }
           
           r = requests.request("POST", url, data=payload, headers=headers, params=querystring)
           
           r.json()
           
           companyId = r.json()["companyId"]
           
           print(type(companyId))
           
           print('Se ha creado la empresa de acuerdo a la persona juridica')
           
           for i in ContactosId:
               
               url = f"https://api.hubapi.com/companies/v2/companies/{companyId}/contacts/{i}"
               
               asociar = requests.request("PUT", url=url, params=querystring)
               
               print(asociar)
               
     if CompanyId != '' and persona == 'Persona natural':
         
         if tipoPersona == 'Contratista/Proveedor' or tipoPersona == 'Contratista' or tipoPersona == 'Proveedor':
           CompanyId = int(CompanyId)
           
           url = f"https://api.hubapi.com/companies/v2/companies/{CompanyId}?hapikey=5a1e6e50-46c5-47f8-aa2d-f35be24978a1"
           
           payload = json.dumps(
               
           {
               
             "properties": [
                 
                   {#info bancaria
                    
                     "name": "razon_social__p__juridica__o_nombres_y_apellidos_p__natural_",
                     
                     "value": nombresPersona
                     
                   },
                   
                   {
                       
                     "name": "name",
                     
                     "value": nombresPersona
                     
                   },
                   
                    {
                        
                      "name": "nit",
                      
                      "value": documentNIT
                      
                    },
                    
                   {
                       
                     "name": "name",
                     
                     "value": nombresPersona
                     
                   },
                   
                   # informacion financiera 
                   {
                       
                     "name": "ingresos_mensuales",
                     
                     "value": ingresoMensual
                     
                   },
                   
                   {
                       
                     "name": "egresos_mensuales",
                     
                     "value": egresosMensuales
                     
                   },
                   
                   {
                       
                     "name": "activos",
                     
                     "value": activos
                     
                   },
                   
                   {
                       
                     "name": "patrimonio",
                     
                     "value": patrimonio
                     
                   },
                   
                    {
                        
                      "name": "domain",
                      
                      "value": Dominio
                      
                    }
             ]
             
           }
           
           );
           
           headers = {
               
               'Content-Type': "application/json",
               
               }
           
           
           r = requests.request("PUT", url, data=payload, headers=headers, params=querystring)
           
           r.json()
           
           print('Se ha actualizado la empresa de la persona natural')
	   
# except Exception:
#      print('fin')