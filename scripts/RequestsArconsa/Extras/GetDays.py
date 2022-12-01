import holidays_co
import warnings
import datetime
from Extras.Colorama import bcolors
import pandas as pd
from datetime import timedelta

warnings.filterwarnings("ignore")
fechaActual = datetime.date.today()
diaActualSemana = datetime.date.today().strftime('%A')
diaFestivoFT = holidays_co.is_holiday_date(fechaActual)
diaAnteriorFestivoFT = holidays_co.is_holiday_date(fechaActual-timedelta(days=1))
diaViernes = holidays_co.is_holiday_date(fechaActual-timedelta(days=3))

def GetProject(p, day):
    Proyectos = list()
    for i in p.index:
        diaSucio = p['Dias'][i]
        dias = diaSucio.split('/')
        for dia in dias:
            if dia.capitalize() == day:
                Proyectos.append(p['# Obra'][i])
    return Proyectos

def getDays(proveedores):
    Pfes = []
    # Define el horario de las obras
    if (diaActualSemana == "Monday" and diaFestivoFT == False):
        Proyectos  = GetProject(proveedores, 'Lunes') #Lunes
    if(diaActualSemana == "Tuesday" and diaFestivoFT == False):
        Proyectos  = GetProject(proveedores, 'Martes')#Martes
    if(diaActualSemana == "Wednesday" and diaFestivoFT == False):
        Proyectos  = GetProject(proveedores, 'Miercoles') #miercoles
    if(diaActualSemana == "Thursday" and diaFestivoFT == False):
        Proyectos  = GetProject(proveedores, 'Jueves') #jueves
    if(diaActualSemana == "Friday" and diaFestivoFT == False):
        Proyectos  = GetProject(proveedores, 'Viernes') #viernes
    if(diaActualSemana == 'Monday' and diaViernes == True):
        Pfes=(GetProject(proveedores, 'Viernes')) #Viernes
    if(diaActualSemana == 'Tuesday' and diaAnteriorFestivoFT == True):
        Pfes=(GetProject(proveedores, 'Lunes')) #Lunes
    if(diaActualSemana == 'Wednesday' and diaAnteriorFestivoFT == True):
        Pfes=(GetProject(proveedores, 'Martes')) #Martes
    if(diaActualSemana == 'Thursday' and diaAnteriorFestivoFT == True):
        Pfes=(GetProject(proveedores, 'Miercoles')) #Miercoles
    if(diaActualSemana == 'Friday' and diaAnteriorFestivoFT == True):
        Pfes=(GetProject(proveedores, 'Jueves')) #Jueves
    if(diaActualSemana == 'Saturday' or diaActualSemana == 'Sunday'):
        raise RuntimeError('Este script solo puede ser ejecutado de lunes a viernes')
    if diaFestivoFT:
        raise RuntimeError('Este script no se puede ser ejecutado los días festivos, las obras de este día serán ejecutadas el proximo día')
    Proyectos = Proyectos+Pfes
    return Proyectos