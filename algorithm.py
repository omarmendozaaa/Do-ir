import pandas as pd
from fuzzywuzzy import process, fuzz
from openpyxl import Workbook, load_workbook
import tkinter as tk

registro_oficial = 'registro_oficial.xlsx'

registro = load_workbook(registro_oficial)

def medio(find, file, sheet_name_ ):
    medios_array =  pd.read_excel(file, sheet_name=sheet_name_)
    tupla = medios_array.loc[medios_array['Unnamed: 1'] == find]
    if tupla['Unnamed: 3'].values[0] == 'WHATSAPP':
        return 'W'
    elif tupla['Unnamed: 3'].values[0] == 'PLATAFORMA ZOOM':
        return 'Z'
    else:
        return 'GM'

def posicion(semana,dia):
    dias =  (int(semana) - 1)*5
    separaciones= int(semana) -1
    columna = 67 + dias + separaciones + dia - 1
    return str(chr(columna))

def algorith(file, file2, semana, dia):
    primero_con = pd.read_excel(file, sheet_name='PRIMERO')
    cuarto_con = pd.read_excel(file, sheet_name='CUARTO')
    quinto_con = pd.read_excel(file, sheet_name='QUINTO')

    r_caridad =  pd.read_excel(file2, sheet_name='1° CARIDAD')
    r_optimismo =  pd.read_excel(file2, sheet_name='1° OPTIMISMO')
    r_honradez =  pd.read_excel(file2, sheet_name='1° HONRADEZ')
    r_fraternidad =  pd.read_excel(file2, sheet_name='4° FRATERNIDAD')
    r_responsabilidad =  pd.read_excel(file2, sheet_name='4° RESPONSABILIDAD')
    r_integracion =  pd.read_excel(file2, sheet_name='4° INTEGRACIÓN')
    r_tolerancia =  pd.read_excel(file2, sheet_name='4° TOLERANCIA')
    r_honestidad =  pd.read_excel(file2, sheet_name='5° HONESTIDAD')
    r_laboriosidad =  pd.read_excel(file2, sheet_name='5° LABORIOSIDAD')
    r_superacion =  pd.read_excel(file2, sheet_name='5° SUPERACIÓN')
    r_respeto =  pd.read_excel(file2, sheet_name='5° RESPETO')

    values_primero = primero_con['Unnamed: 1'].values
    values_cuarto = cuarto_con['Unnamed: 1'].values
    values_quinto = quinto_con['Unnamed: 1'].values

    caridad = r_caridad['Unnamed: 1'].values
    optimismo =  r_optimismo['Unnamed: 1'].values
    honradez =  r_honradez['Unnamed: 1'].values
    fraternidad =  r_fraternidad['Unnamed: 1'].values
    responsabilidad =  r_responsabilidad['Unnamed: 1'].values
    integracion =  r_integracion['Unnamed: 1'].values
    tolerancia =  r_tolerancia['Unnamed: 1'].values
    honestidad =  r_honestidad['Unnamed: 1'].values
    laboriosidad =  r_laboriosidad['Unnamed: 1'].values
    superacion =  r_superacion['Unnamed: 1'].values
    respeto =  r_respeto['Unnamed: 1'].values

    secciones = [(caridad, "1° CARIDAD", 1), (optimismo, "1° OPTIMISMO", 1), (honradez, "1° HONRADEZ", 1),
    (fraternidad, "4° FRATERNIDAD", 4), (responsabilidad, "4° RESPONSABILIDAD", 4),(integracion, "4° INTEGRACIÓN", 4),(tolerancia, "4° TOLERANCIA", 4),
    (honestidad, "5° HONESTIDAD", 5),(laboriosidad, "5° LABORIOSIDAD", 5),(superacion, "5° SUPERACIÓN", 5),(respeto, "5° RESPETO", 5)]


    i = 0
    for seccion in secciones:
        i = 0
        for actual in seccion[0]:
            if type(actual) == str or actual == 'nan':
                if actual ==  'APELLIDOS Y NOMBRES':
                    print("--------------------------------- ", seccion[1], " --------------------------------" )
                    continue
                if seccion[2] == 1:
                    best = process.extractOne(actual,values_primero, scorer=fuzz.token_sort_ratio)
                    mediox = medio(best[0], file, 'PRIMERO')
                elif seccion[2] == 4:
                    best = process.extractOne(actual,values_cuarto, scorer=fuzz.token_sort_ratio)
                    mediox = medio(best[0], file, 'CUARTO')
                else:
                    best = process.extractOne(actual,values_quinto, scorer=fuzz.token_sort_ratio)
                    mediox = medio(best[0], file, 'QUINTO')
            
                i += 1
            else:
                continue
            if best[1] > 80:
                columna = posicion(semana,dia)
                pos = columna + str(i + 4)
                temp_registrer = registro[seccion[1]]
                temp_registrer[pos] = mediox
                #print('{0:2d} {1:44s} {2:8s} {3:5s}'.format(i, actual,'ASISTIÓ', mediox))
            else:
                columna = posicion(semana,dia)
                pos = columna + str(i + 4)
                temp_registrer = registro[seccion[1]]
                temp_registrer[pos] = 'F'
                #print('{0:2d} {1:44s} {2:8s} {3:5s}'.format(i, actual, 'FALTÓ', 'FALTÓ'))
    registro.save('registro_oficial.xlsx')
