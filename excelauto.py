import pandas as pd
import unicodedata
import re

def normalizar_nombre(nombre):
    if pd.isnull(nombre):
        return ''
    nombre = str(nombre).lower().strip()
    nombre = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )
    nombre = re.sub(r'[^a-z\s]', '', nombre)
    nombre = ' '.join(nombre.split())
    return nombre

def obtener_combinaciones(nombre):
    partes = nombre.split()
    combinaciones = set()
    # Genera todas las combinaciones de dos palabras
    for i in range(len(partes)):
        for j in range(len(partes)):
            if i != j:
                combinaciones.add(f"{partes[i]} {partes[j]}")
    # Agrega cada palabra individual
    for parte in partes:
        combinaciones.add(parte)
    return list(combinaciones)

def coincidencia_fuerte(combinaciones, nombre_zoom):
    # Cuenta cuántas combinaciones de dos palabras están en el nombre de Zoom
    return sum(comb in nombre_zoom for comb in combinaciones if ' ' in comb) >= 1

try:
    oficial = pd.read_excel('LISTADO ESTUDIANTES 2025.xlsx', sheet_name='2.E', header=8)
    zoom = pd.read_excel('asistencia_zoom.xlsx')

    oficial['nombre_normalizado'] = oficial['NÓMINA'].apply(normalizar_nombre)
    zoom['nombre_normalizado'] = zoom['Nombre (nombre original)'].apply(normalizar_nombre)

    asistieron = []
    no_asistieron = []
    no_registrados = []

    for idx, nombre_norm in enumerate(oficial['nombre_normalizado']):
        if pd.isnull(oficial['NÓMINA'].iloc[idx]):
            continue
        combinaciones = obtener_combinaciones(nombre_norm)
        encontrado = False
        for nombre_zoom in zoom['nombre_normalizado']:
            if coincidencia_fuerte(combinaciones, nombre_zoom):
                asistieron.append(oficial['NÓMINA'].iloc[idx])
                encontrado = True
                break
        if not encontrado:
            no_asistieron.append(oficial['NÓMINA'].iloc[idx])
    for nombre_zoom in zoom['nombre_normalizado']:
        encontrado = False
        for nombre_norm in oficial['nombre_normalizado']:
            combinaciones = obtener_combinaciones(nombre_norm)
            if coincidencia_fuerte(combinaciones, nombre_zoom):
                encontrado = True
                break
        if not encontrado:
            no_registrados.append(nombre_zoom)

    print("\nAsistieron:")
    for n in asistieron:
        print(n)
    print("\nNo asistieron:")
    for n in no_asistieron:
        print(n)
    print("\nAsistentes no registrados en la lista oficial:")
    for n in no_registrados:
        print(n)

except Exception as e:
    print("Ocurrió un error:", e)