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
    ascii_art = """
        ooooo  oooo ooooooooooo oooooooooo  ooooo ooooooooooo ooooo  oooooooo8     o       oooooooo8 ooooo  ooooooo  oooo   oooo      
         888    88   888    88   888    888  888   888    88   888 o888     88    888    o888     88  888 o888   888o 8888o  88       
          888  88    888ooo8     888oooo88   888   888ooo8     888 888           8  88   888          888 888     888 88 888o88       
           88888     888    oo   888  88o    888   888         888 888o     oo  8oooo88  888o     oo  888 888o   o888 88   8888       
            888     o888ooo8888 o888o  88o8 o888o o888o       o888o 888oooo88 o88o  o888o 888oooo88  o888o  88ooo88  o88o    88       

             ooooooooo  ooooooooooo                                                                                                   
              888    88o 888    88                                                                                                    
              888    888 888ooo8                                                                                                      
              888    888 888    oo                                                                                                    
             o888ooo88  o888ooo8888                                                                                                   

             o       oooooooo8 ooooo  oooooooo8 ooooooooooo ooooooooooo oooo   oooo  oooooooo8 ooooo      o                           
            888     888         888  888        88  888  88  888    88   8888o  88 o888     88  888      888                          
           8  88     888oooooo  888   888oooooo     888      888ooo8     88 888o88 888          888     8  88                         
          8oooo88           888 888          888    888      888    oo   88   8888 888o     oo  888    8oooo88                        
        o88o  o888o o88oooo888 o888o o88oooo888    o888o    o888ooo8888 o88o    88  888oooo88  o888o o88o  o888o                      
    """
    print(ascii_art)
    print("- Por favor, proporciona los archivos necesarios para la verificación.")
    print("- Asegúrate de que los archivos estén en LA MISMA CARPETA que este script.")
    print("- Los nombres de los archivos deben ser EXACTOS, incluyendo mayúsculas y minúsculas.")
    print("-------------------------------------------------------------------------------------------------------------------")
    archivo_oficial = input("Nombre del archivo de la lista oficial (ej: LISTADO ESTUDIANTES 2025.xlsx): ")
    hoja_oficial = input("Nombre de la hoja de la lista oficial (ej: 2.E): ")
    archivo_zoom = input("Nombre del archivo de asistencia Zoom (ej: asistencia_zoom.xlsx): ")

    oficial = pd.read_excel(archivo_oficial, sheet_name=hoja_oficial, header=8)
    zoom = pd.read_excel(archivo_zoom)

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