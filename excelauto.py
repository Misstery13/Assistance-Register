import argparse
import sys
from pathlib import Path
import pandas as pd
import unicodedata
import re
from typing import List, Tuple, Optional, Dict, Set

try:
    from rapidfuzz import fuzz
except Exception:  # pragma: no cover
    fuzz = None  # Will validate at runtime


def normalizar_nombre(nombre: object) -> str:
    """Convierte a minúsculas, elimina tildes y caracteres no alfabéticos, y colapsa espacios."""
    if pd.isnull(nombre):
        return ''
    texto = str(nombre).lower().strip()
    texto = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    texto = re.sub(r'[^a-z\s]', '', texto)
    texto = ' '.join(texto.split())
    return texto


def obtener_combinaciones(nombre_normalizado: str) -> List[str]:
    partes = nombre_normalizado.split()
    combinaciones: Set[str] = set()
    for i in range(len(partes)):
        for j in range(len(partes)):
            if i != j:
                combinaciones.add(f"{partes[i]} {partes[j]}")
    for parte in partes:
        combinaciones.add(parte)
    return list(combinaciones)


def coincidencia_fuerte(combinaciones: List[str], nombre_zoom_normalizado: str) -> bool:
    return sum(comb in nombre_zoom_normalizado for comb in combinaciones if ' ' in comb) >= 1


def coincidencia_fuzzy(nombre_oficial_norm: str, nombre_zoom_norm: str, umbral: int) -> Tuple[bool, int]:
    if not fuzz:
        raise RuntimeError("rapidfuzz no está instalado. Instala 'rapidfuzz' o usa --metodo fuerte.")
    # token_set_ratio maneja orden y duplicados de tokens de forma robusta
    puntaje = fuzz.token_set_ratio(nombre_oficial_norm, nombre_zoom_norm)
    return puntaje >= umbral, puntaje


def emparejar_listas(
    oficial_df: pd.DataFrame,
    zoom_df: pd.DataFrame,
    columna_oficial: str,
    columna_zoom: str,
    metodo: str,
    umbral_fuzzy: int,
) -> Tuple[List[str], List[str], List[Dict[str, object]]]:
    """Devuelve (asistieron, no_asistieron, no_registrados).

    - asistieron: lista de nombres originales de la columna oficial que se encontraron
    - no_asistieron: lista de nombres originales que no se encontraron
    - no_registrados: lista de dicts con {'zoom_original', 'zoom_normalizado', 'mejor_puntaje'}
    """
    oficial_df = oficial_df.copy()
    zoom_df = zoom_df.copy()

    oficial_df['__norm__'] = oficial_df[columna_oficial].apply(normalizar_nombre)
    zoom_df['__norm__'] = zoom_df[columna_zoom].apply(normalizar_nombre)

    asistieron: List[str] = []
    no_asistieron: List[str] = []
    no_registrados: List[Dict[str, object]] = []

    usados_zoom: Set[int] = set()

    for idx_ofi, fila_ofi in oficial_df.iterrows():
        nombre_oficial = fila_ofi.get(columna_oficial)
        if pd.isnull(nombre_oficial):
            continue
        nombre_oficial_norm = fila_ofi['__norm__']
        if not nombre_oficial_norm:
            continue

        encontrado = False
        mejor_puntaje = -1
        mejor_zoom_idx: Optional[int] = None

        # Recorre todos los nombres de Zoom para encontrar un match
        for idx_zoom, fila_zoom in zoom_df.iterrows():
            if idx_zoom in usados_zoom:
                continue
            nombre_zoom_norm = fila_zoom['__norm__']

            if metodo == 'fuerte':
                combinaciones = obtener_combinaciones(nombre_oficial_norm)
                if coincidencia_fuerte(combinaciones, nombre_zoom_norm):
                    encontrado = True
                    mejor_zoom_idx = idx_zoom
                    break
            elif metodo == 'fuzzy':
                ok, puntaje = coincidencia_fuzzy(nombre_oficial_norm, nombre_zoom_norm, umbral_fuzzy)
                if puntaje > mejor_puntaje:
                    mejor_puntaje = puntaje
                    mejor_zoom_idx = idx_zoom
                if ok:
                    encontrado = True
                    break
            else:
                raise ValueError("Método no soportado. Usa 'fuerte' o 'fuzzy'.")

        if encontrado and mejor_zoom_idx is not None:
            asistieron.append(str(nombre_oficial))
            usados_zoom.add(mejor_zoom_idx)
        else:
            no_asistieron.append(str(nombre_oficial))

    # No registrados: entradas de Zoom que no se emparejaron con nadie
    for idx_zoom, fila_zoom in zoom_df.iterrows():
        if idx_zoom in usados_zoom:
            continue
        no_registrados.append({
            'zoom_original': fila_zoom.get(columna_zoom),
            'zoom_normalizado': fila_zoom['__norm__'],
        })

    return asistieron, no_asistieron, no_registrados


def exportar_resultados(
    asistieron: List[str],
    no_asistieron: List[str],
    no_registrados: List[Dict[str, object]],
    ruta_excel: Optional[Path] = None,
    ruta_csv_dir: Optional[Path] = None,
) -> None:
    resumen_df = pd.DataFrame({
        'metrica': ['asistieron', 'no_asistieron', 'no_registrados', 'tasa_asistencia'],
        'valor': [
            len(asistieron),
            len(no_asistieron),
            len(no_registrados),
            round((len(asistieron) / (len(asistieron) + len(no_asistieron))) * 100, 2) if (len(asistieron) + len(no_asistieron)) > 0 else 0.0,
        ]
    })

    if ruta_excel:
        ruta_excel.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            pd.DataFrame({'NOMINA': asistieron}).to_excel(writer, index=False, sheet_name='asistieron')
            pd.DataFrame({'NOMINA': no_asistieron}).to_excel(writer, index=False, sheet_name='no_asistieron')
            pd.DataFrame(no_registrados).to_excel(writer, index=False, sheet_name='no_registrados')
            resumen_df.to_excel(writer, index=False, sheet_name='resumen')

    if ruta_csv_dir:
        ruta_csv_dir.mkdir(parents=True, exist_ok=True)
        pd.DataFrame({'NOMINA': asistieron}).to_csv(ruta_csv_dir / 'asistieron.csv', index=False)
        pd.DataFrame({'NOMINA': no_asistieron}).to_csv(ruta_csv_dir / 'no_asistieron.csv', index=False)
        pd.DataFrame(no_registrados).to_csv(ruta_csv_dir / 'no_registrados.csv', index=False)
        resumen_df.to_csv(ruta_csv_dir / 'resumen.csv', index=False)


def imprimir_resultados(asistieron: List[str], no_asistieron: List[str], no_registrados: List[Dict[str, object]]) -> None:
    print("\n✅ Asistieron:")
    for n in asistieron:
        print(n)
    print("\n❌ No asistieron:")
    for n in no_asistieron:
        print(n)
    print("\n🧾 Asistentes no registrados en la lista oficial:")
    for fila in no_registrados:
        print(fila.get('zoom_original'))


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='Verifica asistencia cruzando lista oficial y reporte de Zoom.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument('--oficial', type=Path, help='Ruta del archivo Excel de la lista oficial (.xlsx)')
    parser.add_argument('--hoja-oficial', type=str, help='Nombre de la hoja en el archivo oficial')
    parser.add_argument('--zoom', type=Path, help='Ruta del archivo Excel de asistencia de Zoom (.xlsx)')
    parser.add_argument('--columna-oficial', type=str, default='NÓMINA', help='Nombre de la columna en el archivo oficial')
    parser.add_argument('--columna-zoom', type=str, default='Nombre (nombre original)', help='Nombre de la columna en el archivo de Zoom')
    parser.add_argument('--header-oficial', type=int, default=8, help='Índice de fila del encabezado (0-index) en el archivo oficial')
    parser.add_argument('--metodo', type=str, default='fuerte', choices=['fuerte', 'fuzzy'], help='Método de emparejamiento')
    parser.add_argument('--umbral-fuzzy', type=int, default=85, help='Umbral de similitud 0-100 para fuzzy')
    parser.add_argument('--excel', type=Path, help='Ruta de salida para Excel (.xlsx)')
    parser.add_argument('--csv-dir', type=Path, help='Directorio de salida para CSVs')
    parser.add_argument('--no-banner', action='store_true', help='Oculta el banner ASCII')
    return parser.parse_args(argv)


def mostrar_banner() -> None:
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


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)

    if not args.no_banner:
        try:
            mostrar_banner()
        except Exception:
            pass

    # Entrada interactiva si faltan parámetros obligatorios
    if not args.oficial:
        args.oficial = Path(input("Nombre del archivo de la lista oficial (ej: LISTADO ESTUDIANTES 2025.xlsx): ").strip())
    if not args.hoja_oficial:
        args.hoja_oficial = input("Nombre de la hoja de la lista oficial (ej: 2.E): ").strip()
    if not args.zoom:
        args.zoom = Path(input("Nombre del archivo de asistencia Zoom (ej: asistencia_zoom.xlsx): ").strip())

    # Validaciones
    if not args.oficial.exists():
        print(f"Error: no existe el archivo oficial: {args.oficial}")
        return 2
    if not args.zoom.exists():
        print(f"Error: no existe el archivo de Zoom: {args.zoom}")
        return 2

    try:
        oficial_df = pd.read_excel(args.oficial, sheet_name=args.hoja_oficial, header=args.header_oficial)
    except Exception as e:
        print(f"Error al leer el archivo oficial: {e}")
        return 2

    try:
        zoom_df = pd.read_excel(args.zoom)
    except Exception as e:
        print(f"Error al leer el archivo de Zoom: {e}")
        return 2

    for columna, nombre in [(args.columna_oficial, 'archivo oficial'), (args.columna_zoom, 'archivo de Zoom')]:
        df = oficial_df if nombre == 'archivo oficial' else zoom_df
        if columna not in df.columns:
            print(f"Error: no se encontró la columna '{columna}' en el {nombre}. Columnas: {list(df.columns)}")
            return 2

    try:
        asistieron, no_asistieron, no_registrados = emparejar_listas(
            oficial_df=oficial_df,
            zoom_df=zoom_df,
            columna_oficial=args.columna_oficial,
            columna_zoom=args.columna_zoom,
            metodo=args.metodo,
            umbral_fuzzy=args.umbral_fuzzy,
        )
    except Exception as e:
        print(f"Ocurrió un error durante el emparejamiento: {e}")
        return 2

    imprimir_resultados(asistieron, no_asistieron, no_registrados)

    if args.excel or args.csv_dir:
        try:
            exportar_resultados(
                asistieron=asistieron,
                no_asistieron=no_asistieron,
                no_registrados=no_registrados,
                ruta_excel=args.excel,
                ruta_csv_dir=args.csv_dir,
            )
            if args.excel:
                print(f"\n💾 Resultados exportados a Excel: {args.excel}")
            if args.csv_dir:
                print(f"💾 Resultados exportados como CSV en: {args.csv_dir}")
        except Exception as e:
            print(f"Error al exportar resultados: {e}")
            return 2

    # Resumen final
    total = len(asistieron) + len(no_asistieron)
    tasa = (len(asistieron) / total * 100) if total else 0
    print("\nResumen:")
    print(f" - Total en lista oficial: {total}")
    print(f" - Asistieron: {len(asistieron)}")
    print(f" - No asistieron: {len(no_asistieron)}")
    print(f" - No registrados (Zoom): {len(no_registrados)}")
    print(f" - Tasa de asistencia: {tasa:.2f}%")

    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("Interrumpido por el usuario")
        sys.exit(130)