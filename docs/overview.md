# 🧠 Visión general

Este script realiza un cruce entre una lista oficial y un reporte de Zoom.

## 🧩 Componentes

- `normalizar_nombre(nombre)`: Estandariza texto (minúsculas, sin tildes, sólo letras y espacios).
- `obtener_combinaciones(nombre)`: Genera combinaciones de dos palabras y palabras individuales.
- `coincidencia_fuerte(combinaciones, nombre_zoom)`: Verdadero si al menos una combinación de dos palabras aparece en el nombre de Zoom.
- 🧱 Bloque principal: lectura de Excel, normalización, cruce y reporte en consola.

## 🔄 Flujo

1. 📘 Leer `archivo_oficial` (hoja `hoja_oficial`, `header=8`).
2. 🎥 Leer `archivo_zoom`.
3. ✨ Crear columnas `nombre_normalizado` a partir de `NÓMINA` y `Nombre (nombre original)`.
4. 🗂️ Construir listas: `asistieron`, `no_asistieron`, `no_registrados`.
5. 🖨️ Imprimir resultados.

## 🚀 Posibles mejoras

- ⚙️ Parametrizar `header` y nombres de columnas via argumentos CLI.
- 📤 Exportar resultados a CSV/Excel además de imprimir.
- 🧠 Medir similitud difusa (fuzzy matching) para nombres parcialmente diferentes.