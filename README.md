# ExcelAuto — Verificación de asistencia (Zoom vs. Lista oficial) 📊

Herramienta en Python para contrastar una lista oficial de estudiantes con el reporte de asistencia de Zoom. Normaliza nombres, busca coincidencias fuertes o difusas y reporta:

- ✅ Asistieron (encontrados en Zoom)
- ❌ No asistieron (no encontrados en Zoom)
- 🧾 Asistentes no registrados (aparecen en Zoom pero no en la lista oficial)

## 🧰 Requisitos

- Python 3.9 o superior
- `pip` para instalar dependencias

## 🛠️ Instalación

Opcional: crear y activar un entorno virtual.

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instala dependencias:

```bash
pip install -r requirements.txt
```

## 🗂️ Archivos esperados

Coloca los archivos en la misma carpeta que `excelauto.py`.

- 📘 Lista oficial (`.xlsx`):
  - Debe contener una columna llamada `NÓMINA` (configurable via `--columna-oficial`).
  - Se leerá la hoja que indiques por su nombre (`--hoja-oficial`).
  - El encabezado está en la fila 9 por defecto (`--header-oficial 8`).

- 🎥 Asistencia Zoom (`.xlsx`):
  - Debe contener la columna `Nombre (nombre original)` (configurable via `--columna-zoom`).

## ▶️ Uso (CLI)

Ayuda:

```bash
python3 excelauto.py --help
```

Ejemplo (método fuerte, sólo consola):

```bash
python3 excelauto.py \
  --oficial "LISTADO ESTUDIANTES 2025.xlsx" \
  --hoja-oficial "2.E" \
  --zoom "asistencia_zoom.xlsx" \
  --metodo fuerte
```

Ejemplo (fuzzy con umbral 88 y exportar a Excel y CSV):

```bash
python3 excelauto.py \
  --oficial "LISTADO ESTUDIANTES 2025.xlsx" \
  --hoja-oficial "2.E" \
  --zoom "asistencia_zoom.xlsx" \
  --metodo fuzzy \
  --umbral-fuzzy 88 \
  --excel resultados/resultados.xlsx \
  --csv-dir resultados/csv
```

Parámetros clave:

- `--metodo`: `fuerte` (regla de combinaciones) o `fuzzy` (similaridad con RapidFuzz).
- `--umbral-fuzzy`: 0–100, por defecto 85.
- `--excel`: ruta de salida para un Excel con 4 hojas (asistieron, no_asistieron, no_registrados, resumen).
- `--csv-dir`: directorio para CSVs.
- `--no-banner`: oculta el banner ASCII.

Si omites parámetros, el script te preguntará interactivamente.

## 🧠 Cómo funciona (resumen)

- ✨ Normalización de nombres: minúsculas, sin acentos, sólo letras y espacios, y colapso de espacios.
- 🧩 Generación de combinaciones: todas las combinaciones de dos palabras y cada palabra individual.
- 🔍 Métodos de coincidencia:
  - `fuerte`: match si al menos una combinación de dos palabras aparece en el nombre de Zoom.
  - `fuzzy`: similitud de tokens con `token_set_ratio` (RapidFuzz) contra un umbral.

## 💡 Consejos y limitaciones

- Verifica que los encabezados y nombres de columnas sean correctos o ajusta los flags.
- Diferencias ortográficas o apodos pueden requerir `--metodo fuzzy` y un umbral más bajo.
- Para encabezados que no están en fila 9, usa `--header-oficial`.

## 🧯 Solución de problemas

- 📦 Error al leer `.xlsx`: verifica `openpyxl` (ya incluido en `requirements.txt`).
- 🔡 Columna no encontrada: revisa `--columna-oficial` y `--columna-zoom`.
- 🧪 Prueba con un umbral diferente: `--umbral-fuzzy 80`.

## 🧪 Desarrollo

- Requisitos: ver `requirements.txt`.
- Ejecuta `ruff`/`flake8` (opcional).

## 🤝 Contribución

Consulta `CONTRIBUTING.md`.

## 📜 Código de conducta

Consulta `CODE_OF_CONDUCT.md`.