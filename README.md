# ExcelAuto — Verificación de asistencia (Zoom vs. Lista oficial) 📊

Herramienta en Python para contrastar una lista oficial de estudiantes con el reporte de asistencia de Zoom. Normaliza nombres, busca coincidencias fuertes y reporta:

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
  - Debe contener una columna llamada `NÓMINA`.
  - Se leerá la hoja que indiques por su nombre.
  - El encabezado está en la fila 9 (índice `header=8` en pandas). Si tu archivo no tiene 8 filas de encabezados, ajusta tu archivo o modifica el script.

- 🎥 Asistencia Zoom (`.xlsx`):
  - Debe contener la columna `Nombre (nombre original)`.

## ▶️ Uso

Ejecuta el script y responde a las indicaciones:

```bash
python3 excelauto.py
```

Se te pedirá:

1) 📝 Nombre del archivo de la lista oficial (por ejemplo: `LISTADO ESTUDIANTES 2025.xlsx`)
2) 📄 Nombre de la hoja dentro del archivo oficial (por ejemplo: `2.E`)
3) 🎯 Nombre del archivo con la asistencia de Zoom (por ejemplo: `asistencia_zoom.xlsx`)

Salida en consola:

- ✅ "Asistieron": lista de nombres de la columna `NÓMINA` encontrados en Zoom.
- ❌ "No asistieron": nombres de `NÓMINA` que no se encontraron en Zoom.
- 🧾 "Asistentes no registrados": nombres normalizados de Zoom que no coinciden con la lista oficial.

## 🧠 Cómo funciona (resumen)

- ✨ Normalización de nombres: minúsculas, sin acentos, sólo letras y espacios, y colapso de espacios.
- 🧩 Generación de combinaciones: todas las combinaciones de dos palabras y cada palabra individual.
- 🔍 Coincidencia fuerte: se considera coincidencia si al menos una combinación de dos palabras aparece en el nombre normalizado de Zoom.

## 💡 Consejos y limitaciones

- Asegúrate de que los encabezados y nombres de columnas sean exactamente los esperados (`NÓMINA`, `Nombre (nombre original)`).
- Diferencias ortográficas, apodos o nombres muy abreviados pueden reducir coincidencias.
- Si tu archivo oficial no tiene 8 filas antes del encabezado real, edítalo o ajusta `header=8` en `excelauto.py`.

## 🧯 Solución de problemas

- 📦 Error al leer `.xlsx`: instala `openpyxl` (ya incluido en `requirements.txt`).
- 🔡 Nombres de columnas no encontrados: verifica mayúsculas/minúsculas y tildes.
- 📁 Archivo u hoja no existe: confirma rutas y nombres exactos.

## 🧪 Desarrollo

- Requisitos: ver `requirements.txt`.
- Ejecuta `ruff`/`flake8` (opcional) si deseas validar estilo (no requerido).

## 🤝 Contribución

Consulta `CONTRIBUTING.md` para pautas y flujo de trabajo.

## 📜 Código de conducta

Consulta `CODE_OF_CONDUCT.md`.