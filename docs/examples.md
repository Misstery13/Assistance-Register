# Ejemplos de uso

## Estructura de archivos

- Oficial (`LISTADO ESTUDIANTES 2025.xlsx`):
  - Hoja: `2.E`
  - Columna: `NÓMINA`
  - Encabezado real en fila 9 (header=8)

- Zoom (`asistencia_zoom.xlsx`):
  - Columna: `Nombre (nombre original)`

## Ejecución

```bash
python3 excelauto.py
```

Entradas interactivas:

- `LISTADO ESTUDIANTES 2025.xlsx`
- `2.E`
- `asistencia_zoom.xlsx`

Salida (ejemplo):

```
Asistieron:
Juan Pérez
María López

No asistieron:
Ana García

Asistentes no registrados en la lista oficial:
juanito perez
```