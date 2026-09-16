# P01 – Taller de optimización y corrección de tablas

Herramienta desarrollada en **Python** para automatizar la detección, limpieza, normalización y corrección de errores en bases de datos tabulares. El proyecto procesa archivos **Excel y CSV**, aplica reglas de calidad de datos, identifica casos que requieren revisión humana y genera archivos corregidos junto con un reporte de auditoría.

##  Demostración en YouTube

[▶ Ver demostración del proyecto P01](https://youtu.be/5P6tQB-oaxk?si=oCSIaQ_-IExEZkXD)

## Objetivo

Reducir el trabajo manual necesario para revisar tablas de datos y establecer un flujo reproducible de limpieza, validación y trazabilidad de cambios.

## Funcionalidades principales

- Lectura de archivos `.xlsx`, `.xls` y `.csv`.
- Normalización de nombres de columnas.
- Eliminación de espacios innecesarios y corrección de capitalización.
- Normalización de correos electrónicos y categorías.
- Conversión y validación de valores numéricos.
- Normalización y validación de fechas.
- Validación de correos y teléfonos.
- Detección de datos importantes faltantes.
- Detección y eliminación de registros exactamente duplicados.
- Identificación de documentos repetidos con información diferente.
- Marcación de casos que requieren revisión humana.
- Generación de un archivo corregido y un reporte detallado en Excel.
- Registro de auditoría en SQLite para conservar trazabilidad de ejecuciones, incidencias y cambios.

## Tecnologías utilizadas

- Python
- Pandas
- OpenPyXL
- SQLite
- Excel / CSV

## Estructura del proyecto

```text
P01_Taller_de_Datos_Etapa1/
├── README.md
├── requirements.txt
├── .gitignore
├── VIDEO_YOUTUBE.url
├── src/
│   └── p01_limpieza.py
├── data/
│   ├── clientes.xlsx
│   └── clientes.csv
├── database/
│   ├── p01_taller_datos.db
│   └── schema_v2.sql
├── clientes corregidos/
│   ├── clientes_corregidos.xlsx
│   └── clientes_reporte_limpieza.xlsx
└── documentacion/
    └── P01_Etapa1_Guia_Aplicacion_y_Video.pdf
```

## Cómo ejecutar el proyecto

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Ejecutar con el archivo predeterminado

```bash
python src/p01_limpieza.py
```

El programa utiliza por defecto:

```text
data/clientes.xlsx
```

### 3. Ejecutar con otro archivo

```bash
python src/p01_limpieza.py --input "ruta/al/archivo.csv"
```

También se pueden utilizar archivos `.xls` y `.xlsx`.

## Resultados generados

Después de ejecutar el programa se generan dos archivos dentro de `clientes corregidos/`:

- `clientes_corregidos.xlsx`: versión procesada de la tabla.
- `clientes_reporte_limpieza.xlsx`: resumen de incidencias, cambios aplicados, elementos pendientes de revisión y perfil de columnas.

Además, la ejecución se registra en la base de datos SQLite ubicada en `database/p01_taller_datos.db`.

## Enfoque del proyecto

El sistema separa las incidencias en dos grupos:

- **Correcciones automáticas seguras:** cambios que pueden aplicarse sin alterar de forma ambigua el significado de los datos.
- **Revisión humana:** casos que no deben corregirse automáticamente porque requieren una decisión del usuario.

Este enfoque permite automatizar tareas repetitivas sin perder control sobre datos dudosos o inconsistentes.

## Autor

**Diego Guerra**

Proyecto incluido en un portafolio orientado a automatización, procesamiento de datos y soluciones empresariales.
