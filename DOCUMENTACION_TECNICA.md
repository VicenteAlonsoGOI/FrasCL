# Documentación Técnica: Sistema de Automatización FrasCL

Esta documentación técnica está dirigida a personal de IT o desarrolladores que necesiten realizar mantenimiento o mejoras en el sistema de generación de facturas.

## 1. Arquitectura del Sistema

El sistema sigue un flujo de procesamiento por lotes (batch processing) lineal:
1.  **Entrada:** Lectura de archivo Excel (.xlsx) mediante `pandas`.
2.  **Transformación:** Limpieza de datos (normalización de moneda, fechas y tipos numéricos).
3.  **Agrupación:** Segmentación de datos por el campo `CLIENTE`.
4.  **Generación:** Creación de documentos PDF mediante el motor `ReportLab`.
5.  **Persistencia:** Almacenamiento organizado en carpetas locales y/o de red.

## 2. Detalles de Implementación (generar_facturas.py)

### 2.1 Procesamiento de Moneda (clean_currency)
Debido a que el Excel fuente contiene importes con el símbolo "€", puntos para miles y comas para decimales, se implementó una lógica de "normalización de strings":
- Se eliminan espacios en blanco y el símbolo de la moneda.
- Se elimina el punto `.` (separador de miles).
- Se sustituye la coma `,` por punto `.` (estándar decimal de Python).
- Se gestionan valores nulos (`NaN`) devolviendo `0.0` para evitar errores en sumatorios.

### 2.2 Gestión de Fechas
El script gestiona la inconsistencia de formatos en el Excel (ej. `19/01/2026` frente a `29.01.2026`):
- Se estandariza el separador a `/`.
- Se utiliza `pd.to_datetime` con `dayfirst=True`.
- Se genera un string formateado `DD/MM/YYYY` para el reporte final.

### 2.3 Generación de PDF y Layout
- **Orientation:** Landscape (A4) para maximizar el espacio horizontal de la tabla.
- **Table Flow:** Se usa `Platypus` (SimpleDocTemplate) para manejar tablas que pueden extenderse a varias páginas.
- **Word Wrap:** Cada celda está envuelta en un objeto `Paragraph`. Esto permite que el motor de ReportLab calcule el salto de línea automático si el texto excede el ancho de la columna.
- **Column Widths:** Definidos estáticamente para asegurar la integridad visual (Total ~25.7cm).

## 3. Lanzador Automatizado (INICIAR_AUTOMATIZACION.bat)

El lanzador está diseñado para ser totalmente autónomo:
- **Soporte UNC:** Utiliza `pushd "%~dp0"` para permitir la ejecución desde carpetas de red sin que CMD falle.
- **Auto-instalación:** Ejecuta `pip install` de forma silenciosa. Si el usuario ya tiene las librerías, el comando no hace nada; si le faltan, las instala antes de lanzar el script.

## 4. Requisitos de Entorno

- **Python 3.12+**
- **Librerías:**
    - `pandas`: Análisis de datos.
    - `reportlab`: Generación de PDF (profesional).
    - `openpyxl`: Motor de lectura Excel.
    - `python-docx`: Generación de manuales en Word.

## 5. Mantenimiento y Cambios
Para cambiar el archivo de entrada, simplemente actualice la variable `excel_file` en la función `generar_pdfs()`. El sistema detectará automáticamente las nuevas columnas siempre que mantengan palabras clave similares (`Cliente`, `Cuantía`, `Exp`, etc.).

---
*Documentación generada por Antigravity AI - v2.0*

---
*Desarrollado por Antigravity AI.*
