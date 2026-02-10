# Documentación Técnica: Proyecto FrasCL

Esta documentación detalla la arquitectura, dependencias y lógica interna del sistema de generación de facturas.

## 🛠️ Tecnologías y Dependencias

- **Lenguaje:** Python 3.12+
- **Librerías principales:**
  - `pandas`: Manipulación y limpieza de datos desde Excel.
  - `openpyxl`: Motor de lectura para archivos `.xlsx`.
  - `reportlab`: Motor de generación de documentos PDF (Platypus).

## 📂 Estructura del Proyecto

- `generar_facturas.py`: Script principal de ejecución.
- `20260209 BORRADOR FRASCL.xlsx`: Fuente de datos de entrada.
- `PROYECTO FrasCL/`: Directorio de salida (generado automáticamente).
- `.gitignore`: Configuración para evitar la subida de datos sensibles o binarios a Git.

## ⚙️ Lógica de Procesamiento

### 1. Extracción de Datos
El script utiliza `pandas` para cargar el Excel. Se realiza una limpieza de los nombres de las columnas para eliminar espacios accidentales.

### 2. Función `clean_currency`
Esta función es crítica para la robustez del sistema. Realiza las siguientes transformaciones:
- Elimina el símbolo `€`.
- Elimina los puntos `.` (separadores de miles en formato español).
- Cambia las comas `,` por puntos `.` (separador decimal para Python).
- Convierte el resultado a `float` para permitir cálculos matemáticos.

### 3. Gestión de Fechas
Se implementó un parsing flexible que detecta fechas tanto con separador de barra `/` como de punto `.`. Se utiliza `dayfirst=True` para respetar el formato europeo.

### 4. Generación de PDF (ReportLab)
- **Word Wrap:** Se utiliza la clase `Paragraph` dentro de las celdas de la `Table` para permitir que los textos largos salten de línea.
- **Anchos Fijos:** Se definen anchos de columna estrictos (`col_widths`) para asegurar que la tabla no exceda los márgenes físicos de una hoja A4 en orientación horizontal (landscape).
- **Estilos:** Se utiliza `TableStyle` para aplicar fondos alternos, encabezados coloreados y el resaltado del sumatorio final.

## 🔄 Mantenimiento

Para actualizar el programa con un nuevo Excel:
1. Reemplazar el archivo `.xlsx` antiguo por el nuevo.
2. Asegurarse de que el nombre del archivo en el código coincide (línea 10 de `generar_facturas.py`).
3. Ejecutar el script.

---
*Desarrollado por Antigravity AI.*
