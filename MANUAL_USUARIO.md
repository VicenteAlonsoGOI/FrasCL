# Manual de Usuario: Automatización de Facturas (FrasCL)

Este manual explica cómo utilizar la herramienta para generar informes PDF de facturas de procuradores a partir de un archivo Excel.

## 📋 Requisitos Previos

1.  **Archivo de Datos:** Debes tener el archivo Excel con las facturas en la misma carpeta que el programa. El archivo debe llamarse exactamente: `20260209 BORRADOR FRASCL.xlsx`.
2.  **Preparación del Excel:** Asegúrate de que el Excel no esté abierto en ese momento para evitar errores de acceso.

## 🚀 Cómo ejecutar el programa

Sigue estos pasos sencillos:

1.  **Abrir la carpeta compartida:** Navega en tu explorador de archivos hasta:
    `\\srv-gesico\Proyectos IA\Herramienta_FrasCL`
2.  **Ejecutar el proceso:**
    - Busca el archivo llamado **`INICIAR_AUTOMATIZACION.bat`** (puede aparecer solo como `INICIAR_AUTOMATIZACION` con un icono de dos engranajes).
    - Haz **doble clic** sobre él. Se abrirá una ventana negra que te informará del progreso.
3.  **Resultado:**
    - El programa creará automáticamente una carpeta llamada `PROYECTO FrasCL` dentro de esa misma carpeta de red.
    - Dentro de esa carpeta encontrarás un archivo PDF por cada cliente.

## ⚠️ Notas Importantes

- **Formato de los datos:** Si el Excel tiene errores de escritura (como las fechas con años incrementados que vimos), el PDF mostrará exactamente lo que ponga el Excel.
- **Nombres de Clientes:** Si un cliente tiene caracteres especiales prohibidos en archivos (como `/`), el programa los sustituirá por un guion bajo `_` para poder guardar el PDF.
- **Columna "Total Factura":** Si el importe es muy grande, ahora el sistema lo ajusta automáticamente para que no se corte.

---
*Para cualquier duda técnica, consulta la Documentación Técnica adjunta.*
