# Manual de Usuario: Automatización de Facturas (FrasCL)

Este manual explica cómo utilizar la herramienta para generar informes PDF de facturas de procuradores a partir de un archivo Excel.

## 📋 Requisitos Previos

1.  **Archivo de Datos:** Debes tener el archivo Excel con las facturas en la misma carpeta que el programa. El archivo debe llamarse exactamente: `20260209 BORRADOR FRASCL.xlsx`.
2.  **Preparación del Excel:** Asegúrate de que el Excel no esté abierto en ese momento para evitar errores de acceso.

## 🚀 Cómo ejecutar el programa

Sigue estos pasos sencillos:

1.  **Abrir la carpeta del proyecto:** Navega hasta `c:\Users\jose.alonso\Documents\FrasCL`.
2.  **Ejecutar el proceso:**
    - Haz clic derecho en el archivo `generar_facturas.py`.
    - Selecciona **"Ejecutar con Python"** (o similar).
    - O simplemente, abre una terminal en esa carpeta y escribe:
      ```powershell
      python generar_facturas.py
      ```
3.  **Resultado:**
    - El programa creará automáticamente una carpeta llamada `PROYECTO FrasCL`.
    - Dentro de esa carpeta encontrarás un archivo PDF por cada cliente.

## ⚠️ Notas Importantes

- **Formato de los datos:** Si el Excel tiene errores de escritura (como las fechas con años incrementados que vimos), el PDF mostrará exactamente lo que ponga el Excel.
- **Nombres de Clientes:** Si un cliente tiene caracteres especiales prohibidos en archivos (como `/`), el programa los sustituirá por un guion bajo `_` para poder guardar el PDF.
- **Columna "Total Factura":** Si el importe es muy grande, ahora el sistema lo ajusta automáticamente para que no se corte.

---
*Para cualquier duda técnica, consulta la Documentación Técnica adjunta.*
