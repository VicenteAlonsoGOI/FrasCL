from docx import Document
from docx.shared import Pt
import os

def create_manual():
    doc = Document()
    doc.add_heading('Manual de Usuario: Automatización de Facturas (FrasCL)', 0)
    
    doc.add_heading('Requisitos Previos', level=1)
    p = doc.add_paragraph()
    p.add_run('1. Archivo de Datos: ').bold = True
    p.add_run('Debes tener el archivo Excel con las facturas en la misma carpeta que el programa. El archivo debe llamarse exactamente: 20260209 BORRADOR FRASCL.xlsx.')
    
    p = doc.add_paragraph()
    p.add_run('2. Preparación del Excel: ').bold = True
    p.add_run('Asegúrate de que el Excel no esté abierto en ese momento para evitar errores de acceso.')
    
    doc.add_heading('Cómo ejecutar el programa', level=1)
    doc.add_paragraph('1. Abrir la carpeta compartida: Navega en tu explorador de archivos hasta: \\\\srv-gesico\\Proyectos IA Vicente\\PROYECTO FrasCL')
    doc.add_paragraph('2. Ejecutar el proceso: Busca el archivo llamado INICIAR_AUTOMATIZACION.bat. Haz doble clic sobre él. Se abrirá una ventana negra que te informará del progreso.')
    doc.add_paragraph('3. Resultado: El programa creará automáticamente una carpeta llamada Facturas_Generadas. Dentro encontrarás un archivo PDF por cada cliente.')
    
    doc.add_heading('Notas Importantes', level=1)
    doc.add_paragraph('- Formato de los datos: Si el Excel tiene errores, el PDF mostrará exactamente lo que ponga el Excel.')
    doc.add_paragraph('- Nombres de Clientes: Caracteres especiales como / se sustituirán por _ .')
    doc.add_paragraph('- Columna Total Factura: Se ajusta automáticamente para evitar cortes.')
    
    doc.save('MANUAL_USUARIO.docx')

def create_tech_doc():
    doc = Document()
    doc.add_heading('Documentación Técnica: Proyecto FrasCL', 0)
    
    doc.add_heading('Tecnologías y Dependencias', level=1)
    doc.add_paragraph('- Lenguaje: Python 3.12+')
    doc.add_paragraph('- Librerías: pandas, openpyxl, reportlab, python-docx.')
    
    doc.add_heading('Estructura del Proyecto', level=1)
    doc.add_paragraph('- generar_facturas.py: Script principal.')
    doc.add_paragraph('- INICIAR_AUTOMATIZACION.bat: Lanzador para usuarios.')
    doc.add_paragraph('- Facturas_Generadas/: Carpeta de salida.')
    
    doc.add_heading('Lógica de Procesamiento', level=1)
    doc.add_paragraph('El sistema limpia importes (moneda española), maneja fechas con puntos o barras, y utiliza ReportLab con Word Wrap para generar los PDFs en formato horizontal.')
    
    doc.save('DOCUMENTACION_TECNICA.docx')

if __name__ == "__main__":
    create_manual()
    create_tech_doc()
    print("Documentos Word generados con éxito.")
