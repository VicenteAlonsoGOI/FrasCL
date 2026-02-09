import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

def generar_pdfs():
    excel_file = '20260209 BORRADOR FRASCL.xlsx'
    if not os.path.exists(excel_file):
        print(f"Error: No se encuentra el archivo {excel_file}")
        return

    # Cargar datos
    df = pd.read_excel(excel_file)
    
    # Limpiar nombres de columnas
    df.columns = [c.strip() for c in df.columns]
    
    # Agrupar por CLIENTE
    clientes = df['CLIENTE'].unique()
    
    # Carpeta solicitada
    output_folder = 'PROYECTO FrasCL'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for cliente in clientes:
        if pd.isna(cliente):
            continue
            
        df_cliente = df[df['CLIENTE'] == cliente].copy()
        output_pdf = os.path.join(output_folder, f"{str(cliente).replace('/', '_')}.pdf")
        
        create_client_pdf(cliente, df_cliente, output_pdf)
        print(f"PDF generado para: {cliente}")

def create_client_pdf(cliente, df_cliente, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    
    styles = getSampleStyleSheet()
    
    # Estilo de Título
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor("#1A5276"),
        alignment=1, # Center
        spaceAfter=20
    )
    
    # Estilo de Subtítulo
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.black,
        alignment=0, # Left
        spaceAfter=15
    )
    
    # Estilo para el cuerpo de la tabla (permite wrap)
    body_style = ParagraphStyle(
        'BodyStyle',
        parent=styles['Normal'],
        fontSize=7,
        leading=8,
        alignment=1, # Center
    )

    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontSize=8,
        leading=9,
        alignment=1, # Center
        textColor=colors.whitesmoke,
        fontName='Helvetica-Bold'
    )
    
    elements.append(Paragraph("LISTADO FACTURAS DE PROCURADORES MES EN CURSO", title_style))
    elements.append(Paragraph(f"CLIENTE: {cliente}", subtitle_style))
    elements.append(Spacer(1, 0.5*cm))
    
    # Mapeo de columnas
    col_mapping = {}
    for actual_col in df_cliente.columns:
        if 'Archivado' in actual_col: col_mapping[actual_col] = 'Fecha'
        elif 'Contrario' in actual_col and 'NIF' not in actual_col: col_mapping[actual_col] = 'Contrario'
        elif 'NIF' in actual_col: col_mapping[actual_col] = 'NIF'
        elif 'Exp' in actual_col: col_mapping[actual_col] = 'Exp.'
        elif 'Cuant' in actual_col: col_mapping[actual_col] = 'Cuantía'
        elif 'Tipo' in actual_col: col_mapping[actual_col] = 'Procedimiento'
        elif 'Autos' in actual_col: col_mapping[actual_col] = 'Autos'
        elif 'PROCURADOR' in actual_col: col_mapping[actual_col] = 'Procurador'
        elif 'Numero' in actual_col or 'factura' in actual_col: col_mapping[actual_col] = 'Fra.'
        elif 'Suma' in actual_col or 'Total' in actual_col: col_mapping[actual_col] = 'Total'

    df_display = df_cliente.rename(columns=col_mapping)
    desired_order = ['Fecha', 'Contrario', 'NIF', 'Exp.', 'Cuantía', 'Procedimiento', 'Autos', 'Procurador', 'Fra.', 'Total']
    df_display = df_display[[col for col in desired_order if col in df_display.columns]]
    
    # Formatear números
    for col in ['Cuantía', 'Total']:
        if col in df_display.columns:
            df_display[col] = pd.to_numeric(df_display[col], errors='coerce')
            df_display[col] = df_display[col].apply(lambda x: f"{float(x):,.2f} \u20ac" if pd.notnull(x) else "0.00 \u20ac")

    # Formatear fechas
    if 'Fecha' in df_display.columns:
        df_display['Fecha'] = pd.to_datetime(df_display['Fecha'], errors='coerce')
        df_display['Fecha'] = df_display['Fecha'].dt.strftime('%d/%m/%Y').fillna("")

    # Convertir encabezados y celdas a Paragraph para permitir wrap
    header_row = [Paragraph(col, header_style) for col in df_display.columns]
    
    formatted_data = [header_row]
    for row in df_display.values.tolist():
        formatted_row = [Paragraph(str(cell), body_style) for cell in row]
        formatted_data.append(formatted_row)
    
    # Sumatorio
    total_val = df_cliente.iloc[:, -1].sum()
    empty_row = [""] * (len(df_display.columns) - 1) + [f"{total_val:,.2f} \u20ac"]
    empty_row[0] = "TOTAL"
    # El sumatorio también debe ser Paragraph para consistencia
    total_row = [Paragraph(str(cell), body_style) for cell in empty_row]
    formatted_data.append(total_row)
    
    # Definir anchos de columna (Total ~25.7cm)
    col_widths = [
        2.0*cm, # Fecha
        4.5*cm, # Contrario
        2.0*cm, # NIF
        1.8*cm, # Exp.
        1.8*cm, # Cuantía
        4.5*cm, # Procedimiento
        1.8*cm, # Autos
        4.5*cm, # Procurador
        1.5*cm, # Fra.
        1.3*cm  # Total
    ]
    
    # Tabla
    table = Table(formatted_data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2E86C1")),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#EBF5FB")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#AED6F1")),
    ])
    table.setStyle(style)
    
    elements.append(table)
    doc.build(elements)

if __name__ == "__main__":
    generar_pdfs()
