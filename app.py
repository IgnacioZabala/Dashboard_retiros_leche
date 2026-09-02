import streamlit as st
import pandas as pd
from fpdf import FPDF
import os

# 1. Configuración de la página
st.set_page_config(page_title="Resumen Semanal de Recolección", layout="wide")
st.title("Panel de retiro de leche por productor")

# --- FUNCIÓN PARA GENERAR EL PDF ---
def generar_pdf(df_productor, tambo_nombre, tambo_id, fecha_inicio, fecha_fin):
    pdf = FPDF()
    pdf.add_page()
    
    # Membrete y Logo
    ruta_logo = "logo.png"
    if os.path.exists(ruta_logo):
        pdf.image(ruta_logo, 10, 8, 30) 
        espacio_izquierdo = 35 
    else:
        espacio_izquierdo = 0
    
    pdf.set_font('Arial', 'B', 16)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, 'Coopagro - Planta Tandil', ln=True, align='L')
    
    pdf.set_font('Arial', '', 12)
    pdf.set_text_color(100, 100, 100)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, 'Resumen Semanal de Recoleccion de Leche', ln=True, align='L')
    
    pdf.ln(5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    
    # Datos del Productor
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f'Productor: {tambo_nombre} (Codigo #{tambo_id})', ln=True)
    pdf.set_font('Arial', '', 11)
    pdf.cell(0, 8, f'Periodo de liquidacion: {fecha_inicio} al {fecha_fin}', ln=True)
    
    # Resumen (KPIs)
    total_litros = df_productor['Litros_Ticket'].sum()
    temp_prom = df_productor['Temperatura'].mean()
    pdf.cell(0, 8, f'Total Recolectado: {total_litros:,.0f} L', ln=True)
    pdf.cell(0, 8, f'Temperatura Promedio: {temp_prom:.1f} C', ln=True)
    
    pdf.ln(10)
    
    # Tabla
    pdf.set_font('Arial', 'B', 10)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(40, 10, 'Fecha', 1, 0, 'C', fill=True)
    pdf.cell(50, 10, 'N Remito', 1, 0, 'C', fill=True)
    pdf.cell(45, 10, 'Litros (Ticket)', 1, 0, 'C', fill=True)
    pdf.cell(45, 10, 'Temp (C)', 1, 1, 'C', fill=True)
    
    pdf.set_font('Arial', '', 10)
    for _, row in df_productor.iterrows():
        fecha_str = row['Fecha'].strftime('%d/%m/%Y')
        remito = str(row['N_Remito']) if pd.notnull(row['N_Remito']) else '-'
        litros = f"{row['Litros_Ticket']:,.0f}" if pd.notnull(row['Litros_Ticket']) else '0'
        temp = f"{row['Temperatura']:.1f}" if pd.notnull(row['Temperatura']) else '-'
        
        pdf.cell(40, 10, fecha_str, 1, 0, 'C')
        pdf.cell(50, 10, remito, 1, 0, 'C')
        pdf.cell(45, 10, litros, 1, 0, 'C')
        pdf.cell(45, 10, temp, 1, 1, 'C')
        
    return bytes(pdf.output(dest='S'), encoding='latin-1')

# 2. Lectura directa del archivo local
ruta_archivo = "Resumen planilla Recibo  OD-PRO-03.xlsx"

if os.path.exists(ruta_archivo):
    # Lectura y limpieza de datos
    df = pd.read_excel(ruta_archivo, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:I")
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura']
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha']) 
    
    st.sidebar.header("Filtros del Resumen")
    lista_tambos = sorted(pd.to_numeric(df['Num_Tambo'], errors='coerce').dropna().astype(int).unique())
    tambo_seleccionado = st.sidebar.selectbox("1. Selecciona el Código de Tambo:", lista_tambos)
    
    df_tambo = df[df['Num_Tambo'] == tambo_seleccionado]
    min_date = df_tambo['Fecha'].min().date()
    max_date = df_tambo['Fecha'].max().date()
    
    fechas_seleccionadas = st.sidebar.date_input("2. Selecciona la semana:", [min_date, max_date])
    
    if len(fechas_seleccionadas) == 2:
        fecha_inicio, fecha_fin = fechas_seleccionadas
        mask = (df_tambo['Fecha'].dt.date >= fecha_inicio) & (df_tambo['Fecha'].dt.date <= fecha_fin)
        df_filtrado = df_tambo.loc[mask]
        
        st.divider()
        nombre_tambo = df_filtrado['Tambo'].iloc[0] if not df_filtrado.empty else "Desconocido"
        st.subheader(f"Resumen Semanal - Tambo {tambo_seleccionado}: {nombre_tambo}")
        
        if not df_filtrado.empty:
            df_mostrar = df_filtrado[['Fecha', 'N_Remito', 'Litros_Ticket', 'Temperatura']].copy()
            df_mostrar['Fecha'] = df_mostrar['Fecha'].dt.strftime('%d/%m/%Y')
            st.dataframe(df_mostrar, hide_index=True, use_container_width=True)
            
            st.write("---")
            pdf_bytes = generar_pdf(
                df_filtrado, nombre_tambo, tambo_seleccionado, 
                fecha_inicio.strftime('%d/%m/%Y'), fecha_fin.strftime('%d/%m/%Y')
            )
            
            st.download_button(
                label="📄 Descargar PDF",
                data=pdf_bytes,
                file_name=f"Resumen_Tambo_{tambo_seleccionado}.pdf",
                mime="application/pdf"
            )
        else:
            st.warning("No hay retiros registrados para este tambo en estas fechas.")
else:
    st.error(f"No se encontró el archivo '{ruta_archivo}'. Asegúrate de que esté en la misma carpeta que este programa.")
