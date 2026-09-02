import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
import zipfile
import io

st.set_page_config(page_title="Resumen Semanal de Recolección - Coopagro", layout="wide")
st.title("Panel de recolección y liquidación por tambo")

# --- FUNCIÓN PARA GENERAR EL PDF ---
def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, fecha_inicio, fecha_fin, num_semana, comp_litros, comp_temp):
    pdf = FPDF()
    pdf.add_page()
    
    ruta_logo = "logo.png"
    espacio_izquierdo = 35 if os.path.exists(ruta_logo) else 0
    if espacio_izquierdo > 0: pdf.image(ruta_logo, 10, 8, 30) 
    
    pdf.set_font('Arial', 'B', 16)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, 'Coopagro - Planta Tandil', ln=True, align='L')
    
    pdf.set_font('Arial', '', 12)
    pdf.set_text_color(100, 100, 100)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, f'Resumen de Recoleccion - Semana {num_semana}', ln=True, align='L')
    
    pdf.ln(5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 11)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 7, f'Productor: {tambo_nombre} (Codigo #{tambo_id})', ln=True)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f'Periodo (Sabado a Viernes): {fecha_inicio} al {fecha_fin}', ln=True)
    
    total_litros = df_productor['Litros_Ticket'].sum()
    temp_prom = df_productor['Temperatura'].mean()
    grasa_prom = df_productor['Grasa'].mean() if 'Grasa' in df_productor.columns else float('nan')
    proteina_prom = df_productor['Proteina'].mean() if 'Proteina' in df_productor.columns else float('nan')
    
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, f'Total Litros: {total_litros:,.0f} L ({comp_litros})', ln=True)
    pdf.cell(0, 6, f'Temperatura Promedio: {temp_prom:.1f} C ({comp_temp})', ln=True)
    
    if pd.notna(grasa_prom) or pd.notna(proteina_prom):
        g_str = f"{grasa_prom:.2f}%" if pd.notna(grasa_prom) else "S/D"
        p_str = f"{proteina_prom:.2f}%" if pd.notna(proteina_prom) else "S/D"
        pdf.cell(0, 6, f'Promedio Solidos -> Grasa: {g_str} | Proteina: {p_str}', ln=True)
    
    pdf.ln(6)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(35, 8, 'Fecha', 1, 0, 'C', fill=True)
    pdf.cell(45, 8, 'N Remito', 1, 0, 'C', fill=True)
    pdf.cell(35, 8, 'Litros (Ticket)', 1, 0, 'C', fill=True)
    pdf.cell(30, 8, 'Temp (C)', 1, 0, 'C', fill=True)
    pdf.cell(45, 8, 'Solidos (G/P)', 1, 1, 'C', fill=True)
    
    pdf.set_font('Arial', '', 9)
    for _, row in df_productor.iterrows():
        fecha_str = row['Fecha'].strftime('%d/%m/%Y')
        remito = str(row['N_Remito']) if pd.notna(row['N_Remito']) else '-'
        litros = f"{row['Litros_Ticket']:,.0f}" if pd.notna(row['Litros_Ticket']) else '0'
        temp = f"{row['Temperatura']:.1f}" if pd.notna(row['Temperatura']) else '-'
        
        g_val = f"{row['Grasa']}%" if ('Grasa' in row and pd.notna(row['Grasa'])) else '-'
        p_val = f"{row['Proteina']}%" if ('Proteina' in row and pd.notna(row['Proteina'])) else '-'
        solidos_str = f"{g_val} / {p_val}" if g_val != '-' or p_val != '-' else '-'
        
        pdf.cell(35, 7, fecha_str, 1, 0, 'C')
        pdf.cell(45, 7, remito, 1, 0, 'C')
        pdf.cell(35, 7, litros, 1, 0, 'C')
        pdf.cell(30, 7, temp, 1, 0, 'C')
        pdf.cell(45, 7, solidos_str, 1, 1, 'C')
        
    return bytes(pdf.output(dest='S'), encoding='latin-1')

# --- CARGA DEL ARCHIVO ---
uploaded_file = st.file_uploader("Sube tu archivo 'Resumen planilla Recibo  OD-PRO-03.xlsx'", type=["xlsx"])

if uploaded_file:
    # Leer datos principales
    df = pd.read_excel(uploaded_file, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:K")
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura', 'Grasa', 'Proteina']
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha']) 
    
    # Lógica de Semana Lechera (Sábado a Viernes)
    # Al sumarle 1 día antes de calcular la semana ISO, logramos que Sábado y Domingo 
    # queden agrupados en la misma semana operativa que arranca el Sábado.
    df['Fecha_Ajustada'] = df['Fecha'] + pd.Timedelta(days=1)
    df['Semana'] = df['Fecha_Ajustada'].dt.isocalendar().week
    
    # --- BARRA LATERAL ---
    st.sidebar.header("Filtros de Reporte")
    
    # 1. Selector de Semana
    semanas_disponibles = sorted(df['Semana'].unique())
    semana_seleccionada = st.sidebar.selectbox("1. Selecciona la Semana:", semanas_disponibles)
    
    # Filtrar datos de la semana seleccionada
    df_semana_actual = df[df['Semana'] == semana_seleccionada]
    
    # 2. Selector de Tambo (Ordenado alfabéticamente por Nombre)
    # Obtenemos los pares únicos de Nombre y Número de Tambo y los ordenamos por Nombre (A-Z)
    mapeo_tambos = df_semana_actual[['Tambo', 'Num_Tambo']].dropna().drop_duplicates()
    mapeo_tambos = mapeo_tambos.sort_values(by='Tambo', ascending=True)
    
    nombres_tambos_ordenados = mapeo_tambos['Tambo'].tolist()
    tambo_nombre_seleccionado = st.sidebar.selectbox("2. Selecciona el Tambo:", nombres_tambos_ordenados)
    
    # Obtener el número de tambo correspondiente al nombre elegido
    tambo_seleccionado = mapeo_tambos[mapeo_tambos['Tambo'] == tambo_nombre_seleccionado]['Num_Tambo'].values[0]

    st.sidebar.divider()
    st.sidebar.subheader("📦 Envío Masivo")
    generar_lote = st.sidebar.button("Generar ZIP con todos los Tambos")

    # --- ACCIÓN DE ENVÍO EN LOTE (ZIP) ---
    if generar_lote:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for _, row_t in mapeo_tambos.iterrows():
                t_name = row_t['Tambo']
                t_id = row_t['Num_Tambo']
                df_t = df_semana_actual[df_semana_actual['Num_Tambo'] == t_id].sort_values('Fecha')
                if not df_t.empty:
                    f_ini = df_t['Fecha'].min().strftime('%d/%m/%Y')
                    f_fin = df_t['Fecha'].max().strftime('%d/%m/%Y')
                    pdf_data = generar_pdf_bytes(df_t, t_name, t_id, f_ini, f_fin, semana_seleccionada, "Sin comp.", "Sin comp.")
                    zip_file.writestr(f"Resumen_Tambo_{t_id}_{t_name.replace(' ', '_')}_Semana_{semana_seleccionada}.pdf", pdf_data)
        
        zip_buffer.seek(0)
        st.sidebar.download_button(
            label="📥 Descargar ZIP con todos los PDFs",
            data=zip_buffer,
            file_name=f"Resumenes_Semana_{semana_seleccionada}.zip",
            mime="application/zip"
        )

    # --- VISTA INDIVIDUAL Y COMPARATIVA ---
    st.divider()
    df_tambo_semana = df_semana_actual[df_semana_actual['Num_Tambo'] == tambo_seleccionado].sort_values('Fecha')
    
    if not df_tambo_semana.empty:
        f_inicio = df_tambo_semana['Fecha'].min()
        f_fin = df_tambo_semana['Fecha'].max()
        
        st.subheader(f"Resumen Semana {semana_seleccionada} - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
        st.caption(f"Período operativo (Sábado a Viernes): {f_inicio.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}")
        
        # Calcular semana anterior para comparativa
        df_tambo_anterior = df[(df['Num_Tambo'] == tambo_seleccionado) & (df['Semana'] == semana_seleccionada - 1)]
        
        litros_actual = df_tambo_semana['Litros_Ticket'].sum()
        temp_actual = df_tambo_semana['Temperatura'].mean()
        grasa_actual = df_tambo_semana['Grasa'].mean() if 'Grasa' in df_tambo_semana.columns else float('nan')
        prot_actual = df_tambo_semana['Proteina'].mean() if 'Proteina' in df_tambo_semana.columns else float('nan')
        
        comp_litros_str = "Sin datos semana previa"
        comp_temp_str = "Sin datos semana previa"
        
        col1, col2, col3, col4 = st.columns(4)
        
        if not df_tambo_anterior.empty:
            litros_anterior = df_tambo_anterior['Litros_Ticket'].sum()
            temp_anterior = df_tambo_anterior['Temperatura'].mean()
            
            diff_litros_pct = ((litros_actual - litros_anterior) / litros_anterior) * 100 if litros_anterior > 0 else 0
            diff_temp = temp_actual - temp_anterior
            
            col1.metric("Litros (Ticket)", f"{litros_actual:,.0f} L", delta=f"{diff_litros_pct:+.1f}% vs. Sem. Ant.")
            col2.metric("Temp. Promedio", f"{temp_actual:.1f} °C", delta=f"{diff_temp:+.1f} °C", delta_color="inverse")
            
            comp_litros_str = f"{diff_litros_pct:+.1f}% vs. Sem. Ant."
            comp_temp_str = f"{diff_temp:+.1f} °C vs. Sem. Ant."
        else:
            col1.metric("Litros (Ticket)", f"{litros_actual:,.0f} L")
            col2.metric("Temp. Promedio", f"{temp_actual:.1f} °C")
            
        col3.metric("Grasa Promedio", f"{grasa_actual:.2f}%" if pd.notna(grasa_actual) else "S/D")
        col4.metric("Proteína Promedio", f"{prot_actual:.2f}%" if pd.notna(prot_actual) else "S/D")
        
        st.markdown("---")
        st.markdown("**Detalle de retiros de la semana:**")
        df_mostrar = df_tambo_semana[['Fecha', 'N_Remito', 'Litros_Ticket', 'Temperatura', 'Grasa', 'Proteina']].copy()
        df_mostrar['Fecha'] = df_mostrar['Fecha'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_mostrar, hide_index=True, use_container_width=True)
        
        pdf_bytes = generar_pdf_bytes(
            df_tambo_semana, tambo_nombre_seleccionado, tambo_seleccionado, 
            f_inicio.strftime('%d/%m/%Y'), f_fin.strftime('%d/%m/%Y'),
            semana_seleccionada, comp_litros_str, comp_temp_str
        )
        
        st.download_button(
            label=f"📄 Descargar PDF de {tambo_nombre_nombre_si_aplica = 'Tambo'} (Semana {semana_seleccionada})".replace("_si_aplica = 'Tambo'", ""),
            data=pdf_bytes,
            file_name=f"Resumen_{tambo_nombre_seleccionado.replace(' ', '_')}_Semana_{semana_seleccionada}.pdf",
            mime="application/pdf"
        )
    else:
        st.warning("No hay registros para este tambo en la semana seleccionada.")
