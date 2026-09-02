import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
import zipfile
import io

st.set_page_config(page_title="Resumen Semanal de Recolección - Coopagro", layout="wide")
st.title("🚜 Panel de Recolección y Liquidación por Tambo")

# --- FUNCIÓN PARA GENERAR EL PDF ---
def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, fecha_inicio, fecha_fin, comp_litros, comp_temp):
    pdf = FPDF()
    pdf.add_page()
    
    # Logo
    ruta_logo = "logo.png"
    espacio_izquierdo = 35 if os.path.exists(ruta_logo) else 0
    if espacio_izquierdo > 0: pdf.image(ruta_logo, 10, 8, 30) 
    
    pdf.set_font('Arial', 'B', 16)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, 'Coopagro - Planta Tandil', ln=True, align='L')
    
    pdf.set_font('Arial', '', 12)
    pdf.set_text_color(100, 100, 100)
    if espacio_izquierdo > 0: pdf.cell(espacio_izquierdo)
    pdf.cell(0, 10, 'Resumen Semanal de Recoleccion', ln=True, align='L')
    
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
    df = pd.read_excel(uploaded_file, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:K")
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura', 'Grasa', 'Proteina']
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha']) 
    
    # Agrupación por Ciclo Operativo (Sábado a Viernes)
    # Calculamos el próximo viernes de cada fecha para usarlo como la "Clave de Cierre" de la semana
    df['Fecha_Cierre_Viernes'] = df['Fecha'] + pd.to_timedelta((4 - df['Fecha'].dt.weekday) % 7, unit='D')
    df['Fecha_Inicio_Sabado'] = df['Fecha_Cierre_Viernes'] - pd.Timedelta(days=6)
    
    # Crear etiqueta amigable para el selector (Ej: "Cierre Viernes 05/09/2025 (Del 30/08 al 05/09)")
    df['Ciclo_Semana'] = df.apply(lambda r: f"Viernes {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')} (Sab {r['Fecha_Inicio_Sabado'].strftime('%d/%m/%Y')} al Vie {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')})", axis=1)
    
    # --- BARRA LATERAL ---
    st.sidebar.header("Filtros de Reporte")
    
    # 1. Selector de Ciclo Semanal ordenado cronológicamente
    ciclos_disponibles = df[['Fecha_Cierre_Viernes', 'Ciclo_Semana']].drop_duplicates().sort_values('Fecha_Cierre_Viernes', ascending=False)['Ciclo_Semana'].tolist()
    ciclo_seleccionado = st.sidebar.selectbox("1. Selecciona el Cierre de Semana (Viernes):", ciclos_disponibles)
    
    df_semana_actual = df[df['Ciclo_Semana'] == ciclo_seleccionado]
    
    # 2. Selector de Tambo ordenado alfabéticamente por Nombre
    mapeo_tambos = df_semana_actual[['Tambo', 'Num_Tambo']].dropna().drop_duplicates()
    mapeo_tambos = mapeo_tambos.sort_values(by='Tambo', ascending=True)
    
    nombres_tambos_ordenados = mapeo_tambos['Tambo'].tolist()
    tambo_nombre_seleccionado = st.sidebar.selectbox("2. Selecciona el Tambo:", nombres_tambos_ordenados)
    
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
                    f_ini = df_t['Fecha_Inicio_Sabado'].iloc[0].strftime('%d/%m/%Y')
                    f_fin = df_t['Fecha_Cierre_Viernes'].iloc[0].strftime('%d/%m/%Y')
                    pdf_data = generar_pdf_bytes(df_t, t_name, t_id, f_ini, f_fin, "Sin comp.", "Sin comp.")
                    zip_file.writestr(f"Resumen_Tambo_{t_id}_{t_name.replace(' ', '_')}_Cierre_{f_fin.replace('/', '-')}.pdf", pdf_data)
        
        zip_buffer.seek(0)
        st.sidebar.download_button(
            label="📥 Descargar ZIP con todos los PDFs",
            data=zip_buffer,
            file_name=f"Resumenes_Cierre_{ciclo_seleccionado[:15].strip()}.zip",
            mime="application/zip"
        )

    # --- VISTA INDIVIDUAL Y COMPARATIVA ---
    st.divider()
    df_tambo_semana = df_semana_actual[df_semana_actual['Num_Tambo'] == tambo_seleccionado].sort_values('Fecha')
    
    if not df_tambo_semana.empty:
        f_inicio = df_tambo_semana['Fecha_Inicio_Sabado'].iloc[0].strftime('%d/%m/%Y')
        f_fin = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0].strftime('%d/%m/%Y')
        
        st.subheader(f"Resumen Cierre Viernes ({f_inicio} al {f_fin}) - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
        
        # Buscar semana anterior para comparativa exacta
        fecha_viernes_actual = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0]
        fecha_viernes_anterior = fecha_viernes_actual - pd.Timedelta(days=7)
        df_tambo_anterior = df[(df['Num_Tambo'] == tambo_seleccionado) & (df['Fecha_Cierre_Viernes'] == fecha_viernes_anterior)]
        
        litros_actual = df_tambo_semana['Litros_Ticket'].sum()
        temp_actual = df_tambo_semana['Temperatura'].mean()
        grasa_actual = df_tambo_semana['Grasa'].mean() if 'Grasa' in df_tambo_semana.columns else float('nan')
        prot_actual = df_tambo_semana['Proteina'].mean() if 'Proteina' in df_tambo_semana.columns else float('nan')
        
        comp_litros_str = "Sin datos periodo previo"
        comp_temp_str = "Sin datos periodo previo"
        
        col1, col2, col3, col4 = st.columns(4)
        
        if not df_tambo_anterior.empty:
            litros_anterior = df_tambo_anterior['Litros_Ticket'].sum()
            temp_anterior = df_tambo_anterior['Temperatura'].mean()
            
            diff_litros_pct = ((litros_actual - litros_anterior) / litros_anterior) * 100 if litros_anterior > 0 else 0
            diff_temp = temp_actual - temp_anterior
            
            col1.metric("Litros (Ticket)", f"{litros_actual:,.0f} L", delta=f"{diff_litros_pct:+.1f}% vs. Per. Ant.")
            col2.metric("Temp. Promedio", f"{temp_actual:.1f} °C", delta=f"{diff_temp:+.1f} °C", delta_color="inverse")
            
            comp_litros_str = f"{diff_litros_pct:+.1f}% vs. Per. Ant."
            comp_temp_str = f"{diff_temp:+.1f} °C vs. Per. Ant."
        else:
            col1.metric("Litros (Ticket)", f"{litros_actual:,.0f} L")
            col2.metric("Temp. Promedio", f"{temp_actual:.1f} °C")
            
        col3.metric("Grasa Promedio", f"{grasa_actual:.2f}%" if pd.notna(grasa_actual) else "S/D")
        col4.metric("Proteína Promedio", f"{prot_actual:.2f}%" if pd.notna(prot_actual) else "S/D")
        
        st.markdown("---")
        st.markdown("**Detalle de retiros del período:**")
        df_mostrar = df_tambo_semana[['Fecha', 'N_Remito', 'Litros_Ticket', 'Temperatura', 'Grasa', 'Proteina']].copy()
        df_mostrar['Fecha'] = df_mostrar['Fecha'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_mostrar, hide_index=True, use_container_width=True)
        
        pdf_bytes = generar_pdf_bytes(
            df_tambo_semana, tambo_nombre_seleccionado, tambo_seleccionado, 
            f_inicio, f_fin, comp_litros_str, comp_temp_str
        )
        
        st.download_button(
            label=f"📄 Descargar PDF de {tambo_nombre_seleccionado} (Cierre {f_fin})",
            data=pdf_bytes,
            file_name=f"Resumen_{tambo_nombre_seleccionado.replace(' ', '_')}_Cierre_{f_fin.replace('/', '-')}.pdf",
            mime="application/pdf"
        )
    else:
        st.warning("No hay registros para este tambo en el período seleccionado.")
