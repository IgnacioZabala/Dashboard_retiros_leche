import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
import zipfile
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Resumen Semanal de Recolección - Coopagro", layout="wide")
st.title("Panel de recolección y liquidación por Tambo")

# --- CONFIGURACIÓN DE GOOGLE DRIVE (Archivo Principal: Remitos y Contactos) ---
FILE_ID = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj" 
url_drive = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

@st.cache_data(ttl=60)
def cargar_datos_drive(url):
    df_remitos = pd.read_excel(url, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:K")
    df_contactos = pd.read_excel(url, sheet_name='Código Tambos')
    return df_remitos, df_contactos

# --- FUNCIONES DE FORMATO ---
def formato_miles(valor):
    return f"{valor:,.0f}".replace(',', '.')

def formato_temp(valor):
    if pd.isna(valor):
        return '-'
    return f"{valor:.1f}".replace('.', ',') + "°"

# --- FUNCIÓN PARA GENERAR EL PDF ---
def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, fecha_inicio, fecha_fin, comp_litros, comp_temp, mostrar_temp, mostrar_grasa, mostrar_prot, mostrar_comp, hay_datos_previos):
    pdf = FPDF()
    pdf.add_page()
    
    ruta_logo = "logo.png"
    if os.path.exists(ruta_logo):
        pdf.image(ruta_logo, x=80, y=10, w=50)
        pdf.set_y(48) 
    else:
        pdf.set_y(15)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 6, 'Resumen semanal de recoleccion', ln=True, align='C')
    
    pdf.set_text_color(0, 0, 0) 
    pdf.ln(4)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y()) 
    
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 7, f'Productor: {tambo_nombre} (Codigo #{tambo_id})', ln=True)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f'Periodo (Sabado a Viernes): {fecha_inicio} al {fecha_fin}', ln=True)
    
    total_litros = df_productor['Litros_Ticket'].sum()
    temp_prom = df_productor['Temperatura'].mean()
    grasa_prom = df_productor['Grasa'].mean() if 'Grasa' in df_productor.columns else float('nan')
    proteina_prom = df_productor['Proteina'].mean() if 'Proteina' in df_productor.columns else float('nan')
    
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 10)
    
    texto_litros = f'Total Litros: {formato_miles(total_litros)} L'
    if mostrar_comp and hay_datos_previos:
        texto_litros += f' ({comp_litros})'
    pdf.cell(0, 6, texto_litros, ln=True)
    
    if mostrar_temp:
        texto_temp = f'Temperatura Promedio: {formato_temp(temp_prom)}'
        if mostrar_comp and hay_datos_previos:
            texto_temp += f' ({comp_temp})'
        pdf.cell(0, 6, texto_temp, ln=True)
    
    if (mostrar_grasa or mostrar_prot) and (pd.notna(grasa_prom) or pd.notna(proteina_prom)):
        partes_solidos = []
        if mostrar_grasa:
            g_str = f"{grasa_prom:.2f}%".replace('.', ',') if pd.notna(grasa_prom) else "S/D"
            partes_solidos.append(f"Grasa: {g_str}")
        if mostrar_prot:
            p_str = f"{proteina_prom:.2f}%".replace('.', ',') if pd.notna(proteina_prom) else "S/D"
            partes_solidos.append(f"Proteina: {p_str}")
        
        pdf.cell(0, 6, f"Promedio Solidos -> {' | '.join(partes_solidos)}", ln=True)
    
    pdf.ln(6)
    
    pdf.set_font('Arial', 'B', 9)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(35, 8, 'Fecha', 1, 0, 'C', fill=True)
    pdf.cell(45, 8, 'N° de remito', 1, 0, 'C', fill=True)
    pdf.cell(35, 8, 'Litros', 1, 0, 'C', fill=True)
    
    if mostrar_temp:
        pdf.cell(30, 8, 'Temp', 1, 0, 'C', fill=True)
    if mostrar_grasa or mostrar_prot:
        pdf.cell(45, 8, 'Solidos (G/P)', 1, 1, 'C', fill=True)
    else:
        pdf.cell(0, 8, '', 0, 1)
    
    pdf.set_font('Arial', '', 9)
    for _, row in df_productor.iterrows():
        fecha_str = row['Fecha'].strftime('%d/%m/%Y')
        remito = str(row['N_Remito']) if pd.notna(row['N_Remito']) else '-'
        litros = formato_miles(row['Litros_Ticket']) if pd.notna(row['Litros_Ticket']) else '0'
        
        pdf.cell(35, 7, fecha_str, 1, 0, 'C')
        pdf.cell(45, 7, remito, 1, 0, 'C')
        pdf.cell(35, 7, litros, 1, 0, 'C')
        
        if mostrar_temp:
            temp_val = formato_temp(row['Temperatura'])
            pdf.cell(30, 7, temp_val, 1, 0, 'C')
            
        if mostrar_grasa or mostrar_prot:
            g_val = f"{row['Grasa']}%".replace('.', ',') if (mostrar_grasa and 'Grasa' in df_productor.columns and pd.notna(row['Grasa'])) else ('-' if mostrar_grasa else '')
            p_val = f"{row['Proteina']}%".replace('.', ',') if (mostrar_prot and 'Proteina' in df_productor.columns and pd.notna(row['Proteina'])) else ('-' if mostrar_prot else '')
            
            if mostrar_grasa and mostrar_prot:
                solidos_str = f"{g_val} / {p_val}"
            else:
                solidos_str = g_val if mostrar_grasa else p_val
                
            pdf.cell(45, 7, solidos_str, 1, 1, 'C')
        else:
            pdf.ln(7)
        
    return bytes(pdf.output(dest='S'), encoding='latin-1')

# --- FUNCIÓN PARA ENVIAR CORREO ---
def enviar_correo_productor(destinatario_email, nombre_contacto, tambo_nombre, pdf_bytes, nombre_archivo):
    try:
        remitente = st.secrets["email"]["remitente"]
        password = st.secrets["email"]["password"]
        
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario_email
        msg['Subject'] = f"Resumen Semanal de Recolección - {tambo_nombre}"
        
        cuerpo_html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
            <p>Buenas tardes, <b>{nombre_contacto}</b>:</p>
            <p>Te adjunto el resumen correspondiente a la recolección de leche y calidad de esta semana para el establecimiento <b>{tambo_nombre}</b>.</p>
            <p>Cualquier consulta quedo a tu disposición.</p>
            <br>
            <p>Saludos cordiales,</p>
            <hr style="border: none; border-top: 1px solid #ccc; width: 300px; text-align: left;">
            <table style="font-size: 13px; color: #555;">
                <tr>
                    <td style="vertical-align: middle; padding-right: 15px;">
                        <img src="https://i.imgur.com/tu_foto_ejemplo.png" width="70" style="border-radius: 50%;">
                    </td>
                    <td style="vertical-align: middle;">
                        <b>Ignacio Zabala</b><br>
                        Coopagro Planta Tandil<br>
                        <i>Gestión y Calidad</i>
                    </td>
                </tr>
            </table>
            <br>
            <img src="https://i.imgur.com/tu_banner_coopagro.png" width="400">
        </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo_html, 'html'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{nombre_archivo}"')
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, destinatario_email, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando correo: {e}")
        return False

# --- CARGA Y PROCESAMIENTO DE DATOS ---
try:
    df_raw, df_contactos_raw = cargar_datos_drive(url_drive)
    
    # Selector en barra lateral para el archivo de laboratorio externo
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔬 Archivo de Laboratorio")
    archivo_lab_subido = st.sidebar.file_uploader("Subir Excel de Lab (Grasa/Proteína)", type=["xlsx", "xls"])

    # Procesar contactos de tambos de forma segura
    df_contactos = df_contactos_raw.copy()
    df_contactos.columns = df_contactos.columns.astype(str).str.strip()
    
    col_codigo = 'Código' if 'Código' in df_contactos.columns else df_contactos.columns[1]
    col_contacto = 'Contacto (nombre)' if 'Contacto (nombre)' in df_contactos.columns else df_contactos.columns[3]
    col_email = 'Email' if 'Email' in df_contactos.columns else df_contactos.columns[4]
    
    df_contactos['Num_Tambo'] = df_contactos[col_codigo].astype(str).str.strip()
    df_contactos['Contacto_Nombre'] = df_contactos[col_contacto]
    df_contactos['Email'] = df_contactos[col_email]
    
    # Procesamiento de Planilla Principal de Remitos
    df = df_raw.copy()
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura', 'Grasa', 'Proteina']
    df['Num_Tambo'] = df['Num_Tambo'].astype(str).str.strip()
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha']) 

 # --- CRUCE SEGURO CON EL ARCHIVO DE LABORATORIO SUBIDO ---
    if archivo_lab_subido is not None:
        try:
            # Leemos el excel de laboratorio
            df_lab = pd.read_excel(archivo_lab_subido)
            df_lab.columns = df_lab.columns.astype(str).str.strip()
            
            # Buscar la columna que contenga las muestras (ej: Sample number)
            matching_cols = [c for c in df_lab.columns if 'sample' in c.lower() or 'number' in c.lower()]
            col_sample = matching_cols[0] if matching_cols else df_lab.columns[0]
            
            lista_tambo = []
            lista_fecha = []
            
            # Recorremos fila por fila de forma segura para evitar index out of bounds
            for val in df_lab[col_sample]:
                if pd.isna(val):
                    lista_tambo.append(None)
                    lista_fecha.append(None)
                    continue
                
                partes = str(val).strip().split()
                # Validamos estrictamente que tenga al menos el formato 'T10 29072026'
                if len(partes) >= 2:
                    tambo_limpio = partes[0].replace('T', '').replace('t', '').strip()
                    try:
                        fecha_limpia = pd.to_datetime(partes[1], format='%d%m%Y', errors='coerce')
                    except:
                        fecha_limpia = None
                    
                    lista_tambo.append(tambo_limpio)
                    lista_fecha.append(fecha_limpia)
                else:
                    lista_tambo.append(None)
                    lista_fecha.append(None)
            
            df_lab['Num_Tambo'] = lista_tambo
            df_lab['Fecha'] = lista_fecha
            
            # Búsqueda flexible de columnas de Grasa (Fat) y Proteína (Protein)
            col_fat = next((c for c in df_lab.columns if 'fat' in c.lower() or 'grasa' in c.lower()), None)
            col_prot = next((c for c in df_lab.columns if 'protein' in c.lower() or 'proteina' in c.lower()), None)
            
            if col_fat:
                df_lab['Grasa_Lab'] = pd.to_numeric(df_lab[col_fat], errors='coerce')
            if col_prot:
                df_lab['Proteina_Lab'] = pd.to_numeric(df_lab[col_prot], errors='coerce')
                
            cols_merge = ['Num_Tambo', 'Fecha']
            if col_fat: cols_merge.append('Grasa_Lab')
            if col_prot: cols_merge.append('Proteina_Lab')
            
            df_lab_clean = df_lab[cols_merge].dropna(subset=['Fecha', 'Num_Tambo'])
            
            # Cruzar con la tabla principal
            df = pd.merge(df, df_lab_clean, on=['Num_Tambo', 'Fecha'], how='left')
            
            if 'Grasa_Lab' in df.columns:
                df['Grasa'] = df['Grasa_Lab'].combine_first(df['Grasa'])
            if 'Proteina_Lab' in df.columns:
                df['Proteina'] = df['Proteina_Lab'].combine_first(df['Proteina'])
                
        except Exception as err_lab:
            st.sidebar.error(f"Error procesando lab: {err_lab}")
    
    # Agrupación por Ciclo Operativo (Sábado a Viernes)
    df['Fecha_Cierre_Viernes'] = df['Fecha'] + pd.to_timedelta((4 - df['Fecha'].dt.weekday) % 7, unit='D')
    df['Fecha_Inicio_Sabado'] = df['Fecha_Cierre_Viernes'] - pd.Timedelta(days=6)
    
    df['Ciclo_Semana'] = df.apply(lambda r: f"Viernes {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')} (Sab {r['Fecha_Inicio_Sabado'].strftime('%d/%m/%Y')} al Vie {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')})", axis=1)
    
    # --- BARRA LATERAL ---
    st.sidebar.header("Filtros de Reporte")
    
    ciclos_disponibles = df[['Fecha_Cierre_Viernes', 'Ciclo_Semana']].drop_duplicates().sort_values('Fecha_Cierre_Viernes', ascending=False)['Ciclo_Semana'].tolist()
    ciclo_seleccionado = st.sidebar.selectbox("1. Selecciona el Cierre de Semana (Viernes):", ciclos_disponibles)
    
    df_semana_actual = df[df['Ciclo_Semana'] == ciclo_seleccionado]
    
    mapeo_tambos = df_semana_actual[['Tambo', 'Num_Tambo']].dropna().drop_duplicates()
    mapeo_tambos = mapeo_tambos.sort_values(by='Tambo', ascending=True)
    
    nombres_tambos_ordenados = mapeo_tambos['Tambo'].tolist()
    tambo_nombre_seleccionado = st.sidebar.selectbox("2. Selecciona el Tambo:", nombres_tambos_ordenados)
    
    tambo_seleccionado = mapeo_tambos[mapeo_tambos['Tambo'] == tambo_nombre_seleccionado]['Num_Tambo'].values[0]

    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Elementos del Reporte")
    ver_temperatura = st.sidebar.checkbox("Incluir Temperatura", value=True)
    ver_grasa = st.sidebar.checkbox("Incluir Grasa", value=True)
    ver_proteina = st.sidebar.checkbox("Incluir Proteína", value=True)
    ver_comparacion = st.sidebar.checkbox("Incluir Comparativa vs. Semana Ant.", value=True)

    st.sidebar.divider()
    st.sidebar.subheader("📦 Envío Masivo & ZIP")
    generar_lote = st.sidebar.button("Generar ZIP con todos los Tambos")

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
                    
                    f_v_act = df_t['Fecha_Cierre_Viernes'].iloc[0]
                    f_v_ant = f_v_act - pd.Timedelta(days=7)
                    df_t_ant = df[(df['Num_Tambo'] == t_id) & (df['Fecha_Cierre_Viernes'] == f_v_ant)]
                    tiene_prev = not df_t_ant.empty
                    
                    comp_l = f"{((df_t['Litros_Ticket'].sum() - df_t_ant['Litros_Ticket'].sum()) / df_t_ant['Litros_Ticket'].sum()) * 100:+.1f}% vs. semana ant.".replace('.', ',') if tiene_prev and df_t_ant['Litros_Ticket'].sum() > 0 else ""
                    comp_t = f"{df_t['Temperatura'].mean() - df_t_ant['Temperatura'].mean():+.1f}° vs. semana ant.".replace('.', ',') if tiene_prev else ""
                    
                    pdf_data = generar_pdf_bytes(df_t, t_name, t_id, f_ini, f_fin, comp_l, comp_t, ver_temperatura, ver_grasa, ver_proteina, ver_comparacion, tiene_prev)
                    zip_file.writestr(f"Resumen_Tambo_{t_id}_{t_name.replace(' ', '_')}_Cierre_{f_fin.replace('/', '-')}.pdf", pdf_data)
            
        zip_buffer.seek(0)
        st.sidebar.download_button(
            label="📥 Descargar ZIP con todos los PDFs",
            data=zip_buffer,
            file_name=f"Resumenes_Cierre_{ciclo_seleccionado[:15].strip()}.zip",
            mime="application/zip"
        )

    # --- VISTA INDIVIDUAL Y ENVÍO DE MAIL ---
    st.divider()
    df_tambo_semana = df_semana_actual[df_semana_actual['Num_Tambo'] == str(tambo_seleccionado)].sort_values('Fecha')
    
    if not df_tambo_semana.empty:
        f_inicio = df_tambo_semana['Fecha_Inicio_Sabado'].iloc[0].strftime('%d/%m/%Y')
        f_fin = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0].strftime('%d/%m/%Y')
        
        st.subheader(f"Resumen Cierre Viernes ({f_inicio} al {f_fin}) - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
        
        info_contacto = df_contactos[df_contactos['Num_Tambo'] == str(tambo_seleccionado)]
        email_tambo = ""
        nombre_contacto = "Productor"
        if not info_contacto.empty:
            val_email = info_contacto['Email'].values[0]
            val_nombre = info_contacto['Contacto_Nombre'].values[0]
            if pd.notna(val_email) and "@" in str(val_email):
                email_tambo = str(val_email).strip()
            if pd.notna(val_nombre):
                nombre_contacto = str(val_nombre).strip()

        fecha_viernes_actual = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0]
        fecha_viernes_anterior = fecha_viernes_actual - pd.Timedelta(days=7)
        df_tambo_anterior = df[(df['Num_Tambo'] == str(tambo_seleccionado)) & (df['Fecha_Cierre_Viernes'] == fecha_viernes_anterior)]
        
        litros_actual = df_tambo_semana['Litros_Ticket'].sum()
        temp_actual = df_tambo_semana['Temperatura'].mean()
        grasa_actual = df_tambo_semana['Grasa'].mean() if 'Grasa' in df_tambo_semana.columns else float('nan')
        prot_actual = df_tambo_semana['Proteina'].mean() if 'Proteina' in df_tambo_semana.columns else float('nan')
        
        hay_datos_previos = not df_tambo_anterior.empty
        comp_litros_str = ""
        comp_temp_str = ""
        
        cols_a_mostrar = sum([1, ver_temperatura, ver_grasa, ver_proteina])
        c1 = st.columns(cols_a_mostrar)
        
        col_idx = 0
        if hay_datos_previos:
            litros_anterior = df_tambo_anterior['Litros_Ticket'].sum()
            temp_anterior = df_tambo_anterior['Temperatura'].mean()
            
            diff_litros_pct = ((litros_actual - litros_anterior) / litros_anterior) * 100 if litros_anterior > 0 else 0
            diff_temp = temp_actual - temp_anterior
            
            delta_litros_val = f"{diff_litros_pct:+.1f}% vs. semana ant.".replace('.', ',') if ver_comparacion else None
            delta_temp_val = f"{diff_temp:+.1f}° vs. semana ant.".replace('.', ',') if ver_comparacion else None
            
            comp_litros_str = f"{diff_litros_pct:+.1f}% vs. semana ant.".replace('.', ',')
            comp_temp_str = f"{diff_temp:+.1f}° vs. semana ant.".replace('.', ',')
            
            c1[col_idx].metric("Litros", f"{formato_miles(litros_actual)} L", delta=delta_litros_val)
            col_idx += 1
            if ver_temperatura:
                c1[col_idx].metric("Temp. Promedio", formato_temp(temp_actual), delta=delta_temp_val, delta_color="inverse")
                col_idx += 1
        else:
            c1[col_idx].metric("Litros", f"{formato_miles(litros_actual)} L")
            col_idx += 1
            if ver_temperatura:
                c1[col_idx].metric("Temp. Promedio", formato_temp(temp_actual))
                col_idx += 1
                
        if ver_grasa:
            c1[col_idx].metric("Grasa Promedio", f"{grasa_actual:.2f}%".replace('.', ',') if pd.notna(grasa_actual) else "S/D")
            col_idx += 1
        if ver_proteina:
            c1[col_idx].metric("Proteína Promedio", f"{prot_actual:.2f}%".replace('.', ',') if pd.notna(prot_actual) else "S/D")
            col_idx += 1
        
        st.markdown("---")
        st.markdown("**Detalle de retiros del período:**")
        
        columnas_visibles = ['Fecha', 'N_Remito', 'Litros_Ticket']
        if ver_temperatura:
            columnas_visibles.append('Temperatura')
        if ver_grasa:
            columnas_visibles.append('Grasa')
        if ver_proteina:
            columnas_visibles.append('Proteina')
            
        df_mostrar = df_tambo_semana[columnas_visibles].copy()
        df_mostrar['Litros_Ticket'] = df_mostrar['Litros_Ticket'].apply(lambda x: formato_miles(x) if pd.notna(x) else '0')
        
        if ver_temperatura:
            df_mostrar['Temperatura'] = df_mostrar['Temperatura'].apply(lambda x: formato_temp(x))
            
        if 'Grasa' in df_mostrar.columns:
            df_mostrar['Grasa'] = df_mostrar['Grasa'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
        if 'Proteina' in df_mostrar.columns:
            df_mostrar['Proteina'] = df_mostrar['Proteina'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
            
        df_mostrar = df_mostrar.rename(columns={
            'Litros_Ticket': 'Litros',
            'N_Remito': 'N° de remito',
            'Temperatura': 'Temp'
        })
        
        df_mostrar['Fecha'] = df_mostrar['Fecha'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_mostrar, hide_index=True, use_container_width=True)
        
        pdf_bytes = generar_pdf_bytes(
            df_tambo_semana, tambo_nombre_seleccionado, tambo_seleccionado, 
            f_inicio, f_fin, comp_litros_str, comp_temp_str, 
            ver_temperatura, ver_grasa, ver_proteina, ver_comparacion, hay_datos_previos
        )
        
        nombre_pdf_salida = f"Resumen_{tambo_nombre_seleccionado.replace(' ', '_')}_Cierre_{f_fin.replace('/', '-')}.pdf"

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            st.download_button(
                label=f"📄 Descargar PDF de {tambo_nombre_seleccionado}",
                data=pdf_bytes,
                file_name=nombre_pdf_salida,
                mime="application/pdf"
            )
        with col_btn2:
            st.info(f"📧 Correo registrado: **{email_tambo if email_tambo else 'No cargado o sin formato de email'}**")
            if st.button(f"📤 Enviar mail a {nombre_contacto}"):
                if not email_tambo:
                    st.error("No se puede enviar el correo porque este tambo no tiene un email válido registrado en la solapa 'Código Tambos'.")
                else:
                    with st.spinner("Enviando correo electrónico..."):
                        exito = enviar_correo_productor(email_tambo, nombre_contacto, tambo_nombre_seleccionado, pdf_bytes, nombre_pdf_salida)
                        if exito:
                            st.success(f"¡Correo enviado exitosamente a {email_tambo}!")
                        else:
                            st.error("Hubo un error al enviar el correo. Revisa la configuración de las credenciales SMTP.")
    else:
        st.warning("No hay registros para este tambo en el período seleccionado.")
        
except Exception as e:
    st.error(f"Error al cargar o procesar el archivo desde Google Drive o Laboratorio. Detalle técnico: {e}")
