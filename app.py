import streamlit as st
import pandas as pd
import re
import traceback
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
st.title("🚜 Panel de Recolección y Liquidación por Tambo")

# --- CONFIGURACIÓN DE GOOGLE DRIVE ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj" 
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"      

url_remitos = f"https://drive.google.com/uc?export=download&id={FILE_ID_REMITOS}"
url_lab = f"https://drive.google.com/uc?export=download&id={FILE_ID_LAB}"

@st.cache_data(ttl=60)
def cargar_datos_drive(u_remitos, u_lab):
    df_remitos_raw = pd.DataFrame()
    df_contactos = pd.DataFrame()
    df_lab = pd.DataFrame()
    
    try:
        # Volvemos al rango exacto de columnas B:K y skiprows=4 que requiere tu planilla original
        df_remitos_raw = pd.read_excel(u_remitos, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:K")
    except Exception as e:
        st.error(f"Error al leer la solapa de Remitos en Drive: {e}")
        
    try:
        df_contactos = pd.read_excel(u_remitos, sheet_name='Código Tambos')
    except Exception as e:
        st.warning(f"No se pudo leer la solapa 'Código Tambos': {e}")
        
    try:
        xls_lab = pd.ExcelFile(u_lab)
        df_lab_temp = pd.read_excel(u_lab, sheet_name=xls_lab.sheet_names[0], header=None)
        header_row = 0
        for idx, row in df_lab_temp.iterrows():
            row_str = " ".join([str(x).lower() for x in row.dropna() if pd.notna(x)])
            if 'sample' in row_str or 'fat' in row_str or 'protein' in row_str or 'grasa' in row_str:
                header_row = idx
                break
        df_lab = pd.read_excel(u_lab, sheet_name=xls_lab.sheet_names[0], header=header_row)
        df_lab.columns = [str(c).strip() for c in df_lab.columns]
    except Exception as e:
        st.warning(f"Advertencia al leer laboratorio: {e}")
        
    return df_remitos_raw, df_contactos, df_lab

def formato_miles(valor):
    return f"{valor:,.0f}".replace(',', '.')

def formato_temp(valor):
    if pd.isna(valor): return '-'
    return f"{valor:.1f}".replace('.', ',') + "°"

def limpiar_tambo(val):
    if pd.isna(val): return ""
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    if s.isdigit(): s = 'T' + s
    return s

def extraer_fecha_lab(texto):
    if pd.isna(texto): return None
    s = str(texto).strip()
    match = re.search(r'(\d{2})[-/]?(\d{2})[-/]?(\d{4})', s)
    if match:
        d, m, a = match.groups()
        try:
            return pd.to_datetime(f"{a}-{m}-{d}").normalize()
        except:
            pass
    return None

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
        fecha_str = row['Fecha'].strftime('%d/%m/%Y') if pd.notna(row['Fecha']) else ''
        remito = str(row['N_Remito']) if pd.notna(row['N_Remito']) else '-'
        litros = formato_miles(row['Litros_Ticket']) if pd.notna(row['Litros_Ticket']) else '0'
        
        pdf.cell(35, 7, fecha_str, 1, 0, 'C')
        pdf.cell(45, 7, remito, 1, 0, 'C')
        pdf.cell(35, 7, litros, 1, 0, 'C')
        
        if mostrar_temp:
            temp_val = formato_temp(row['Temperatura'])
            pdf.cell(30, 7, temp_val, 1, 0, 'C')
            
        if mostrar_grasa or mostrar_prot:
            g_val = f"{row['Grasa']:.2f}%".replace('.', ',') if (mostrar_grasa and 'Grasa' in df_productor.columns and pd.notna(row['Grasa'])) else ('-' if mostrar_grasa else '')
            p_val = f"{row['Proteina']:.2f}%".replace('.', ',') if (mostrar_prot and 'Proteina' in df_productor.columns and pd.notna(row['Proteina'])) else ('-' if mostrar_prot else '')
            
            if mostrar_grasa and mostrar_prot:
                solidos_str = f"{g_val} / {p_val}"
            else:
                solidos_str = g_val if mostrar_grasa else p_val
                
            pdf.cell(45, 7, solidos_str, 1, 1, 'C')
        else:
            pdf.ln(7)
        
    return bytes(pdf.output(dest='S'), encoding='latin-1')

def enviar_correo_productor(destinatario_email, nombre_contacto, tambo_nombre, pdf_bytes, nombre_archivo):
    try:
        remitente = st.secrets["email"]["remitente"]
        password = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario_email
        msg['Subject'] = f"Resumen Semanal de Recolección - {tambo_nombre}"
        
        cuerpo_html = f"<html><body><p>Buenas tardes, <b>{nombre_contacto}</b>:</p><p>Te adjunto el resumen de la semana.</p></body></html>"
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
        return False

# --- CARGA Y PROCESAMIENTO DE DATOS ---
try:
    df_raw, df_contactos_raw, df_lab_raw = cargar_datos_drive(url_remitos, url_lab)
    
    if df_raw.empty:
        st.error("El archivo de remitos se cargó vacío o no se pudo acceder a través de Google Drive.")
        st.stop()

    df_contactos = df_contactos_raw.copy()
    df_contactos.columns = [str(c).strip() for c in df_contactos.columns]
    cols_c = list(df_contactos.columns)
    
    col_codigo = next((c for c in cols_c if 'código' in c.lower() or 'codigo' in c.lower() and 'viejo' not in c.lower()), cols_c[1] if len(cols_c) > 1 else cols_c[0])
    col_contacto = next((c for c in cols_c if 'contacto' in c.lower() or 'nombre' in c.lower()), cols_c[3] if len(cols_c) > 3 else cols_c[min(2, len(cols_c)-1)])
    col_email = next((c for c in cols_c if 'email' in c.lower() or 'correo' in c.lower()), cols_c[4] if len(cols_c) > 4 else cols_c[min(len(cols_c)-1, len(cols_c)-1)])
    
    df_contactos['Num_Tambo'] = df_contactos[col_codigo].apply(limpiar_tambo)
    df_contactos['Contacto_Nombre'] = df_contactos[col_contacto]
    df_contactos['Email'] = df_contactos[col_email]
    
    # Asignación estricta de las 10 columnas correspondientes a B:K
    df = df_raw.iloc[:, :10].copy()
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura', 'Grasa', 'Proteina']
    
    df['Num_Tambo'] = df['Num_Tambo'].apply(limpiar_tambo)
    df = df.dropna(subset=['Fecha'])
    
    # Conversión estricta y segura de fechas
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = df['Fecha'].dt.normalize()

    # CLAVES DE CRUCE (TEXTO PLANO)
    df['merge_tambo'] = df['Num_Tambo'].astype(str).str.strip().str.upper()
    df['merge_fecha'] = df['Fecha'].dt.strftime('%Y-%m-%d')

    if not df_lab_raw.empty:
        try:
            df_lab = df_lab_raw.copy()
            df_lab.columns = [str(c).strip() for c in df_lab.columns]
            
            col_sample = next((c for c in df_lab.columns if 'sample' in c.lower() or 'number' in c.lower() or 'tambo' in c.lower()), df_lab.columns[0])
            
            lista_tambo, lista_fecha = [], []
            for val in df_lab[col_sample]:
                if pd.isna(val):
                    lista_tambo.append(None)
                    lista_fecha.append(None)
                    continue
                texto = str(val).strip()
                t_limpio = limpiar_tambo(texto.split()[0]) if len(texto.split()) > 0 else None
                f_limpia = extraer_fecha_lab(texto)
                lista_tambo.append(t_limpio)
                lista_fecha.append(f_limpia)
            
            df_lab['Num_Tambo'] = lista_tambo
            s_fechas_lab = pd.to_datetime(pd.Series(lista_fecha), errors='coerce')
            df_lab['Fecha'] = s_fechas_lab.dt.normalize()
            
            col_fat = next((c for c in df_lab.columns if 'fat' in c.lower() or 'grasa' in c.lower()), None)
            col_prot = next((c for c in df_lab.columns if 'protein' in c.lower() or 'proteina' in c.lower()), None)
            
            if col_fat: df_lab['Grasa_Lab'] = pd.to_numeric(df_lab[col_fat].astype(str).str.replace(',', '.'), errors='coerce')
            if col_prot: df_lab['Proteina_Lab'] = pd.to_numeric(df_lab[col_prot].astype(str).str.replace(',', '.'), errors='coerce')
                
            cols_agg = {}
            if col_fat: cols_agg['Grasa_Lab'] = 'mean'
            if col_prot: cols_agg['Proteina_Lab'] = 'mean'
            
            if cols_agg:
                df_lab_clean = df_lab.dropna(subset=['Fecha', 'Num_Tambo']).groupby(['Num_Tambo', 'Fecha'], as_index=False).agg(cols_agg)
                
                df_lab_clean['merge_tambo'] = df_lab_clean['Num_Tambo'].astype(str).str.strip().str.upper()
                df_lab_clean['merge_fecha'] = df_lab_clean['Fecha'].dt.strftime('%Y-%m-%d')
                
                df = pd.merge(df, df_lab_clean, on=['merge_tambo', 'merge_fecha'], how='left')
                
                if 'Grasa_Lab' in df.columns: df['Grasa'] = df['Grasa_Lab'].combine_first(df['Grasa'])
                if 'Proteina_Lab' in df.columns: df['Proteina'] = df['Proteina_Lab'].combine_first(df['Proteina'])
                
        except Exception as err_lab:
            st.sidebar.error(f"Error procesando lab interno: {err_lab}")
    
    df['Fecha_Cierre_Viernes'] = df['Fecha'] + pd.to_timedelta((4 - df['Fecha'].dt.weekday) % 7, unit='D')
    df['Fecha_Inicio_Sabado'] = df['Fecha_Cierre_Viernes'] - pd.Timedelta(days=6)
    df['Ciclo_Semana'] = df.apply(lambda r: f"Viernes {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')} (Sab {r['Fecha_Inicio_Sabado'].strftime('%d/%m/%Y')} al Vie {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')})", axis=1)
    
    st.sidebar.header("Filtros de Reporte")
    ciclos_disponibles = df[['Fecha_Cierre_Viernes', 'Ciclo_Semana']].drop_duplicates().sort_values('Fecha_Cierre_Viernes', ascending=False)['Ciclo_Semana'].tolist()
    if not ciclos_disponibles:
        st.warning("No hay ciclos de semana disponibles en los datos cargados.")
        st.stop()
        
    ciclo_seleccionado = st.sidebar.selectbox("1. Selecciona el Cierre de Semana:", ciclos_disponibles)
    df_semana_actual = df[df['Ciclo_Semana'] == ciclo_seleccionado]
    
    mapeo_tambos = df_semana_actual[['Tambo', 'Num_Tambo']].dropna().drop_duplicates().sort_values(by='Tambo', ascending=True)
    if mapeo_tambos.empty:
        st.warning("No hay tambos en esta semana.")
        st.stop()
        
    tambo_nombre_seleccionado = st.sidebar.selectbox("2. Selecciona el Tambo:", mapeo_tambos['Tambo'].tolist())
    tambo_seleccionado = mapeo_tambos[mapeo_tambos['Tambo'] == tambo_nombre_seleccionado]['Num_Tambo'].values[0]

    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Elementos del Reporte")
    ver_temperatura = st.sidebar.checkbox("Incluir Temperatura", value=True)
    ver_grasa = st.sidebar.checkbox("Incluir Grasa", value=True)
    ver_proteina = st.sidebar.checkbox("Incluir Proteína", value=True)
    ver_comparacion = st.sidebar.checkbox("Incluir Comparativa vs. Semana Ant.", value=True)

    st.divider()
    df_tambo_semana = df_semana_actual[df_semana_actual['Num_Tambo'] == str(tambo_seleccionado)].sort_values('Fecha')
    
    if not df_tambo_semana.empty:
        f_inicio = df_tambo_semana['Fecha_Inicio_Sabado'].iloc[0].strftime('%d/%m/%Y')
        f_fin = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0].strftime('%d/%m/%Y')
        st.subheader(f"Resumen Cierre Viernes ({f_inicio} al {f_fin}) - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
        
        info_contacto = df_contactos[df_contactos['Num_Tambo'] == str(tambo_seleccionado)]
        email_tambo = info_contacto['Email'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Email'].values[0]) else ""
        nombre_contacto = info_contacto['Contacto_Nombre'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Contacto_Nombre'].values[0]) else "Productor"

        f_viernes_ant = df_tambo_semana['Fecha_Cierre_Viernes'].iloc[0] - pd.Timedelta(days=7)
        df_tambo_anterior = df[(df['Num_Tambo'] == str(tambo_seleccionado)) & (df['Fecha_Cierre_Viernes'] == f_viernes_ant)]
        
        litros_actual = df_tambo_semana['Litros_Ticket'].sum()
        temp_actual = df_tambo_semana['Temperatura'].mean()
        grasa_actual = df_tambo_semana['Grasa'].mean() if 'Grasa' in df_tambo_semana.columns else float('nan')
        prot_actual = df_tambo_semana['Proteina'].mean() if 'Proteina' in df_tambo_semana.columns else float('nan')
        
        hay_datos_previos = not df_tambo_anterior.empty
        c1 = st.columns(sum([1, ver_temperatura, ver_grasa, ver_proteina]))
        col_idx = 0
        
        comp_litros_str, comp_temp_str = "", ""
        if hay_datos_previos:
            diff_pct = ((litros_actual - df_tambo_anterior['Litros_Ticket'].sum()) / df_tambo_anterior['Litros_Ticket'].sum()) * 100 if df_tambo_anterior['Litros_Ticket'].sum() > 0 else 0
            comp_litros_str = f"{diff_pct:+.1f}% vs. semana ant.".replace('.', ',')
            c1[col_idx].metric("Litros", f"{formato_miles(litros_actual)} L", delta=comp_litros_str if ver_comparacion else None)
            col_idx += 1
            if ver_temperatura:
                diff_t = temp_actual - df_tambo_anterior['Temperatura'].mean()
                comp_temp_str = f"{diff_t:+.1f}° vs. semana ant.".replace('.', ',')
                c1[col_idx].metric("Temp. Promedio", formato_temp(temp_actual), delta=comp_temp_str if ver_comparacion else None, delta_color="inverse")
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
            
        st.markdown("---")
        st.markdown("**Detalle de retiros:**")
        cols_vis = ['Fecha', 'N_Remito', 'Litros_Ticket'] + (['Temperatura'] if ver_temperatura else []) + (['Grasa'] if ver_grasa else []) + (['Proteina'] if ver_proteina else [])
        df_show = df_tambo_semana[cols_vis].copy()
        df_show['Litros_Ticket'] = df_show['Litros_Ticket'].apply(lambda x: formato_miles(x) if pd.notna(x) else '0')
        if ver_temperatura: df_show['Temperatura'] = df_show['Temperatura'].apply(formato_temp)
        if 'Grasa' in df_show.columns: df_show['Grasa'] = df_show['Grasa'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
        if 'Proteina' in df_show.columns: df_show['Proteina'] = df_show['Proteina'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
        df_show['Fecha'] = df_show['Fecha'].dt.strftime('%d/%m/%Y')
        st.dataframe(df_show.rename(columns={'Litros_Ticket': 'Litros', 'N_Remito': 'N° de remito', 'Temperatura': 'Temp'}), hide_index=True, use_container_width=True)
    else:
        st.warning("No hay registros para este tambo.")
        
except Exception as e:
    st.error(f"Error procesando los archivos: {e}")
    with st.expander("Ver detalles técnicos del error"):
        st.code(traceback.format_exc())
