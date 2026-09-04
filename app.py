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

st.set_page_config(page_title="Resumen de Recolección - Coopagro", layout="wide")
st.title("Panel de recolección y liquidación por Tambo")

# --- CONFIGURACIÓN DE GOOGLE DRIVE PARA LOS 3 ARCHIVOS ---
FILE_ID_REMITOS = "16Uh0EwP8tyW79TfJlvcjE8li5Lc6RSLj" 
FILE_ID_LAB = "1NNYjM5Aqg9iDdJ85UoALRim8P2A1kaUD"      
FILE_ID_BACSOMATIC = "1KeTle24zxjK-clKAuXsAOUzGkfBNXgI8"      

url_remitos = f"https://drive.google.com/uc?export=download&id={FILE_ID_REMITOS}"
url_lab = f"https://drive.google.com/uc?export=download&id={FILE_ID_LAB}"
url_bacsomatic = f"https://drive.google.com/uc?export=download&id={FILE_ID_BACSOMATIC}"

# Diccionario para meses en español
MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

@st.cache_data(ttl=60)
def cargar_datos_drive(u_remitos, u_lab, u_bacsomatic):
    df_remitos_raw = pd.DataFrame()
    df_contactos = pd.DataFrame()
    df_lab = pd.DataFrame()
    df_bacsomatic = pd.DataFrame()
    
    try:
        df_remitos_raw = pd.read_excel(u_remitos, sheet_name='Résumen OD-PRO-03', skiprows=4, usecols="B:K")
    except Exception as e:
        pass
        
    try:
        df_contactos = pd.read_excel(u_remitos, sheet_name='Código Tambos')
    except Exception as e:
        pass
        
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
        st.sidebar.warning(f"No se pudo cargar el archivo Milko: {e}")
        
    try:
        xls_bac = pd.ExcelFile(u_bacsomatic)
        df_bac_temp = pd.read_excel(u_bacsomatic, sheet_name=xls_bac.sheet_names[0], header=None)
        header_row_bac = 0
        for idx, row in df_bac_temp.iterrows():
            row_str = " ".join([str(x).lower() for x in row.dropna() if pd.notna(x)])
            if 'id usuario' in row_str or 'ufc' in row_str or 'scc' in row_str:
                header_row_bac = idx
                break
        df_bacsomatic = pd.read_excel(u_bacsomatic, sheet_name=xls_bac.sheet_names[0], header=header_row_bac)
        df_bacsomatic.columns = [str(c).strip() for c in df_bacsomatic.columns]
    except Exception as e:
        st.sidebar.warning(f"No se pudo cargar el archivo Bacsomatic: {e}")
        
    return df_remitos_raw, df_contactos, df_lab, df_bacsomatic

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

def extraer_fecha_texto(texto):
    if pd.isna(texto): return None
    s = str(texto).strip()
    
    match_compacto = re.search(r'(\d{2})(\d{2})(\d{4})', s)
    if match_compacto:
        d, m, a = match_compacto.groups()
        try:
            return pd.to_datetime(f"{a}-{m}-{d}").normalize()
        except:
            pass
            
    match = re.search(r'(\d{2})[-/]?(\d{2})[-/]?(\d{4})', s)
    if match:
        d, m, a = match.groups()
        try:
            return pd.to_datetime(f"{a}-{m}-{d}").normalize()
        except:
            pass
    return None

def generar_pdf_bytes(df_productor, tambo_nombre, tambo_id, periodo_texto, comp_litros, comp_temp, mostrar_temp, mostrar_grasa, mostrar_prot, mostrar_crios, mostrar_ufc, mostrar_scc, mostrar_comp, hay_datos_previos, es_mensual=False):
    pdf = FPDF()
    pdf.add_page()
    
    ruta_logo = "logo.png"
    if os.path.exists(ruta_logo):
        pdf.image(ruta_logo, x=65, y=10, w=80)
        pdf.set_y(52) 
    else:
        pdf.set_y(15)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.set_text_color(100, 100, 100)
    titulo_reporte = 'Resumen mensual de recoleccion' if es_mensual else 'Resumen semanal de recoleccion'
    pdf.cell(0, 6, titulo_reporte, ln=True, align='C')
    pdf.set_text_color(0, 0, 0) 
    pdf.ln(4)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y()) 
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(0, 7, f'Productor: {tambo_nombre} (Codigo #{tambo_id})', ln=True)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f'Periodo: {periodo_texto}', ln=True)
    
    total_litros = df_productor['Litros_Ticket'].sum()
    temp_prom = df_productor['Temperatura'].mean()
    grasa_prom = df_productor['Grasa'].mean() if 'Grasa' in df_productor.columns else float('nan')
    proteina_prom = df_productor['Proteina'].mean() if 'Proteina' in df_productor.columns else float('nan')
    crios_prom = df_productor['Crioscopia'].mean() if 'Crioscopia' in df_productor.columns else float('nan')
    ufc_prom = df_productor['UFC'].mean() if 'UFC' in df_productor.columns else float('nan')
    scc_prom = df_productor['SCC'].mean() if 'SCC' in df_productor.columns else float('nan')
    
    pdf.ln(3)
    pdf.set_font('Arial', 'B', 10)
    
    texto_litros = f'Total Litros: {formato_miles(total_litros)} L'
    if not es_mensual and mostrar_comp and hay_datos_previos:
        texto_litros += f' ({comp_litros})'
    pdf.cell(0, 6, texto_litros, ln=True)
    
    if mostrar_temp:
        texto_temp = f'Temperatura Promedio: {formato_temp(temp_prom)}'
        if not es_mensual and mostrar_comp and hay_datos_previos:
            texto_temp += f' ({comp_temp})'
        pdf.cell(0, 6, texto_temp, ln=True)
    
    partes_solidos = []
    if mostrar_grasa and pd.notna(grasa_prom): partes_solidos.append(f"Grasa: {grasa_prom:.2f}%".replace('.', ','))
    if mostrar_prot and pd.notna(proteina_prom): partes_solidos.append(f"Proteina: {proteina_prom:.2f}%".replace('.', ','))
    if mostrar_crios and pd.notna(crios_prom): partes_solidos.append(f"Crioscopia: {crios_prom:.3f}".replace('.', ','))
    if mostrar_ufc and pd.notna(ufc_prom): partes_solidos.append(f"UFC: {formato_miles(ufc_prom)}")
    if mostrar_scc and pd.notna(scc_prom): partes_solidos.append(f"Células Somáticas: {formato_miles(scc_prom)}")
        
    if partes_solidos:
        pdf.cell(0, 6, f"Promedios Lab -> {' | '.join(partes_solidos)}", ln=True)
    
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 9)
    pdf.set_fill_color(200, 220, 255)
    
    cols_header = [('Fecha', 26), ('N° de remito', 38), ('Litros', 26)]
    if mostrar_temp: cols_header.append(('Temp', 18))
    if mostrar_grasa: cols_header.append(('Grasa', 20))
    if mostrar_prot: cols_header.append(('Proteína', 20))
    if mostrar_crios: cols_header.append(('Crioscopia', 25))
    if mostrar_ufc: cols_header.append(('UFC', 18))
    if mostrar_scc: cols_header.append(('Células Somáticas', 28))
        
    for i, (col_name, col_w) in enumerate(cols_header):
        is_last = (i == len(cols_header) - 1)
        pdf.cell(col_w, 8, col_name, 1, 1 if is_last else 0, 'C', fill=True)
    
    pdf.set_font('Arial', '', 9)
    for _, row in df_productor.iterrows():
        fecha_str = row['Fecha'].strftime('%d/%m/%Y') if pd.notna(row['Fecha']) else ''
        remito = str(row['N_Remito']) if pd.notna(row['N_Remito']) else '-'
        litros = formato_miles(row['Litros_Ticket']) if pd.notna(row['Litros_Ticket']) else '0'
        
        row_cells = [(fecha_str, 26), (remito, 38), (litros, 26)]
        
        if mostrar_temp: row_cells.append((formato_temp(row['Temperatura']), 18))
        if mostrar_grasa: row_cells.append((f"{row['Grasa']:.2f}%".replace('.', ',') if ('Grasa' in df_productor.columns and pd.notna(row['Grasa'])) else '-', 20))
        if mostrar_prot: row_cells.append((f"{row['Proteina']:.2f}%".replace('.', ',') if ('Proteina' in df_productor.columns and pd.notna(row['Proteina'])) else '-', 20))
        if mostrar_crios: row_cells.append((f"{row['Crioscopia']:.3f}".replace('.', ',') if ('Crioscopia' in df_productor.columns and pd.notna(row['Crioscopia'])) else '-', 25))
        if mostrar_ufc: row_cells.append((formato_miles(row['UFC']) if ('UFC' in df_productor.columns and pd.notna(row['UFC'])) else '-', 18))
        if mostrar_scc: row_cells.append((formato_miles(row['SCC']) if ('SCC' in df_productor.columns and pd.notna(row['SCC'])) else '-', 28))
            
        for i, (val, col_w) in enumerate(row_cells):
            is_last = (i == len(row_cells) - 1)
            pdf.cell(col_w, 7, val, 1, 1 if is_last else 0, 'C')
        
    return bytes(pdf.output(dest='S'), encoding='latin-1')

def enviar_correo_productor(destinatario_email, nombre_contacto, tambo_nombre, pdf_bytes, nombre_archivo, tipo_reporte="semanal"):
    try:
        remitente = st.secrets["email"]["remitente"]
        password = st.secrets["email"]["password"]
        msg = MIMEMultipart()
        msg['From'] = remitente
        msg['To'] = destinatario_email
        msg['Subject'] = f"Resumen {tipo_reporte.capitalize()} de Recolección - {tambo_nombre}"
        
        cuerpo_html = f"<html><body><p>Buenas tardes, <b>{nombre_contacto}</b>:</p><p>Te adjunto el resumen {tipo_reporte} de recolección y calidad de leche.</p></body></html>"
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
    df_raw, df_contactos_raw, df_lab_raw, df_bac_raw = cargar_datos_drive(url_remitos, url_lab, url_bacsomatic)
    
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
    
    df = df_raw.iloc[:, :10].copy()
    df.columns = ['Fecha', 'N_Remito', 'Num_Tambo', 'Tambo', 'Litros_Ticket', 'Litros_Planilla', 'Diferencia', 'Temperatura', 'Grasa', 'Proteina']
    
    df['Num_Tambo'] = df['Num_Tambo'].apply(limpiar_tambo)
    df = df.dropna(subset=['Fecha'])
    
    df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
    df = df.dropna(subset=['Fecha'])
    df['Fecha'] = df['Fecha'].dt.normalize()

    # Inicializar columnas de laboratorio vacías por defecto para evitar KeyErrors
    for col_lab_init in ['Grasa', 'Proteina', 'Crioscopia', 'UFC', 'SCC']:
        if col_lab_init not in df.columns:
            df[col_lab_init] = float('nan')

    # Claves de cruce base en el DataFrame principal
    df['merge_tambo'] = df['Num_Tambo'].astype(str).str.strip().str.upper()
    df['merge_fecha'] = df['Fecha'].dt.strftime('%Y-%m-%d')

    # Asignación secuencial por orden de remito (Dobles retiros diarios)
    df = df.sort_values(by=['Num_Tambo', 'Fecha', 'N_Remito'])
    df['orden_remito'] = df.groupby(['Num_Tambo', 'Fecha']).cumcount() + 1

    # 1. PROCESAR MILFOSCAN CON SECUENCIA
    if not df_lab_raw.empty:
        try:
            df_lab = df_lab_raw.copy()
            df_lab.columns = [str(c).strip() for c in df_lab.columns]
            col_sample = next((c for c in df_lab.columns if 'sample' in c.lower() or 'number' in c.lower() or 'tambo' in c.lower() or 'muestra' in c.lower()), df_lab.columns[0])
            col_date_lab = next((c for c in df_lab.columns if 'fecha' in c.lower() or 'date' in c.lower() or 'time' in c.lower()), None)
            
            lista_tambo, lista_fecha = [], []
            for _, r_lab in df_lab.iterrows():
                val_sample = r_lab[col_sample]
                if pd.isna(val_sample):
                    lista_tambo.append(None); lista_fecha.append(None); continue
                texto = str(val_sample).strip()
                partes = texto.split()
                t_limpio = limpiar_tambo(partes[0]) if len(partes) > 0 else None
                f_limpia = None
                if col_date_lab and pd.notna(r_lab[col_date_lab]):
                    f_limpia = pd.to_datetime(r_lab[col_date_lab], errors='coerce')
                    if pd.notna(f_limpia): f_limpia = f_limpia.normalize()
                if f_limpia is None:
                    for p in partes:
                        f_int = extraer_fecha_texto(p)
                        if f_int is not None: f_limpia = f_int; break
                    if f_limpia is None: f_limpia = extraer_fecha_texto(texto)
                lista_tambo.append(t_limpio); lista_fecha.append(f_limpia)
            
            df_lab['Num_Tambo'] = lista_tambo
            df_lab['Fecha'] = pd.to_datetime(pd.Series(lista_fecha), errors='coerce').dt.normalize()
            
            df_lab = df_lab.sort_values(by=['Num_Tambo', 'Fecha'])
            df_lab['orden_remito'] = df_lab.groupby(['Num_Tambo', 'Fecha']).cumcount() + 1
            
            col_fat = next((c for c in df_lab.columns if 'fat' in c.lower() or 'grasa' in c.lower()), None)
            col_prot = next((c for c in df_lab.columns if 'protein' in c.lower() or 'proteina' in c.lower()), None)
            col_fp = next((c for c in df_lab.columns if c.lower() == 'fp' or 'crios' in c.lower() or 'congelacion' in c.lower()), None)
            
            if col_fat: df_lab['Grasa_Lab'] = pd.to_numeric(df_lab[col_fat].astype(str).str.replace(',', '.'), errors='coerce')
            if col_prot: df_lab['Proteina_Lab'] = pd.to_numeric(df_lab[col_prot].astype(str).str.replace(',', '.'), errors='coerce')
            if col_fp: df_lab['Crioscopia_Lab'] = pd.to_numeric(df_lab[col_fp].astype(str).str.replace(',', '.'), errors='coerce')
                
            cols_milko = []
            if col_fat: cols_milko.append('Grasa_Lab')
            if col_prot: cols_milko.append('Proteina_Lab')
            if col_fp: cols_milko.append('Crioscopia_Lab')
            
            if cols_milko:
                df_milko_clean = df_lab.dropna(subset=['Fecha', 'Num_Tambo'])[['Num_Tambo', 'Fecha', 'orden_remito'] + cols_milko].copy()
                df_milko_clean['merge_tambo'] = df_milko_clean['Num_Tambo'].astype(str).str.strip().str.upper()
                df_milko_clean['merge_fecha'] = df_milko_clean['Fecha'].dt.strftime('%Y-%m-%d')
                
                df = pd.merge(df, df_milko_clean[['merge_tambo', 'merge_fecha', 'orden_remito'] + cols_milko], on=['merge_tambo', 'merge_fecha', 'orden_remito'], how='left')
                if 'Grasa_Lab' in df.columns: df['Grasa'] = df['Grasa_Lab'].combine_first(df.get('Grasa', pd.Series(dtype=float)))
                if 'Proteina_Lab' in df.columns: df['Proteina'] = df['Proteina_Lab'].combine_first(df.get('Proteina', pd.Series(dtype=float)))
                if 'Crioscopia_Lab' in df.columns: df['Crioscopia'] = df['Crioscopia_Lab'].combine_first(df.get('Crioscopia', pd.Series(dtype=float)))
        except Exception as e:
            st.sidebar.error(f"Error procesando Milko: {e}")

    # 2. PROCESAR BACSOMATIC CON SECUENCIA
    if not df_bac_raw.empty:
        try:
            df_bac = df_bac_raw.copy()
            df_bac.columns = [str(c).strip() for c in df_bac.columns]
            
            col_id_user = next((c for c in df_bac.columns if 'id usuario' in c.lower() or 'sample' in c.lower() or 'tambo' in c.lower() or 'muestra' in c.lower()), df_bac.columns[0])
            
            lista_tambo_bac, lista_fecha_bac = [], []
            for _, r_bac in df_bac.iterrows():
                val_id = r_bac[col_id_user]
                if pd.isna(val_id):
                    lista_tambo_bac.append(None); lista_fecha_bac.append(None); continue
                
                texto = str(val_id).strip()
                partes = texto.split()
                t_limpio = limpiar_tambo(partes[0]) if len(partes) > 0 else None
                
                f_limpia = extraer_fecha_texto(texto)
                if f_limpia is None:
                    for p in partes:
                        f_int = extraer_fecha_texto(p)
                        if f_int is not None: f_limpia = f_int; break
                
                lista_tambo_bac.append(t_limpio)
                lista_fecha_bac.append(f_limpia)
            
            df_bac['Num_Tambo'] = lista_tambo_bac
            df_bac['Fecha'] = pd.to_datetime(pd.Series(lista_fecha_bac), errors='coerce').dt.normalize()
            
            df_bac = df_bac.sort_values(by=['Num_Tambo', 'Fecha'])
            df_bac['orden_remito'] = df_bac.groupby(['Num_Tambo', 'Fecha']).cumcount() + 1
            
            col_ufc = next((c for c in df_bac.columns if 'ufc' in c.lower()), None)
            col_scc = next((c for c in df_bac.columns if 'scc' in c.lower() or 'celulas' in c.lower() or 'somáticas' in c.lower()), None)
            
            if col_ufc: df_bac['UFC_Val'] = pd.to_numeric(df_bac[col_ufc].astype(str).str.replace(',', '.'), errors='coerce')
            if col_scc: df_bac['SCC_Val'] = pd.to_numeric(df_bac[col_scc].astype(str).str.replace(',', '.'), errors='coerce')
            
            cols_bac = []
            if col_ufc: cols_bac.append('UFC_Val')
            if col_scc: cols_bac.append('SCC_Val')
            
            if cols_bac:
                df_bac_clean = df_bac.dropna(subset=['Fecha', 'Num_Tambo'])[['Num_Tambo', 'Fecha', 'orden_remito'] + cols_bac].copy()
                df_bac_clean['merge_tambo'] = df_bac_clean['Num_Tambo'].astype(str).str.strip().str.upper()
                df_bac_clean['merge_fecha'] = df_bac_clean['Fecha'].dt.strftime('%Y-%m-%d')
                
                df = pd.merge(df, df_bac_clean[['merge_tambo', 'merge_fecha', 'orden_remito'] + cols_bac], on=['merge_tambo', 'merge_fecha', 'orden_remito'], how='left')
                if 'UFC_Val' in df.columns: df['UFC'] = df['UFC_Val'].combine_first(df.get('UFC', pd.Series(dtype=float)))
                if 'SCC_Val' in df.columns: df['SCC'] = df['SCC_Val'].combine_first(df.get('SCC', pd.Series(dtype=float)))
        except Exception as e:
            st.sidebar.error(f"Error procesando Bacsomatic: {e}")
    
    df['Fecha_Cierre_Viernes'] = df['Fecha'] + pd.to_timedelta((4 - df['Fecha'].dt.weekday) % 7, unit='D')
    df['Fecha_Inicio_Sabado'] = df['Fecha_Cierre_Viernes'] - pd.Timedelta(days=6)
    df['Ciclo_Semana'] = df.apply(lambda r: f"Viernes {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')} (Sáb {r['Fecha_Inicio_Sabado'].strftime('%d/%m/%Y')} al Vie {r['Fecha_Cierre_Viernes'].strftime('%d/%m/%Y')})", axis=1)
    
    # --- MENÚ LATERAL: TIPO DE REPORTE (SEMANAL VS MENSUAL) ---
    st.sidebar.header("Tipo de Reporte")
    tipo_reporte_opcion = st.sidebar.radio("Seleccione la modalidad:", ["Semanal", "Mensual"])

    st.sidebar.markdown("---")
    st.sidebar.header("Filtros de Reporte")
    
    if tipo_reporte_opcion == "Semanal":
        ciclos_disponibles = df[['Fecha_Cierre_Viernes', 'Ciclo_Semana']].drop_duplicates().sort_values('Fecha_Cierre_Viernes', ascending=False)['Ciclo_Semana'].tolist()
        if not ciclos_disponibles:
            st.warning("No hay ciclos de semana disponibles en los datos cargados.")
            st.stop()
            
        ciclo_seleccionado = st.sidebar.selectbox("1. Seleccione el Cierre de Semana:", ciclos_disponibles)
        df_filtrado_periodo = df[df['Ciclo_Semana'] == ciclo_seleccionado]
        
        mapeo_tambos = df_filtrado_periodo[['Tambo', 'Num_Tambo']].dropna().drop_duplicates().sort_values(by='Tambo', ascending=True)
        if mapeo_tambos.empty:
            st.warning("No hay tambos en este periodo.")
            st.stop()
            
        tambo_nombre_seleccionado = st.sidebar.selectbox("2. Seleccione el Tambo:", mapeo_tambos['Tambo'].tolist())
        tambo_seleccionado = mapeo_tambos[mapeo_tambos['Tambo'] == tambo_nombre_seleccionado]['Num_Tambo'].values[0]

        st.sidebar.markdown("---")
        st.sidebar.subheader("⚙️ Elementos del Reporte")
        ver_temperatura = st.sidebar.checkbox("Incluir Temperatura", value=True)
        ver_grasa = st.sidebar.checkbox("Incluir Grasa", value=True)
        ver_proteina = st.sidebar.checkbox("Incluir Proteína", value=True)
        ver_crioscopia = st.sidebar.checkbox("Incluir Crioscopia", value=True)
        ver_ufc = st.sidebar.checkbox("Incluir UFC", value=True)
        ver_scc = st.sidebar.checkbox("Incluir Células Somáticas (SCC)", value=True)
        ver_comparacion = st.sidebar.checkbox("Incluir Comparativa vs. Semana Ant.", value=True)

        st.divider()
        df_tambo_periodo = df_filtrado_periodo[df_filtrado_periodo['Num_Tambo'] == str(tambo_seleccionado)].sort_values(['Fecha', 'N_Remito'])
        
        if not df_tambo_periodo.empty:
            f_inicio = df_tambo_periodo['Fecha_Inicio_Sabado'].iloc[0].strftime('%d/%m/%Y')
            f_fin = df_tambo_periodo['Fecha_Cierre_Viernes'].iloc[0].strftime('%d/%m/%Y')
            periodo_texto_pdf = f"{f_inicio} al {f_fin}"
            st.subheader(f"Resumen Cierre Viernes ({periodo_texto_pdf}) - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
            
            info_contacto = df_contactos[df_contactos['Num_Tambo'] == str(tambo_seleccionado)]
            email_tambo = info_contacto['Email'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Email'].values[0]) else ""
            nombre_contacto = info_contacto['Contacto_Nombre'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Contacto_Nombre'].values[0]) else "Productor"

            f_viernes_ant = df_tambo_periodo['Fecha_Cierre_Viernes'].iloc[0] - pd.Timedelta(days=7)
            df_tambo_anterior = df[(df['Num_Tambo'] == str(tambo_seleccionado)) & (df['Fecha_Cierre_Viernes'] == f_viernes_ant)]
            
            litros_actual = df_tambo_periodo['Litros_Ticket'].sum()
            temp_actual = df_tambo_periodo['Temperatura'].mean()
            grasa_actual = df_tambo_periodo['Grasa'].mean() if 'Grasa' in df_tambo_periodo.columns else float('nan')
            prot_actual = df_tambo_periodo['Proteina'].mean() if 'Proteina' in df_tambo_periodo.columns else float('nan')
            crios_actual = df_tambo_periodo['Crioscopia'].mean() if 'Crioscopia' in df_tambo_periodo.columns else float('nan')
            ufc_actual = df_tambo_periodo['UFC'].mean() if 'UFC' in df_tambo_periodo.columns else float('nan')
            scc_actual = df_tambo_periodo['SCC'].mean() if 'SCC' in df_tambo_periodo.columns else float('nan')
            
            hay_datos_previos = not df_tambo_anterior.empty
            
            num_metrics = 1 + int(ver_temperatura) + int(ver_grasa) + int(ver_proteina) + int(ver_crioscopia) + int(ver_ufc) + int(ver_scc)
            c1 = st.columns(num_metrics)
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
                col_idx += 1
            if ver_crioscopia:
                c1[col_idx].metric("Crioscopia Promedio", f"{crios_actual:.3f}".replace('.', ',') if pd.notna(crios_actual) else "S/D")
                col_idx += 1
            if ver_ufc:
                c1[col_idx].metric("UFC Promedio", f"{formato_miles(ufc_actual)}" if pd.notna(ufc_actual) else "S/D")
                col_idx += 1
            if ver_scc:
                c1[col_idx].metric("SCC Promedio", f"{formato_miles(scc_actual)}" if pd.notna(scc_actual) else "S/D")
                
            st.markdown("---")
            st.markdown("**Detalle de retiros:**")
            
            cols_vis = ['Fecha', 'N_Remito', 'Litros_Ticket'] \
                       + (['Temperatura'] if ver_temperatura else []) \
                       + (['Grasa'] if ver_grasa else []) \
                       + (['Proteina'] if ver_proteina else []) \
                       + (['Crioscopia'] if ver_crioscopia else []) \
                       + (['UFC'] if ver_ufc else []) \
                       + (['SCC'] if ver_scc else [])
                       
            df_show = df_tambo_periodo[cols_vis].copy()
            df_show['Litros_Ticket'] = df_show['Litros_Ticket'].apply(lambda x: formato_miles(x) if pd.notna(x) else '0')
            if ver_temperatura: df_show['Temperatura'] = df_show['Temperatura'].apply(formato_temp)
            if 'Grasa' in df_show.columns: df_show['Grasa'] = df_show['Grasa'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
            if 'Proteina' in df_show.columns: df_show['Proteina'] = df_show['Proteina'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
            if 'Crioscopia' in df_show.columns: df_show['Crioscopia'] = df_show['Crioscopia'].apply(lambda x: f"{x:.3f}".replace('.', ',') if pd.notna(x) else '-')
            if 'UFC' in df_show.columns: df_show['UFC'] = df_show['UFC'].apply(lambda x: formato_miles(x) if pd.notna(x) else '-')
            if 'SCC' in df_show.columns: df_show['SCC'] = df_show['SCC'].apply(lambda x: formato_miles(x) if pd.notna(x) else '-')
            df_show['Fecha'] = df_show['Fecha'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(df_show.rename(columns={'Litros_Ticket': 'Litros', 'N_Remito': 'N° de remito', 'Temperatura': 'Temp', 'Crioscopia': 'Crioscopia', 'SCC': 'Células Somáticas'}), hide_index=True, use_container_width=True)
            
            st.markdown("---")
            col_btn1, col_btn2 = st.columns(2)
            nombre_archivo_pdf = f"Resumen_Semanal_{tambo_nombre_seleccionado.replace(' ', '_')}_{f_inicio.replace('/', '-')}.pdf"
            
            pdf_bytes = generar_pdf_bytes(
                df_tambo_periodo, 
                tambo_nombre_seleccionado, 
                tambo_seleccionado, 
                periodo_texto_pdf, 
                comp_litros_str, 
                comp_temp_str, 
                ver_temperatura, 
                ver_grasa, 
                ver_proteina, 
                ver_crioscopia,
                ver_ufc,
                ver_scc,
                ver_comparacion, 
                hay_datos_previos,
                es_mensual=False
            )
            
            with col_btn1:
                st.download_button(label="📥 Descargar Resumen en PDF", data=pdf_bytes, file_name=nombre_archivo_pdf, mime="application/pdf", use_container_width=True)
            with col_btn2:
                if email_tambo:
                    if st.button(f"📧 Enviar por Mail a {nombre_contacto}", use_container_width=True):
                        exito = enviar_correo_productor(email_tambo, nombre_contacto, tambo_nombre_seleccionado, pdf_bytes, nombre_archivo_pdf, tipo_reporte="semanal")
                        if exito: st.success(f"¡Correo enviado con éxito a {email_tambo}!")
                        else: st.error("Hubo un error al enviar el correo.")
                else:
                    st.warning("⚠️ Este tambo no tiene un email cargado.")

    else:
        # --- MODO REPORTE MENSUAL (EJ. AGOSTO) ---
        df['AnioMes'] = df['Fecha'].dt.to_period('M')
        meses_disponibles = sorted(df['AnioMes'].unique(), reverse=True)
        if not meses_disponibles:
            st.warning("No hay meses disponibles en los datos cargados.")
            st.stop()
            
        def formatear_mes_es(periodo):
            mes_num = periodo.month
            anio = periodo.year
            return f"{MESES_ES.get(mes_num, periodo.strftime('%B'))} {anio}"

        mes_seleccionado = st.sidebar.selectbox("1. Seleccione el Mes:", meses_disponibles, format_func=formatear_mes_es)
        df_mes_actual = df[df['AnioMes'] == mes_seleccionado]
        
        mapeo_tambos = df_mes_actual[['Tambo', 'Num_Tambo']].dropna().drop_duplicates().sort_values(by='Tambo', ascending=True)
        if mapeo_tambos.empty:
            st.warning("No hay tambos en este mes.")
            st.stop()
            
        tambo_nombre_seleccionado = st.sidebar.selectbox("2. Seleccione el Tambo:", mapeo_tambos['Tambo'].tolist())
        tambo_seleccionado = mapeo_tambos[mapeo_tambos['Tambo'] == tambo_nombre_seleccionado]['Num_Tambo'].values[0]

        st.sidebar.markdown("---")
        st.sidebar.subheader("⚙️ Elementos del Reporte Mensual")
        ver_temperatura = st.sidebar.checkbox("Incluir Temperatura", value=True)
        ver_grasa = st.sidebar.checkbox("Incluir Grasa", value=True)
        ver_proteina = st.sidebar.checkbox("Incluir Proteína", value=True)
        ver_crioscopia = st.sidebar.checkbox("Incluir Crioscopia", value=True)
        ver_ufc = st.sidebar.checkbox("Incluir UFC", value=True)
        ver_scc = st.sidebar.checkbox("Incluir Células Somáticas (SCC)", value=True)

        st.divider()
        df_tambo_mes = df_mes_actual[df_mes_actual['Num_Tambo'] == str(tambo_seleccionado)].sort_values(['Fecha', 'N_Remito'])
        
        if not df_tambo_mes.empty:
            nombre_mes_str = formatear_mes_es(mes_seleccionado)
            st.subheader(f"Resumen Mensual ({nombre_mes_str}) - {tambo_nombre_seleccionado} (Código #{tambo_seleccionado})")
            
            info_contacto = df_contactos[df_contactos['Num_Tambo'] == str(tambo_seleccionado)]
            email_tambo = info_contacto['Email'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Email'].values[0]) else ""
            nombre_contacto = info_contacto['Contacto_Nombre'].values[0] if not info_contacto.empty and pd.notna(info_contacto['Contacto_Nombre'].values[0]) else "Productor"

            litros_actual = df_tambo_mes['Litros_Ticket'].sum()
            temp_actual = df_tambo_mes['Temperatura'].mean()
            grasa_actual = df_tambo_mes['Grasa'].mean() if 'Grasa' in df_tambo_mes.columns else float('nan')
            prot_actual = df_tambo_mes['Proteina'].mean() if 'Proteina' in df_tambo_mes.columns else float('nan')
            crios_actual = df_tambo_mes['Crioscopia'].mean() if 'Crioscopia' in df_tambo_mes.columns else float('nan')
            ufc_actual = df_tambo_mes['UFC'].mean() if 'UFC' in df_tambo_mes.columns else float('nan')
            scc_actual = df_tambo_mes['SCC'].mean() if 'SCC' in df_tambo_mes.columns else float('nan')
            
            num_metrics = 1 + int(ver_temperatura) + int(ver_grasa) + int(ver_proteina) + int(ver_crioscopia) + int(ver_ufc) + int(ver_scc)
            c1 = st.columns(num_metrics)
            col_idx = 0
            
            c1[col_idx].metric("Litros Totales", f"{formato_miles(litros_actual)} L")
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
            if ver_crioscopia:
                c1[col_idx].metric("Crioscopia Promedio", f"{crios_actual:.3f}".replace('.', ',') if pd.notna(crios_actual) else "S/D")
                col_idx += 1
            if ver_ufc:
                c1[col_idx].metric("UFC Promedio", f"{formato_miles(ufc_actual)}" if pd.notna(ufc_actual) else "S/D")
                col_idx += 1
            if ver_scc:
                c1[col_idx].metric("SCC Promedio", f"{formato_miles(scc_actual)}" if pd.notna(scc_actual) else "S/D")
                
            st.markdown("---")
            st.markdown("**Detalle diario del mes:**")
            
            cols_vis = ['Fecha', 'N_Remito', 'Litros_Ticket'] \
                       + (['Temperatura'] if ver_temperatura else []) \
                       + (['Grasa'] if ver_grasa else []) \
                       + (['Proteina'] if ver_proteina else []) \
                       + (['Crioscopia'] if ver_crioscopia else []) \
                       + (['UFC'] if ver_ufc else []) \
                       + (['SCC'] if ver_scc else [])
                       
            df_show = df_tambo_mes[cols_vis].copy()
            df_show['Litros_Ticket'] = df_show['Litros_Ticket'].apply(lambda x: formato_miles(x) if pd.notna(x) else '0')
            if ver_temperatura: df_show['Temperatura'] = df_show['Temperatura'].apply(formato_temp)
            if 'Grasa' in df_show.columns: df_show['Grasa'] = df_show['Grasa'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
            if 'Proteina' in df_show.columns: df_show['Proteina'] = df_show['Proteina'].apply(lambda x: f"{x:.2f}%".replace('.', ',') if pd.notna(x) else '-')
            if 'Crioscopia' in df_show.columns: df_show['Crioscopia'] = df_show['Crioscopia'].apply(lambda x: f"{x:.3f}".replace('.', ',') if pd.notna(x) else '-')
            if 'UFC' in df_show.columns: df_show['UFC'] = df_show['UFC'].apply(lambda x: formato_miles(x) if pd.notna(x) else '-')
            if 'SCC' in df_show.columns: df_show['SCC'] = df_show['SCC'].apply(lambda x: formato_miles(x) if pd.notna(x) else '-')
            df_show['Fecha'] = df_show['Fecha'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(df_show.rename(columns={'Litros_Ticket': 'Litros', 'N_Remito': 'N° de remito', 'Temperatura': 'Temp', 'Crioscopia': 'Crioscopia', 'SCC': 'Células Somáticas'}), hide_index=True, use_container_width=True)
            
            st.markdown("---")
            col_btn1, col_btn2 = st.columns(2)
            nombre_archivo_pdf = f"Resumen_Mensual_{tambo_nombre_seleccionado.replace(' ', '_')}_{mes_seleccionado.strftime('%Y_%m')}.pdf"
            
            pdf_bytes = generar_pdf_bytes(
                df_tambo_mes, 
                tambo_nombre_seleccionado, 
                tambo_seleccionado, 
                nombre_mes_str, 
                "", 
                "", 
                ver_temperatura, 
                ver_grasa, 
                ver_proteina, 
                ver_crioscopia,
                ver_ufc,
                ver_scc,
                False, 
                False,
                es_mensual=True
            )
            
            with col_btn1:
                st.download_button(label="📥 Descargar Resumen Mensual en PDF", data=pdf_bytes, file_name=nombre_archivo_pdf, mime="application/pdf", use_container_width=True)
            with col_btn2:
                if email_tambo:
                    if st.button(f"📧 Enviar Reporte Mensual por Mail a {nombre_contacto}", use_container_width=True):
                        exito = enviar_correo_productor(email_tambo, nombre_contacto, tambo_nombre_seleccionado, pdf_bytes, nombre_archivo_pdf, tipo_reporte="mensual")
                        if exito: st.success(f"¡Correo mensual enviado con éxito a {email_tambo}!")
                        else: st.error("Hubo un error al enviar el correo.")
                else:
                    st.warning("⚠️ Este tambo no tiene un email cargado.")
        else:
            st.warning("No hay registros para este tambo en el mes seleccionado.")

except Exception as e:
    st.error(f"Error procesando los archivos: {e}")
    with st.expander("Ver detalles técnicos del error"):
        st.code(traceback.format_exc())
