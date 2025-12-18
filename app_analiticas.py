# app_analiticas.py - VERSIÓN MODIFICADA CON DOS PROCESOS SEPARADOS
import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import io

st.set_page_config(
    page_title="Sistema de Analíticas - Laboratorio",
    page_icon="🔬",
    layout="wide"
)

# ======================= FUNCIONES =======================

def cargar_datos_analisis_recibidos():
    """Carga los datos de análisis recibidos (con resultados)"""
    if not os.path.exists("datos"):
        os.makedirs("datos")
    
    if not os.path.exists("datos/analisis_recibidos.csv"):
        df = pd.DataFrame(columns=[
            'ID', 'N_BOLETIN', 'Fecha', 'DENOMINACION_MUESTRA',
            'AEROBIOS_SOLICITADO', 'SALMONELLA_SOLICITADO', 
            'ENTEROBACTERIAS_SOLICITADO', 'MOHOS_LEVADURAS_SOLICITADO',
            'ESTAFILOCOCOS_SOLICITADO', 'OTRAS_DETERMINACIONES_SOLICITADO',
            'OBSERVACIONES_GENERALES',
            'AEROBIOS_RESULTADO', 'SALMONELLA_RESULTADO',
            'ENTEROBACTERIAS_RESULTADO', 'MOHOS_LEVADURAS_RESULTADO',
            'ESTAFILOCOCOS_RESULTADO', 'OBSERVACIONES_RESULTADOS',
            'FECHA_ENVIO', 'DESTINO_LABORATORIO', 'DIRECCION_LABORATORIO'
        ])
        df.to_csv("datos/analisis_recibidos.csv", index=False, encoding='utf-8')
    
    return pd.read_csv("datos/analisis_recibidos.csv", encoding='utf-8', 
                       dtype={'N_BOLETIN': str})

def cargar_datos_muestras_enviadas():
    """Carga los datos de muestras enviadas (sin resultados)"""
    if not os.path.exists("datos"):
        os.makedirs("datos")
    
    if not os.path.exists("datos/muestras_enviadas.csv"):
        df = pd.DataFrame(columns=[
            'ID', 'Fecha', 'DENOMINACION_MUESTRA',
            'AEROBIOS_SOLICITADO', 'SALMONELLA_SOLICITADO', 
            'ENTEROBACTERIAS_SOLICITADO', 'MOHOS_LEVADURAS_SOLICITADO',
            'ESTAFILOCOCOS_SOLICITADO', 'OTRAS_DETERMINACIONES_SOLICITADO',
            'OBSERVACIONES_GENERALES',
            'FECHA_ENVIO', 'DESTINO_LABORATORIO', 'DIRECCION_LABORATORIO'
        ])
        df.to_csv("datos/muestras_enviadas.csv", index=False, encoding='utf-8')
    
    return pd.read_csv("datos/muestras_enviadas.csv", encoding='utf-8')

def guardar_analisis_recibido(nuevo_registro):
    """Guarda un nuevo análisis recibido"""
    df = cargar_datos_analisis_recibidos()
    nuevo_id = 1 if len(df) == 0 else df['ID'].max() + 1
    nuevo_registro['ID'] = nuevo_id
    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    df.to_csv("datos/analisis_recibidos.csv", index=False, encoding='utf-8')
    return nuevo_id

def guardar_muestra_enviada(nuevo_registro):
    """Guarda una nueva muestra enviada"""
    df = cargar_datos_muestras_enviadas()
    nuevo_id = 1 if len(df) == 0 else df['ID'].max() + 1
    nuevo_registro['ID'] = nuevo_id
    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
    df.to_csv("datos/muestras_enviadas.csv", index=False, encoding='utf-8')
    return nuevo_id

def cargar_configuracion():
    with open("datos/config.json", "r", encoding='utf-8') as f:
        return json.load(f)

def guardar_configuracion(config):
    with open("datos/config.json", "w", encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

def actualizar_registro_recibido(id_registro, registro_actualizado):
    df = cargar_datos_analisis_recibidos()
    idx = df[df['ID'] == id_registro].index[0]
    for col in registro_actualizado:
        df.loc[idx, col] = registro_actualizado[col]
    df.to_csv("datos/analisis_recibidos.csv", index=False, encoding='utf-8')

def actualizar_registro_enviado(id_registro, registro_actualizado):
    df = cargar_datos_muestras_enviadas()
    idx = df[df['ID'] == id_registro].index[0]
    for col in registro_actualizado:
        df.loc[idx, col] = registro_actualizado[col]
    df.to_csv("datos/muestras_enviadas.csv", index=False, encoding='utf-8')

def eliminar_registro_recibido(id_registro):
    df = cargar_datos_analisis_recibidos()
    df = df[df['ID'] != id_registro]
    df.to_csv("datos/analisis_recibidos.csv", index=False, encoding='utf-8')

def eliminar_registro_enviado(id_registro):
    df = cargar_datos_muestras_enviadas()
    df = df[df['ID'] != id_registro]
    df.to_csv("datos/muestras_enviadas.csv", index=False, encoding='utf-8')

def generar_pdf_albaran(muestras_seleccionadas, fecha_envio, direcciones):
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    
    normal_style = ParagraphStyle(
        'CustomNormal', parent=styles['Normal'],
        fontSize=9, leading=11, spaceAfter=6
    )
    
    header_style = ParagraphStyle(
        'CustomHeader', parent=styles['Heading3'],
        fontSize=10, leading=12, spaceAfter=8,
        textColor=colors.HexColor('#333333'),
        fontName='Helvetica-Bold'
    )
    
    doc = SimpleDocTemplate(buffer, pagesize=letter,
        topMargin=40, bottomMargin=40, leftMargin=30, rightMargin=30)
    elements = []
    
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Title'],
        fontSize=14, textColor=colors.HexColor('#000000'),
        alignment=TA_CENTER, spaceAfter=20
    )
    title = Paragraph("<b>ALBARÁN DE ENVÍO DE MUESTRAS</b>", title_style)
    elements.append(title)
    
    fecha_style = ParagraphStyle(
        'FechaStyle', parent=normal_style,
        fontSize=9, alignment=TA_LEFT, spaceAfter=25
    )
    fecha_texto = Paragraph(f"<b>Fecha de envío:</b> {fecha_envio.strftime('%d/%m/%Y')}", fecha_style)
    elements.append(fecha_texto)
    
    elements.append(Paragraph("<b>ORIGEN - DATOS DE ENVÍO</b>", header_style))
    for linea in direcciones['origen'].split('\n'):
        if linea.strip():
            elements.append(Paragraph(linea.strip(), normal_style))
    
    elements.append(Spacer(1, 15))
    elements.append(Paragraph("<b>DESTINO - LABORATORIO</b>", header_style))
    for linea in direcciones['destino'].split('\n'):
        if linea.strip():
            elements.append(Paragraph(linea.strip(), normal_style))
    
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("<b>MUESTRAS ENVIADAS PARA ANÁLISIS</b>", header_style))
    elements.append(Spacer(1, 8))
    
    tabla_datos = [['Nº', 'PRODUCTO / MUESTRA', 'DESCRIPCIÓN', 'ANÁLISIS SOLICITADOS']]
    
    for i, muestra in enumerate(muestras_seleccionadas, 1):
        analisis_solicitados = []
        if muestra.get('aerobios_solicitado', False): analisis_solicitados.append("Aerobios")
        if muestra.get('salmonella_solicitado', False): analisis_solicitados.append("Salmonella")
        if muestra.get('enterobacterias_solicitado', False): analisis_solicitados.append("Enterobacterias")
        if muestra.get('mohos_levaduras_solicitado', False): analisis_solicitados.append("Mohos y Levaduras")
        if muestra.get('estafilococos_solicitado', False): analisis_solicitados.append("Estafilococos Aureus")
        
        otras_det = muestra.get('otras_determinaciones_solicitado', '')
        analisis_texto = ', '.join(analisis_solicitados)
        if otras_det:
            analisis_texto = (analisis_texto + f" | {otras_det}") if analisis_texto else otras_det
        
        nombre_muestra = str(muestra['denominacion_muestra'])[:40]
        descripcion = str(muestra.get('descripcion', ''))[:60]
        
        fila = [
            str(i),
            Paragraph(nombre_muestra, normal_style),
            Paragraph(descripcion if descripcion else '-', normal_style),
            Paragraph(analisis_texto[:80], normal_style)
        ]
        tabla_datos.append(fila)
    
    tabla = Table(tabla_datos, colWidths=[20, 140, 160, 220])
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4B6584')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (1, 1), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9FA')]),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
    ]))
    
    elements.append(tabla)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("<b>OBSERVACIONES:</b>", normal_style))
    for i, muestra in enumerate(muestras_seleccionadas, 1):
        obs_muestra = muestra.get('observaciones_generales', '')
        if obs_muestra:
            elements.append(Paragraph(f"• Muestra {i}: {obs_muestra[:100]}", normal_style))
    elements.append(Paragraph("• Conservar las muestras refrigeradas hasta su análisis.", normal_style))
    elements.append(Paragraph("• Los resultados serán enviados por correo electrónico.", normal_style))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ======================= INTERFAZ =======================

def main():
    # Cargar configuraciones
    if not os.path.exists("datos/config.json"):
        config = {
            "direccion_origen": "Avd. de la Madera Nº 5\n45520 Villaluenga de la Sagra - TOLEDO\n925355925\nMIGUEL MARTÍN\nmiguel@catalysis.es",
            "direccion_destino": "LABORATORIO AINPROT S.A.\nC/ Valentín Beato, Nº 11, 3ºC\n28037 - MADRID",
            "productos_frecuentes": ["DIAMEL PETS C-4", "OCOXIN PETS C-5", "VIUSID PETS C-7"]
        }
        os.makedirs("datos", exist_ok=True)
        with open("datos/config.json", "w", encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    
    config = cargar_configuracion()
    
    if 'muestras_albaran' not in st.session_state:
        st.session_state.muestras_albaran = []
    
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/1998/1998678.png", width=100)
        st.title("🔬 Sistema de Analíticas")
        
        # Contar registros
        df_recibidos = cargar_datos_analisis_recibidos()
        df_enviados = cargar_datos_muestras_enviadas()
        
        opcion = st.radio("Menú principal:",
            ["📋 Registrar análisis recibido", "📤 Enviar muestras para análisis",
             "📦 Generar albarán de envío",
             "📊 Historial análisis recibidos", "📋 Historial muestras enviadas",
             "⚙️ Configurar direcciones"])
        
        st.markdown("---")
        st.caption(f"📥 {len(df_recibidos)} análisis recibidos")
        st.caption(f"📤 {len(df_enviados)} muestras enviadas")
    
    # ===== REGISTRAR ANÁLISIS RECIBIDO =====
    if opcion == "📋 Registrar análisis recibido":
        st.title("📋 Registrar Análisis Recibido")
        with st.form("form_analisis_recibido"):
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                st.subheader("Datos del Albarán")
                n_boletin = st.text_input("Nº de Boletín*", placeholder="Ej: LAB-2024-001")
                fecha = st.date_input("Fecha del análisis*", datetime.now())
                denominacion_muestra = st.text_input("DENOMINACIÓN DE LA MUESTRA*")
                
                st.subheader("Análisis Solicitados")
                aerobios_solicitado = st.checkbox("Aerobios", value=True, key="aerobios_sol")
                salmonella_solicitado = st.checkbox("Salmonella", value=True, key="salmonella_sol")
                enterobacterias_solicitado = st.checkbox("Enterobacterias", value=True, key="enterobacterias_sol")
                mohos_levaduras_solicitado = st.checkbox("Mohos y Levaduras", value=True, key="mohos_sol")
                estafilococos_solicitado = st.checkbox("Estafilococos Aureus", value=True, key="estafilococos_sol")
                otras_determinaciones_solicitado = st.text_area("OTRAS DETERMINACIONES", height=60)
                observaciones_generales = st.text_area("Observaciones", height=80)
            
            with col2:
                st.subheader("Información del Envío")
                fecha_envio = st.date_input("Fecha de envío*", datetime.now())
                destino_laboratorio = st.text_input("Destino Laboratorio*", value="LABORATORIO AINPROT S.A.")
                direccion_laboratorio = st.text_area("Dirección Laboratorio*", 
                    value="C/ Valentín Beato, Nº 11, 3ºC\n28037 - MADRID", height=100)
            
            with col3:
                st.subheader("Resultados del Análisis")
                st.markdown("---")
                aerobios_resultado = st.text_input("AEROBIOS Resultado", placeholder="Ej: < 10 ufc/g")
                salmonella_resultado = st.text_input("SALMONELLA Resultado", placeholder="Ej: Ausente en 25g")
                enterobacterias_resultado = st.text_input("ENTEROBACTERIAS Resultado", placeholder="Ej: < 10 ufc/g")
                mohos_levaduras_resultado = st.text_input("MOHOS Y LEVADURAS Resultado", placeholder="Ej: < 10 ufc/g")
                estafilococos_resultado = st.text_input("ESTAFILOCOCOS AUREUS Resultado", placeholder="Ej: < 10 ufc/g")
                observaciones_resultados = st.text_area("Observaciones resultados", height=100)
            
            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                enviar = st.form_submit_button("💾 Guardar análisis recibido", type="primary", use_container_width=True)
            
            if enviar:
                if n_boletin and denominacion_muestra:
                    nuevo_registro = {
                        'N_BOLETIN': n_boletin,
                        'Fecha': fecha.strftime('%Y-%m-%d'),
                        'DENOMINACION_MUESTRA': denominacion_muestra,
                        'AEROBIOS_SOLICITADO': aerobios_solicitado,
                        'SALMONELLA_SOLICITADO': salmonella_solicitado,
                        'ENTEROBACTERIAS_SOLICITADO': enterobacterias_solicitado,
                        'MOHOS_LEVADURAS_SOLICITADO': mohos_levaduras_solicitado,
                        'ESTAFILOCOCOS_SOLICITADO': estafilococos_solicitado,
                        'OTRAS_DETERMINACIONES_SOLICITADO': otras_determinaciones_solicitado,
                        'OBSERVACIONES_GENERALES': observaciones_generales,
                        'AEROBIOS_RESULTADO': aerobios_resultado,
                        'SALMONELLA_RESULTADO': salmonella_resultado,
                        'ENTEROBACTERIAS_RESULTADO': enterobacterias_resultado,
                        'MOHOS_LEVADURAS_RESULTADO': mohos_levaduras_resultado,
                        'ESTAFILOCOCOS_RESULTADO': estafilococos_resultado,
                        'OBSERVACIONES_RESULTADOS': observaciones_resultados,
                        'FECHA_ENVIO': fecha_envio.strftime('%Y-%m-%d'),
                        'DESTINO_LABORATORIO': destino_laboratorio,
                        'DIRECCION_LABORATORIO': direccion_laboratorio
                    }
                    id_guardado = guardar_analisis_recibido(nuevo_registro)
                    st.success(f"✅ Análisis recibido registrado (ID: {id_guardado})")
                else:
                    st.error("⚠️ Completa los campos obligatorios (*)")
    
    # ===== ENVIAR MUESTRAS PARA ANÁLISIS =====
    elif opcion == "📤 Enviar muestras para análisis":
        st.title("📤 Enviar Muestras para Análisis")
        
        with st.form("form_muestra_envio"):
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.subheader("Datos de la Muestra")
                fecha = st.date_input("Fecha*", datetime.now())
                denominacion_muestra = st.text_input("DENOMINACIÓN DE LA MUESTRA*")
                descripcion = st.text_area("Descripción del envío:", height=100,
                    placeholder="Ej: 150 ML PRINCIPIO DE ENVASADO")
                otras_determinaciones = st.text_area("Otras determinaciones:", height=60,
                    placeholder="Ej: Recuento total, pH, etc.")
                observaciones_generales = st.text_area("Observaciones:", height=80,
                    placeholder="Ej: Muestra tomada del lote X")
            
            with col2:
                st.subheader("Información del Envío")
                fecha_envio = st.date_input("Fecha de envío*", datetime.now())
                destino_laboratorio = st.text_input("Destino Laboratorio*", value="LABORATORIO AINPROT S.A.")
                direccion_laboratorio = st.text_area("Dirección Laboratorio*", 
                    value="C/ Valentín Beato, Nº 11, 3ºC\n28037 - MADRID", height=100)
                
                st.subheader("Análisis Solicitados")
                st.markdown("---")
                a_aerobios = st.checkbox("Aerobios", value=True, key="a_aerobios_env")
                a_salmonella = st.checkbox("Salmonella", value=True, key="a_salmonella_env")
                a_enterobacterias = st.checkbox("Enterobacterias", value=True, key="a_enterobacterias_env")
                a_mohos = st.checkbox("Mohos y Levaduras", value=True, key="a_mohos_env")
                a_estafilococos = st.checkbox("Estafilococos Aureus", value=True, key="a_estafilococos_env")
                st.caption("✅ Marca los análisis requeridos")
            
            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                enviar = st.form_submit_button("📤 Guardar muestra para envío", type="primary", use_container_width=True)
            
            if enviar:
                if denominacion_muestra:
                    nuevo_registro = {
                        'Fecha': fecha.strftime('%Y-%m-%d'),
                        'DENOMINACION_MUESTRA': denominacion_muestra,
                        'AEROBIOS_SOLICITADO': a_aerobios,
                        'SALMONELLA_SOLICITADO': a_salmonella,
                        'ENTEROBACTERIAS_SOLICITADO': a_enterobacterias,
                        'MOHOS_LEVADURAS_SOLICITADO': a_mohos,
                        'ESTAFILOCOCOS_SOLICITADO': a_estafilococos,
                        'OTRAS_DETERMINACIONES_SOLICITADO': otras_determinaciones,
                        'OBSERVACIONES_GENERALES': observaciones_generales,
                        'FECHA_ENVIO': fecha_envio.strftime('%Y-%m-%d'),
                        'DESTINO_LABORATORIO': destino_laboratorio,
                        'DIRECCION_LABORATORIO': direccion_laboratorio
                    }
                    id_guardado = guardar_muestra_enviada(nuevo_registro)
                    st.success(f"✅ Muestra para envío registrada (ID: {id_guardado})")
                else:
                    st.error("⚠️ Completa los campos obligatorios (*)")
    
    # ===== GENERAR ALBARÁN =====
    elif opcion == "📦 Generar albarán de envío":
        st.title("📦 Generar Albarán de Envío")
        
        with st.expander("✏️ Editar direcciones", expanded=False):
            col_dir1, col_dir2 = st.columns(2)
            with col_dir1:
                origen_edit = st.text_area("Dirección origen:", value=config['direccion_origen'], height=150)
            with col_dir2:
                destino_edit = st.text_area("Dirección destino:", value=config['direccion_destino'], height=150)
            
            if st.button("💾 Guardar direcciones"):
                config['direccion_origen'] = origen_edit
                config['direccion_destino'] = destino_edit
                guardar_configuracion(config)
                st.success("✅ Direcciones guardadas")
        
        fecha_envio = st.date_input("Fecha prevista de envío", datetime.now())
        st.markdown("---")
        st.subheader("Añadir muestras al albarán")
        
        with st.form("form_muestra_albaran"):
            col_muestra1, col_muestra2 = st.columns([2, 1])
            
            with col_muestra1:
                denominacion_muestra = st.text_input("DENOMINACIÓN DE LA MUESTRA*")
                descripcion = st.text_area("Descripción del envío:", height=100,
                    placeholder="Ej: 150 ML PRINCIPIO DE ENVASADO")
                otras_determinaciones = st.text_area("Otras determinaciones:", height=60,
                    placeholder="Ej: Recuento total, pH, etc.")
                observaciones = st.text_area("Observaciones:", height=80,
                    placeholder="Ej: Muestra tomada del lote X")
            
            with col_muestra2:
                st.markdown("**Análisis solicitados:**")
                st.markdown("---")
                a_aerobios = st.checkbox("Aerobios", value=True, key="a_aerobios_albaran")
                a_salmonella = st.checkbox("Salmonella", value=True, key="a_salmonella_albaran")
                a_enterobacterias = st.checkbox("Enterobacterias", value=True, key="a_enterobacterias_albaran")
                a_mohos = st.checkbox("Mohos y Levaduras", value=True, key="a_mohos_albaran")
                a_estafilococos = st.checkbox("Estafilococos Aureus", value=True, key="a_estafilococos_albaran")
                st.markdown("---")
                st.caption("✅ Marca los análisis requeridos")
            
            anadir_presionado = st.form_submit_button("➕ Añadir muestra",
                use_container_width=True, type="primary")
        
        if anadir_presionado:
            if denominacion_muestra:
                nueva_muestra = {
                    'denominacion_muestra': denominacion_muestra,
                    'descripcion': descripcion,
                    'otras_determinaciones_solicitado': otras_determinaciones,
                    'observaciones_generales': observaciones,
                    'aerobios_solicitado': a_aerobios,
                    'salmonella_solicitado': a_salmonella,
                    'enterobacterias_solicitado': a_enterobacterias,
                    'mohos_levaduras_solicitado': a_mohos,
                    'estafilococos_solicitado': a_estafilococos
                }
                st.session_state.muestras_albaran.append(nueva_muestra)
                st.success(f"✅ '{denominacion_muestra}' añadida")
            else:
                st.error("⚠️ El nombre es obligatorio")
        
        st.markdown("---")
        st.subheader(f"📋 Muestras ({len(st.session_state.muestras_albaran)})")
        
        if st.session_state.muestras_albaran:
            for i, muestra in enumerate(st.session_state.muestras_albaran, 1):
                with st.expander(f"{i}. {muestra['denominacion_muestra']}"):
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.write(f"**Descripción:** {muestra.get('descripcion', 'N/A')}")
                        if muestra.get('otras_determinaciones_solicitado'):
                            st.write(f"**Otras:** {muestra['otras_determinaciones_solicitado']}")
                        if muestra.get('observaciones_generales'):
                            st.write(f"**Obs:** {muestra['observaciones_generales']}")
                    
                    with col2:
                        analisis = []
                        if muestra['aerobios_solicitado']: analisis.append("✓ Aerobios")
                        if muestra['salmonella_solicitado']: analisis.append("✓ Salmonella")
                        if muestra['enterobacterias_solicitado']: analisis.append("✓ Enterobacterias")
                        if muestra['mohos_levaduras_solicitado']: analisis.append("✓ Mohos/Lev")
                        if muestra['estafilococos_solicitado']: analisis.append("✓ Estafilococos")
                        
                        st.write("**Análisis:**")
                        for a in analisis:
                            st.write(a)
                    
                    if st.button(f"🗑️ Eliminar", key=f"eliminar_{i}"):
                        st.session_state.muestras_albaran.pop(i-1)
                        st.rerun()
            
            st.markdown("---")
            col_gen1, col_gen2 = st.columns([1, 1])
            
            with col_gen1:
                if st.button("🖨️ Generar PDF", type="primary", use_container_width=True):
                    direcciones = {
                        'origen': config['direccion_origen'],
                        'destino': config['direccion_destino']
                    }
                    pdf_buffer = generar_pdf_albaran(st.session_state.muestras_albaran, fecha_envio, direcciones)
                    st.download_button(
                        label="📥 Descargar albarán",
                        data=pdf_buffer,
                        file_name=f"albaran_{fecha_envio.strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
            with col_gen2:
                if st.button("🗑️ Limpiar albarán", use_container_width=True):
                    st.session_state.muestras_albaran = []
                    st.rerun()
        else:
            st.info("ℹ️ Añade muestras al albarán")
    
    # ===== HISTORIAL ANÁLISIS RECIBIDOS =====
    elif opcion == "📊 Historial análisis recibidos":
        st.title("📊 Historial de Análisis Recibidos")
        df_recibidos = cargar_datos_analisis_recibidos()
        st.subheader(f"📥 {len(df_recibidos)} análisis recibidos")
        
        if len(df_recibidos) > 0:
            tab_tabla, tab_editar = st.tabs(["📊 Ver tabla", "✏️ Editar/Eliminar"])
            
            with tab_tabla:
                columnas_mostrar = ['ID', 'N_BOLETIN', 'Fecha', 'DENOMINACION_MUESTRA', 
                                   'AEROBIOS_RESULTADO', 'SALMONELLA_RESULTADO',
                                   'ENTEROBACTERIAS_RESULTADO', 'MOHOS_LEVADURAS_RESULTADO',
                                   'ESTAFILOCOCOS_RESULTADO', 'OTRAS_DETERMINACIONES_SOLICITADO',
                                   'OBSERVACIONES_GENERALES']
                df_mostrar = df_recibidos[columnas_mostrar].copy()
                df_mostrar.columns = ['ID', 'Nº Boletín', 'Fecha', 'Muestra', 'Aerobios', 'Salmonella',
                                     'Enterobacterias', 'Mohos/Levaduras', 'Estafilococos',
                                     'Otras Determ.', 'Observaciones']
                st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                
                # Crear botones de descarga en columnas
                col_csv, col_xlsx = st.columns(2)
                
                with col_csv:
                    csv = df_recibidos.to_csv(index=False, encoding='utf-8')
                    st.download_button(
                        "📥 Descargar CSV", 
                        csv, 
                        f"analisis_recibidos_{datetime.now().strftime('%Y%m%d')}.csv", 
                        "text/csv",
                        use_container_width=True
                    )
                
                with col_xlsx:
                    excel_buffer = io.BytesIO()
                    # Escribir a Excel usando to_excel directamente
                    df_recibidos.to_excel(excel_buffer, index=False, sheet_name='Analisis_Recibidos', engine='openpyxl')
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        "📊 Descargar Excel", 
                        excel_buffer, 
                        f"analisis_recibidos_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="Descargar en formato Excel XLSX"
                    )
            
            with tab_editar:
                st.subheader("Editar o Eliminar Registros")
                
                id_seleccionado = st.selectbox(
                    "Selecciona un análisis para editar/eliminar:",
                    df_recibidos['ID'].values,
                    format_func=lambda x: f"ID {x} - {df_recibidos[df_recibidos['ID']==x]['DENOMINACION_MUESTRA'].values[0]}"
                )
                
                registro = df_recibidos[df_recibidos['ID'] == id_seleccionado].iloc[0]
                
                st.divider()
                
                col_acciones1, col_acciones2 = st.columns(2)
                
                with col_acciones1:
                    if st.button("✏️ Editar Registro", type="primary", use_container_width=True, key="btn_editar_rec"):
                        st.session_state.editando_recibido = id_seleccionado
                
                with col_acciones2:
                    if st.button("🗑️ Eliminar Registro", type="secondary", use_container_width=True, key="btn_eliminar_rec"):
                        eliminar_registro_recibido(id_seleccionado)
                        st.success(f"✅ Registro ID {id_seleccionado} eliminado")
                        st.rerun()
                
                st.divider()
                
                if 'editando_recibido' in st.session_state and st.session_state.editando_recibido == id_seleccionado:
                    st.subheader("📝 Editar Registro")
                    
                    with st.form("form_editar_recibido"):
                        col1, col2, col3 = st.columns([1, 1, 1])
                        
                        with col1:
                            st.subheader("Datos del Análisis")
                            n_boletin_edit = st.text_input("Nº de Boletín*", 
                                value=str(registro['N_BOLETIN']), key="nboletin_ed")
                            fecha_edit = st.date_input("Fecha del análisis*", 
                                value=pd.to_datetime(registro['Fecha']).date(), key="fecha_ed_rec")
                            denominacion_edit = st.text_input("DENOMINACIÓN*", 
                                value=registro['DENOMINACION_MUESTRA'], key="denom_ed_rec")
                            
                            st.subheader("Análisis Solicitados")
                            aerobios_sol_edit = st.checkbox("Aerobios", 
                                value=bool(registro['AEROBIOS_SOLICITADO']), key="aero_ed_rec")
                            salmonella_sol_edit = st.checkbox("Salmonella", 
                                value=bool(registro['SALMONELLA_SOLICITADO']), key="salm_ed_rec")
                            enterobacterias_sol_edit = st.checkbox("Enterobacterias", 
                                value=bool(registro['ENTEROBACTERIAS_SOLICITADO']), key="ent_ed_rec")
                            mohos_sol_edit = st.checkbox("Mohos y Levaduras", 
                                value=bool(registro['MOHOS_LEVADURAS_SOLICITADO']), key="moh_ed_rec")
                            estafilococos_sol_edit = st.checkbox("Estafilococos Aureus", 
                                value=bool(registro['ESTAFILOCOCOS_SOLICITADO']), key="est_ed_rec")
                            otras_det_edit = st.text_area("OTRAS DETERMINACIONES", 
                                value=str(registro['OTRAS_DETERMINACIONES_SOLICITADO']), height=60, key="otras_ed_rec")
                            obs_gen_edit = st.text_area("Observaciones", 
                                value=str(registro['OBSERVACIONES_GENERALES']), height=80, key="obs_ed_rec")
                        
                        with col2:
                            st.subheader("Información del Envío")
                            fecha_envio_edit = st.date_input("Fecha de envío*", 
                                value=pd.to_datetime(registro['FECHA_ENVIO']).date(), key="fecha_env_ed_rec")
                            destino_edit = st.text_input("Destino Laboratorio*", 
                                value=registro['DESTINO_LABORATORIO'], key="dest_ed_rec")
                            direccion_edit = st.text_area("Dirección Laboratorio*", 
                                value=registro['DIRECCION_LABORATORIO'], height=100, key="dir_ed_rec")
                        
                        with col3:
                            st.subheader("Resultados del Análisis")
                            st.markdown("---")
                            aero_res_edit = st.text_input("AEROBIOS Resultado", 
                                value=str(registro['AEROBIOS_RESULTADO']), key="aero_res_ed_rec")
                            salm_res_edit = st.text_input("SALMONELLA Resultado", 
                                value=str(registro['SALMONELLA_RESULTADO']), key="salm_res_ed_rec")
                            ent_res_edit = st.text_input("ENTEROBACTERIAS Resultado", 
                                value=str(registro['ENTEROBACTERIAS_RESULTADO']), key="ent_res_ed_rec")
                            mohos_res_edit = st.text_input("MOHOS Y LEVADURAS Resultado", 
                                value=str(registro['MOHOS_LEVADURAS_RESULTADO']), key="mohos_res_ed_rec")
                            est_res_edit = st.text_input("ESTAFILOCOCOS AUREUS Resultado", 
                                value=str(registro['ESTAFILOCOCOS_RESULTADO']), key="est_res_ed_rec")
                            obs_res_edit = st.text_area("Observaciones", 
                                value=str(registro['OBSERVACIONES_RESULTADOS']), height=100, key="obs_res_ed_rec")
                        
                        col_btn_edit1, col_btn_edit2 = st.columns([1, 1])
                        with col_btn_edit1:
                            guardar_edit = st.form_submit_button("💾 Guardar cambios", type="primary", use_container_width=True)
                        with col_btn_edit2:
                            cancelar_edit = st.form_submit_button("❌ Cancelar", use_container_width=True)
                        
                        if guardar_edit:
                            if n_boletin_edit:
                                registro_actualizado = {
                                    'N_BOLETIN': n_boletin_edit,
                                    'Fecha': fecha_edit.strftime('%Y-%m-%d'),
                                    'DENOMINACION_MUESTRA': denominacion_edit,
                                    'AEROBIOS_SOLICITADO': aerobios_sol_edit,
                                    'SALMONELLA_SOLICITADO': salmonella_sol_edit,
                                    'ENTEROBACTERIAS_SOLICITADO': enterobacterias_sol_edit,
                                    'MOHOS_LEVADURAS_SOLICITADO': mohos_sol_edit,
                                    'ESTAFILOCOCOS_SOLICITADO': estafilococos_sol_edit,
                                    'OTRAS_DETERMINACIONES_SOLICITADO': otras_det_edit,
                                    'OBSERVACIONES_GENERALES': obs_gen_edit,
                                    'AEROBIOS_RESULTADO': aero_res_edit,
                                    'SALMONELLA_RESULTADO': salm_res_edit,
                                    'ENTEROBACTERIAS_RESULTADO': ent_res_edit,
                                    'MOHOS_LEVADURAS_RESULTADO': mohos_res_edit,
                                    'ESTAFILOCOCOS_RESULTADO': est_res_edit,
                                    'OBSERVACIONES_RESULTADOS': obs_res_edit,
                                    'FECHA_ENVIO': fecha_envio_edit.strftime('%Y-%m-%d'),
                                    'DESTINO_LABORATORIO': destino_edit,
                                    'DIRECCION_LABORATORIO': direccion_edit
                                }
                                actualizar_registro_recibido(id_seleccionado, registro_actualizado)
                                st.session_state.editando_recibido = None
                                st.success(f"✅ Registro ID {id_seleccionado} actualizado")
                                st.rerun()
                            else:
                                st.error("⚠️ El Nº de Boletín es obligatorio")
                        
                        if cancelar_edit:
                            st.session_state.editando_recibido = None
                            st.rerun()
        else:
            st.info("ℹ️ No hay análisis recibidos registrados")
    
    # ===== HISTORIAL MUESTRAS ENVIADAS =====
    elif opcion == "📋 Historial muestras enviadas":
        st.title("📋 Historial de Muestras Enviadas")
        df_enviados = cargar_datos_muestras_enviadas()
        st.subheader(f"📤 {len(df_enviados)} muestras enviadas")
        
        if len(df_enviados) > 0:
            tab_tabla, tab_editar = st.tabs(["📊 Ver tabla", "✏️ Editar/Eliminar"])
            
            with tab_tabla:
                columnas_mostrar = ['ID', 'Fecha', 'DENOMINACION_MUESTRA', 
                                   'AEROBIOS_SOLICITADO', 'SALMONELLA_SOLICITADO',
                                   'ENTEROBACTERIAS_SOLICITADO', 'MOHOS_LEVADURAS_SOLICITADO',
                                   'ESTAFILOCOCOS_SOLICITADO', 'OTRAS_DETERMINACIONES_SOLICITADO',
                                   'OBSERVACIONES_GENERALES', 'FECHA_ENVIO']
                df_mostrar = df_enviados[columnas_mostrar].copy()
                df_mostrar.columns = ['ID', 'Fecha', 'Muestra', 'Aerobios', 'Salmonella',
                                     'Enterobacterias', 'Mohos/Levaduras', 'Estafilococos',
                                     'Otras Determ.', 'Observaciones', 'Fecha Envío']
                st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
                
                # Crear botones de descarga en columnas
                col_csv, col_xlsx = st.columns(2)
                
                with col_csv:
                    csv = df_enviados.to_csv(index=False, encoding='utf-8')
                    st.download_button(
                        "📥 Descargar CSV", 
                        csv, 
                        f"muestras_enviadas_{datetime.now().strftime('%Y%m%d')}.csv", 
                        "text/csv",
                        use_container_width=True
                    )
                
                with col_xlsx:
                    excel_buffer = io.BytesIO()
                    # Método directo sin usar with statement para ExcelWriter
                    df_enviados.to_excel(excel_buffer, index=False, sheet_name='Muestras_Enviadas')
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        "📊 Descargar Excel", 
                        excel_buffer, 
                        f"muestras_enviadas_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="Descargar en formato Excel XLSX"
                    )
            
            with tab_editar:
                st.subheader("Editar o Eliminar Registros")
                
                id_seleccionado = st.selectbox(
                    "Selecciona una muestra para editar/eliminar:",
                    df_enviados['ID'].values,
                    format_func=lambda x: f"ID {x} - {df_enviados[df_enviados['ID']==x]['DENOMINACION_MUESTRA'].values[0]}"
                )
                
                registro = df_enviados[df_enviados['ID'] == id_seleccionado].iloc[0]
                
                st.divider()
                
                col_acciones1, col_acciones2 = st.columns(2)
                
                with col_acciones1:
                    if st.button("✏️ Editar Registro", type="primary", use_container_width=True, key="btn_editar_env"):
                        st.session_state.editando_enviado = id_seleccionado
                
                with col_acciones2:
                    if st.button("🗑️ Eliminar Registro", type="secondary", use_container_width=True, key="btn_eliminar_env"):
                        eliminar_registro_enviado(id_seleccionado)
                        st.success(f"✅ Registro ID {id_seleccionado} eliminado")
                        st.rerun()
                
                st.divider()
                
                if 'editando_enviado' in st.session_state and st.session_state.editando_enviado == id_seleccionado:
                    st.subheader("📝 Editar Registro")
                    
                    with st.form("form_editar_enviado"):
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.subheader("Datos de la Muestra")
                            fecha_edit = st.date_input("Fecha*", 
                                value=pd.to_datetime(registro['Fecha']).date(), key="fecha_ed_env")
                            denominacion_edit = st.text_input("DENOMINACIÓN DE LA MUESTRA*", 
                                value=registro['DENOMINACION_MUESTRA'], key="denom_ed_env")
                            descripcion_edit = st.text_area("Descripción del envío:", 
                                value=str(registro.get('descripcion', '')), height=100, key="desc_ed_env")
                            otras_det_edit = st.text_area("Otras determinaciones:", 
                                value=str(registro['OTRAS_DETERMINACIONES_SOLICITADO']), height=60, key="otras_ed_env")
                            obs_gen_edit = st.text_area("Observaciones:", 
                                value=str(registro['OBSERVACIONES_GENERALES']), height=80, key="obs_ed_env")
                        
                        with col2:
                            st.subheader("Información del Envío")
                            fecha_envio_edit = st.date_input("Fecha de envío*", 
                                value=pd.to_datetime(registro['FECHA_ENVIO']).date(), key="fecha_env_ed_env")
                            destino_edit = st.text_input("Destino Laboratorio*", 
                                value=registro['DESTINO_LABORATORIO'], key="dest_ed_env")
                            direccion_edit = st.text_area("Dirección Laboratorio*", 
                                value=registro['DIRECCION_LABORATORIO'], height=100, key="dir_ed_env")
                            
                            st.subheader("Análisis Solicitados")
                            aerobios_sol_edit = st.checkbox("Aerobios", 
                                value=bool(registro['AEROBIOS_SOLICITADO']), key="aero_ed_env")
                            salmonella_sol_edit = st.checkbox("Salmonella", 
                                value=bool(registro['SALMONELLA_SOLICITADO']), key="salm_ed_env")
                            enterobacterias_sol_edit = st.checkbox("Enterobacterias", 
                                value=bool(registro['ENTEROBACTERIAS_SOLICITADO']), key="ent_ed_env")
                            mohos_sol_edit = st.checkbox("Mohos y Levaduras", 
                                value=bool(registro['MOHOS_LEVADURAS_SOLICITADO']), key="moh_ed_env")
                            estafilococos_sol_edit = st.checkbox("Estafilococos Aureus", 
                                value=bool(registro['ESTAFILOCOCOS_SOLICITADO']), key="est_ed_env")
                        
                        col_btn_edit1, col_btn_edit2 = st.columns([1, 1])
                        with col_btn_edit1:
                            guardar_edit = st.form_submit_button("💾 Guardar cambios", type="primary", use_container_width=True)
                        with col_btn_edit2:
                            cancelar_edit = st.form_submit_button("❌ Cancelar", use_container_width=True)
                        
                        if guardar_edit:
                            if denominacion_edit:
                                registro_actualizado = {
                                    'Fecha': fecha_edit.strftime('%Y-%m-%d'),
                                    'DENOMINACION_MUESTRA': denominacion_edit,
                                    'AEROBIOS_SOLICITADO': aerobios_sol_edit,
                                    'SALMONELLA_SOLICITADO': salmonella_sol_edit,
                                    'ENTEROBACTERIAS_SOLICITADO': enterobacterias_sol_edit,
                                    'MOHOS_LEVADURAS_SOLICITADO': mohos_sol_edit,
                                    'ESTAFILOCOCOS_SOLICITADO': estafilococos_sol_edit,
                                    'OTRAS_DETERMINACIONES_SOLICITADO': otras_det_edit,
                                    'OBSERVACIONES_GENERALES': obs_gen_edit,
                                    'FECHA_ENVIO': fecha_envio_edit.strftime('%Y-%m-%d'),
                                    'DESTINO_LABORATORIO': destino_edit,
                                    'DIRECCION_LABORATORIO': direccion_edit
                                }
                                actualizar_registro_enviado(id_seleccionado, registro_actualizado)
                                st.session_state.editando_enviado = None
                                st.success(f"✅ Registro ID {id_seleccionado} actualizado")
                                st.rerun()
                            else:
                                st.error("⚠️ La denominación de la muestra es obligatoria")
                        
                        if cancelar_edit:
                            st.session_state.editando_enviado = None
                            st.rerun()
        else:
            st.info("ℹ️ No hay muestras enviadas registradas")
    
    # ===== CONFIGURAR =====
    elif opcion == "⚙️ Configurar direcciones":
        st.title("⚙️ Configuración")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Dirección Origen")
            origen = st.text_area("Dirección:", value=config['direccion_origen'], height=200)
        
        with col2:
            st.subheader("Dirección Destino")
            destino = st.text_area("Dirección:", value=config['direccion_destino'], height=200)
        
        if st.button("💾 Guardar configuración", type="primary", use_container_width=True):
            config['direccion_origen'] = origen
            config['direccion_destino'] = destino
            guardar_configuracion(config)
            st.success("✅ Configuración guardada")

if __name__ == "__main__":
    main()