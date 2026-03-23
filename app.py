import streamlit as st
import json
import re
import os
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. UTILIDADES DE PROCESAMIENTO
# ==========================================

def limpiar_nombre_archivo(titulo):
    """
    Convierte el título en un nombre de archivo seguro:
    'Mi Libro: Edición 1' -> 'mi_libro_edicion_1'
    """
    # Eliminar acentos
    a, b = 'áéíóúüñÁÉÍÓÚÜÑ', 'aeiouunAEIOUUN'
    trans = str.maketrans(a, b)
    limpio = titulo.translate(trans)
    # Reemplazar caracteres no alfanuméricos por guiones bajos
    limpio = re.sub(r'[^a-zA-Z0-9]', '_', limpio)
    # Evitar múltiples guiones bajos seguidos y pasar a minúsculas
    return re.sub(r'_+', '_', limpio).strip('_').lower()

# (Las funciones de bajo nivel como set_mirror_margins, add_page_number, 
# add_toc_field y add_formatted_text se mantienen exactamente igual que antes)

# ==========================================
# 2. CONFIGURACIÓN DE ESTILOS Y MOTOR
# ==========================================

def setup_styles(doc):
    styles = doc.styles
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.2
    body.paragraph_format.first_line_indent = Inches(0.25)

    h1 = styles['Heading 1']
    h1.font.name, h1.font.size = 'Aptos', Pt(16)
    h1.font.bold = True
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before, h1.paragraph_format.space_after = Inches(2.0), Inches(1.0)
    return doc

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # Lógica de maquetación (Simplificada para el ejemplo)
    section = doc.sections[0]
    # (Aquí iría tu lógica de apply_layout definida anteriormente)
    
    p_title = doc.add_paragraph(meta['title'])
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.runs[0].font.size = Pt(32)
    p_title.runs[0].font.bold = True
    
    # ... Resto del proceso de generación de párrafos ...
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 3. INTERFAZ DE USUARIO (STREAMLIT)
# ==========================================

st.set_page_config(page_title="Editor Editorial", layout="wide")
st.title("📚 Generador Editorial Profesional")

# Inicialización de estado
if 'meta_data' not in st.session_state:
    st.session_state.meta_data = {
        'title': "Sin Titulo", 'subtitle': "", 'author': "Autor",
        'year': str(datetime.now().year), 'copyright': ""
    }

with st.sidebar:
    st.header("1. Ficha Editorial JSON")
    json_upload = st.file_uploader("Cargar JSON", type=["json"])
    
    if json_upload:
        try:
            data = json.load(json_upload)
            # Actualizamos el estado con los datos del JSON
            st.session_state.meta_data['title'] = data.get('titulo', data.get('title', "Sin Titulo"))
            st.session_state.meta_data['author'] = data.get('autor', data.get('author', "Autor"))
            # ... otros campos ...
            st.success("Metadatos cargados")
        except:
            st.error("Error al leer el JSON")

    # Inputs vinculados
    m_title = st.text_input("Título del Libro", value=st.session_state.meta_data['title'])
    m_author = st.text_input("Autor", value=st.session_state.meta_data['author'])
    size_mode = st.selectbox("Formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

# Carga de Manuscrito
st.header("2. Manuscrito")
md_file = st.file_uploader("Subir Markdown", type=["md", "txt"])
md_content = md_file.read().decode("utf-8") if md_file else ""
final_text = st.text_area("Contenido", value=md_content, height=300)

# ==========================================
# 4. GENERACIÓN Y DESCARGA DINÁMICA
# ==========================================

if st.button("🚀 Generar y Descargar", use_container_width=True):
    if final_text:
        meta = {
            'title': m_title, 
            'author': m_author, 
            'subtitle': st.session_state.meta_data['subtitle'],
            'copyright': f"© {m_author}"
        }
        
        # 1. Generamos el binario del DOCX
        docx_out = run_book_conversion(final_text, meta, size_mode)
        
        # 2. Creamos el nombre de archivo dinámico
        nombre_limpio = limpiar_nombre_archivo(m_title)
        nombre_final = f"{nombre_limpio}.docx"
        
        st.success(f"Documento '{nombre_final}' listo.")
        
        # 3. Botón de descarga con el nombre dinámico
        st.download_button(
            label="📥 Presione aquí para descargar su libro",
            data=docx_out,
            file_name=nombre_final,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    else:
        st.error("No hay contenido para procesar.")
