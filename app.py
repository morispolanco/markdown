import streamlit as st
import json
import re
from datetime import datetime
from io import BytesIO
# ... (tus otras importaciones de docx se mantienen igual)

# -------------------------------
# LÓGICA DE EXTRACCIÓN JSON
# -------------------------------

def cargar_datos_desde_json(archivo_json):
    """
    Intenta leer el JSON y normalizar las llaves a la estructura interna.
    Soporta sinónimos en español e inglés.
    """
    try:
        data = json.load(archivo_json)
        
        # Diccionario de normalización (Mapeo de llaves)
        # Permite que el JSON use "titulo" o "title", "autor" o "author", etc.
        mapeo = {
            'title': data.get('titulo', data.get('title', 'Sin Título')),
            'subtitle': data.get('subtitulo', data.get('subtitle', '')),
            'author': data.get('autor', data.get('author', 'Autor Anónimo')),
            'publisher': data.get('editorial', data.get('publisher', 'Editorial Independiente')),
            'year': str(data.get('año', data.get('year', datetime.now().year))),
            'copyright': data.get('copyright', '')
        }
        
        # Generar copyright automático si el campo viene vacío en el JSON
        if not mapeo['copyright']:
            mapeo['copyright'] = f"© {mapeo['year']} {mapeo['author']}. Todos los derechos reservados."
            
        return mapeo, True
    except Exception as e:
        return None, f"Error al leer el JSON: {e}"

# -------------------------------
# CONFIGURACIÓN INICIAL DE INTERFAZ
# -------------------------------

st.set_page_config(page_title="Generador Editorial", layout="wide")
st.title("📚 Generador Editorial Profesional")

# Inicializamos el estado de la sesión para persistir los metadatos
if 'metadata' not in st.session_state:
    st.session_state.metadata = {
        'title': "Título del Libro",
        'subtitle': "",
        'author': "Nombre del Autor",
        'publisher': "Mi Editorial",
        'year': str(datetime.now().year),
        'copyright': ""
    }

# -------------------------------
# SIDEBAR: FICHA EDITORIAL (JSON)
# -------------------------------

with st.sidebar:
    st.header("1. Ficha Técnica")
    st.info("Sube un archivo .json para autocompletar los campos.")
    
    archivo_ficha = st.file_uploader("Cargar ficha JSON", type=["json"])
    
    if archivo_ficha is not None:
        datos_extraidos, exito = cargar_datos_desde_json(archivo_ficha)
        if exito:
            st.session_state.metadata = datos_extraidos
            st.success("¡Metadatos cargados!")
        else:
            st.error(datos_extraidos)

    st.divider()
    
    # Formulario de edición (se precarga con lo que haya en session_state)
    st.subheader("Edición de Metadatos")
    m_title = st.text_input("Título", value=st.session_state.metadata['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.metadata['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.metadata['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.metadata['publisher'])
    m_year = st.text_input("Año", value=st.session_state.metadata['year'])
    m_copy = st.text_area("Copyright", value=st.session_state.metadata['copyright'])
    
    size_mode = st.selectbox("Formato de impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    m_file_name = st.text_input("Nombre del archivo de salida", "manuscrito_maquetado")

# -------------------------------
# CUERPO PRINCIPAL: MANUSCRITO
# -------------------------------

st.header("2. Contenido del Manuscrito")
archivo_md = st.file_uploader("Sube tu archivo Markdown (.md o .txt)", type=["md", "txt"])

# Manejo del contenido de texto
content = ""
if archivo_md:
    content = archivo_md.read().decode("utf-8")
    st.success(f"Archivo '{archivo_md.name}' listo para procesar.")

editor_text = st.text_area(
    "Editor/Previsualización del Contenido",
    value=content,
    height=400
)

# -------------------------------
# PROCESAMIENTO FINAL
# -------------------------------

if st.button("🚀 Generar Documento Editorial", use_container_width=True):
    if editor_text:
        # Empaquetamos los metadatos actuales de la interfaz
        meta_final = {
            'title': m_title,
            'subtitle': m_sub,
            'author': m_author,
            'publisher': m_pub,
            'year': m_year,
            'copyright': m_copy
        }
        
        # Llamada a tu función original (que definiste arriba en tu script)
        try:
            docx_bundle = run_book_conversion(editor_text, meta_final, size_mode)
            st.success("¡Documento generado con éxito!")
            st.download_button(
                label="📥 Descargar Manuscrito (.docx)",
                data=docx_bundle,
                file_name=f"{m_file_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error durante la conversión: {e}")
    else:
        st.warning("El contenido está vacío. Sube un archivo o escribe en el editor.")
