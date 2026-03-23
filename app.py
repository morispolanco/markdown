import json  # Nueva dependencia estándar

# -------------------------------
# NUEVA FUNCIÓN: PROCESAMIENTO DE JSON
# -------------------------------

def process_editorial_json(json_file):
    """
    Lee el archivo JSON y mapea los campos a la estructura interna.
    Si falta algún campo, utiliza valores por defecto.
    """
    try:
        data = json.load(json_file)
        
        # Mapeo y limpieza de datos con valores por defecto (fallback)
        clean_meta = {
            'title': data.get('titulo', data.get('title', "Título del Libro")),
            'subtitle': data.get('subtitulo', data.get('subtitle', "")),
            'author': data.get('autor', data.get('author', "Nombre del Autor")),
            'publisher': data.get('editorial', data.get('publisher', "Mi Editorial")),
            'year': str(data.get('año', data.get('year', datetime.now().year))),
            'copyright': data.get('copyright', "")
        }
        
        # Generación automática de Copyright si el JSON no lo trae
        if not clean_meta['copyright']:
            clean_meta['copyright'] = f"© {clean_meta['year']} {clean_meta['author']}. Todos los derechos reservados."
            
        return clean_meta, True
    except Exception as e:
        return None, f"Error en el formato JSON: {e}"

# -------------------------------
# INTERFAZ DE USUARIO (STREAMLIT)
# -------------------------------

st.title("📚 Generador Editorial Profesional")

# Inicialización de metadatos en el estado de la sesión (Session State)
if 'meta_data' not in st.session_state:
    st.session_state.meta_data = {
        'title': "Título del Libro", 'subtitle': "", 'author': "Nombre del Autor",
        'publisher': "Mi Editorial", 'year': str(datetime.now().year),
        'copyright': f"© {datetime.now().year} Nombre del Autor. Todos los derechos reservados."
    }

with st.sidebar:
    st.header("1. Configuración Editorial")
    
    # Carga de la Ficha JSON
    json_upload = st.file_uploader("Subir Ficha Editorial (JSON)", type=["json"])
    
    if json_upload:
        new_meta, status = process_editorial_json(json_upload)
        if new_meta:
            st.session_state.meta_data = new_meta
            st.success("✅ Datos cargados desde JSON")
        else:
            st.error(status)

    st.divider()
    
    # Campos de texto vinculados al session_state
    m_title = st.text_input("Título del Libro", value=st.session_state.meta_data['title'])
    m_sub = st.text_input("Subtítulo (opcional)", value=st.session_state.meta_data['subtitle'])
    m_author = st.text_input("Nombre del Autor", value=st.session_state.meta_state['author'])
    m_pub = st.text_input("Sello Editorial", value=st.session_state.meta_data['publisher'])
    m_year = st.text_input("Año de Edición", value=st.session_state.meta_data['year'])
    size_mode = st.selectbox("Formato de impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    m_copy = st.text_area("Texto de Copyright", value=st.session_state.meta_data['copyright'])
    m_file = st.text_input("Nombre del archivo de salida", "manuscrito_maquetado")

# ... (Continúa con la carga del archivo Markdown y el botón de generación)
