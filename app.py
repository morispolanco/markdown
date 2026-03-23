import yaml # Opcional, pero usaremos Regex para no añadir dependencias externas

# -------------------------------
# NUEVA FUNCIÓN: EXTRACCIÓN DE METADATOS
# -------------------------------

def extract_metadata(text):
    """
    Busca patrones tipo 'Clave: Valor' en el texto.
    Soporta: Título, Subtítulo, Autor, Editorial, Año, Copyright.
    """
    metadata = {
        'title': "Título del Libro",
        'subtitle': "",
        'author': "Nombre del Autor",
        'publisher': "Mi Editorial",
        'year': str(datetime.now().year),
        'copyright': ""
    }
    
    # Mapeo de términos que el usuario podría usar en la ficha
    patterns = {
        'title': r'(?i)T[íi]tulo:\s*(.*)',
        'subtitle': r'(?i)Subt[íi]tulo:\s*(.*)',
        'author': r'(?i)Autor:\s*(.*)',
        'publisher': r'(?i)Editorial|Sello:\s*(.*)',
        'year': r'(?i)A[ñn]o:\s*(\d{4})',
        'copyright': r'(?i)Copyright:\s*(.*)'
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            metadata[key] = match.group(1).strip()
            
    # Si no hay copyright explícito, lo generamos
    if not metadata['copyright']:
        metadata['copyright'] = f"© {metadata['year']} {metadata['author']}. Todos los derechos reservados."
        
    return metadata

# -------------------------------
# INTERFAZ DE USUARIO MODIFICADA
# -------------------------------

st.title("📚 Generador Editorial Profesional")

# 1. CARGA DE FICHA TÉCNICA (METADATOS)
st.sidebar.header("1. Cargar Ficha Editorial")
meta_file = st.sidebar.file_uploader("Sube la ficha (.txt o .md)", type=["txt", "md"], key="meta_upload")

# Diccionario inicial por defecto
extracted_meta = {
    'title': "Título del Libro", 'subtitle': "", 'author': "Nombre del Autor",
    'publisher': "Mi Editorial", 'year': str(datetime.now().year), 'copyright': ""
}

if meta_file:
    meta_text = meta_file.read().decode("utf-8")
    extracted_meta = extract_metadata(meta_text)
    st.sidebar.success("✅ Metadatos extraídos")

# 2. CAMPOS DE EDICIÓN EN SIDEBAR (Se llenan con lo extraído)
with st.sidebar:
    st.header("2. Confirmar Información")
    m_title = st.text_input("Título del Libro", value=extracted_meta['title'])
    m_sub = st.text_input("Subtítulo (opcional)", value=extracted_meta['subtitle'])
    m_author = st.text_input("Nombre del Autor", value=extracted_meta['author'])
    m_pub = st.text_input("Sello Editorial", value=extracted_meta['publisher'])
    m_year = st.text_input("Año de Edición", value=extracted_meta['year'])
    size_mode = st.selectbox("Formato de impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    m_copy = st.text_area("Texto de Copyright", value=extracted_meta['copyright'])
    m_file = st.text_input("Nombre del archivo de salida", "manuscrito_maquetado")

# 3. CARGA DEL MANUSCRITO (CONTENIDO)
st.header("Carga de Contenido")
uploaded_file = st.file_uploader("Sube tu manuscrito (Markdown)", type=["md", "txt"], key="content_upload")

# ... (El resto del código de procesamiento se mantiene igual)
