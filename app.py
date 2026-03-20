import streamlit as st
from io import BytesIO
import tempfile
import os
import re
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Intentar importar motores opcionales
try:
    import pypandoc
    PANDOC_AVAILABLE = True
except ImportError:
    PANDOC_AVAILABLE = False

# -------------------------------
# UTILIDADES DE BAJO NIVEL (XML)
# -------------------------------

def set_mirror_margins(section):
    """
    Habilita los márgenes simétricos en el XML de Word.
    Asegura que el 'gutter' (canal de encuadernación) esté siempre en el interior.
    """
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_formatted_text(paragraph, text):
    """
    Analiza el texto en busca de marcas de Markdown inline (negrita y cursiva)
    y las aplica como 'runs' de Word.
    """
    # Regex para detectar negrita (** o __) y cursiva (* o _)
    # Priorizamos la combinación de ambas (*** o ___)
    pattern = re.compile(r'(\*\*\*.*?\*\*\*|___.*?___|\*\*.*?\*\*|__.*?__|Simplified\*.*?\*|_.*?_)')
    
    # Dividir el texto conservando los delimitadores
    parts = re.split(r'(\*\*\*.*?\*\*\*|___.*?___|\*\*.*?\*\*|__.*?__|\*.*?\*|_.*?_)', text)
    
    for part in parts:
        if not part:
            continue
            
        is_bold = False
        is_italic = False
        clean_part = part
        
        if part.startswith('***') and part.endswith('***'):
            is_bold = is_italic = True
            clean_part = part[3:-3]
        elif part.startswith('___') and part.endswith('___'):
            is_bold = is_italic = True
            clean_part = part[3:-3]
        elif (part.startswith('**') and part.endswith('**')) or (part.startswith('__') and part.endswith('__')):
            is_bold = True
            clean_part = part[2:-2]
        elif (part.startswith('*') and part.endswith('*')) or (part.startswith('_') and part.endswith('_')):
            is_italic = True
            clean_part = part[1:-1]
            
        run = paragraph.add_run(clean_part)
        run.bold = is_bold
        run.italic = is_italic

def clean_text(paragraph):
    """Limpia espacios dobles y artefactos de párrafo básicos."""
    # Nota: No usamos join(split()) aquí porque romperíamos los runs formateados.
    # La limpieza se hace ahora a nivel de string antes de crear los runs.
    pass

# -------------------------------
# CONFIGURACIÓN Y CONSTANTES
# -------------------------------
st.set_page_config(
    page_title="Markdown Editorial Pro", 
    page_icon="📚", 
    layout="wide"
)

# -------------------------------
# LÓGICA DE PROCESAMIENTO TEXTUAL
# -------------------------------

def fix_spanish_casing(text, user_exceptions=""):
    """Aplica la norma de capitalización española a encabezados."""
    if not text: return text
    
    base_excepciones = ["España", "México", "Dios", "Python", "Streamlit", "Word", "Markdown", "I", "II", "III", "IV", "V"]
    user_list = [ex.strip() for ex in user_exceptions.split(",") if ex.strip()]
    lista_total = set(base_excepciones + user_list)

    def process_segment(segment):
        palabras = segment.split()
        if not palabras: return ""
        resultado = []
        for i, word in enumerate(palabras):
            clean_word = re.sub(r'[^\wáéíóúÁÉÍÓÚñÑ]', '', word)
            if i == 0:
                resultado.append(word[0].upper() + word[1:] if len(word) > 1 else word.upper())
            elif clean_word in lista_total or clean_word.istitle():
                resultado.append(word)
            else:
                resultado.append(word.lower())
        return " ".join(resultado)

    match = re.match(r'^(#+)\s+(.+)$', text)
    if match:
        hashes = match.group(1)
        content = match.group(2)
        return f"{hashes} {process_segment(content)}"
    return process_segment(text)

# -------------------------------
# GESTIÓN DE ESTILOS PROFESIONALES
# -------------------------------

def apply_book_layout(section, size_option="Trade Paperback (5.5x8.5)"):
    """Aplica tamaño de papel y márgenes profesionales."""
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width = Inches(5.5)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.875)
        section.left_margin = Inches(0.875)  # Interior (Gutter)
        section.right_margin = Inches(0.75)  # Exterior
    else:
        # A4 estándar
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1)

    set_mirror_margins(section)

def setup_styles(doc):
    """Define la tipografía Garamond y jerarquías de libro."""
    styles = doc.styles

    # --- Texto Base (Cuerpo con sangría) ---
    if 'Body Text' not in styles:
        styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name = 'Garamond'
    body.font.size = Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    body.paragraph_format.line_spacing = 1.2
    body.paragraph_format.first_line_indent = Inches(0.25)
    body.paragraph_format.widow_control = True
    body.paragraph_format.space_after = Pt(0)

    # --- Primer Párrafo (Sin sangría) ---
    if 'First Paragraph' not in styles:
        styles.add_style('First Paragraph', 1)
    fp = styles['First Paragraph']
    fp.font.name = 'Garamond'
    fp.font.size = Pt(11)
    fp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fp.paragraph_format.line_spacing = 1.2
    fp.paragraph_format.first_line_indent = 0
    fp.paragraph_format.space_after = Pt(0)

    # --- Título de Capítulo (Heading 1) ---
    h1 = styles['Heading 1']
    h1.font.name = 'Garamond'
    h1.font.size = Pt(22)
    h1.font.bold = False
    h1.font.color.rgb = RGBColor(0,0,0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Inches(2.0)
    h1.paragraph_format.space_after = Inches(1.0)
    h1.paragraph_format.keep_with_next = True

    # --- Título de Obra (Portada) ---
    if 'BookTitle' not in styles:
        s = styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
        s.font.name = 'Garamond'; s.font.size = Pt(28); s.font.bold = True
        s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        s.paragraph_format.space_before = Inches(1.5)

    return doc

# -------------------------------
# MOTOR DE CONVERSIÓN MEJORADO
# -------------------------------

def run_book_conversion(md_text, title, author, user_exceptions, size_option):
    """
    Motor principal que implementa la lógica de libro físico y procesa Markdown inline.
    """
    doc = Document()
    setup_styles(doc)
    
    # Configurar primera sección (Portada)
    current_section = doc.sections[0]
    apply_book_layout(current_section, size_option)

    # 1. Portada
    title_p = doc.add_paragraph(style='BookTitle')
    add_formatted_text(title_p, fix_spanish_casing(title, user_exceptions))
    
    auth_p = doc.add_paragraph()
    auth_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = auth_p.add_run(author)
    run.font.name = 'Garamond'
    run.font.size = Pt(14)
    run.font.italic = True
    
    # 2. Procesar Contenido
    lines = md_text.split('\n')
    is_first_para_after_heading = False
    
    for line in lines:
        text = line.strip()
        if not text: 
            continue

        # Detección de Capítulos
        if text.startswith('# '):
            # Nuevo capítulo -> Nueva sección en página IMPAR (Recto)
            new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
            apply_book_layout(new_sect, size_option)
            
            p = doc.add_paragraph(style='Heading 1')
            add_formatted_text(p, text[2:])
            is_first_para_after_heading = True
            
        elif text.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(18)
            add_formatted_text(p, text[3:])
            # Aplicar formato de subencabezado a los runs creados
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(14)
            is_first_para_after_heading = True
            
        else:
            # Párrafos de texto con soporte para Markdown inline
            style = 'First Paragraph' if is_first_para_after_heading else 'Body Text'
            p = doc.add_paragraph(style=style)
            
            # Limpiar espacios dobles antes de procesar runs
            clean_line = " ".join(text.split())
            add_formatted_text(p, clean_line)
            
            is_first_para_after_heading = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ DE USUARIO (UI)
# -------------------------------

st.title("📚 Generador de Libros Profesionales")
st.markdown("Transforma tu Markdown en un archivo Word maquetado para imprenta, ahora con soporte para **negrita** y *cursiva*.")

with st.sidebar:
    st.header("⚙️ Configuración del Libro")
    
    with st.expander("Metadatos", expanded=True):
        doc_title = st.text_input("Título de la Obra", "Título del Libro")
        doc_author = st.text_input("Autor/a", "Nombre del Autor")
        file_name = st.text_input("Nombre de archivo", "manuscrito_final")

    with st.expander("Maquetación Física"):
        size_mode = st.selectbox("Formato de impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
        st.info("El formato Trade Paperback incluye márgenes simétricos y capítulos en páginas derechas.")

    with st.expander("Normativa Española"):
        apply_fix = st.checkbox("Corregir mayúsculas en títulos", value=True)
        exceptions = st.text_area("Excepciones", placeholder="Ej: Madrid, Cervantes, ONU...")

    with st.expander("Motor de Conversión"):
        engine_opt = ["Motor Editorial (Analizador de Markdown nativo)"]
        if PANDOC_AVAILABLE: engine_opt.append("Pandoc (General)")
        selected_engine = st.selectbox("Motor", engine_opt)

# Pestañas
tab_edit, tab_preview = st.tabs(["📝 Manuscrito", "🔍 Previsualización de Títulos"])

with tab_edit:
    uploaded_md = st.file_uploader("Cargar .md", type=["md", "txt"])
    content = ""
    if uploaded_md:
        content = uploaded_md.read().decode("utf-8")
    content = st.text_area("Contenido Markdown", value=content, height=450, 
                           placeholder="Escribe aquí... Usa **negrita** y *cursiva* para probar.")

with tab_preview:
    if content:
        st.subheader("Tratamiento de Títulos (Norma Española)")
        for l in content.split('\n'):
            if l.startswith('#'):
                st.write(f"Original: `{l}`")
                st.write(f"Corregido: `{fix_spanish_casing(l, exceptions)}`")
                st.divider()

st.divider()
if st.button("🚀 Generar Archivo para Imprenta", use_container_width=True):
    if not content.strip():
        st.error("Escribe contenido antes de continuar.")
    else:
        with st.spinner("Convirtiendo Markdown y maquetando páginas..."):
            # Procesar capitalización si aplica
            final_content = content
            if apply_fix:
                lines = content.split('\n')
                final_content = '\n'.join([fix_spanish_casing(l, exceptions) if l.startswith('#') else l for l in lines])

            try:
                if "Motor Editorial" in selected_engine:
                    result = run_book_conversion(final_content, doc_title, doc_author, exceptions, size_mode)
                else:
                    # Fallback a Pandoc si está seleccionado y disponible
                    result = pypandoc.convert_text(final_content, "docx", format="md")

                st.success("✅ ¡Manuscrito convertido y generado con éxito!")
                st.download_button(
                    label="📥 Descargar DOCX Formateado",
                    data=result,
                    file_name=f"{file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error: {e}")
