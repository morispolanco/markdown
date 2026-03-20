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

def clean_text(paragraph):
    """Limpia espacios dobles y artefactos de párrafo."""
    if paragraph.text:
        paragraph.text = " ".join(paragraph.text.split())

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
    Motor principal que implementa la lógica de libro físico:
    - Márgenes de espejo.
    - Capítulos en páginas impares.
    - Alternancia de sangrías.
    """
    doc = Document()
    setup_styles(doc)
    
    # Configurar primera sección (Portada)
    current_section = doc.sections[0]
    apply_book_layout(current_section, size_option)

    # 1. Portada
    doc.add_paragraph(fix_spanish_casing(title, user_exceptions), style='BookTitle')
    auth_p = doc.add_paragraph(author)
    auth_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    auth_p.runs[0].font.name = 'Garamond'
    auth_p.runs[0].font.size = Pt(14)
    auth_p.runs[0].font.italic = True
    
    # 2. Procesar Contenido
    lines = md_text.split('\n')
    is_first_para_after_heading = False
    
    for line in lines:
        text = line.strip()
        if not text: continue

        # Detección de Capítulos
        if text.startswith('# '):
            # Nuevo capítulo -> Nueva sección en página IMPAR (Recto)
            new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
            apply_book_layout(new_sect, size_option)
            
            doc.add_paragraph(text[2:], style='Heading 1')
            is_first_para_after_heading = True
            
        elif text.startswith('## '):
            p = doc.add_paragraph(text[3:])
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(14)
            p.paragraph_format.space_before = Pt(18)
            is_first_para_after_heading = True
            
        else:
            # Párrafos de texto
            style = 'First Paragraph' if is_first_para_after_heading else 'Body Text'
            p = doc.add_paragraph(text, style=style)
            clean_text(p)
            is_first_para_after_heading = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ DE USUARIO (UI)
# -------------------------------

st.title("📚 Generador de Libros Profesionales")
st.markdown("Transforma tu Markdown en un archivo Word listo para imprenta (5.5\" x 8.5\").")

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
        engine_opt = ["Motor Editorial (Recomendado para libros)"]
        if PANDOC_AVAILABLE: engine_opt.append("Pandoc (General)")
        selected_engine = st.selectbox("Motor", engine_opt)

# Pestañas
tab_edit, tab_preview = st.tabs(["📝 Manuscrito", "🔍 Previsualización de Títulos"])

with tab_edit:
    uploaded_md = st.file_uploader("Cargar .md", type=["md", "txt"])
    content = ""
    if uploaded_md:
        content = uploaded_md.read().decode("utf-8")
    content = st.text_area("Contenido Markdown", value=content, height=450)

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
        with st.spinner("Maquetando páginas y ajustando márgenes..."):
            # Procesar capitalización si aplica
            final_content = content
            if apply_fix:
                lines = content.split('\n')
                final_content = '\n'.join([fix_spanish_casing(l, exceptions) if l.startswith('#') else l for l in lines])

            try:
                if "Motor Editorial" in selected_engine:
                    result = run_book_conversion(final_content, doc_title, doc_author, exceptions, size_mode)
                else:
                    # Fallback a Pandoc si está seleccionado
                    result = pypandoc.convert_text(final_content, "docx", format="md")

                st.success("✅ ¡Manuscrito generado con éxito!")
                st.download_button(
                    label="📥 Descargar DOCX para Imprenta",
                    data=result,
                    file_name=f"{file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error: {e}")
