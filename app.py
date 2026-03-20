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
    # Dividir el texto conservando los delimitadores para negrita y cursiva
    parts = re.split(r'(\*\*\*.*?\*\*\*|___.*?___|\*\*.*?\*\*|__.*?__|\*.*?\*|_.*?_)', text)
    
    for part in parts:
        if not part:
            continue
            
        is_bold = False
        is_italic = False
        clean_part = part
        
        if (part.startswith('***') and part.endswith('***')) or (part.startswith('___') and part.endswith('___')):
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

# -------------------------------
# LÓGICA DE PROCESAMIENTO TEXTUAL
# -------------------------------

def fix_spanish_casing(text, user_exceptions=""):
    """
    Aplica la norma de capitalización española a encabezados, 
    preservando los hashtags para el motor de conversión.
    """
    if not text: return text
    
    # Nombres propios y siglas
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

    # Regex mejorada para detectar encabezados con o sin espacio
    match = re.match(r'^(#+)\s*(.*)$', text)
    if match:
        hashes = match.group(1)
        content = match.group(2).strip()
        if not content: return text # Línea de solo hashes
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
        section.left_margin = Inches(0.875)
        section.right_margin = Inches(0.75)
    else:
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

    # Cuerpo base
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name = 'Garamond'; body.font.size = Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.2
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Primer párrafo tras título
    if 'First Paragraph' not in styles: styles.add_style('First Paragraph', 1)
    fp = styles['First Paragraph']
    fp.font.name = 'Garamond'; fp.font.size = Pt(11)
    fp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fp.paragraph_format.first_line_indent = 0

    # H1 - Capítulos
    h1 = styles['Heading 1']
    h1.font.name = 'Garamond'; h1.font.size = Pt(22); h1.font.bold = False
    h1.font.color.rgb = RGBColor(0,0,0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Inches(2.0)
    h1.paragraph_format.space_after = Inches(1.0)
    h1.paragraph_format.keep_with_next = True

    # Portada
    if 'BookTitle' not in styles:
        s = styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
        s.font.name = 'Garamond'; s.font.size = Pt(28); s.font.bold = True
        s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        s.paragraph_format.space_before = Inches(1.5)

    return doc

# -------------------------------
# MOTOR DE CONVERSIÓN CORREGIDO
# -------------------------------

def run_book_conversion(md_text, title, author, user_exceptions, size_option):
    doc = Document()
    setup_styles(doc)
    
    current_section = doc.sections[0]
    apply_book_layout(current_section, size_option)

    # 1. Portada
    title_p = doc.add_paragraph(style='BookTitle')
    add_formatted_text(title_p, fix_spanish_casing(title, user_exceptions))
    
    auth_p = doc.add_paragraph()
    auth_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = auth_p.add_run(author)
    run.font.name = 'Garamond'; run.font.size = Pt(14); run.font.italic = True
    
    # 2. Procesar Contenido
    lines = md_text.split('\n')
    is_first_para_after_heading = False
    
    for line in lines:
        raw_text = line.strip()
        if not raw_text: continue

        # Detección mejorada de encabezados (Hashtags)
        header_match = re.match(r'^(#+)\s*(.*)$', raw_text)
        
        if header_match:
            hashes = header_match.group(1)
            level = len(hashes)
            content = header_match.group(2).strip()
            
            if level == 1:
                # Nivel 1: Nuevo capítulo y nueva sección
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_book_layout(new_sect, size_option)
                p = doc.add_paragraph(style='Heading 1')
                add_formatted_text(p, content)
            elif level == 2:
                # Nivel 2: Subtítulo destacado
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(18)
                p.paragraph_format.space_after = Pt(6)
                add_formatted_text(p, content)
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(14)
            elif level == 3:
                # Nivel 3: Subtítulo menor
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                add_formatted_text(p, content)
                for r in p.runs:
                    r.bold = True
                    r.font.italic = True
                    r.font.size = Pt(12)
            else:
                # Niveles 4+ (Estilo negrita simple)
                p = doc.add_paragraph()
                add_formatted_text(p, content)
                for r in p.runs: r.bold = True

            is_first_para_after_heading = True
            
        else:
            # Párrafos normales
            style = 'First Paragraph' if is_first_para_after_heading else 'Body Text'
            p = doc.add_paragraph(style=style)
            # Limpiar espacios extra pero mantener formato inline
            clean_line = " ".join(raw_text.split())
            add_formatted_text(p, clean_line)
            is_first_para_after_heading = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ DE USUARIO (UI)
# -------------------------------

st.title("📚 Generador Editorial Pro")
st.markdown("Conversión precisa de Markdown a Word para impresión profesional.")

with st.sidebar:
    st.header("⚙️ Configuración")
    doc_title = st.text_input("Título de la Obra", "Título del Libro")
    doc_author = st.text_input("Autor/a", "Nombre del Autor")
    file_name = st.text_input("Nombre de archivo", "manuscrito_final")
    size_mode = st.selectbox("Formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    apply_fix = st.checkbox("Corregir mayúsculas en títulos", value=True)
    exceptions = st.text_area("Excepciones", placeholder="Ej: Madrid, Cervantes...")

tab_edit, tab_preview = st.tabs(["📝 Manuscrito", "🔍 Análisis de Títulos"])

with tab_edit:
    uploaded_md = st.file_uploader("Cargar .md", type=["md", "txt"])
    content = uploaded_md.read().decode("utf-8") if uploaded_md else ""
    content = st.text_area("Contenido Markdown", value=content, height=450, placeholder="Escribe aquí... # Capítulo 1")

with tab_preview:
    if content:
        st.subheader("Detección de Estructura")
        for l in content.split('\n'):
            if l.strip().startswith('#'):
                m = re.match(r'^(#+)\s*(.*)$', l.strip())
                if m:
                    level = len(m.group(1))
                    txt = m.group(2)
                    st.write(f"Nivel {level}: **{txt}**")

if st.button("🚀 Generar para Impresión", use_container_width=True):
    if not content.strip():
        st.error("No hay contenido para procesar.")
    else:
        with st.spinner("Procesando hashtags y maquetación..."):
            final_content = content
            if apply_fix:
                lines = content.split('\n')
                final_content = '\n'.join([fix_spanish_casing(l, exceptions) if l.strip().startswith('#') else l for l in lines])

            result = run_book_conversion(final_content, doc_title, doc_author, exceptions, size_mode)
            st.success("✅ ¡Conversión exitosa!")
            st.download_button(
                label="📥 Descargar DOCX",
                data=result,
                file_name=f"{file_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
