import streamlit as st
from io import BytesIO
import tempfile
import os
import re
from datetime import datetime
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

def add_toc_field(paragraph):
    """
    Inserta un campo de Tabla de Contenidos (TOC) que Word puede actualizar.
    """
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    run = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._r.append(instrText)

    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'separate')
    run._r.append(fldChar)

    run = paragraph.add_run("Índice de contenidos (Actualizar en Word)")
    
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)

def add_formatted_text(paragraph, text):
    """
    Analiza marcas de Markdown inline (**negrita**, *cursiva*) y las aplica al documento.
    """
    parts = re.split(r'(\*\*\*.*?\*\*\*|___.*?___|\*\*.*?\*\*|__.*?__|\*.*?\*|_.*?_)', text)
    
    for part in parts:
        if not part: continue
        is_bold, is_italic = False, False
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
        run.bold, run.italic = is_bold, is_italic

# -------------------------------
# LÓGICA DE TEXTO Y CAPITALIZACIÓN
# -------------------------------

def fix_spanish_casing(text, user_exceptions=""):
    """
    Aplica la norma de capitalización española a títulos (Sentence case).
    """
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

    match = re.match(r'^(#+)\s*(.*)$', text)
    if match:
        hashes, content = match.group(1), match.group(2).strip()
        return f"{hashes} {process_segment(content)}" if content else text
    return process_segment(text)

# -------------------------------
# CONFIGURACIÓN DE ESTILOS
# -------------------------------

def apply_book_layout(section, size_option="Trade Paperback (5.5x8.5)"):
    """
    Aplica el tamaño de página y los márgenes de espejo.
    """
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.875)
        section.left_margin, section.right_margin = Inches(0.875), Inches(0.75)
    else:
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1.25), Inches(1)
    set_mirror_margins(section)

def setup_styles(doc):
    """
    Configura la tipografía Garamond y los estilos editoriales.
    """
    styles = doc.styles
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing, body.paragraph_format.first_line_indent = 1.2, Inches(0.25)

    if 'First Paragraph' not in styles: styles.add_style('First Paragraph', 1)
    fp = styles['First Paragraph']
    fp.font.name, fp.font.size = 'Garamond', Pt(11)
    fp.paragraph_format.alignment, fp.paragraph_format.first_line_indent = WD_ALIGN_PARAGRAPH.JUSTIFY, 0

    h1 = styles['Heading 1']
    h1.font.name, h1.font.size, h1.font.bold = 'Garamond', Pt(22), False
    h1.font.color.rgb = RGBColor(0,0,0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before, h1.paragraph_format.space_after = Inches(2.0), Inches(1.0)
    h1.paragraph_format.keep_with_next = True

    if 'BookTitle' not in styles:
        s = styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
        s.font.name, s.font.size, s.font.bold = 'Garamond', Pt(32), True
        s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        s.paragraph_format.space_before = Inches(1.0)

    return doc

# -------------------------------
# MOTOR DE GENERACIÓN DE LIBRO
# -------------------------------

def run_book_conversion(md_text, meta, size_option):
    """
    Genera la estructura completa del libro: Portada, Copyright, TOC, Blanca y Contenido.
    """
    doc = Document()
    setup_styles(doc)
    
    # 1. PORTADA (Página 1)
    current_section = doc.sections[0]
    apply_book_layout(current_section, size_option)
    
    p_title = doc.add_paragraph(style='BookTitle')
    add_formatted_text(p_title, fix_spanish_casing(meta['title'], meta['ex']))
    
    if meta['subtitle']:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_s = p_sub.add_run(meta['subtitle'])
        run_s.font.name, run_s.font.size = 'Garamond', Pt(18)
    
    for _ in range(5): doc.add_paragraph() # Espaciado
    
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.name, run_a.font.size, run_a.font.italic = 'Garamond', Pt(16), True
    
    for _ in range(3): doc.add_paragraph()
    
    p_pub = doc.add_paragraph()
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_pub.add_run(f"{meta['publisher']}\n{meta['year']}")
    run_p.font.name, run_p.font.size = 'Garamond', Pt(12)
    
    # 2. PÁGINA DE CRÉDITOS / COPYRIGHT (Página 2)
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.0)
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.name, run_c.font.size = 'Garamond', Pt(9)
    
    # 3. TABLA DE CONTENIDOS (Página 3)
    doc.add_page_break()
    doc.add_paragraph("Índice de contenidos", style='Heading 1').paragraph_format.space_before = Pt(24)
    toc_p = doc.add_paragraph()
    add_toc_field(toc_p)
    
    # 4. PÁGINA EN BLANCO (Página 4)
    doc.add_page_break()
    doc.add_paragraph("") 
    
    # 5. INICIO DEL CONTENIDO (Página 5+)
    lines = md_text.split('\n')
    is_first_para_after_heading = False
    
    for line in lines:
        raw_text = line.strip()
        if not raw_text: continue

        header_match = re.match(r'^(#+)\s*(.*)$', raw_text)
        if header_match:
            hashes, content = header_match.group(1), header_match.group(2).strip()
            level = len(hashes)
            
            if level == 1:
                # Los capítulos siempre empiezan en página IMPAR (Recto)
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_book_layout(new_sect, size_option)
                p = doc.add_paragraph(style='Heading 1')
                add_formatted_text(p, content)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                add_formatted_text(p, content)
                for r in p.runs: 
                    r.bold = (level <= 3)
                    r.font.size = Pt(14 if level == 2 else 12)
            is_first_para_after_heading = True
        else:
            style = 'First Paragraph' if is_first_para_after_heading else 'Body Text'
            p = doc.add_paragraph(style=style)
            add_formatted_text(p, " ".join(raw_text.split()))
            is_first_para_after_heading = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ DE USUARIO (ESP)
# -------------------------------

st.set_page_config(page_title="Markdown a Word Profesional", page_icon="📚")
st.title("📚 Generador Editorial Profesional")
st.markdown("Convierte tu manuscrito Markdown en un documento Word maquetado para impresión física.")

with st.sidebar:
    st.header("⚙️ Configuración del Libro")
    
    with st.expander("Metadatos del Manuscrito", expanded=True):
        m_title = st.text_input("Título del Libro", "Mi Obra Maestra")
        m_sub = st.text_input("Subtítulo", "Una historia inolvidable")
        m_author = st.text_input("Nombre del Autor/a", "Nombre Apellido")
        m_pub = st.text_input("Sello Editorial", "Ediciones del Sol")
        m_year = st.text_input("Año de publicación", str(datetime.now().year))
        m_file = st.text_input("Nombre del archivo de salida", "manuscrito_final")

    with st.expander("Créditos y Copyright"):
        m_copy = st.text_area("Texto de copyright", 
                             f"© {datetime.now().year} {m_author}.\nTodos los derechos reservados.\nQueda prohibida la reproducción total o parcial...")

    size_mode = st.selectbox("Formato de página", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    apply_fix = st.checkbox("Corregir mayúsculas en títulos", value=True)
    exceptions = st.text_area("Excepciones de capitalización", placeholder="Cervantes, Madrid, ONU...")

tab_edit, tab_preview = st.tabs(["📝 Editor Markdown", "🔍 Análisis de Estructura"])

with tab_edit:
    uploaded_md = st.file_uploader("Subir archivo .md", type=["md", "txt"])
    content = uploaded_md.read().decode("utf-8") if uploaded_md else ""
    content = st.text_area("Pega o edita tu contenido aquí", value=content, height=400)

with tab_preview:
    if content:
        st.subheader("Capítulos y Secciones Detectadas")
        for l in content.split('\n'):
            if l.strip().startswith('#'):
                st.info(f"Encabezado detectado: {l}")
    else:
        st.write("Escribe algo en el editor para ver el análisis.")

if st.button("🚀 Generar Libro para Impresión", use_container_width=True):
    if not content.strip():
        st.error("Por favor, introduce el contenido del manuscrito.")
    else:
        with st.spinner("Generando maquetación profesional..."):
            meta = {
                'title': m_title, 'subtitle': m_sub, 'author': m_author,
                'publisher': m_pub, 'year': m_year, 'copyright': m_copy, 'ex': exceptions
            }
            
            final_content = content
            if apply_fix:
                lines = content.split('\n')
                final_content = '\n'.join([fix_spanish_casing(l, exceptions) if l.strip().startswith('#') else l for l in lines])

            result = run_book_conversion(final_content, meta, size_mode)
            st.success("✅ ¡Documento generado correctamente!")
            st.download_button(
                label="📥 Descargar archivo Word (.docx)",
                data=result,
                file_name=f"{m_file}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
