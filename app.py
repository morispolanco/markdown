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

# -------------------------------
# UTILIDADES DE BAJO NIVEL (XML)
# -------------------------------

def set_mirror_margins(section):
    """Habilita márgenes simétricos (espejo) en el XML de Word."""
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_page_number(footer):
    """Inserta numeración de página automática en el centro del pie de página."""
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Iniciar campo
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    # Definir que es el campo PAGE
    run = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    run._r.append(instrText)

    # Finalizar campo
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)

def add_toc_field(paragraph):
    """Inserta el código de campo para la Tabla de Contenidos."""
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

    run = paragraph.add_run("Haga clic derecho aquí y seleccione 'Actualizar campo' para generar el índice.")
    
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)

def add_formatted_text(paragraph, text):
    """Procesa negritas y cursivas de Markdown inline."""
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
# ESTILOS Y MAQUETACIÓN
# -------------------------------

def apply_layout(section, size_option):
    """Configura tamaño, márgenes y numeración de página para una sección."""
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.875)
        section.left_margin = Inches(0.875)
        section.right_margin = Inches(0.75)
    else:
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1.25), Inches(1)
    
    set_mirror_margins(section)
    add_page_number(section.footer)

def setup_styles(doc):
    """Configura los estilos Heading 1, 2, 3 y Body Text."""
    styles = doc.styles
    
    # Texto normal
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.2
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Heading 1 (Capítulos)
    h1 = styles['Heading 1']
    h1.font.name, h1.font.size, h1.font.bold = 'Garamond', Pt(24), False
    h1.font.color.rgb = RGBColor(0,0,0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before, h1.paragraph_format.space_after = Inches(2.0), Inches(1.0)
    h1.paragraph_format.keep_with_next = True

    # Heading 2 (Secciones)
    h2 = styles['Heading 2']
    h2.font.name, h2.font.size, h2.font.bold = 'Garamond', Pt(16), True
    h2.font.color.rgb = RGBColor(0,0,0)
    h2.paragraph_format.space_before, h2.paragraph_format.space_after = Pt(18), Pt(12)

    # Heading 3 (Subsecciones)
    h3 = styles['Heading 3']
    h3.font.name, h3.font.size, h3.font.italic = 'Garamond', Pt(13), True
    h3.font.color.rgb = RGBColor(0,0,0)
    h3.paragraph_format.space_before, h3.paragraph_format.space_after = Pt(12), Pt(6)

    return doc

# -------------------------------
# MOTOR DE CONVERSIÓN
# -------------------------------

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # 1. PORTADA
    section = doc.sections[0]
    apply_layout(section, size_option)
    
    p_title = doc.add_paragraph(meta['title'])
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.runs[0].font.size = Pt(32)
    p_title.runs[0].font.bold = True
    
    if meta['subtitle']:
        p_sub = doc.add_paragraph(meta['subtitle'])
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.runs[0].font.size = Pt(18)
    
    for _ in range(6): doc.add_paragraph()
    
    p_author = doc.add_paragraph(meta['author'])
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.runs[0].font.size = Pt(16)
    p_author.runs[0].font.italic = True

    # 2. COPYRIGHT
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.5)
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.size = Pt(9)
    
    # 3. ÍNDICE
    doc.add_page_break()
    doc.add_paragraph("Índice de contenidos", style='Heading 1').paragraph_format.space_before = Pt(24)
    toc_p = doc.add_paragraph()
    add_toc_field(toc_p)
    
    # 4. PÁGINA BLANCA
    doc.add_page_break()
    doc.add_paragraph("")
    
    # 5. CONTENIDO
    lines = md_text.split('\n')
    is_first_para = False
    
    for line in lines:
        raw_text = line.strip()
        if not raw_text: continue
        
        header_match = re.match(r'^(#+)\s*(.*)$', raw_text)
        if header_match:
            hashes, content = header_match.group(1), header_match.group(2).strip()
            level = len(hashes)
            
            if level == 1:
                # Nueva sección en página IMPAR para capítulos
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_layout(new_sect, size_option)
                p = doc.add_paragraph(style='Heading 1')
            else:
                p = doc.add_paragraph(style=f'Heading {min(level, 3)}')
            
            add_formatted_text(p, content)
            is_first_para = True
        else:
            p = doc.add_paragraph(style='Body Text')
            if is_first_para:
                p.paragraph_format.first_line_indent = 0
            add_formatted_text(p, " ".join(raw_text.split()))
            is_first_para = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# UI DE STREAMLIT
# -------------------------------

st.title("📚 Generador Editorial Profesional")
st.markdown("Crea documentos Word con secciones Recto/Verso, numeración y estilos nativos.")

with st.sidebar:
    st.header("Metadatos")
    m_title = st.text_input("Título", "Título del Libro")
    m_sub = st.text_input("Subtítulo", "")
    m_author = st.text_input("Autor", "Nombre del Autor")
    m_pub = st.text_input("Editorial", "Mi Editorial")
    m_year = st.text_input("Año", str(datetime.now().year))
    m_file = st.text_input("Nombre de archivo", "libro_maquetado")
    size_mode = st.selectbox("Formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])
    m_copy = st.text_area("Copyright", f"© {datetime.now().year} {m_author}. Todos los derechos reservados.")

content = st.text_area("Contenido Markdown", height=400, placeholder="# Capítulo 1\nTexto del capítulo...")

if st.button("🚀 Generar y Descargar", use_container_width=True):
    if content:
        meta = {'title': m_title, 'subtitle': m_sub, 'author': m_author, 
                'publisher': m_pub, 'year': m_year, 'copyright': m_copy}
        result = run_book_conversion(content, meta, size_mode)
        st.success("Documento generado. Al abrirlo, actualice el índice.")
        st.download_button("📥 Descargar .docx", result, f"{m_file}.docx")
    else:
        st.error("Introduce contenido Markdown.")
