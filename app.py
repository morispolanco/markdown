import streamlit as st
import json
import re
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
# 1. UTILIDADES DE BAJO NIVEL
# ==========================================

def set_mirror_margins(section):
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_page_number(footer):
    """Añade número de página al footer (se ocultará en portada por configuración de sección)."""
    if not footer.paragraphs:
        footer.add_paragraph()
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)
    run = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    run._r.append(instrText)
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)

def add_formatted_text(paragraph, text):
    parts = re.split(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*)', text)
    for part in parts:
        if not part: continue
        is_bold, is_italic = False, False
        clean_part = part
        if part.startswith('***') and part.endswith('***'):
            is_bold = is_italic = True
            clean_part = part[3:-3]
        elif part.startswith('**') and part.endswith('**'):
            is_bold = True
            clean_part = part[2:-2]
        elif part.startswith('*') and part.endswith('*'):
            is_italic = True
            clean_part = part[1:-1]
        run = paragraph.add_run(clean_part)
        run.bold, run.italic = is_bold, is_italic

# ==========================================
# 2. CONFIGURACIÓN DE ESTILOS Y MOTOR
# ==========================================

def setup_styles(doc):
    styles = doc.styles
    # Cuerpo
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Título 1 (Capítulos)
    h1 = styles['Heading 1']
    h1.font.name, h1.font.size = 'Aptos', Pt(24)
    h1.font.bold = True
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(0) # Cero espacio antes solicitado
    h1.paragraph_format.space_after = Pt(24)
    return doc

def apply_layout(section, size_option, is_first=False):
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.875)
        section.left_margin, section.right_margin = Inches(0.875), Inches(0.75)
    else:
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1.25), Inches(1)
    
    set_mirror_margins(section)
    
    # Configuración para que la primera página no tenga número
    if is_first:
        section.different_first_page_header_footer = True
        # El footer de la primera página se deja vacío automáticamente al no llamarlo
    else:
        add_page_number(section.footer)

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # --- 1. PORTADA ---
    section = doc.sections[0]
    apply_layout(section, size_option, is_first=True)
    
    # Título (Sin espacios antes)
    p_title = doc.add_paragraph(meta['title'], style='Heading 1')
    p_title.runs[0].font.size = Pt(36)
    
    if meta.get('subtitle'):
        p_sub = doc.add_paragraph(meta['subtitle'])
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.runs[0].font.size = Pt(18)

    # Espacio flexible para empujar autor y editorial hacia abajo
    for _ in range(10): doc.add_paragraph()

    # Autor
    p_author = doc.add_paragraph(meta['author'])
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.runs[0].font.size = Pt(18)
    p_author.runs[0].italic = True

    # Editorial (Debajo del autor en la primera página)
    p_pub = doc.add_paragraph(meta['publisher'])
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_pub.runs[0].font.size = Pt(14)
    p_pub.runs[0].bold = True

    # --- 2. CRÉDITOS ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.0)
    p_copy.add_run(meta['copyright']).font.size = Pt(9)
    
    # --- 3. CONTENIDO ---
    is_after_heading = False
    for line in md_text.split('\n'):
        clean = line.strip()
        if not clean: continue
        
        header_match = re.match(r'^(#+)\s*(.*)$', clean)
        if header_match:
            level = len(header_match.group(1))
            if level == 1:
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_layout(new_sect, size_option, is_first=False)
                p = doc.add_paragraph(style='Heading 1')
            else:
                p = doc.add_paragraph(style=f'Heading {min(level, 5)}')
            add_formatted_text(p, header_match.group(2))
            is_after_heading = True
        else:
            p = doc.add_paragraph(style='Body Text')
            if is_after_heading:
                p.paragraph_format.first_line_indent = 0
                is_after_heading = False
            add_formatted_text(p, clean)

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def limpiar_nombre(t):
    return re.sub(r'[^a-z0-9]', '_', t.lower().replace(' ','_'))

# ==========================================
# 3. INTERFAZ STREAMLIT
# ==========================================

st.set_page_config(page_title="Maquetador Pro")
st.title("📚 Generador Editorial")

if 'meta' not in st.session_state:
    st.session_state.meta = {'title': "", 'subtitle': "", 'author': "", 'publisher': "", 'year': "2026", 'copyright': ""}

with st.sidebar:
    st.header("Ficha JSON")
    json_up = st.file_uploader("Subir Ficha", type=["json"])
    if json_up:
        data = json.load(json_up)
        st.session_state.meta.update({
            'title': data.get('titulo', data.get('title', "")),
            'subtitle': data.get('subtitulo', data.get('subtitle', "")),
            'author': data.get('autor', data.get('author', "")),
            'publisher': data.get('editorial', data.get('publisher', "")),
            'year': str(data.get('año', data.get('year', "2026"))),
        })
        st.session_state.meta['copyright'] = f"© {st.session_state.meta['year']} {st.session_state.meta['author']}"

    m_title = st.text_input("Título", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.meta['publisher'])
    m_copy = st.text_area("Copyright", value=st.session_state.meta['copyright'])
    size_mode = st.selectbox("Formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

md_file = st.file_uploader("Manuscrito Markdown", type=["md", "txt"])
text_input = md_file.read().decode("utf-8") if md_file else ""
editor = st.text_area("Contenido", value=text_input, height=300)

if st.button("Generar Libro", use_container_width=True):
    if editor and m_title:
        bundle = {'title': m_title, 'subtitle': m_sub, 'author': m_author, 'publisher': m_pub, 'copyright': m_copy}
        out = run_book_conversion(editor, bundle, size_mode)
        st.download_button(f"📥 Descargar {m_title}", out, f"{limpiar_nombre(m_title)}.docx")
    else:
        st.warning("Título y contenido son obligatorios.")
