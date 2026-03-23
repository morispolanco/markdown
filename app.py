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
# 1. UTILIDADES DE BAJO NIVEL (XML)
# ==========================================

def set_mirror_margins(section):
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_page_number(footer):
    """Inserta número de página centrado en el footer."""
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.clear()
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    run2 = paragraph.add_run()
    instrText = OxmlElement('w:instrText'); instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    run2._r.append(instrText)
    run3 = paragraph.add_run()
    fldChar2 = OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar2)

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
# 2. CONFIGURACIÓN EDITORIAL Y ESTILOS
# ==========================================

def setup_styles(doc):
    styles = doc.styles
    # Estilo Normal
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.15
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Estilos de Títulos (Color automático/Negro)
    for i in range(1, 6):
        style_name = f'Heading {i}'
        h = styles[style_name] if style_name in styles else styles.add_style(style_name, 1)
        h.font.name = 'Aptos'
        h.font.bold = True
        h.font.color.rgb = RGBColor(0, 0, 0)
        if i == 1:
            h.font.size = Pt(24)
            h.paragraph_format.space_before = Pt(0)
            h.paragraph_format.space_after = Pt(36)
            h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            h.font.size = Pt(16 - i)
            h.paragraph_format.space_before = Pt(18)
            h.paragraph_format.space_after = Pt(12)
    return doc

def apply_layout(section, size_option, has_number=True):
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.875)
        section.left_margin, section.right_margin = Inches(0.875), Inches(0.75)
    else:
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1.25), Inches(1)
    set_mirror_margins(section)
    if has_number:
        add_page_number(section.footer)

# ==========================================
# 3. MOTOR DE CONVERSIÓN
# ==========================================

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # --- PÁGINA 1: PORTADA ---
    section = doc.sections[0]
    apply_layout(section, size_option, has_number=False)
    section.different_first_page_header_footer = True 
    
    # Título (Cero espacio antes)
    p_title = doc.add_paragraph(meta['title'], style='Heading 1')
    p_title.runs[0].font.size = Pt(34)
    
    if meta.get('subtitle'):
        p_sub = doc.add_paragraph(meta['subtitle'])
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.add_run().font.size = Pt(18)

    # CENTRADO VERTICAL DEL AUTOR (Aprox. mitad de página)
    for _ in range(12): doc.add_paragraph()
    
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.name = 'Garamond'
    run_a.font.size = Pt(20)
    run_a.italic = True

    # EDITORIAL AL FINAL DE LA PÁGINA
    for _ in range(10): doc.add_paragraph()
    
    p_pub = doc.add_paragraph()
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_pub.add_run(meta['publisher'])
    run_p.font.name = 'Aptos'
    run_p.font.size = Pt(14)
    run_p.bold = True

    # --- PÁGINA 2: COPYRIGHT ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.5)
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.size = Pt(9)
    run_c.font.name = 'Garamond'
    
    # --- PÁGINA 3: CONTENIDO ---
    doc.add_page_break()
    p_cont = doc.add_paragraph()
    run_cont = p_cont.add_run("Contenido")
    run_cont.font.name = 'Aptos'
    run_cont.font.size = Pt(20)
    p_cont.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_cont.paragraph_format.first_line_indent = 0
    doc.add_paragraph() # Espacio para el índice manual

    # --- CAPÍTULOS (PÁGINA IMPAR) ---
    is_after_heading = False
    lines = md_text.split('\n')
    
    for line in lines:
        clean = line.strip()
        if not clean: continue
        
        header_match = re.match(r'^(#+)\s*(.*)$', clean)
        if header_match:
            level = len(header_match.group(1))
            title_text = header_match.group(2)
            
            if level == 1:
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_layout(new_sect, size_option, has_number=True)
                p = doc.add_paragraph(style='Heading 1')
            else:
                p = doc.add_paragraph(style=f'Heading {min(level, 5)}')
            
            add_formatted_text(p, title_text)
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

def slugify(text):
    return re.sub(r'[^a-z0-9]', '_', text.lower().replace(' ','_'))

# ==========================================
# 4. INTERFAZ STREAMLIT
# ==========================================

st.set_page_config(page_title="Maquetador Editorial Pro")
st.title("📚 Generador Editorial")

if 'meta' not in st.session_state:
    st.session_state.meta = {'title': "Título", 'subtitle': "", 'author': "Autor", 'publisher': "Editorial", 'year': "2026", 'copyright': ""}

with st.sidebar:
    st.header("Cargar Ficha JSON")
    json_up = st.file_uploader("Subir JSON", type=["json"])
    if json_up:
        try:
            data = json.load(json_up)
            st.session_state.meta.update({
                'title': data.get('titulo', data.get('title', "")),
                'subtitle': data.get('subtitulo', data.get('subtitle', "")),
                'author': data.get('autor', data.get('author', "")),
                'publisher': data.get('editorial', data.get('publisher', "")),
                'year': str(data.get('año', data.get('year', "2026"))),
                'copyright': data.get('copyright', "")
            })
            if not st.session_state.meta['copyright']:
                st.session_state.meta['copyright'] = f"© {st.session_state.meta['year']} {st.session_state.meta['author']}."
            st.success("JSON cargado con éxito")
        except:
            st.error("Archivo JSON inválido")

    st.divider()
    m_title = st.text_input("Título", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.meta['publisher'])
    m_copy = st.text_area("Copyright", value=st.session_state.meta['copyright'], height=100)
    size_mode = st.selectbox("Tamaño", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

md_file = st.file_uploader("Manuscrito (.md)", type=["md", "txt"])
text_input = md_file.read().decode("utf-8") if md_file else ""
editor = st.text_area("Contenido", value=text_input, height=300)

if st.button("🚀 Generar Libro", use_container_width=True):
    if editor and m_title:
        bundle = {'title': m_title, 'subtitle': m_sub, 'author': m_author, 'publisher': m_pub, 'copyright': m_copy}
        docx_bytes = run_book_conversion(editor, bundle, size_mode)
        st.download_button(f"📥 Descargar {m_title}", docx_bytes, f"{slugify(m_title)}.docx")
    else:
        st.warning("Debe proporcionar un título y contenido.")
