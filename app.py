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
    style_normal = styles['Normal']
    style_normal.font.name = 'Garamond'
    style_normal.font.size = Pt(11)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style_normal.paragraph_format.line_spacing = 1.0

    # Estilo Body Text
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.15
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Estilos de Títulos
    for i in range(1, 6):
        style_name = f'Heading {i}'
        h = styles[style_name] if style_name in styles else styles.add_style(style_name, 1)
        h.font.name = 'Aptos'
        h.font.bold = True
        h.font.color.rgb = RGBColor(0, 0, 0)
        
        if i == 1:
            h.font.size = Pt(24)
            h.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_before = Pt(24)
            h.paragraph_format.space_after = Pt(48)
            h.paragraph_format.keep_with_next = True
        else:
            h.font.size = Pt(16 - i)
            h.paragraph_format.space_before = Pt(18)
            h.paragraph_format.space_after = Pt(12)
    return doc

def apply_layout(section, size_option, has_number=True):
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.75)
        section.left_margin, section.right_margin = Inches(0.75), Inches(0.65)
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
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(40)
    run_t = p_title.add_run(meta['title'])
    run_t.font.size = Pt(36)
    run_t.bold = False 

    if meta.get('subtitle'):
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_s = p_sub.add_run(meta['subtitle'])
        run_s.font.size = Pt(16)
        run_s.italic = True

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Inches(1.8)

    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.name = 'Garamond'
    run_a.font.size = Pt(18)
    run_a.italic = True

    p_spacer_pub = doc.add_paragraph()
    p_spacer_pub.paragraph_format.space_before = Inches(1.2)
    
    p_pub = doc.add_paragraph()
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_pub.add_run(meta['publisher'])
    run_p.font.name = 'Aptos'
    run_p.font.size = Pt(12)
    run_p.bold = True

    # --- PÁGINA 2: COPYRIGHT ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(2.0)
    p_copy.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.size = Pt(9)
    run_c.font.name = 'Garamond'
    
    # --- PÁGINAS DE CONTENIDO ---
    doc.add_page_break()
    lines = md_text.split('\n')
    is_after_heading = False
    
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
                if ":" in title_text:
                    parts = title_text.split(":", 1)
                    add_formatted_text(p, parts[0].strip() + ":")
                    p.add_run().add_break()
                    add_formatted_text(p, parts[1].strip())
                else:
                    add_formatted_text(p, title_text)
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

    # --- ÚLTIMA PÁGINA: BIOGRAFÍA DEL AUTOR ---
    if meta.get('author_bio'):
        # Forzar inicio en página nueva
        doc.add_page_break()
        # Opcionalmente, puedes querer que empiece en página impar como los capítulos:
        # doc.add_section(WD_SECTION_START.ODD_PAGE)
        
        p_bio_title = doc.add_paragraph("Biografía del Autor", style='Heading 1')
        p_bio_content = doc.add_paragraph(style='Body Text')
        p_bio_content.paragraph_format.first_line_indent = 0
        add_formatted_text(p_bio_content, meta['author_bio'])

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 4. INTERFAZ STREAMLIT
# ==========================================

def slugify(text):
    return re.sub(r'[^a-z0-9]', '_', text.lower().replace(' ','_'))

st.set_page_config(page_title="Maquetador Editorial Pro", layout="wide")
st.title("📚 Generador Editorial")

# Inicialización de metadatos en sesión
if 'meta' not in st.session_state:
    st.session_state.meta = {
        'title': "Título", 
        'subtitle': "", 
        'author': "Autor", 
        'publisher': "Editorial", 
        'year': "2026", 
        'copyright': "",
        'author_bio': "" # Espacio para la biografía
    }

with st.sidebar:
    st.header("⚙️ Configuración")
    json_up = st.file_uploader("Cargar ficha JSON", type=["json"])
    if json_up:
        try:
            data = json.load(json_up)
            st.session_state.meta.update({
                'title': data.get('titulo', data.get('title', st.session_state.meta['title'])),
                'subtitle': data.get('subtitulo', data.get('subtitle', st.session_state.meta['subtitle'])),
                'author': data.get('autor', data.get('author', st.session_state.meta['author'])),
                'publisher': data.get('editorial', data.get('publisher', st.session_state.meta['publisher'])),
                'year': str(data.get('año', data.get('year', "2026"))),
                'copyright': data.get('copyright', ""),
                'author_bio': data.get('biografia', data.get('author_bio', ""))
            })
            st.success("JSON cargado")
        except:
            st.error("Error en JSON")

    st.divider()
    m_title = st.text_input("Título", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.meta['publisher'])
    m_copy = st.text_area("Copyright", value=st.session_state.meta['copyright'], height=80)
    
    # NUEVO CAMPO PARA BIOGRAFÍA
    m_bio = st.text_area("Biografía del autor", value=st.session_state.meta['author_bio'], height=150, help="Aparecerá en la última página del libro.")
    
    size_mode = st.selectbox("Formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

md_file = st.file_uploader("Manuscrito (.md)", type=["md", "txt"])
text_input = md_file.read().decode("utf-8") if md_file else ""
editor = st.text_area("Contenido", value=text_input, height=300)

if st.button("🚀 Generar Libro", use_container_width=True):
    if editor and m_title:
        bundle = {
            'title': m_title, 
            'subtitle': m_sub, 
            'author': m_author, 
            'publisher': m_pub, 
            'copyright': m_copy,
            'author_bio': m_bio # Pasar biografía al motor
        }
        docx_bytes = run_book_conversion(editor, bundle, size_mode)
        st.download_button(f"📥 Guardar {m_title}.docx", docx_bytes, f"{slugify(m_title)}.docx")
    else:
        st.warning("Faltan datos obligatorios.")
