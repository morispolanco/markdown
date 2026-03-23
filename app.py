import streamlit as st
import json
import re
import os
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
# 1. UTILIDADES DE BAJO NIVEL (XML / DOCX)
# ==========================================

def set_mirror_margins(section):
    """Habilita márgenes simétricos para impresión (encuadernación)."""
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_page_number(footer):
    """Inserta numeración automática en el centro del pie de página."""
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

def add_toc_field(paragraph):
    """Inserta el código de campo para la Tabla de Contenidos."""
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar)

    run = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-5" \\h \\z \\u'
    run._r.append(instrText)

    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'separate')
    run._r.append(fldChar)

    run = paragraph.add_run("Actualice este campo en Word para ver el índice.")
    
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar)

def add_formatted_text(paragraph, text):
    """Analiza negritas y cursivas de Markdown básico."""
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
# 2. CONFIGURACIÓN DE MAQUETACIÓN Y ESTILOS
# ==========================================

def apply_layout(section, size_option):
    """Aplica medidas físicas y márgenes de libro."""
    if size_option == "Trade Paperback (5.5x8.5)":
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.875)
        section.left_margin = Inches(0.875) # Lado interior (Gutter)
        section.right_margin = Inches(0.75)
    else:
        section.page_width, section.page_height = Inches(8.27), Inches(11.69)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1)
    
    set_mirror_margins(section)
    add_page_number(section.footer)

def setup_styles(doc):
    """Define la tipografía y jerarquía visual."""
    styles = doc.styles
    
    # Cuerpo del texto (Garamond)
    if 'Body Text' not in styles: styles.add_style('Body Text', WD_STYLE_TYPE.PARAGRAPH)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.15
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Título 1 (Capítulos)
    h1 = styles['Heading 1']
    h1.font.name, h1.font.size = 'Aptos', Pt(20)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Inches(1.5)
    h1.paragraph_format.space_after = Inches(0.5)
    
    return doc

# ==========================================
# 3. MOTOR DE CONVERSIÓN EDITORIAL
# ==========================================

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # --- PORTADA ---
    section = doc.sections[0]
    apply_layout(section, size_option)
    
    for _ in range(6): doc.add_paragraph()
    p_title = doc.add_paragraph(meta['title'], style='Heading 1')
    p_title.runs[0].font.size = Pt(32)
    
    if meta['subtitle']:
        p_sub = doc.add_paragraph(meta['subtitle'])
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.add_run().font.size = Pt(18)
    
    for _ in range(8): doc.add_paragraph()
    p_author = doc.add_paragraph(meta['author'])
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.runs[0].font.italic = True
    p_author.runs[0].font.size = Pt(16)

    # --- CRÉDITOS ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.0)
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.size = Pt(9)
    
    # --- ÍNDICE ---
    doc.add_page_break()
    doc.add_paragraph("Índice de Contenidos", style='Heading 1')
    toc_p = doc.add_paragraph()
    add_toc_field(toc_p)
    
    # --- CONTENIDO ---
    is_after_heading = False
    lines = md_text.split('\n')
    
    for line in lines:
        clean_line = line.strip()
        if not clean_line: continue
        
        header_match = re.match(r'^(#+)\s*(.*)$', clean_line)
        if header_match:
            level = len(header_match.group(1))
            title_text = header_match.group(2)
            
            if level == 1:
                # Nuevo capítulo en página IMPAR
                new_sect = doc.add_section(WD_SECTION_START.ODD_PAGE)
                apply_layout(new_sect, size_option)
                p = doc.add_paragraph(style='Heading 1')
            else:
                p = doc.add_paragraph(style=f'Heading {min(level, 5)}')
            
            add_formatted_text(p, title_text)
            is_after_heading = True
        else:
            p = doc.add_paragraph(style='Body Text')
            if is_after_heading:
                p.paragraph_format.first_line_indent = 0 # Sin sangría tras título
                is_after_heading = False
            add_formatted_text(p, clean_line)

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def limpiar_nombre_archivo(titulo):
    """Genera un nombre de archivo seguro a partir del título."""
    s = titulo.lower()
    s = re.sub(r'[áéíóúüñ]', lambda m: {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ü':'u','ñ':'n'}[m.group()], s)
    s = re.sub(r'[^a-z0-9]', '_', s)
    return re.sub(r'_+', '_', s).strip('_')

# ==========================================
# 4. INTERFAZ DE USUARIO (STREAMLIT)
# ==========================================

st.set_page_config(page_title="Maquetador Editorial", layout="centered")
st.title("📚 Generador Editorial Pro")

# Inicialización de estado
if 'meta' not in st.session_state:
    st.session_state.meta = {
        'title': "Título del Libro", 'subtitle': "", 'author': "Autor",
        'year': str(datetime.now().year), 'copyright': ""
    }

with st.sidebar:
    st.header("1. Metadatos (JSON)")
    json_upload = st.file_uploader("Ficha Editorial", type=["json"])
    
    if json_upload:
        try:
            data = json.load(json_upload)
            st.session_state.meta.update({
                'title': data.get('titulo', data.get('title', st.session_state.meta['title'])),
                'subtitle': data.get('subtitulo', data.get('subtitle', "")),
                'author': data.get('autor', data.get('author', st.session_state.meta['author'])),
                'year': str(data.get('año', data.get('year', datetime.now().year))),
                'copyright': data.get('copyright', "")
            })
            if not st.session_state.meta['copyright']:
                st.session_state.meta['copyright'] = f"© {st.session_state.meta['year']} {st.session_state.meta['author']}"
            st.success("JSON cargado")
        except:
            st.error("Archivo JSON inválido")

    st.divider()
    m_title = st.text_input("Título Final", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_copy = st.text_area("Copyright", value=st.session_state.meta['copyright'])
    size_mode = st.selectbox("Formato Impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

st.header("2. Manuscrito")
md_file = st.file_uploader("Subir Markdown (.md)", type=["md", "txt"])
md_content = md_file.read().decode("utf-8") if md_file else ""
final_text = st.text_area("Editor de Contenido", value=md_content, height=300)

if st.button("🚀 Procesar Libro Completo", use_container_width=True):
    if final_text:
        meta_bundle = {
            'title': m_title, 'subtitle': m_sub, 
            'author': m_author, 'copyright': m_copy
        }
        with st.spinner("Maquetando documento..."):
            docx_bytes = run_book_conversion(final_text, meta_bundle, size_mode)
            nombre_archivo = f"{limpiar_nombre_archivo(m_title)}.docx"
            
            st.success("¡Maquetación terminada!")
            st.download_button(
                label=f"📥 Descargar {nombre_archivo}",
                data=docx_bytes,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Por favor, cargue el contenido del manuscrito.")
