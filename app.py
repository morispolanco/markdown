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
    
    # Estilo Normal (Base para el texto y el título según solicitud)
    style_normal = styles['Normal']
    style_normal.font.name = 'Garamond'
    style_normal.font.size = Pt(11)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Estilo Body Text (Para párrafos del libro)
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.line_spacing = 1.15
    body.paragraph_format.first_line_indent = Inches(0.25)

    # Estilos de Títulos (Heading 1 para capítulos, etc.)
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
    
    # Título (Estilo Normal, 36pt, centrado)
    # Usamos style=None para heredar del estilo base "Normal"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(24)
    run_t = p_title.add_run(meta['title'])
    run_t.font.size = Pt(36)
    run_t.bold = False # Estilo normal, no negrita a menos que se desee

    # Subtítulo
    if meta.get('subtitle'):
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_s = p_sub.add_run(meta['subtitle'])
        run_s.font.size = Pt(18)
        run_s.italic = True

    # Espaciado controlado para asegurar que todo quepa en la pág 1
    # Usamos espaciado de párrafo en lugar de muchas líneas vacías para mayor estabilidad
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Inches(2.5)

    # Autor
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.name = 'Garamond'
    run_a.font.size = Pt(20)
    run_a.italic = True

    # Editorial al fondo de la página 1
    p_spacer_pub = doc.add_paragraph()
    p_spacer_pub.paragraph_format.space_before = Inches(1.5)
    
    p_pub = doc.add_paragraph()
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_pub.add_run(meta['publisher'])
    run_p.font.name = 'Aptos'
    run_p.font.size = Pt(14)
    run_p.bold = True

    # --- PÁGINA 2: COPYRIGHT ---
    doc.add_page_break()
    # Forzar que el copyright esté en la parte inferior de la página 2
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(6.0)
    p_copy.alignment = WD_ALIGN_PARAGRAPH.LEFT
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
    doc.add_paragraph() 

    # --- CAPÍTULOS ---
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
                # Los capítulos nuevos empiezan en página impar
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

st.set_page_config(page_title="Maquetador Editorial Pro", layout="wide")
st.title("📚 Generador Editorial Profesional")

if 'meta' not in st.session_state:
    st.session_state.meta = {'title': "Título del Libro", 'subtitle': "Subtítulo de la obra", 'author': "Nombre del Autor", 'publisher': "Nombre de la Editorial", 'year': "2026", 'copyright': ""}

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
                'copyright': data.get('copyright', "")
            })
            if not st.session_state.meta['copyright']:
                st.session_state.meta['copyright'] = f"© {st.session_state.meta['year']} {st.session_state.meta['author']}. Todos los derechos reservados."
            st.success("¡Datos cargados!")
        except:
            st.error("Error al leer el JSON.")

    st.divider()
    m_title = st.text_input("Título", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.meta['publisher'])
    m_copy = st.text_area("Aviso de Copyright", value=st.session_state.meta['copyright'], height=100)
    size_mode = st.selectbox("Tamaño del formato", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

col1, col2 = st.columns([1, 1])
with col1:
    md_file = st.file_uploader("Subir Manuscrito (.md)", type=["md", "txt"])
    text_input = md_file.read().decode("utf-8") if md_file else ""
with col2:
    editor = st.text_area("Vista previa / Edición manual", value=text_input, height=300)

if st.button("🚀 Generar y Descargar Documento", use_container_width=True):
    if editor and m_title:
        bundle = {'title': m_title, 'subtitle': m_sub, 'author': m_author, 'publisher': m_pub, 'copyright': m_copy}
        docx_bytes = run_book_conversion(editor, bundle, size_mode)
        st.download_button(
            label=f"📥 Guardar {m_title}.docx",
            data=docx_bytes,
            file_name=f"{slugify(m_title)}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Por favor, asegúrate de tener un título y contenido en el manuscrito.")
