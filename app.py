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
# 2. CONFIGURACIÓN DE ESTILOS Y MAQUETACIÓN
# ==========================================

def setup_styles(doc):
    styles = doc.styles
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.first_line_indent = Inches(0.25)
    body.paragraph_format.line_spacing = 1.15

    h1 = styles['Heading 1']
    h1.font.name, h1.font.size = 'Aptos', Pt(24)
    h1.font.bold = True
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(0) 
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
    section.different_first_page_header_footer = is_first
    if not is_first:
        add_page_number(section.footer)

# ==========================================
# 3. MOTOR DE CONVERSIÓN
# ==========================================

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # --- PÁGINA 1: PORTADA ---
    section = doc.sections[0]
    apply_layout(section, size_option, is_first=True)
    
    # Título (Cero espacio antes)
    p_title = doc.add_paragraph(meta['title'], style='Heading 1')
    p_title.runs[0].font.size = Pt(34)
    
    if meta.get('subtitle'):
        p_sub = doc.add_paragraph(meta['subtitle'])
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.add_run().font.size = Pt(18)

    # Espacio para empujar autor y editorial
    for _ in range(12): doc.add_paragraph()

    # Autor
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.size = Pt(18)
    run_a.italic = True

    # Editorial (Debajo del autor)
    p_pub = doc.add_paragraph()
    p_pub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_p = p_pub.add_run(meta['publisher'])
    run_p.font.size = Pt(14)
    run_p.bold = True

    # --- PÁGINA 2: CRÉDITOS ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(5.0)
    # Insertar el texto completo de copyright
    run_c = p_copy.add_run(meta['copyright'])
    run_c.font.size = Pt(9)
    run_c.font.name = 'Garamond'
    
    # --- CONTENIDO ---
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
# 4. INTERFAZ STREAMLIT
# ==========================================

st.set_page_config(page_title="Maquetador Pro")
st.title("📚 Generador Editorial")

# Estado inicial para evitar pérdidas de datos al recargar
if 'meta' not in st.session_state:
    st.session_state.meta = {
        'title': "Título del Libro", 
        'subtitle': "", 
        'author': "Autor", 
        'publisher': "Editorial", 
        'year': "2026", 
        'copyright': ""
    }

with st.sidebar:
    st.header("1. Cargar Ficha JSON")
    json_up = st.file_uploader("Subir archivo .json", type=["json"])
    
    if json_up:
        try:
            data = json.load(json_up)
            st.session_state.meta.update({
                'title': data.get('titulo', data.get('title', st.session_state.meta['title'])),
                'subtitle': data.get('subtitulo', data.get('subtitle', "")),
                'author': data.get('autor', data.get('author', st.session_state.meta['author'])),
                'publisher': data.get('editorial', data.get('publisher', st.session_state.meta['publisher'])),
                'year': str(data.get('año', data.get('year', "2026"))),
                'copyright': data.get('copyright', "")
            })
            # Si el JSON no trae copyright, lo generamos por defecto
            if not st.session_state.meta['copyright']:
                st.session_state.meta['copyright'] = f"© {st.session_state.meta['year']} {st.session_state.meta['author']}. Todos los derechos reservados."
            st.success("Metadatos cargados")
        except:
            st.error("Error al procesar el JSON")

    st.divider()
    st.header("2. Confirmar Datos")
    m_title = st.text_input("Título", value=st.session_state.meta['title'])
    m_sub = st.text_input("Subtítulo", value=st.session_state.meta['subtitle'])
    m_author = st.text_input("Autor", value=st.session_state.meta['author'])
    m_pub = st.text_input("Editorial", value=st.session_state.meta['publisher'])
    # Este es el campo que ahora capturará el mensaje completo
    m_copy = st.text_area("Texto de Copyright", value=st.session_state.meta['copyright'], height=100)
    size_mode = st.selectbox("Formato de impresión", ["Trade Paperback (5.5x8.5)", "Estándar A4"])

st.header("3. Manuscrito")
md_file = st.file_uploader("Sube el archivo Markdown (.md)", type=["md", "txt"])
text_input = md_file.read().decode("utf-8") if md_file else ""
editor = st.text_area("Contenido del Manuscrito", value=text_input, height=300)

if st.button("🚀 Generar y Descargar Libro", use_container_width=True):
    if editor and m_title:
        # Empaquetar datos finales desde los inputs de la UI
        final_meta = {
            'title': m_title, 
            'subtitle': m_sub, 
            'author': m_author, 
            'publisher': m_pub, 
            'copyright': m_copy # Aquí se envía el texto completo del área de texto
        }
        
        docx_out = run_book_conversion(editor, final_meta, size_mode)
        nombre_file = f"{limpiar_nombre(m_title)}.docx"
        
        st.success(f"Libro '{nombre_file}' generado.")
        st.download_button("📥 Descargar Archivo DOCX", docx_out, nombre_file)
    else:
        st.warning("Asegúrate de tener un título y contenido en el manuscrito.")
