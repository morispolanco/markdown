import streamlit as st
import json
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. UTILIDADES DE BAJO NIVEL (XML)
# ==========================================

def set_mirror_margins(section):
    """Activa los márgenes simétricos en el XML del documento para impresión."""
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        mirror_margins = OxmlElement('w:mirrorMargins')
        sectPr.insert(sectPr.index(cols[0]), mirror_margins)

def add_page_number(footer):
    """Inserta un campo de número de página centrado en el pie de página."""
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.clear()
    
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)

    run2 = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    run2._r.append(instrText)

    run3 = paragraph.add_run()
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run3._r.append(fldChar2)

def setup_headers(section, author, title):
    """Configura encabezados: Autor en pares, Título en impares. Times New Roman 9pt Versalita."""
    # Encabezado páginas IMPARES -> TÍTULO DEL LIBRO
    header_odd = section.header
    p_odd = header_odd.paragraphs[0] if header_odd.paragraphs else header_odd.add_paragraph()
    p_odd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_odd.clear()
    run_odd = p_odd.add_run(title)
    run_odd.font.name = 'Times New Roman'
    run_odd.font.size = Pt(9)
    run_odd.font.small_caps = True

    # Encabezado páginas PARES -> NOMBRE DEL AUTOR
    header_even = section.even_page_header
    p_even = header_even.paragraphs[0] if header_even.paragraphs else header_even.add_paragraph()
    p_even.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_even.clear()
    run_even = p_even.add_run(author)
    run_even.font.name = 'Times New Roman'
    run_even.font.size = Pt(9)
    run_even.font.small_caps = True

def add_formatted_text(paragraph, text):
    """Maneja negritas y cursivas básicas de Markdown (** y *)."""
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
    """Establece la tipografía y espaciados basados en la plantilla 5.25x8."""
    styles = doc.styles
    
    # Estilo Normal (Garamond para el cuerpo)
    style_normal = styles['Normal']
    style_normal.font.name = 'Garamond'
    style_normal.font.size = Pt(11)
    style_normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    style_normal.paragraph_format.line_spacing = 1.15

    # Estilo Body Text (Con sangría de primera línea)
    if 'Body Text' not in styles: styles.add_style('Body Text', 1)
    body = styles['Body Text']
    body.font.name, body.font.size = 'Garamond', Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    body.paragraph_format.first_line_indent = Inches(0.25)
    body.paragraph_format.space_after = Pt(0)

    # Estilo First Paragraph (Sin sangría)
    if 'First Paragraph' not in styles: styles.add_style('First Paragraph', 1)
    first_p = styles['First Paragraph']
    first_p.font.name, first_p.font.size = 'Garamond', Pt(11)
    first_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    first_p.paragraph_format.first_line_indent = 0

    # Estilo Heading 1 (Capítulos - Aptos/Sans)
    h1 = styles['Heading 1'] if 'Heading 1' in styles else styles.add_style('Heading 1', 1)
    h1.font.name = 'Aptos'
    h1.font.size = Pt(18)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.space_before = Pt(72)
    h1.paragraph_format.space_after = Pt(36)
    h1.paragraph_format.keep_with_next = True

    return doc

def apply_layout(section, size_option, meta=None):
    """Aplica el tamaño de página y los márgenes de impresión."""
    if size_option == "Pocket (5.25 x 8 in)":
        section.page_width, section.page_height = Inches(5.25), Inches(8.0)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)  # Interior para el lomo
        section.right_margin = Inches(0.5)  # Exterior
    else:
        section.page_width, section.page_height = Inches(5.5), Inches(8.5)
        section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.75)
        section.left_margin, section.right_margin = Inches(0.75), Inches(0.6)
    
    set_mirror_margins(section)
    if meta:
        setup_headers(section, meta['author'], meta['title'])

# ==========================================
# 3. MOTOR DE CONVERSIÓN
# ==========================================

def run_book_conversion(md_text, meta, size_option):
    doc = Document()
    setup_styles(doc)
    
    # Activar encabezados diferentes para pares e impares
    doc.settings.odd_and_even_pages_header_footer = True
    
    # --- PÁGINA 1: PORTADA INTERNA ---
    section = doc.sections[0]
    apply_layout(section, size_option)
    section.different_first_page_header_footer = True 
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(120)
    run_t = p_title.add_run(meta['title'].upper())
    run_t.font.size = Pt(28)
    run_t.bold = True

    if meta.get('subtitle'):
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_s = p_sub.add_run(meta['subtitle'])
        run_s.font.size = Pt(14)
        run_s.italic = True

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Inches(2.0)

    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_a = p_author.add_run(meta['author'])
    run_a.font.size = Pt(16)

    # --- PÁGINA 2: COPYRIGHT ---
    doc.add_page_break()
    p_copy = doc.add_paragraph()
    p_copy.paragraph_format.space_before = Inches(4.5)
    p_copy.alignment = WD_ALIGN_PARAGRAPH.LEFT
    copyright_text = f"{meta['title']}\nCopyright © {meta['year']} {meta['author']}\nAll rights reserved.\n\nISBN: {meta.get('isbn', '________________')}"
    run_c = p_copy.add_run(copyright_text)
    run_c.font.size = Pt(9)
    
    # --- PÁGINA 3: DEDICATORIA ---
    doc.add_page_break()
    p_ded = doc.add_paragraph()
    p_ded.paragraph_format.space_before = Inches(2.0)
    p_ded.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_d = p_ded.add_run(meta.get('dedication', "Insert dedication text here."))
    run_d.italic = True

    # --- CUERPO DEL LIBRO ---
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
                apply_layout(new_sect, size_option, meta)
                add_page_number(new_sect.footer)
                new_sect.different_first_page_header_footer = True 
                
                p = doc.add_paragraph(style='Heading 1')
                add_formatted_text(p, title_text.upper())
            else:
                p = doc.add_paragraph(style=f'Heading {min(level, 5)}')
                add_formatted_text(p, title_text)
            is_after_heading = True
        else:
            style = 'First Paragraph' if is_after_heading else 'Body Text'
            p = doc.add_paragraph(style=style)
            add_formatted_text(p, clean)
            is_after_heading = False

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ==========================================
# 4. INTERFAZ STREAMLIT
# ==========================================

st.set_page_config(page_title="Maquetador Editorial Pro", layout="centered")

if 'book_data' not in st.session_state:
    st.session_state.book_data = {
        'title': "Mi Gran Novela",
        'author': "Nombre del Autor",
        'year': "2025",
        'subtitle': "",
        'isbn': "",
        'dedication': "Para aquellos que creen en la magia de las palabras.",
        'manuscript': ""
    }

st.title("📖 Maquetador Editorial Profesional")

with st.sidebar:
    st.header("📂 Importación de Archivos")
    
    # 1. Carga de metadatos JSON
    uploaded_json = st.file_uploader("1. Cargar ficha (JSON)", type=["json"])
    if uploaded_json is not None:
        try:
            data = json.load(uploaded_json)
            mapping = {
                'titulo': 'title', 'title': 'title',
                'autor': 'author', 'author': 'author',
                'año': 'year', 'year': 'year',
                'subtitulo': 'subtitle', 'subtitle': 'subtitle',
                'isbn': 'isbn',
                'dedicatoria': 'dedication', 'dedication': 'dedication'
            }
            for key, val in data.items():
                if key.lower() in mapping:
                    st.session_state.book_data[mapping[key.lower()]] = str(val)
            st.success("¡Metadatos cargados!")
        except:
            st.error("Error al leer el JSON.")

    st.divider()

    # 2. Carga del manuscrito Markdown
    uploaded_md = st.file_uploader("2. Cargar Manuscrito (.md, .txt)", type=["md", "txt"])
    if uploaded_md is not None:
        try:
            content = uploaded_md.read().decode("utf-8")
            st.session_state.book_data['manuscript'] = content
            st.success("¡Manuscrito cargado!")
        except:
            st.error("Error al leer el archivo de texto.")

with st.expander("📝 Formulario de Metadatos", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        m_title = st.text_input("Título del Libro", value=st.session_state.book_data['title'])
        m_author = st.text_input("Nombre del Autor", value=st.session_state.book_data['author'])
    with col2:
        m_year = st.text_input("Año", value=st.session_state.book_data['year'])
        size_mode = st.selectbox("Formato", ["Pocket (5.25 x 8 in)", "Trade (5.5 x 8.5 in)"])

    m_sub = st.text_input("Subtítulo", value=st.session_state.book_data['subtitle'])
    m_isbn = st.text_input("ISBN", value=st.session_state.book_data['isbn'])
    m_dedication = st.text_area("Dedicatoria", value=st.session_state.book_data['dedication'])

st.subheader("🖋️ Editor del Manuscrito")
# El editor se sincroniza con el archivo subido pero permite edición manual
editor_content = st.text_area(
    "Contenido (puedes editarlo aquí directamente)", 
    value=st.session_state.book_data['manuscript'], 
    height=400, 
    placeholder="# CAPÍTULO 1\nEscribe o sube tu archivo..."
)

if st.button("🚀 Generar Libro Profesional", use_container_width=True):
    if editor_content and m_title:
        with st.spinner("Procesando maquetación..."):
            bundle = {
                'title': m_title, 'subtitle': m_sub, 'author': m_author, 
                'year': m_year, 'isbn': m_isbn, 'dedication': m_dedication
            }
            docx_bytes = run_book_conversion(editor_content, bundle, size_mode)
            
            st.success("¡Libro generado correctamente!")
            st.download_button(
                label=f"📥 Descargar {m_title}.docx",
                data=docx_bytes,
                file_name=f"{m_title.replace(' ', '_')}_Final.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Es necesario un título y contenido para generar el libro.")
