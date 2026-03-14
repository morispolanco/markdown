import streamlit as st
from io import BytesIO
import tempfile
import os
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Markdown → Word (Plantilla de Libro)", page_icon="📚", layout="centered")

# -------------------------------
# LÓGICA DE CAPITALIZACIÓN (NORMA ESPAÑOLA)
# -------------------------------
def corregir_capitalizacion_espanola(texto_md):
    """
    Transforma los títulos (#) y subtítulos (##, ###) a minúsculas
    excepto la primera letra y nombres propios conocidos.
    """
    # Lista básica de excepciones (puedes ampliarla)
    excepciones = ["España", "México", "Python", "Streamlit", "Pandoc", "Word", "Markdown", "Dios", "Europa"]

    def transformar_linea(match):
        hashes = match.group(1) # Los #
        contenido = match.group(2).strip()
        
        palabras = contenido.split()
        nuevas_palabras = []
        
        for i, word in enumerate(palabras):
            clean_word = re.sub(r'[^\w]', '', word) # Quitar puntos/comas para comparar
            if i == 0:
                nuevas_palabras.append(word.capitalize())
            elif clean_word in excepciones:
                nuevas_palabras.append(word)
            else:
                nuevas_palabras.append(word.lower())
        
        return f"{hashes} {' '.join(nuevas_palabras)}"

    # Busca líneas que inicien con uno o más '#'
    return re.sub(r'^(#+)\s+(.+)$', transformar_linea, texto_md, flags=re.MULTILINE)

# -------------------------------
# PLANTILLA Y MOTORES (TU CÓDIGO ORIGINAL)
# -------------------------------
def apply_book_template(doc):
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    try:
        # Estilo para Título de Libro
        if 'BookTitle' not in doc.styles:
            title_style = doc.styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
            title_font = title_style.font
            title_font.name = 'Times New Roman'
            title_font.size = Pt(24)
            title_font.bold = True
            title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_style.paragraph_format.space_after = Pt(24)

        # Estilo para Título de Capítulo
        if 'ChapterTitle' not in doc.styles:
            chapter_style = doc.styles.add_style('ChapterTitle', WD_STYLE_TYPE.PARAGRAPH)
            chapter_font = chapter_style.font
            chapter_font.name = 'Times New Roman'
            chapter_font.size = Pt(18)
            chapter_font.bold = True
            chapter_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            chapter_style.paragraph_format.space_before = Pt(24)
            
        # Estilo para Párrafo
        if 'BookParagraph' not in doc.styles:
            para_style = doc.styles.add_style('BookParagraph', WD_STYLE_TYPE.PARAGRAPH)
            para_font = para_style.font
            para_font.name = 'Times New Roman'
            para_font.size = Pt(11)
            para_style.paragraph_format.first_line_indent = Inches(0.25)
            para_style.paragraph_format.line_spacing = 1.15
    except:
        pass
    return doc

def convert_with_pandoc(md_text, template_bytes=None):
    import pypandoc
    extra_args = ["--standalone"]
    tmp_template_path = None
    
    if template_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_template:
            tmp_template.write(template_bytes)
            tmp_template_path = tmp_template.name
        extra_args.append(f"--reference-doc={tmp_template_path}")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
        out_path = tmp_out.name
        pypandoc.convert_text(md_text, "docx", format="md", outputfile=out_path, extra_args=extra_args)
        with open(out_path, "rb") as f:
            data = f.read()
    
    if os.path.exists(out_path): os.remove(out_path)
    if tmp_template_path and os.path.exists(tmp_template_path): os.remove(tmp_template_path)
    return data

def convert_with_python(md_text, template_bytes=None):
    import markdown
    from htmldocx import HtmlToDocx
    md_html = markdown.markdown(md_text, extensions=["extra", "fenced_code", "toc"])
    
    doc = Document(BytesIO(template_bytes)) if template_bytes else apply_book_template(Document())
    HtmlToDocx().add_html_to_document(md_html, doc)
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_book_document(md_text, title, author, fix_titles):
    doc = Document()
    doc = apply_book_template(doc)
    
    # Portada
    t_portada = corregir_capitalizacion_espanola(f"# {title}").replace("# ", "") if fix_titles else title
    p_title = doc.add_paragraph(t_portada, style='BookTitle')
    
    # Contenido
    lines = md_text.split('\n')
    for line in lines:
        if line.startswith('# '):
            p = doc.add_paragraph(line[2:], style='ChapterTitle')
        elif line.startswith('## '):
            p = doc.add_paragraph(line[3:])
            p.bold = True
        elif line.strip():
            doc.add_paragraph(line.strip(), style='BookParagraph')
            
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ STREAMLIT
# -------------------------------
st.title("📚 Markdown → Word")

with st.sidebar:
    st.header("⚙️ Configuración")
    motor = st.radio("Motor", ["Pandoc (recomendado)", "Motor ligero (Python)", "Plantilla predefinida"])
    corregir_títulos = st.checkbox("Aplicar norma española de títulos", value=True)
    template_file = st.file_uploader("Plantilla .docx", type=["docx"])
    nombre_archivo = st.text_input("Nombre de salida", "mi_libro")

archivo_md = st.file_uploader("Sube tu .md", type=["md", "txt"])
texto_area = st.text_area("O pega el texto aquí")

contenido = archivo_md.read().decode("utf-8") if archivo_md else texto_area

if st.button("Convertir y Descargar"):
    if contenido:
        # 1. CORRECCIÓN DE MAYÚSCULAS (Si está activo)
        if corregir_títulos:
            contenido = corregir_capitalizacion_espanola(contenido)
        
        # 2. PROCESAMIENTO SEGÚN MOTOR
        t_bytes = template_file.read() if template_file else None
        
        try:
            if motor.startswith("Pandoc"):
                resultado = convert_with_pandoc(contenido, t_bytes)
            elif motor.startswith("Motor ligero"):
                resultado = convert_with_python(contenido, t_bytes)
            else:
                resultado = create_book_document(contenido, "Título", "Autor", corregir_títulos)
            
            st.success("¡Documento generado!")
            st.download_button("⬇️ Descargar Word", resultado, f"{nombre_archivo}.docx")
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("No hay contenido para convertir.")
