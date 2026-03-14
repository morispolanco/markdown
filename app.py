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

st.set_page_config(page_title="Markdown → Word (Plantilla de Libro)", page_icon="📚", layout="centered")

# -------------------------------
# Lógica de Capitalización Española
# -------------------------------
def fix_spanish_casing(text):
    """
    Convierte un título a minúsculas excepto la primera letra y nombres propios definidos.
    """
    if not text:
        return text
    
    # Lista de palabras que siempre deben mantener su capitalización (nombres propios)
    # Puedes ampliar esta lista según tus necesidades
    proper_nouns = ["España", "México", "Colombia", "Argentina", "Word", "Markdown", "Pandoc", "Python", "Streamlit"]
    
    # Dividir por espacios
    words = text.split()
    new_words = []
    
    for i, word in enumerate(words):
        # Limpiar puntuación para comparar con la lista de nombres propios
        clean_word = re.sub(r'[^\w]', '', word)
        
        if i == 0:
            # La primera palabra siempre lleva mayúscula inicial
            new_words.append(word.capitalize())
        elif clean_word in proper_nouns:
            # Si es nombre propio, se deja como está
            new_words.append(word)
        else:
            # El resto a minúsculas
            new_words.append(word.lower())
            
    return " ".join(new_words)

def process_markdown_titles(md_text):
    """
    Busca líneas que empiezan con # y aplica la norma española de capitalización.
    """
    def replace_header(match):
        hashes = match.group(1)
        title_content = match.group(2).strip()
        return f"{hashes} {fix_spanish_casing(title_content)}"

    # Regex para detectar headers de Markdown (# Título)
    return re.sub(r'^(#+)\s+(.+)$', replace_header, md_text, flags=re.MULTILINE)

# -------------------------------
# UI y Configuración
# -------------------------------
st.title("📚 Markdown → Word (Plantilla de Libro)")

with st.sidebar:
    st.header("⚙️ Configuración")
    
    fix_titles = st.checkbox("Corregir mayúsculas en títulos (Norma Española)", value=True, 
                             help="Convierte 'Título De Mi Libro' a 'Título de mi libro'.")
    
    motor = st.radio(
        "Motor de conversión",
        options=["Pandoc (mejor compatibilidad)", "Motor ligero (Python)", "Plantilla de libro (predefinida)"],
        index=0
    )
    
    template_file = st.file_uploader("Sube tu plantilla .docx (opcional)", type=["docx"])
    
    if motor == "Plantilla de libro (predefinida)":
        book_title = st.text_input("Título del libro", value="Book Title")
        book_author = st.text_input("Autor del libro", value="Author Name")
    
    nombre_salida = st.text_input("Nombre del archivo", value="documento_markdown")

# --- Las funciones apply_book_template, convert_with_pandoc y convert_with_python 
# se mantienen igual que en tu código original, solo cambia la llamada final ---

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
        title_style = doc.styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
        title_font = title_style.font
        title_font.name = 'Times New Roman'
        title_font.size = Pt(24)
        title_font.bold = True
        title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        chapter_style = doc.styles.add_style('ChapterTitle', WD_STYLE_TYPE.PARAGRAPH)
        chapter_font = chapter_style.font
        chapter_font.name = 'Times New Roman'
        chapter_font.size = Pt(18)
        chapter_font.bold = True
        chapter_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        para_style = doc.styles.add_style('BookParagraph', WD_STYLE_TYPE.PARAGRAPH)
        para_font = para_style.font
        para_font.name = 'Times New Roman'
        para_font.size = Pt(11)
        para_style.paragraph_format.first_line_indent = Inches(0.25)
    except:
        pass
    return doc

# (Aquí irían convert_with_pandoc y convert_with_python de tu código original)

def create_book_document(md_text, title, author):
    doc = Document()
    doc = apply_book_template(doc)
    
    # Aplicar corrección al título principal si está activo
    final_title = fix_spanish_casing(title) if fix_titles else title
    
    title_para = doc.add_paragraph()
    title_para.style = doc.styles['BookTitle']
    title_para.add_run(final_title)
    
    # ... resto del procesamiento de líneas ...
    lines = md_text.split('\n')
    for line in lines:
        if line.startswith('# '):
            t = fix_spanish_casing(line[2:].strip()) if fix_titles else line[2:].strip()
            p = doc.add_paragraph(t, style=doc.styles['ChapterTitle'])
        elif line.startswith('## '):
            t = fix_spanish_casing(line[3:].strip()) if fix_titles else line[3:].strip()
            p = doc.add_paragraph(t) # O estilo Heading 2
        else:
            if line.strip():
                p = doc.add_paragraph(line.strip(), style=doc.styles['BookParagraph'])
                
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# -------------------------------
# Ejecución Principal
# -------------------------------
archivo = st.file_uploader("Sube Markdown", type=["md", "txt"])
texto_md = st.text_area("O pega aquí")

contenido = archivo.read().decode("utf-8") if archivo else texto_md

if st.button("Convertir a .docx"):
    if contenido:
        # PROCESAMIENTO DE TÍTULOS SEGÚN NORMA ESPAÑOLA
        if fix_titles:
            contenido = process_markdown_titles(contenido)
        
        # Selección de motor (simplificado para el ejemplo)
        if motor.startswith("Plantilla de libro"):
            res = create_book_document(contenido, book_title, book_author)
        else:
            # Aquí llamarías a convert_with_pandoc o convert_with_python pasándole 'contenido' ya procesado
            st.warning("Motor seleccionado. En una app completa, aquí se generaría el archivo.")
            res = None 
            
        if res:
            st.download_button("Descargar", res, file_name=f"{nombre_salida}.docx")
