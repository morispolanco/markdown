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

st.title("📚 Markdown → Word (Plantilla de Libro)")

st.markdown(
    "- Pega tu Markdown o sube un archivo .md\n"
    "- Elige el motor de conversión\n"
    "- Descarga el documento .docx formateado como libro"
)

# -------------------------------
# Función para aplicar formato de plantilla de libro
# -------------------------------
def apply_book_template(doc):
    # Configurar márgenes (5x8 pulgadas aprox)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    # Configurar fuentes
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # Crear estilos personalizados si no existen
    try:
        # Estilo para Título de Libro
        title_style = doc.styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
        title_font = title_style.font
        title_font.name = 'Times New Roman'
        title_font.size = Pt(24)
        title_font.bold = True
        title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_style.paragraph_format.space_after = Pt(24)
        
        # Estilo para Subtítulo
        subtitle_style = doc.styles.add_style('BookSubtitle', WD_STYLE_TYPE.PARAGRAPH)
        subtitle_font = subtitle_style.font
        subtitle_font.name = 'Times New Roman'
        subtitle_font.size = Pt(14)
        subtitle_font.italic = True
        subtitle_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_style.paragraph_format.space_after = Pt(36)
        
        # Estilo para Autor
        author_style = doc.styles.add_style('BookAuthor', WD_STYLE_TYPE.PARAGRAPH)
        author_font = author_style.font
        author_font.name = 'Times New Roman'
        author_font.size = Pt(16)
        author_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        author_style.paragraph_format.space_after = Pt(48)
        
        # Estilo para Título de Capítulo (Level 1)
        chapter_style = doc.styles.add_style('ChapterTitle', WD_STYLE_TYPE.PARAGRAPH)
        chapter_font = chapter_style.font
        chapter_font.name = 'Times New Roman'
        chapter_font.size = Pt(18)
        chapter_font.bold = True
        chapter_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        chapter_style.paragraph_format.space_before = Pt(24)
        chapter_style.paragraph_format.space_after = Pt(18)
        
        # Estilo para Encabezado de Nivel 2
        heading2_style = doc.styles.add_style('BookHeading2', WD_STYLE_TYPE.PARAGRAPH)
        heading2_font = heading2_style.font
        heading2_font.name = 'Times New Roman'
        heading2_font.size = Pt(14)
        heading2_font.bold = True
        heading2_style.paragraph_format.space_before = Pt(12)
        heading2_style.paragraph_format.space_after = Pt(6)
        
        # Estilo para Encabezado de Nivel 3
        heading3_style = doc.styles.add_style('BookHeading3', WD_STYLE_TYPE.PARAGRAPH)
        heading3_font = heading3_style.font
        heading3_font.name = 'Times New Roman'
        heading3_font.size = Pt(12)
        heading3_font.bold = True
        heading3_font.italic = True
        heading3_style.paragraph_format.space_before = Pt(6)
        heading3_style.paragraph_format.space_after = Pt(3)
        
        # Estilo para Párrafo normal
        para_style = doc.styles.add_style('BookParagraph', WD_STYLE_TYPE.PARAGRAPH)
        para_font = para_style.font
        para_font.name = 'Times New Roman'
        para_font.size = Pt(11)
        para_style.paragraph_format.first_line_indent = Inches(0.25)
        para_style.paragraph_format.space_after = Pt(6)
        para_style.paragraph_format.line_spacing = 1.15
        
        # Estilo para Copyright
        copyright_style = doc.styles.add_style('BookCopyright', WD_STYLE_TYPE.PARAGRAPH)
        copyright_font = copyright_style.font
        copyright_font.name = 'Times New Roman'
        copyright_font.size = Pt(9)
        copyright_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Estilo para Tabla de Contenidos
        toc_style = doc.styles.add_style('BookTOC', WD_STYLE_TYPE.PARAGRAPH)
        toc_font = toc_style.font
        toc_font.name = 'Times New Roman'
        toc_font.size = Pt(11)
        toc_style.paragraph_format.left_indent = Inches(0.25)
        
    except:
        pass  # Los estilos ya existen
    
    return doc

# -------------------------------
# Conversión con Pandoc (recomendado)
# -------------------------------
def convert_with_pandoc(md_text: str) -> bytes:
    try:
        import pypandoc
        # Creamos archivo temporal de salida porque pypandoc no devuelve bytes directamente para docx
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "salida.docx")
            pypandoc.convert_text(
                md_text,
                "docx",
                format="md",
                outputfile=out_path,
                extra_args=["--standalone"]  # documento completo
            )
            
            # Cargar el documento y aplicar la plantilla de libro
            doc = Document(out_path)
            doc = apply_book_template(doc)
            
            # Guardar en memoria
            bio = BytesIO()
            doc.save(bio)
            bio.seek(0)
            return bio.getvalue()
    except Exception as e:
        raise RuntimeError(f"Pandoc no disponible o falló la conversión: {e}")

# -------------------------------
# Conversión con motor ligero (Markdown → HTML → DOCX)
# -------------------------------
def convert_with_python(md_text: str) -> bytes:
    # 1) Markdown → HTML
    import markdown

    # Extensiones para soportar sintaxis amplia (listas, tablas, bloques de código, etc.)
    md_html = markdown.markdown(
        md_text,
        extensions=[
            "extra",        # incluye tables, abbr, attr_list, def_list, etc.
            "fenced_code",
            "sane_lists",
            "toc",
            "admonition",
            "footnotes",
        ],
        output_format="html5",
    )

    # 2) HTML → DOCX
    # Usamos htmldocx (ligero). No cubre el 100% de HTML, pero funciona bien para Markdown típico.
    from docx import Document as DocxDocument
    from htmldocx import HtmlToDocx

    doc = DocxDocument()
    doc = apply_book_template(doc)
    
    converter = HtmlToDocx()
    # Inserta el HTML en el documento
    converter.add_html_to_document(md_html, doc)

    # 3) Guardar en memoria
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# -------------------------------
# Función para crear documento con estructura de libro
# -------------------------------
def create_book_document(md_text: str, title: str = "Book Title", author: str = "Author Name") -> bytes:
    # Crear un nuevo documento
    doc = Document()
    doc = apply_book_template(doc)
    
    # Página de título
    title_para = doc.add_paragraph()
    title_para.style = doc.styles['BookTitle']
    title_para.add_run(title)
    
    # Subtítulo
    subtitle_para = doc.add_paragraph()
    subtitle_para.style = doc.styles['BookSubtitle']
    subtitle_para.add_run("You can write a brief sub title here")
    
    # Autor
    author_para = doc.add_paragraph()
    author_para.style = doc.styles['BookAuthor']
    author_para.add_run(f"By\n{author}")
    
    # Página de copyright
    doc.add_page_break()
    copyright_para = doc.add_paragraph()
    copyright_para.style = doc.styles['BookCopyright']
    copyright_para.add_run(f"Title of Your Book\n©Copyright 2022 {author}, Title\n\nALL RIGHTS RESERVED\nNo part of this publication may be reproduced, stored in a retrieval system, or transmitted, in any form or by any means, electronic, mechanical, photocopying, recording or otherwise, without the express written permission of the author.\n\nName of Printer Goes Here\nISBN: 000-1234567890\n\nYour Organization Title here\nAn Example Incorporated Company\n0000 Example Street, Sample Suite\nState, City & Zip\n0123-456-7890\n\nFree book template downloaded from:\nhttps://usedtotech.com\nadmin@usedtotech.com")
    
    # Tabla de contenidos
    doc.add_page_break()
    toc_heading = doc.add_paragraph()
    toc_heading.style = doc.styles['BookTitle']
    toc_heading.add_run("Contents")
    
    # Analizar el markdown para extraer capítulos
    chapters = re.findall(r'^# (.+)$', md_text, re.MULTILINE)
    
    for i, chapter in enumerate(chapters, 1):
        toc_item = doc.add_paragraph()
        toc_item.style = doc.styles['BookTOC']
        toc_item.add_run(f"{chapter}\t{5+i*4}")
    
    # Procesar el contenido markdown
    doc.add_page_break()
    lines = md_text.split('\n')
    current_chapter = None
    
    for line in lines:
        if line.startswith('# '):
            # Nuevo capítulo
            chapter_title = line[2:].strip()
            chapter_para = doc.add_paragraph()
            chapter_para.style = doc.styles['ChapterTitle']
            chapter_para.add_run(f"Chapter {len(chapters) - chapters.index(chapter_title)}\n{chapter_title}")
            current_chapter = chapter_title
        elif line.startswith('## '):
            # Encabezado de nivel 2
            heading2 = doc.add_paragraph()
            heading2.style = doc.styles['BookHeading2']
            heading2.add_run(line[3:].strip())
        elif line.startswith('### '):
            # Encabezado de nivel 3
            heading3 = doc.add_paragraph()
            heading3.style = doc.styles['BookHeading3']
            heading3.add_run(line[4:].strip())
        elif line.strip() == '':
            # Línea en blanco
            doc.add_paragraph()
        else:
            # Párrafo normal
            para = doc.add_paragraph()
            para.style = doc.styles['BookParagraph']
            para.add_run(line.strip())
    
    # Guardar en memoria
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# -------------------------------
# UI
# -------------------------------
with st.sidebar:
    motor = st.radio(
        "Motor de conversión",
        options=["Pandoc (mejor compatibilidad)", "Motor ligero (Python)", "Plantilla de libro"],
        index=2,
        help="Pandoc soporta prácticamente toda la sintaxis Markdown. El motor ligero funciona sin Pandoc. La plantilla de libro aplica el formato de la plantilla proporcionada."
    )
    
    if motor == "Plantilla de libro":
        book_title = st.text_input("Título del libro", value="Book Title")
        book_author = st.text_input("Autor del libro", value="Author Name")
    
    nombre_salida = st.text_input("Nombre del archivo (sin extensión)", value="documento_markdown")
    vista_previa = st.checkbox("Mostrar vista previa del Markdown", value=True)
    st.markdown("—")
    st.caption("Consejo: con Pandoc puedes convertir tablas, listas de tareas, tachados, notas al pie y más.")

archivo = st.file_uploader("Sube un archivo Markdown (.md, .markdown, .txt)", type=["md", "markdown", "txt"])
texto_md = st.text_area("O pega tu Markdown aquí", height=300, placeholder="# Título\n\n**Negrita**, *cursiva*, listas, tablas, etc.")

if archivo is not None:
    try:
        contenido = archivo.read().decode("utf-8", errors="ignore")
        st.info("Se usará el contenido del archivo subido.")
    except Exception:
        contenido = ""
        st.error("No se pudo leer el archivo subido.")
else:
    contenido = texto_md

if vista_previa and contenido.strip():
    with st.expander("Vista previa renderizada del Markdown", expanded=False):
        st.markdown(contenido)

col1, col2 = st.columns([1, 2])
with col1:
    convertir = st.button("Convertir a .docx", type="primary", use_container_width=True)

if convertir:
    if not contenido.strip():
        st.warning("Por favor, proporciona contenido Markdown (archivo o texto).")
    else:
        try:
            if motor.startswith("Pandoc"):
                bytes_docx = convert_with_pandoc(contenido)
                etiqueta_motor = "Pandoc"
            elif motor.startswith("Motor ligero"):
                bytes_docx = convert_with_python(contenido)
                etiqueta_motor = "motor ligero (Python)"
            else:  # Plantilla de libro
                bytes_docx = create_book_document(contenido, book_title, book_author)
                etiqueta_motor = "plantilla de libro"
            
            st.success(f"Conversión exitosa con {etiqueta_motor}. ¡Listo para descargar!")
            st.download_button(
                "Descargar .docx",
                data=bytes_docx,
                file_name=f"{nombre_salida or 'documento'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Ocurrió un error durante la conversión: {e}")
            if motor.startswith("Pandoc"):
                st.info(
                    "Si no tienes Pandoc instalado, puedes:\n"
                    "- Instalarlo localmente desde https://pandoc.org/installing.html\n"
                    "- O instalar la rueda que lo incluye: pip install pypandoc-binary"
                )
