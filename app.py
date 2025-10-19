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
    "- (Opcional) Sube tu propia plantilla .docx para personalizar el estilo\n"
    "- Descarga el documento .docx formateado"
)

# -------------------------------
# Función para aplicar formato de plantilla de libro (por defecto)
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
        
    except Exception as e:
        # Los estilos ya existen o hubo un error, lo ignoramos para continuar
        pass
    
    return doc

# -------------------------------
# Conversión con Pandoc (recomendado)
# -------------------------------
def convert_with_pandoc(md_text: str, template_bytes: bytes = None) -> bytes:
    try:
        import pypandoc
        
        # Argumentos extra para Pandoc
        extra_args = ["--standalone"]
        
        # Si se proporciona una plantilla, guardarla temporalmente y usarla como referencia
        if template_bytes:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_template:
                tmp_template.write(template_bytes)
                tmp_template_path = tmp_template.name
            extra_args.append(f"--reference-doc={tmp_template_path}")

        # Crear archivo temporal de salida
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_out:
            out_path = tmp_out.name
            
            pypandoc.convert_text(
                md_text,
                "docx",
                format="md",
                outputfile=out_path,
                extra_args=extra_args
            )
            
            # Leer el resultado final a memoria
            with open(out_path, "rb") as f:
                bio = BytesIO(f.read())
            
            # Limpiar archivos temporales
            os.remove(out_path)
            if template_bytes:
                os.remove(tmp_template_path)
            
            bio.seek(0)
            return bio.getvalue()

    except ImportError:
        raise RuntimeError("La librería 'pypandoc' no está instalada. Ejecuta `pip install pypandoc`.")
    except Exception as e:
        # Limpiar archivos temporales en caso de error
        if 'out_path' in locals() and os.path.exists(out_path):
            os.remove(out_path)
        if 'tmp_template_path' in locals() and os.path.exists(tmp_template_path):
            os.remove(tmp_template_path)
        raise RuntimeError(f"Pandoc no disponible o falló la conversión: {e}")

# -------------------------------
# Conversión con motor ligero (Markdown → HTML → DOCX)
# -------------------------------
def convert_with_python(md_text: str, template_bytes: bytes = None) -> bytes:
    # 1) Markdown → HTML
    import markdown

    md_html = markdown.markdown(
        md_text,
        extensions=[
            "extra", "fenced_code", "sane_lists", "toc", "admonition", "footnotes",
        ],
        output_format="html5",
    )

    # 2) HTML → DOCX
    from docx import Document as DocxDocument
    from htmldocx import HtmlToDocx

    # Cargar la plantilla si se proporciona, si no, crear un nuevo documento
    if template_bytes:
        doc = DocxDocument(BytesIO(template_bytes))
    else:
        doc = DocxDocument()
        doc = apply_book_template(doc) # Aplicar plantilla por defecto solo si no hay una personalizada
    
    converter = HtmlToDocx()
    converter.add_html_to_document(md_html, doc)

    # 3) Guardar en memoria
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# -------------------------------
# Función para crear documento con estructura de libro predefinida
# -------------------------------
def create_book_document(md_text: str, title: str = "Book Title", author: str = "Author Name") -> bytes:
    # NOTA: Esta función ignora la plantilla subida por el usuario, ya que genera
    # una estructura de libro fija y muy específica desde cero.
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
    
    for line in lines:
        if line.startswith('# '):
            chapter_title = line[2:].strip()
            chapter_para = doc.add_paragraph()
            chapter_para.style = doc.styles['ChapterTitle']
            chapter_para.add_run(f"Chapter {len(chapters) - chapters.index(chapter_title)}\n{chapter_title}")
        elif line.startswith('## '):
            heading2 = doc.add_paragraph()
            heading2.style = doc.styles['BookHeading2']
            heading2.add_run(line[3:].strip())
        elif line.startswith('### '):
            heading3 = doc.add_paragraph()
            heading3.style = doc.styles['BookHeading3']
            heading3.add_run(line[4:].strip())
        elif line.strip() == '':
            doc.add_paragraph()
        else:
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
    st.header("⚙️ Configuración")
    
    motor = st.radio(
        "Motor de conversión",
        options=["Pandoc (mejor compatibilidad)", "Motor ligero (Python)", "Plantilla de libro (predefinida)"],
        index=0,
        help="**Pandoc**: Usa tu plantilla .docx como referencia para estilos. **Motor ligero**: Añade contenido a tu plantilla. **Plantilla de libro**: Crea una estructura fija y usa su propio estilo."
    )
    
    template_file = st.file_uploader(
        "Sube tu plantilla .docx (opcional)",
        type=["docx"],
        help="Sube un archivo .docx para usarlo como base de estilos. Se usará con los motores Pandoc y Ligero. La opción 'Plantilla de libro' la ignorará."
    )
    
    if motor == "Plantilla de libro (predefinida)":
        book_title = st.text_input("Título del libro", value="Book Title")
        book_author = st.text_input("Autor del libro", value="Author Name")
    
    nombre_salida = st.text_input("Nombre del archivo (sin extensión)", value="documento_markdown")
    vista_previa = st.checkbox("Mostrar vista previa del Markdown", value=True)
    
    st.markdown("---")
    st.caption("Consejo: con Pandoc puedes convertir tablas, listas de tareas, tachados, notas al pie y más.")

# Leer la plantilla subida a memoria
template_bytes = None
if template_file is not None:
    try:
        template_bytes = template_file.read()
        st.sidebar.success("✅ Plantilla .docx cargada.")
    except Exception as e:
        st.sidebar.error(f"Error al leer la plantilla: {e}")
        template_bytes = None

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
                bytes_docx = convert_with_pandoc(contenido, template_bytes)
                etiqueta_motor = "Pandoc"
            elif motor.startswith("Motor ligero"):
                bytes_docx = convert_with_python(contenido, template_bytes)
                etiqueta_motor = "motor ligero (Python)"
            else:  # Plantilla de libro
                bytes_docx = create_book_document(contenido, book_title, book_author)
                etiqueta_motor = "plantilla de libro predefinida"
            
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
                    "- O instalar la rueda que lo incluye: `pip install pypandoc-binary`"
                )
