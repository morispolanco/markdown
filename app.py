import streamlit as st
from io import BytesIO
import tempfile
import os
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

# -------------------------------
# CONFIGURACIÓN Y ESTILOS
# -------------------------------
st.set_page_config(page_title="Markdown → Word (Norma Española)", page_icon="📚", layout="centered")

def apply_book_template(doc):
    """Configuración estética por defecto para estilo libro."""
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)

    try:
        # Estilo Título de Libro (Portada)
        if 'BookTitle' not in doc.styles:
            s = doc.styles.add_style('BookTitle', WD_STYLE_TYPE.PARAGRAPH)
            s.font.name = 'Times New Roman'; s.font.size = Pt(24); s.font.bold = True
            s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            s.paragraph_format.space_after = Pt(24)

        # Estilo Títulos de Capítulos (#)
        if 'ChapterTitle' not in doc.styles:
            s = doc.styles.add_style('ChapterTitle', WD_STYLE_TYPE.PARAGRAPH)
            s.font.name = 'Times New Roman'; s.font.size = Pt(18); s.font.bold = True
            s.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            s.paragraph_format.space_before = Pt(24); s.paragraph_format.space_after = Pt(18)

        # Estilo Párrafo Libro
        if 'BookParagraph' not in doc.styles:
            s = doc.styles.add_style('BookParagraph', WD_STYLE_TYPE.PARAGRAPH)
            s.font.name = 'Times New Roman'; s.font.size = Pt(11)
            s.paragraph_format.first_line_indent = Inches(0.25)
            s.paragraph_format.line_spacing = 1.15
            s.paragraph_format.space_after = Pt(6)
    except:
        pass
    return doc

# -------------------------------
# LÓGICA DE CAPITALIZACIÓN ESPAÑOLA
# -------------------------------
def fix_spanish_casing(text, excepciones_usuario=""):
    """
    Convierte títulos a minúsculas excepto:
    1. La primera palabra.
    2. Nombres propios o palabras en la lista de excepciones.
    """
    if not text: return text
    
    # Lista base de nombres propios
    base_excepciones = ["España", "México", "Dios", "Python", "Streamlit", "Pandoc", "Word", "Markdown"]
    # Unir con excepciones que el usuario escriba en la UI
    lista_final = base_excepciones + [ex.strip() for ex in excepciones_usuario.split(",") if ex.strip()]

    def procesar_frase(frase):
        palabras = frase.split()
        resultado = []
        for i, word in enumerate(palabras):
            clean_word = re.sub(r'[^\w]', '', word)
            if i == 0:
                resultado.append(word.capitalize())
            elif clean_word in lista_final:
                resultado.append(word)
            else:
                resultado.append(word.lower())
        return " ".join(resultado)

    # Si el texto es una línea de Markdown con #, preservamos los #
    match = re.match(r'^(#+)\s+(.+)$', text)
    if match:
        return f"{match.group(1)} {procesar_frase(match.group(2))}"
    return procesar_frase(text)

# -------------------------------
# MOTORES DE CONVERSIÓN
# -------------------------------
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
    md_html = markdown.markdown(md_text, extensions=["extra", "fenced_code", "sane_lists", "toc"])
    
    doc = Document(BytesIO(template_bytes)) if template_bytes else apply_book_template(Document())
    HtmlToDocx().add_html_to_document(md_html, doc)
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_book_document(md_text, title, author, ex_words):
    """Motor predefinido con estructura de libro."""
    doc = apply_book_template(Document())
    
    # Título principal con capitalización corregida
    doc.add_paragraph(fix_spanish_casing(title, ex_words), style='BookTitle')
    doc.add_paragraph(f"Por: {author}", style='Normal').paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    lines = md_text.split('\n')
    for line in lines:
        if line.startswith('# '):
            doc.add_paragraph(line[2:].strip(), style='ChapterTitle')
        elif line.startswith('## '):
            p = doc.add_paragraph(line[3:].strip())
            p.bold = True
        elif line.strip():
            doc.add_paragraph(line.strip(), style='BookParagraph')
            
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# -------------------------------
# INTERFAZ DE USUARIO (UI)
# -------------------------------
st.title("📚 Markdown a Word")

with st.sidebar:
    st.header("⚙️ Configuración")
    motor = st.radio("Motor de conversión", ["Pandoc (Recomendado)", "Motor ligero (Python)", "Plantilla de libro"])
    
    st.subheader("Norma Española")
    corregir = st.checkbox("Corregir mayúsculas en títulos", value=True)
    excepciones_input = st.text_area("Excepciones (Nombres propios, títulos de libros)", 
                                     placeholder="Don Quijote, Harry Potter, Madrid...")
    
    template_file = st.file_uploader("Subir plantilla .docx base", type=["docx"])
    nombre_archivo = st.text_input("Nombre del archivo de salida", "mi_documento")

archivo_md = st.file_uploader("Sube tu archivo Markdown", type=["md", "txt"])
texto_area = st.text_area("O pega tu contenido aquí", height=250)

contenido = archivo_md.read().decode("utf-8") if archivo_md else texto_area

if st.button("🚀 Convertir y Descargar"):
    if contenido.strip():
        # 1. PROCESAR CAPITALIZACIÓN SI ESTÁ ACTIVO
        if corregir:
            lineas = contenido.split('\n')
            contenido_procesado = []
            for l in lineas:
                if l.startswith('#'):
                    contenido_procesado.append(fix_spanish_casing(l, excepciones_input))
                else:
                    contenido_procesado.append(l)
            contenido = '\n'.join(contenido_procesado)

        # 2. GENERAR DOCX
        t_bytes = template_file.read() if template_file else None
        
        try:
            if motor.startswith("Pandoc"):
                res = convert_with_pandoc(contenido, t_bytes)
            elif motor.startswith("Motor ligero"):
                res = convert_with_python(contenido, t_bytes)
            else:
                res = create_book_document(contenido, "Título de Obra", "Autor", excepciones_input)

            st.success("¡Documento generado con éxito!")
            st.download_button(
                label="⬇️ Descargar archivo Word",
                data=res,
                file_name=f"{nombre_archivo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Error en la conversión: {e}")
            if "pypandoc" in str(e):
                st.info("Nota: Pandoc requiere que el software esté instalado en el servidor/PC.")
    else:
        st.warning("Escribe o sube algún contenido antes de convertir.")
