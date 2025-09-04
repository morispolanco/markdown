# app.py
import streamlit as st
from io import BytesIO
import tempfile
import os

st.set_page_config(page_title="Markdown → Word", page_icon="📝", layout="centered")

st.title("📝 Markdown → Word (.docx)")

st.markdown(
    "- Pega tu Markdown o sube un archivo .md\n"
    "- Elige el motor de conversión (Pandoc recomendado para máxima compatibilidad)\n"
    "- Descarga el documento .docx"
)

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
            with open(out_path, "rb") as f:
                return f.read()
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
    converter = HtmlToDocx()
    # Inserta el HTML en el documento
    converter.add_html_to_document(md_html, doc)

    # 3) Guardar en memoria
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
        options=["Pandoc (mejor compatibilidad)", "Motor ligero (Python)"],
        index=0,
        help="Pandoc soporta prácticamente toda la sintaxis Markdown. El motor ligero funciona sin Pandoc."
    )
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
            else:
                bytes_docx = convert_with_python(contenido)
                etiqueta_motor = "motor ligero (Python)"
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
