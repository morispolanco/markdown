# app.py
import streamlit as st
import subprocess
import os
import tempfile

# --- Configuración de la página ---
st.set_page_config(
    page_title="Conversor Markdown a Word",
    page_icon="📄",
    layout="centered"
)

# --- Título y descripción ---
st.title("📄 Conversor de Markdown a Word con Plantilla")
st.markdown("""
Esta aplicación convierte tu texto **Markdown** en un documento de **Word (.docx)**,
aplicando los estilos de una plantilla que tú proporcionas.
""")

# --- Widgets de entrada de usuario ---
markdown_text = st.text_area(
    "1. Introduce tu texto Markdown:",
    height=300,
    placeholder="# Título Principal\n\n## Subtítulo\n\nEste es un párrafo con **negrita** y *cursiva*.\n\n- Elemento de lista 1\n- Elemento de lista 2\n\n> Esto es una cita."
)

template_file = st.file_uploader(
    "2. Sube tu plantilla de Word (.docx):",
    type=['docx'],
    help="Sube un archivo .docx que contenga los estilos que quieres aplicar (Título 1, Título 2, Normal, etc.)."
)

# --- Botón de conversión ---
if st.button("🚀 Convertir y Descargar", type="primary"):
    
    # Validación de entradas
    if not markdown_text.strip():
        st.error("Por favor, introduce algún texto en el área de Markdown.")
        st.stop()
        
    if template_file is None:
        st.error("Por favor, sube un archivo de plantilla .docx.")
        st.stop()

    with tempfile.TemporaryDirectory() as temp_dir:
        st.info("Procesando... por favor, espera.")
        
        try:
            # 1. Guardar el texto Markdown en un archivo temporal
            md_path = os.path.join(temp_dir, "input.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(markdown_text)

            # 2. Guardar la plantilla subida en un archivo temporal
            template_path = os.path.join(temp_dir, template_file.name)
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())

            # 3. Definir la ruta del archivo de salida
            output_path = os.path.join(temp_dir, "output.docx")

            # 4. Construir y ejecutar el comando de Pandoc
            # Usamos la ruta al binario local que incluimos en el repositorio
            pandoc_path = os.path.join(os.path.dirname(__file__), "bin", "pandoc")

            # Asegurarnos de que el binario tenga permisos de ejecución
            os.chmod(pandoc_path, 0o755)
            
            command = [
                pandoc_path,  # <-- CAMBIO CLAVE: Usar la ruta local
                md_path,
                "-o", output_path,
                "--reference-doc", template_path
            ]
            
            # Ejecutamos el comando. 'check=True' lanzará una excepción si Pandoc falla.
            result = subprocess.run(command, check=True, capture_output=True, text=True)

            # 5. Leer el archivo .docx generado para la descarga
            with open(output_path, "rb") as f:
                binary_data = f.read()

            # 6. Mostrar el botón de descarga
            st.success("¡Conversión completada con éxito!")
            st.download_button(
                label="📥 Descargar documento.docx",
                data=binary_data,
                file_name="documento_convertido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except FileNotFoundError:
            st.error("""
            **Error Crítico:** No se encontró el binario de Pandoc en la carpeta `bin/`.
            Asegúrate de haber seguido los pasos para incluir el binario en tu repositorio.
            """)
        except subprocess.CalledProcessError as e:
            st.error(f"""
            **Error durante la conversión con Pandoc:**
            ```
            {e.stderr}
            ```
            Revisa tu texto Markdown y tu plantilla .docx e inténtalo de nuevo.
            """)
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {e}")
