def corregir_capitalizacion_espanola(texto_md):
    """
    Aplica la norma española: Mayúscula en la primera palabra y nombres propios.
    Mantiene mayúsculas en palabras específicas (nombres de libros o marcas).
    """
    # Lista ampliada de nombres propios y títulos que deben mantener su mayúscula
    excepciones = [
        "España", "México", "Python", "Streamlit", "Pandoc", "Word", 
        "Markdown", "Biblia", "Quijote", "Cien", "Años", "Soledad"
    ]

    def transformar_linea(match):
        hashes = match.group(1)
        contenido = match.group(2).strip()
        
        palabras = contenido.split()
        nuevas_palabras = []
        
        for i, word in enumerate(palabras):
            # Limpiamos signos de puntuación para comparar con la lista de excepciones
            clean_word = re.sub(r'[^\w]', '', word)
            
            # REGLA 1: La primera palabra SIEMPRE con mayúscula inicial
            if i == 0:
                nuevas_palabras.append(word.capitalize())
            
            # REGLA 2: Si la palabra está en nuestra lista de excepciones/títulos, se respeta
            elif clean_word in excepciones:
                nuevas_palabras.append(word)
            
            # REGLA 3: El resto a minúsculas
            else:
                nuevas_palabras.append(word.lower())
        
        return f"{hashes} {' '.join(nuevas_palabras)}"

    return re.sub(r'^(#+)\s+(.+)$', transformar_linea, texto_md, flags=re.MULTILINE)
