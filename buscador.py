import streamlit as st
import openpyxl
from scholarly import scholarly
from fake_useragent import UserAgent

# Configurar fallback para evitar el error de lista vacía
ua = UserAgent(fallback="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

def search_and_save(query, filename):
    # Crear y configurar el libro de Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Título', 'Autores', 'Año', 'Revista', 'Enlace ePrint'])

    # Realizar la búsqueda y extraer los resultados
    try:
        search_query = scholarly.search_pubs(query)
        for i, article in enumerate(search_query):
            if i >= 100:  # Limitar el número de resultados
                break

            # Extraer los datos del campo 'bib' y 'eprint_url'
            bib_data = article.get('bib', {})
            title = bib_data.get('title', 'N/A')
            authors = ", ".join(bib_data.get('author', []))
            year = bib_data.get('pub_year', 'N/A')
            journal = bib_data.get('venue', 'N/A')
            eprint_url = article.get('eprint_url', 'N/A')

            # Agregar la información en la hoja de Excel
            ws.append([title, authors, year, journal, eprint_url])

        # Guardar el archivo Excel
        wb.save(filename)
        return f"Resultados guardados en {filename}"

    except Exception as e:
        return f"Ocurrió un error al realizar la búsqueda: {e}"

# Interfaz de usuario con Streamlit
st.title("Buscador de Google Scholar")

# Campo de entrada para la consulta de búsqueda
query = st.text_input("Ingrese la consulta de búsqueda:")

# Campo de entrada para el nombre del archivo
filename = st.text_input("Ingrese el nombre del archivo (sin extensión):")

# Botón para ejecutar la búsqueda y guardar resultados
if st.button("Buscar y Guardar"):
    if not query or not filename:
        st.warning("Por favor complete ambos campos.")
    else:
        result_message = search_and_save(query, f"{filename}.xlsx")
        st.success(result_message) if "guardados" in result_message else st.error(result_message)
