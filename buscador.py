import streamlit as st
import openpyxl
from scholarly import scholarly
from fake_useragent import UserAgent
from io import BytesIO

# Configurar fallback para evitar el error de lista vacía
ua = UserAgent(fallback="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

def search_and_save(query, max_results):
    # Crear y configurar el libro de Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Título', 'Autores', 'Año', 'Revista', 'Enlace ePrint'])

    # Realizar la búsqueda y extraer los resultados
    try:
        search_query = scholarly.search_pubs(query)
        for i, article in enumerate(search_query):
            if i >= max_results:  # Limitar el número de resultados
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

        # Guardar el archivo Excel en memoria
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    except Exception as e:
        return str(e)

# Interfaz de usuario con Streamlit
st.title("Buscador de Google Scholar")

# Campo de entrada para la consulta de búsqueda
query = st.text_input("Ingrese la consulta de búsqueda:")

# Campo de entrada para la cantidad de resultados
max_results = st.number_input("Cantidad de resultados a guardar:", min_value=1, max_value=100, value=10, step=1)

# Botón para ejecutar la búsqueda y descargar resultados
if st.button("Buscar y Descargar"):
    if not query:
        st.warning("Por favor complete el campo de búsqueda.")
    else:
        result = search_and_save(query, max_results)
        if isinstance(result, BytesIO):
            st.success("Búsqueda completada. Descargue el archivo a continuación.")
            st.download_button(
                label="Descargar resultados",
                data=result,
                file_name="resultados_scholar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error(f"Ocurrió un error: {result}")


