import sys
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox
from scholarly import scholarly
from fake_useragent import UserAgent



# Configurar fallback para evitar el error de lista vacía
ua = UserAgent(fallback="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

class ScholarSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Configuración básica de la ventana
        self.setWindowTitle('Buscador de Google Scholar')
        self.setGeometry(100, 100, 400, 200)

        # Crear layout principal
        layout = QVBoxLayout()

        # Etiqueta y campo de entrada para la consulta de búsqueda
        self.query_label = QLabel("Ingrese la consulta de búsqueda:", self)
        layout.addWidget(self.query_label)
        self.query_input = QLineEdit(self)
        layout.addWidget(self.query_input)

        # Etiqueta y campo de entrada para el nombre del archivo
        self.filename_label = QLabel("Ingrese el nombre del archivo (sin extensión):", self)
        layout.addWidget(self.filename_label)
        self.filename_input = QLineEdit(self)
        layout.addWidget(self.filename_input)

        # Botón para ejecutar la búsqueda y guardar resultados
        self.search_button = QPushButton("Buscar y Guardar", self)
        self.search_button.clicked.connect(self.search_and_save)
        layout.addWidget(self.search_button)

        # Establecer layout
        self.setLayout(layout)

    def search_and_save(self):
        # Obtener la consulta y el nombre de archivo del usuario
        query = self.query_input.text()
        filename = self.filename_input.text() + ".xlsx"

        # Validar que los campos no estén vacíos
        if not query or not filename:
            QMessageBox.warning(self, "Advertencia", "Por favor complete ambos campos.")
            return

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
            QMessageBox.information(self, "Éxito", f"Resultados guardados en {filename}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al realizar la búsqueda: {e}")

# Crear la aplicación y ejecutar la interfaz
app = QApplication(sys.argv)
window = ScholarSearchApp()
window.show()
sys.exit(app.exec_())
