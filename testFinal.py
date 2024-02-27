import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QComboBox, QPushButton, QFileDialog, QHeaderView, QMessageBox
from PyQt5.QtCore import Qt
import openpyxl
import pandas as pd
import os



class MiVentana(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Filtrado de Excel - Moodle')
        self.setStyleSheet("background-color: #f0f0f0; font-family: Arial, sans-serif;")

        encabezado = QLabel('<h1>Instituto Superior Tecnológico España</h1>', self)
        encabezado.setStyleSheet("color: #1e3b72; margin: 15px; font-size: 18px; font-weight: bold; text-align: center;")
        zona_combobox_botones = QWidget(self)

        self.combo_box_carrera = QComboBox(zona_combobox_botones)
        self.combo_box_carrera.setEditable(True)
        self.combo_box_carrera.addItems([
                "ADMINISTRACIÓN DE EMPRESAS Y EMPRENDIMIENTO",
                "ADMINISTRACION FINANCIERA",
                "ADMINISTRACIÓN",
                "ADMINISTRACION DE EMPRESAS E INTELIGENCIA DE NEGOCIOS",
                "DESARROLLO DE APLICACIONES WEB",
                "DESARROLLO INFANTIL INTEGRAL",
                "ENFERMERÍA",
                "GESTIÓN ESTRATÉGICA DEL MARKETING DIGITAL",
                "GESTIÓN DE FINANZAS Y RIESGOS FINANCIEROS",
                "LABORATORIO CLÍNICO",
                "LOGÍSTICA Y TRANSPORTE",
                "MARKETING",
                "REHABILITACIÓN FÍSICA",
                "SISTEMAS DE INFORMACIÓN Y CIBERSEGURIDAD"
            ])

        self.combo_box_carrera.setStyleSheet("background-color: #fff;")
        self.combo_box_carrera.setFixedSize(370, 30)
        self.combo_box_carrera.activated.connect(self.filtrarPorCarreraSeleccionada)


        #Filtro modalidad (EL/PRE)
        self.combo_box_modalidad = QComboBox(zona_combobox_botones)
        self.combo_box_modalidad.setEditable(True)
        self.combo_box_modalidad.addItems([
                "EN LÍNEA",
                "PRESENCIAL"
            ])

        self.combo_box_modalidad.setStyleSheet("background-color: #fff; margin-left: 10px;")
        self.combo_box_modalidad.setFixedSize(100, 30)
        self.combo_box_modalidad.currentIndexChanged.connect(self.filtrarPorModalidadSeleccionada)

        #Filtro curso
        self.combo_box_curso = QComboBox(zona_combobox_botones)
        self.combo_box_curso.setEditable(True)
        self.combo_box_curso.setStyleSheet("background-color: #fff;")
        self.combo_box_curso.setFixedSize(370, 30)
        self.combo_box_curso.currentIndexChanged.connect(self.contar_coincidencias)

        btn_cargar_excel = QPushButton("Cargar", zona_combobox_botones)
        btn_cargar_excel.clicked.connect(self.cargarArchivoExcel)
        btn_cargar_excel.setStyleSheet("background-color: #0781D8; color: #fff; padding: 10px 15px; border: none; border-radius: 5px; font-size: 14px;")

        #btn_limpiar_tabla = QPushButton("FCC", zona_combobox_botones)
        #btn_limpiar_tabla.clicked.connect(self.filtrarPorCarreraYCurso)
        #btn_limpiar_tabla.setStyleSheet("background-color: #ff9800; color: #fff; padding: 10px 15px; border: none; border-radius: 5px; font-size: 14px;")

        self.label_filtrado = QLabel("Filtrado:")
        self.label_filtrado.setStyleSheet("margin: 10px;")
        self.label_total = QLabel("Total:")
        self.label_total.setStyleSheet("margin: 10px;")

        layout_combobox_botones = QHBoxLayout(zona_combobox_botones)
        layout_combobox_botones.addWidget(QLabel("Modalidad:"))
        layout_combobox_botones.addWidget(self.combo_box_modalidad)
        #layout_combobox_botones.addSpacing(10)
        layout_combobox_botones.addWidget(QLabel("Carrera:"))
        layout_combobox_botones.addWidget(self.combo_box_carrera)
        #layout_combobox_botones.addSpacing(10)
        layout_combobox_botones.addWidget(QLabel("Curso:"))
        layout_combobox_botones.addWidget(self.combo_box_curso)
        #layout_combobox_botones.addSpacing(10)
        layout_combobox_botones.addWidget(self.label_filtrado)
        #layout_combobox_botones.addSpacing(10)
        layout_combobox_botones.addWidget(self.label_total)
        layout_combobox_botones.addStretch(1)
        #layout_combobox_botones.addWidget(btn_limpiar_tabla)
        layout_combobox_botones.addWidget(btn_cargar_excel)

        self.tabla = QTableWidget(self)
        self.tabla.setColumnCount(5)
        self.tabla.setHorizontalHeaderLabels(["Email", "Nombres", "Apellidos", "Carrera", "Curso"])
        # Ajuste para que los encabezados de las columnas se estiren automáticamente para ocupar el espacio disponible
        self.tabla.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Aplicar estilos al encabezado horizontal
        self.tabla.horizontalHeader().setStyleSheet(
            "QHeaderView::section {"
            "color: #FFF;"  # Color de la fuente
            "background-color: #1E3B72;"  # Color de fondo
            "padding: 5px;"  # Ajuste de padding si es necesario
            "border: 1px solid #FFF;"  # Eliminar bordes si se prefiere
            "}"
        )
        self.tabla.setStyleSheet("margin-top: 10px; margin-bottom: 15px")
        # O, para estirar solo la última sección, puedes comentar la línea anterior y descomentar la siguiente:
        # self.tabla.horizontalHeader().setStretchLastSection(True)
        #limpiar datos de memoria
        self.limpiarTabla

        layout_vertical = QVBoxLayout(self)
        layout_vertical.addWidget(encabezado, alignment=Qt.AlignVCenter | Qt.AlignHCenter)
        layout_vertical.addWidget(zona_combobox_botones)
        layout_vertical.addWidget(self.tabla, stretch=1)

        self.cargarArchivoExcel()

    def cargarArchivoExcel(self):
        opciones = QFileDialog.Options()
        opciones |= QFileDialog.ReadOnly
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel", "", "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)", options=opciones)
        if archivo:
            self.convertirExcelACSV(archivo)

    def convertirExcelACSV(self, excel_path):
        csv_path = 'UsuariosSinAcceder.csv'

        # Verificar que el archivo tenga la extensión .xlsx
        if not excel_path.lower().endswith('.xlsx'):
            QMessageBox.warning(self, 'Error', "El archivo no tiene la extensión .xlsx. Por favor, introduce un archivo Excel válido.")
            return

        # Verificar que el archivo no esté vacío
        if os.path.getsize(excel_path) == 0:
            QMessageBox.warning(self, 'Error', "El archivo está vacío. Por favor, introduce un archivo Excel válido.")
            return

        try:
            # Leer el archivo Excel
            libro_excel = openpyxl.load_workbook(excel_path)
            hoja_excel = libro_excel.active

            # Crear DataFrame con los datos de Excel
            data = []
            for fila_excel in hoja_excel.iter_rows(min_row=2, values_only=True):
                fila_excel_str = [str(valor) for valor in fila_excel]
                data.append(list(fila_excel_str))
            df = pd.DataFrame(data, columns=["category id", "category name", "course id", "course shortname", 
                            "course fullname", "user id", "username", "first name", "last name"])
            
             # Validar la columna category id
            if df['category id'].str.contains('-- ROW LIMIT EXCEEDED --').any():
                # Eliminar las filas que contienen el valor '-- ROW LIMIT EXCEEDED --'
                df = df[~df['category id'].str.contains('-- ROW LIMIT EXCEEDED --')]

            # Convertir el DataFrame a CSV
            df.to_csv(csv_path, index=False)
            QMessageBox.information(self, 'Éxito', f"Archivo convertido con éxito y guardado como {csv_path}")
            # Cargar los datos del CSV en el ComboBox y la tabla
            self.cargar_categorias(csv_path)
            self.cargarDatosDesdeCSV(csv_path)
            total_data = df.shape[0]
            self.label_total.setText(f"TOTAL:  {total_data}")
            self.label_filtrado.setText(f"FILTRADO:  0")
            
        except Exception as e:
            QMessageBox.warning(self, 'Error', f"Ocurrió un error al convertir el archivo: {e}")
            print(f"Ocurrió un error al convertir el archivo: {e}")

    def cargar_categorias(self, archivo):
        df = pd.read_csv(archivo)
        categorias = df['course shortname'].apply(lambda x: f"{x.split(' - ')[0].strip()} - [{x.split('- [')[1].strip()}").unique()
        self.combo_box_curso.addItems(categorias)

    def contar_coincidencias(self):
        categoria_seleccionada = self.combo_box_curso.currentText()
        archivo_csv = 'UsuariosSinAcceder.csv'

        try:
            df = pd.read_csv(archivo_csv)
            #conteo_categoria = (df['course shortname'].apply(lambda x: f"{x.split(' - ')[0].strip()} - [{x.split('- [')[1].strip()}") == categoria_seleccionada).sum()
            #self.resultLabel.setText(f"Coincidencias de '{categoria_seleccionada}': {conteo_categoria}")
            #print(f"Coincidencias de '{categoria_seleccionada}': {conteo_categoria}")
            # Mostrar los resultados
            #QMessageBox.information(self, 'Resultados', f"Coincidencias de '{categoria_seleccionada}': {conteo_categoria}")

            # Filtrar el DataFrame por la categoría seleccionada
            df_filtrado = df[df['course shortname'].apply(lambda x: f"{x.split(' - ')[0].strip()} - [{x.split('- [')[1].strip()}") == categoria_seleccionada]
            total_data = df_filtrado.shape[0]
            self.label_filtrado.setText(f"FILTRADO:  {total_data}")

            # Actualizar la tabla con los datos filtrados
            self.mostrarResultadosFiltrados(df_filtrado)
            #conteo_categorias = df_filtrado['categoria'].value_counts()
            #QMessageBox.information(self, 'Resultados', f"Total de coincidencias '{conteo_categorias}': LOL")

        except Exception as e:
            QMessageBox.warning(self, 'Error', f"Ocurrió un error al cargar el archivo CSV: {e}")

    def cargarDatosDesdeCSV(self, archivo):
        df = pd.read_csv(archivo)
        self.tabla.setRowCount(len(df))
        for i, row in df.iterrows():
            # Asumiendo que 'username', 'first name', 'last name', 'category name', 'course fullname'
            # son los nombres de las columnas en tu DataFrame.
            self.tabla.setItem(i, 0, QTableWidgetItem(str(row['username'])))  # Email
            self.tabla.setItem(i, 1, QTableWidgetItem(str(row['first name'])))  # Nombres
            self.tabla.setItem(i, 2, QTableWidgetItem(str(row['last name'])))  # Apellidos
            self.tabla.setItem(i, 3, QTableWidgetItem(str(row['category name'])))  # Carrera
            self.tabla.setItem(i, 4, QTableWidgetItem(str(row['course fullname'])))  # Nombre del Curso


    def limpiarTabla(self):
        self.tabla.clearContents
        self.tabla.setRowCount(0)
        # Restablecer cualquier otro estado de filtro aquí


    def filtrarPorCarreraSeleccionada(self):
        # Limpiar la tabla
        self.tabla.setRowCount(0)
        # Asegúrate de reemplazar 'ruta_al_archivo.csv' con la ruta real de tu archivo CSV
        ruta_al_archivo = 'UsuariosSinAcceder.csv'
        df = pd.read_csv(ruta_al_archivo)

        # Equivalencias entre las abreviaturas y los nombres completos de las carreras
        equivalencias = {
            "ADMEE": "ADMINISTRACIÓN DE EMPRESAS Y EMPRENDIMIENTO",
            "ADMIN FIN": "ADMINISTRACION FINANCIERA",
            "ADMIN": "ADMINISTRACIÓN",
            "AEIN": "ADMINISTRACION DE EMPRESAS E INTELIGENCIA DE NEGOCIOS",
            "DAW": "DESARROLLO DE APLICACIONES WEB",
            "DES INFANTIL": "DESARROLLO INFANTIL INTEGRAL",
            "ENFERMERIA": "ENFERMERÍA",
            "GEMD": "GESTIÓN ESTRATÉGICA DEL MARKETING DIGITAL",
            "GESRFIN": "GESTIÓN DE FINANZAS Y RIESGOS FINANCIEROS",
            "LAB CLINICO": "LABORATORIO CLÍNICO",
            "LOG-TRAN": "LOGÍSTICA Y TRANSPORTE",
            "MARKETING": "MARKETING",
            "R FISICA": "REHABILITACIÓN FÍSICA",
            "SICS": "SISTEMAS DE INFORMACIÓN Y CIBERSEGURIDAD"
        }

        # Obtener la selección actual del usuario
        eleccion = self.combo_box_carrera.currentIndex()
        categoria_seleccionada = list(equivalencias.keys())[eleccion]

        # Definir una expresión regular para buscar 'ADMIN' seguido por cualquier cosa excepto 'FIN'
        patron = f"^{categoria_seleccionada}(?! FIN)"

        # Filtrar el DataFrame por la presencia del patrón en 'category name'
        filtro = df['category name'].str.contains(patron, case=False, na=False, regex=True)
        filtered_df = df[filtro]

        # Contar las ocurrencias de la categoría seleccionada
        total_data = filtered_df.shape[0]

        # Mostrar los resultados
        self.label_filtrado.setText(f"FILTRADO:  {total_data}")

        # Actualizar la tabla con los datos filtrados
        self.mostrarResultadosFiltrados(filtered_df)

    def mostrarResultadosFiltrados(self, filtered_df):
        # Limpiar la tabla
        self.tabla.setRowCount(0)

        # Filtrar el DataFrame para seleccionar solo las columnas necesarias
        columnas_necesarias = ['username', 'first name', 'last name', 'category name', 'course shortname']
        filtered_df = filtered_df[columnas_necesarias]

        # Establecer las etiquetas del encabezado de las columnas en la tabla
        self.tabla.setHorizontalHeaderLabels(["Email", "Nombres", "Apellidos", "Carrera", "Curso"])

        # Obtener el número de filas del DataFrame filtrado
        num_filas = len(filtered_df)

        # Agregar las filas al QTableWidget
        for fila in range(num_filas):
            self.tabla.insertRow(fila)
            # Actualizar las celdas de la fila con los valores del DataFrame filtrado
            for columna, valor in enumerate(filtered_df.iloc[fila]):
                self.tabla.setItem(fila, columna, QTableWidgetItem(str(valor)))

    def filtrarPorModalidadSeleccionada(self):
        # Limpiar la tabla
        self.tabla.setRowCount(0)
        # Asegúrate de reemplazar 'ruta_al_archivo.csv' con la ruta real de tu archivo CSV
        ruta_al_archivo = 'UsuariosSinAcceder.csv'
        df = pd.read_csv(ruta_al_archivo)

        # Equivalencias entre las abreviaturas y los nombres completos de las carreras
        equivalencias = {
            "EL": "EN LÍNEA",
            "PRE": "PRESENCIAL",
        }

        # Obtener la elección del usuario
        eleccion = self.combo_box_modalidad.currentIndex()
        categoria_seleccionada = list(equivalencias.keys())[eleccion]

        # Filtrar el DataFrame por la presencia de la categoría seleccionada en 'category name'
        filtro = df['category name'].str.contains(categoria_seleccionada, case=False, na=False)
        filtered_df = df[filtro]

        # Contar las ocurrencias de la categoría seleccionada
        occurrences_count = filtered_df.shape[0]  # Número de filas que cumplen el criterio

        # Mostrar los resultados
        #QMessageBox.information(self, 'Resultados', f"Total de coincidencias para '{equivalencias[categoria_seleccionada]}': {occurrences_count}")
        total_data = filtered_df.shape[0]
        self.label_filtrado.setText(f"FILTRADO:  {total_data}")
        # Actualizar la tabla con los datos filtrados
        self.mostrarResultadosFiltrados(filtered_df)

    def filtrarPorCarreraYCurso(self):
        try:
            # Limpia la tabla
            self.tabla.setRowCount(0)
            
            # Ruta al archivo CSV
            ruta_al_archivo = 'UsuariosSinAcceder.csv'
            
            # Lectura del archivo CSV
            df = pd.read_csv(ruta_al_archivo)
            
            # Equivalencias entre las abreviaturas y los nombres completos de las carreras
            equivalencias = {
                "ADMEE": "ADMINISTRACIÓN DE EMPRESAS Y EMPRENDIMIENTO",
                "ADMIN FIN": "ADMINISTRACION FINANCIERA",
                "ADMIN": "ADMINISTRACIÓN",
                "AEIN": "ADMINISTRACION DE EMPRESAS E INTELIGENCIA DE NEGOCIOS",
                "DAW": "DESARROLLO DE APLICACIONES WEB",
                "DES INFANTIL": "DESARROLLO INFANTIL INTEGRAL",
                "ENFERMERIA": "ENFERMERÍA",
                "GEMD": "GESTIÓN ESTRATÉGICA DEL MARKETING DIGITAL",
                "GESRFIN": "GESTIÓN DE FINANZAS Y RIESGOS FINANCIEROS",
                "LAB CLINICO": "LABORATORIO CLÍNICO",
                "LOG-TRAN": "LOGÍSTICA Y TRANSPORTE",
                "MARKETING": "MARKETING",
                "R FISICA": "REHABILITACIÓN FÍSICA",
                "SICS": "SISTEMAS DE INFORMACIÓN Y CIBERSEGURIDAD"
            }
            
            # Elección del usuario
            eleccion_carrera = self.combo_box_carrera.currentIndex()
            categoria_seleccionada_carrera = list(equivalencias.keys())[eleccion_carrera]
            
            # Definir una expresión regular para buscar 'ADMIN' seguido por cualquier cosa excepto 'FIN'
            patron = f"^{categoria_seleccionada_carrera}(?! FIN)"
            
            # Filtrar el DataFrame por la presencia del patrón en 'category name'
            filtro_carrera = df['category name'].str.contains(patron, case=False, na=False, regex=True)
            df_filtrado_carrera = df[filtro_carrera]
            
            # Categoría seleccionada para el filtro de curso
            categoria_seleccionada_curso = self.combo_box_curso.currentText()
            
            # Aplicar filtro de curso sobre el DataFrame resultante de la carrera
            filtro_curso = df_filtrado_carrera['course shortname'].apply(lambda x: f"{x.split(' - ')[0].strip()} - [{x.split('- [')[1].strip()}") == categoria_seleccionada_curso
            df_filtrado_final = df_filtrado_carrera[filtro_curso]
            
            # Contar las ocurrencias de la categoría seleccionada
            occurrences_count = df_filtrado_final.shape[0]  # Número de filas que cumplen el criterio
            
            # Mostrar los resultados
            QMessageBox.information(self, 'Resultados', f"Total de coincidencias para '{equivalencias[categoria_seleccionada_carrera]}' y '{categoria_seleccionada_curso}': {occurrences_count}")
            
            # Actualizar la tabla con los datos filtrados
            self.mostrarResultadosFiltrados(df_filtrado_final)

        except Exception as e:
            QMessageBox.warning(self, 'Error', f"Ocurrió un error al filtrar los datos: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ventana = MiVentana()
    ventana.showMaximized()
    #ventana.setWindowFlag(Qt.FramelessWindowHint,False)
    sys.exit(app.exec_())
