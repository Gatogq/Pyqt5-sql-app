import sys
from os.path import dirname, abspath, join, basename
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QComboBox, QLineEdit, QPushButton, QMessageBox, QVBoxLayout, QHBoxLayout, QWidget, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtGui import QFont

from reportes import reporte_retardos

class ReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Retardos Supracare")

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        main_layout = QHBoxLayout()
        left_side = QVBoxLayout()

        font = self.font()
        font.setPointSize(10)
        combo_font = QFont()
        combo_font.setPointSize(10)

        # Option 1: Select Excel File 1
        self.file1_label = QLabel("Archivo Excel 1", self)
        self.file1_label.setFont(font)
        left_side.addWidget(self.file1_label, alignment=Qt.AlignCenter)

        self.file1_button = QPushButton("...", self)
        self.file1_button.clicked.connect(self.select_file1)
        self.file1_button.setFont(font)
        left_side.addWidget(self.file1_button, alignment=Qt.AlignCenter)
        self.file1_path_label = QLabel("")  # Label to display the selected file name
        self.file1_path_label.setFont(font)
        left_side.addWidget(self.file1_path_label, alignment=Qt.AlignCenter)

        self.file1_path = ""

        # Option 2: Select Excel File 2
        self.file2_label = QLabel("Archivo Excel 2", self)
        self.file2_label.setFont(font)
        left_side.addWidget(self.file2_label, alignment=Qt.AlignCenter)

        self.file2_button = QPushButton("...", self)
        self.file2_button.clicked.connect(self.select_file2)
        self.file2_button.setFont(font)
        left_side.addWidget(self.file2_button, alignment=Qt.AlignCenter)
        self.file2_path_label = QLabel("")  # Label to display the selected file name
        self.file2_path_label.setFont(font)
        left_side.addWidget(self.file2_path_label, alignment=Qt.AlignCenter)

        self.file2_path = ""

        # Option 3: Select Excel File 3
        self.file3_label = QLabel("Directorio", self)
        self.file3_label.setFont(font)
        left_side.addWidget(self.file3_label, alignment=Qt.AlignCenter)

        self.file3_button = QPushButton("...", self)
        self.file3_button.clicked.connect(self.select_file3)
        self.file3_button.setFont(font)
        left_side.addWidget(self.file3_button, alignment=Qt.AlignCenter)
        self.file3_path_label = QLabel("")  # Label to display the selected file name
        self.file3_path_label.setFont(font)
        left_side.addWidget(self.file3_path_label, alignment=Qt.AlignCenter)

        self.file3_path = ""

        self.month_label = QLabel("Mes", self)
        self.month_label.setFont(font)
        left_side.addWidget(self.month_label, alignment=Qt.AlignCenter)

        self.month_combo = QComboBox(self)
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setFont(combo_font)
        left_side.addWidget(self.month_combo, alignment=Qt.AlignCenter)

        # Option 4: Sheet Name (Same as before)
        self.sheet_name_label = QLabel("Hoja", self)
        self.sheet_name_label.setFont(font)
        left_side.addWidget(self.sheet_name_label, alignment=Qt.AlignCenter)

        self.sheet_name_entry = QLineEdit(self)
        self.sheet_name_entry.setFont(font)
        self.sheet_name_entry.setAlignment(Qt.AlignCenter)
        left_side.addWidget(self.sheet_name_entry, alignment=Qt.AlignCenter)

        # Generate Button
        self.generate_button = QPushButton("Generar", self)
        self.generate_button.clicked.connect(self.generate_report)
        self.generate_button.setFont(font)
        left_side.addWidget(self.generate_button, alignment=Qt.AlignCenter)

        # Additional image (same as before)
        self.additional_image_label = QLabel(self)
        left_side.addWidget(self.additional_image_label, alignment=Qt.AlignLeft | Qt.AlignTop)

        if getattr(sys, 'frozen', False):
            b_dir = sys._MEIPASS
        else:
            b_dir = dirname(abspath(__file__))

        additional_image_path = join(b_dir, 'data', 'supracare.png')
        additional_pixmap = QPixmap(additional_image_path)
        additional_pixmap = additional_pixmap.scaled(110, 110, Qt.KeepAspectRatio)
        self.additional_image_label.setPixmap(additional_pixmap)

        # Right side (image, same as before)
        right_side = QVBoxLayout()
        self.image_label = QLabel(self)
        right_side.addWidget(self.image_label, alignment=Qt.AlignCenter)

        image_path = join(b_dir, 'data', 'aguiladeoro.png')
        pixmap = QPixmap(image_path)
        pixmap = pixmap.scaled(250, 250, Qt.KeepAspectRatio)
        self.image_label.setPixmap(pixmap)
        self.image_label.setAlignment(Qt.AlignCenter)

        main_layout.addLayout(left_side)
        main_layout.addLayout(right_side)

        layout.addLayout(main_layout)

        self.setGeometry(200, 200, 550, 300)

    def select_file1(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file1, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel 1", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file1:
            self.file1_path = file1
            self.file1_path_label.setText(basename(file1))  # Display the selected file name
            self.file1_button.setEnabled(False)

    def select_file2(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file2, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel 2", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file2:
            self.file2_path = file2
            self.file2_path_label.setText(basename(file2))  # Display the selected file name
            self.file2_button.setEnabled(False)

    def select_file3(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file3, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo Excel 3", "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file3:
            self.file3_path = file3
            self.file3_path_label.setText(basename(file3))  # Display the selected file name
            self.file3_button.setEnabled(False)

    def generate_report(self):
        # Get the selected sheet name
        sheet_name = self.sheet_name_entry.text()
        month = int(self.month_combo.currentText())

        # Check if all three file paths are selected
        if not self.file1_path or not self.file2_path or not self.file3_path:
            self.show_error_message("Por favor, seleccione los tres archivos Excel.")
            return

        try:
            # Call the reporte_retardos function with the selected file paths and sheet_name
            rec, sheet_key = reporte_retardos(self.file1_path, self.file2_path, self.file3_path, sheet_name,month)

            # Show a success message with the number of records processed and the sheet key
            self.show_success_message(f"Se ha descargado la información correctamente, se procesaron {rec} registros."
                                  f"<br>Para acceder al reporte <a href='https://docs.google.com/spreadsheets/d/{sheet_key}/'>click aquí.</a> "
                                  f"<br>Para acceder al Dashboard <a href='https://lookerstudio.google.com/u/0/reporting/2fc11923-94e3-4454-ab7c-bf9916a0744d'>click aquí</a>.")
        except Exception as e:
            # Show an error message if an exception occurs
            self.show_error_message(f"Error: {str(e)}")

    def show_success_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(message)
        msg_box.setWindowTitle("Success")
        msg_box.exec_()

    def show_error_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec_()

def main():
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
