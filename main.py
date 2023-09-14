import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QComboBox, QLineEdit, QPushButton, QMessageBox, QVBoxLayout, QHBoxLayout, QWidget,QProgressBar
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtGui import QFont

from reportes import reporte_traspasos

class ReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Traspasos Supracare")

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
    
        self.month_label = QLabel("Mes", self)
        self.month_label.setFont(font)
        left_side.addWidget(self.month_label, alignment=Qt.AlignCenter)

        self.month_combo = QComboBox(self)
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setFont(combo_font)
        left_side.addWidget(self.month_combo, alignment=Qt.AlignCenter)

        self.year_label = QLabel("Año", self)
        self.year_label.setFont(combo_font)
        left_side.addWidget(self.year_label, alignment=Qt.AlignCenter)

        self.year_combo = QComboBox(self)
        self.year_combo.addItems([str(i) for i in range(2023, 2031)])
        self.year_combo.setFont(combo_font)
        left_side.addWidget(self.year_combo, alignment=Qt.AlignCenter)

        self.sheet_name_label = QLabel("Hoja", self)
        self.sheet_name_label.setFont(font)
        left_side.addWidget(self.sheet_name_label, alignment=Qt.AlignCenter)

        self.sheet_name_entry = QLineEdit(self)
        self.sheet_name_entry.setFont(font)
        self.sheet_name_entry.setAlignment(Qt.AlignCenter)
        left_side.addWidget(self.sheet_name_entry, alignment=Qt.AlignCenter)

        self.generate_button = QPushButton("Generar", self)
        self.generate_button.clicked.connect(self.generate_report)
        self.generate_button.setFont(font)
        left_side.addWidget(self.generate_button, alignment=Qt.AlignCenter)
        
        self.additional_image_label = QLabel(self)
        left_side.addWidget(self.additional_image_label, alignment=Qt.AlignLeft | Qt.AlignTop)

        if getattr(sys, 'frozen', False):
            b_dir = sys._MEIPASS
        else:
            b_dir = os.path.dirname(os.path.abspath(__file__))

        additional_image_path = os.path.join(b_dir,'data','supracare.png')
        additional_pixmap = QPixmap(additional_image_path)
        additional_pixmap = additional_pixmap.scaled(110, 110, Qt.KeepAspectRatio)
        self.additional_image_label.setPixmap(additional_pixmap)

        right_side = QVBoxLayout()
        self.image_label = QLabel(self)
        right_side.addWidget(self.image_label, alignment=Qt.AlignCenter)

        image_path = os.path.join(b_dir,'data','aguiladeoro.png')
        pixmap = QPixmap(image_path)
        pixmap = pixmap.scaled(250, 250, Qt.KeepAspectRatio)
        self.image_label.setPixmap(pixmap)
        self.image_label.setAlignment(Qt.AlignCenter)

        main_layout.addLayout(left_side)
        main_layout.addLayout(right_side)

        layout.addLayout(main_layout)

        self.setGeometry(200, 200, 550, 300)

    def generate_report(self):
        month = int(self.month_combo.currentText())  # Get the selected month
        year = int(self.year_combo.currentText())    # Get the selected year
        sheet_name = self.sheet_name_entry.text()    # Get the sheet name from the QLineEdit
        
        try:
            rec, key = reporte_traspasos(month, year, sheet_name)
            self.show_success_message(f"Se ha descargado la información correctamente, se procesaron {rec} registros.<br>Para acceder al reporte <a href='https://docs.google.com/spreadsheets/d/{key}/'>click aquí.</a> <br>Para acceder al Dashboard <a href='https://lookerstudio.google.com/u/0/reporting/9de240ac-c295-45c5-9e5e-9204688d523d/page/478ID'>click aquí</a>.")
        except Exception as e:
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
