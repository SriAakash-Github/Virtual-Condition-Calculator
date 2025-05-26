import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QGridLayout, QFileDialog, QComboBox, QMessageBox,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QHeaderView
)
from PyQt5.QtCore import Qt, QRegExp
from PyQt5.QtGui import QDoubleValidator, QRegExpValidator, QPalette, QColor
import pandas as pd

class VirtualConditionCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Virtual Condition Calculator")
        self.entries = []
        self.init_ui()

    def init_ui(self):
        validator = QDoubleValidator()
        letter_validator = QRegExpValidator(QRegExp("[A-Za-z]"))

        main_layout = QVBoxLayout()
        form_layout = QGridLayout()

        self.nominal_input = QLineEdit()
        self.nominal_input.setValidator(validator)
        self.upper_limit_input = QLineEdit()
        self.upper_limit_input.setValidator(validator)
        self.lower_limit_input = QLineEdit()
        self.lower_limit_input.setValidator(validator)
        self.tolerance_input = QLineEdit()
        self.tolerance_input.setValidator(validator)
        self.datum_input = QLineEdit()
        self.datum_input.setMaxLength(1)
        self.datum_input.setValidator(letter_validator)

        self.feature_type = QComboBox()
        self.feature_type.addItems(["Pin Size", "Hole Size"])

        # Focus movement
        self.nominal_input.returnPressed.connect(self.upper_limit_input.setFocus)
        self.upper_limit_input.returnPressed.connect(self.lower_limit_input.setFocus)
        self.lower_limit_input.returnPressed.connect(self.tolerance_input.setFocus)
        self.tolerance_input.returnPressed.connect(self.datum_input.setFocus)
        self.datum_input.returnPressed.connect(self.feature_type.setFocus)
        self.feature_type.activated.connect(self.focus_add_button)

        # Live update
        for field in [self.nominal_input, self.upper_limit_input,
                      self.lower_limit_input, self.tolerance_input]:
            field.textChanged.connect(self.calculate_virtual_condition)
        self.feature_type.currentIndexChanged.connect(self.calculate_virtual_condition)

        # Labels
        self.vc_75 = QLabel("VC @ 75%: ")
        self.vc_80 = QLabel("VC @ 80%: ")
        self.vc_90 = QLabel("VC @ 90%: ")
        self.vc_100 = QLabel("VC @ 100%: ")

        form_layout.addWidget(QLabel("Nominal Size"), 0, 0)
        form_layout.addWidget(self.nominal_input, 0, 1)
        form_layout.addWidget(QLabel("Upper Limit (+)"), 1, 0)
        form_layout.addWidget(self.upper_limit_input, 1, 1)
        form_layout.addWidget(QLabel("Lower Limit (-)"), 2, 0)
        form_layout.addWidget(self.lower_limit_input, 2, 1)
        form_layout.addWidget(QLabel("Position Tolerance"), 3, 0)
        form_layout.addWidget(self.tolerance_input, 3, 1)
        form_layout.addWidget(QLabel("Datum (Letter)"), 4, 0)
        form_layout.addWidget(self.datum_input, 4, 1)
        form_layout.addWidget(QLabel("Feature Type"), 5, 0)
        form_layout.addWidget(self.feature_type, 5, 1)

        form_layout.addWidget(self.vc_75, 6, 0, 1, 2)
        form_layout.addWidget(self.vc_80, 7, 0, 1, 2)
        form_layout.addWidget(self.vc_90, 8, 0, 1, 2)
        form_layout.addWidget(self.vc_100, 9, 0, 1, 2)

        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Entry")
        self.add_btn.clicked.connect(self.add_entry)
        self.add_btn.setStyleSheet("background-color: #2ecc71; color: white; font-weight: bold;")

        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self.delete_selected_entry)
        delete_btn.setStyleSheet("background-color: #e74c3c; color: white; font-weight: bold;")

        save_btn = QPushButton("Save All to Excel")
        save_btn.clicked.connect(self.save_results_to_excel)
        save_btn.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")

        btn_layout.addWidget(self.add_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(save_btn)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "Nominal Size", "Upper Limit (+)", "Lower Limit (-)",
            "position Tolerance", "Feature Type", "Datum",
            "VC @ 75%", "VC @ 80%", "VC @ 90%", "VC @ 100%"
        ])
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.setSortingEnabled(True)
        self.table.cellChanged.connect(self.edit_table_entry)

        # Force column size to show "Specified Tolerance"
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setMinimumSectionSize(150)

        main_layout.addLayout(form_layout)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(QLabel("Stored Entries:"))
        main_layout.addWidget(self.table)

        self.setLayout(main_layout)

        self.setStyleSheet("""
            QLineEdit, QComboBox {
                background-color: #2c2c2c;
                color: white;
                border: 1px solid #555;
                padding: 4px;
                border-radius: 4px;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #00bc8c;
                background-color: #3a3a3a;
            }
            QLabel {
                color: white;
            }
            QTableWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                gridline-color: #555;
            }
            QHeaderView::section {
                background-color: #333;
                color: white;
                padding: 4px;
                border: 1px solid #444;
            }
        """)

    def focus_add_button(self):
        self.add_btn.setFocus()

    def calculate_virtual_condition(self):
        try:
            self.nominal = float(self.nominal_input.text() or 0)
            self.upper = float(self.upper_limit_input.text() or 0)
            self.lower = float(self.lower_limit_input.text() or 0)
            self.tolerance = float(self.tolerance_input.text() or 0)
            feature_type = self.feature_type.currentText()

            self.mmc_size = (
                self.nominal - self.lower if feature_type == "Pin Size"
                else self.nominal + self.lower
            )

            self.vc_75_val = self.mmc_size - self.tolerance * 0.75
            self.vc_80_val = self.mmc_size - self.tolerance * 0.80
            self.vc_90_val = self.mmc_size - self.tolerance * 0.90
            self.vc_100_val = self.mmc_size - self.tolerance * 1.00

            self.vc_75.setText(f"VC @ 75%: {self.vc_75_val:.3f}")
            self.vc_80.setText(f"VC @ 80%: {self.vc_80_val:.3f}")
            self.vc_90.setText(f"VC @ 90%: {self.vc_90_val:.3f}")
            self.vc_100.setText(f"VC @ 100%: {self.vc_100_val:.3f}")
        except ValueError:
            self.vc_75.setText("VC @ 75%: —")
            self.vc_80.setText("VC @ 80%: —")
            self.vc_90.setText("VC @ 90%: —")
            self.vc_100.setText("VC @ 100%: —")

    def add_entry(self):
        try:
            self.calculate_virtual_condition()
            datum = self.datum_input.text().upper() or "-"
            entry = [
                self.nominal, self.upper, self.lower, self.tolerance,
                self.feature_type.currentText(), datum,
                round(self.vc_75_val, 3), round(self.vc_80_val, 3),
                round(self.vc_90_val, 3), round(self.vc_100_val, 3)
            ]
            self.entries.append(entry)
            self.update_table()
            QMessageBox.information(self, "Entry Added", "Entry successfully added.")
        except AttributeError:
            QMessageBox.warning(self, "Invalid Data", "Please fill all fields properly.")

    def update_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.entries))
        for row, entry in enumerate(self.entries):
            for col, value in enumerate(entry):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, col, item)
        self.table.blockSignals(False)

    def delete_selected_entry(self):
        selected = self.table.currentRow()
        if selected >= 0:
            self.entries.pop(selected)
            self.update_table()
        else:
            QMessageBox.warning(self, "No Selection", "Please select a row to delete.")

    def edit_table_entry(self, row, col):
        try:
            new_value = self.table.item(row, col).text()
            if col in [0, 1, 2, 3, 6, 7, 8, 9]:
                new_value = float(new_value)
            self.entries[row][col] = new_value
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid number.")
            self.update_table()

    def save_results_to_excel(self):
        if not self.entries:
            QMessageBox.warning(self, "No Entries", "Add entries before saving.")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save All Entries", "", "Excel Files (*.xlsx)")
        if save_path:
            df = pd.DataFrame(self.entries, columns=[
                "Nominal Size", "Upper Limit (+)", "Lower Limit (-)",
                "Position Tolerance", "Feature Type", "Datum",
                "VC @ 75%", "VC @ 80%", "VC @ 90%", "VC @ 100%"
            ])
            df.to_excel(save_path, index=False)
            QMessageBox.information(self, "Saved", f"{len(self.entries)} entries saved.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    QApplication.setStyle("Fusion")

    dark_palette = QPalette()
    dark_palette.setColor(QPalette.Window, QColor(30, 30, 30))
    dark_palette.setColor(QPalette.WindowText, Qt.white)
    dark_palette.setColor(QPalette.Base, QColor(20, 20, 20))
    dark_palette.setColor(QPalette.AlternateBase, QColor(30, 30, 30))
    dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
    dark_palette.setColor(QPalette.ToolTipText, Qt.white)
    dark_palette.setColor(QPalette.Text, Qt.white)
    dark_palette.setColor(QPalette.Button, QColor(45, 45, 45))
    dark_palette.setColor(QPalette.ButtonText, Qt.white)
    dark_palette.setColor(QPalette.Highlight, QColor(0, 188, 140))
    dark_palette.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(dark_palette)

    window = VirtualConditionCalculator()
    window.show()
    sys.exit(app.exec_())
