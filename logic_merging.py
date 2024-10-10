import pandas as pd
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QVBoxLayout, QWidget, QComboBox, QMessageBox, QLineEdit, QLabel, QHBoxLayout

class ExcelMerger(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Merger with Find and Replace")

        # ComboBoxes to select Key columns from main and secondary Excel files
        self.main_key_combobox = QComboBox()
        self.secondary_key_combobox = QComboBox()

        # Input fields for Find and Replace
        self.find_input = QLineEdit()
        self.replace_input = QLineEdit()

        # Labels for the inputs
        find_label = QLabel("Find:")
        replace_label = QLabel("Replace:")

        # Button to load main Excel file
        self.load_main_button = QPushButton("Load Main Excel")
        self.load_main_button.clicked.connect(self.load_main_excel)

        # Button to load secondary Excel file(s)
        self.load_secondary_button = QPushButton("Load Secondary Excel")
        self.load_secondary_button.clicked.connect(self.load_secondary_excel)

        # Button to merge files
        self.merge_button = QPushButton("Merge and Apply Replacements")
        self.merge_button.clicked.connect(self.merge_excels)

        # Layout for the find-replace inputs
        find_replace_layout = QHBoxLayout()
        find_replace_layout.addWidget(find_label)
        find_replace_layout.addWidget(self.find_input)
        find_replace_layout.addWidget(replace_label)
        find_replace_layout.addWidget(self.replace_input)

        # Main layout
        layout = QVBoxLayout()
        layout.addWidget(self.load_main_button)
        layout.addWidget(self.main_key_combobox)
        layout.addWidget(self.load_secondary_button)
        layout.addWidget(self.secondary_key_combobox)
        layout.addLayout(find_replace_layout)
        layout.addWidget(self.merge_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Placeholder for loaded Excel data
        self.main_data = None
        self.secondary_data = None

    def load_main_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Main Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                self.main_data = pd.read_excel(file_path)
                self.main_key_combobox.clear()
                self.main_key_combobox.addItems(self.main_data.columns)
                QMessageBox.information(self, "Success", "Main Excel file loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load main Excel file: {str(e)}")

    def load_secondary_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Secondary Excel File", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                self.secondary_data = pd.read_excel(file_path)
                self.secondary_key_combobox.clear()
                self.secondary_key_combobox.addItems(self.secondary_data.columns)
                QMessageBox.information(self, "Success", "Secondary Excel file loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load secondary Excel file: {str(e)}")

    def merge_excels(self):
        if self.main_data is None or self.secondary_data is None:
            QMessageBox.warning(self, "Warning", "Please load both main and secondary Excel files first!")
            return

        # Get selected key columns
        main_key = self.main_key_combobox.currentText()
        secondary_key = self.secondary_key_combobox.currentText()

        if not main_key or not secondary_key:
            QMessageBox.warning(self, "Warning", "Please select a valid key column from both files!")
            return

        # Merge the data on the selected keys
        merged_data = self.main_data.merge(self.secondary_data[[secondary_key, 'Profile', 'Value']], left_on=main_key, right_on=secondary_key, how='left')

        # Apply find and replace logic if Profile matches
        find_value = self.find_input.text()
        replace_value = self.replace_input.text()

        if find_value and replace_value:
            # Perform the replacement only if "Profile" values match
            merged_data.loc[merged_data['Profile_x'] == merged_data['Profile_y'], 'Value'] = merged_data['Value'].replace(find_value, replace_value)

        # Save the merged and modified data
        merged_data.to_excel("merged_file_with_replacements.xlsx", index=False)
        QMessageBox.information(self, "Success", "Excel files merged and saved as merged_file_with_replacements.xlsx!")

if __name__ == "__main__":
    app = QApplication([])
    window = ExcelMerger()
    window.show()
    app.exec()
