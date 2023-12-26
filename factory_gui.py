import json
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
                             QSpacerItem, QSizePolicy, QComboBox, QLineEdit, QLabel, QFileDialog, QMessageBox)
import sys
import engine
import os

# Name of json file that will be created somewhere.
CONFIG_FILE = "FactoryConfig.json"


class FactoryGui(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Image Factory")

        # Lock window to this size.
        self.setFixedSize(800, 300)

        # Create buttons.
        process_button = QPushButton("Process", self)
        process_button.setFixedSize(200, 30)
        template_button = QPushButton("Browse", self)
        xlsx_button = QPushButton("Browse", self)
        out_button = QPushButton("Browse", self)
        ink_button = QPushButton("Browse", self)

        # Create a label and a dropdown list.
        extension_label = QLabel("File extension:", self)
        dropdown_list = QComboBox(self)
        dropdown_list.setObjectName("extensions")
        dropdown_list.addItem("png")
        dropdown_list.addItem("pdf")
        dropdown_list.addItem("svg")

        # Create a label and a field for entering an integer.
        dpi_label = QLabel("DPI:", self)
        dpi_field = QLineEdit(self)
        dpi_field.setObjectName("dpi")
        dpi_field.setPlaceholderText(" e.g. 500")

        # Create a label and a field for entering text.
        text_label = QLabel("File name:", self)
        text_field = QLineEdit(self)
        text_field.setObjectName("text")
        text_field.setPlaceholderText(" %VAR_Name% or different")

        # Create a label and a field for browsing files.
        svg_label = QLabel("Template directory:", self)
        svg_field = QLineEdit(self)
        svg_field.setObjectName("template")
        svg_field.setPlaceholderText("Path to svg file")

        # Create a label and a field for browsing files.
        xlsx_label = QLabel("Data directory:", self)
        xlsx_field = QLineEdit(self)
        xlsx_field.setObjectName("xlsx")
        xlsx_field.setPlaceholderText("Path to xlsx file")

        # Create a label and a field for browsing files.
        out_label = QLabel("Output directory:", self)
        out_field = QLineEdit(self)
        out_field.setObjectName("out")
        out_field.setPlaceholderText("Path to folder/directory for created files")

        # Create a label and a field for browsing files.
        ink_label = QLabel("Inkscape directory:", self)
        ink_field = QLineEdit(self)
        ink_field.setObjectName("ink")
        ink_field.setPlaceholderText("Path to inkscape.exe")

        # Connect button signals to functions.
        process_button.clicked.connect(self.process_button_work)
        template_button.clicked.connect(lambda: self.browse_button_file(svg_field, "template"))
        xlsx_button.clicked.connect(lambda: self.browse_button_file(xlsx_field, "xlsx"))
        out_button.clicked.connect(lambda: self.browse_button_directory(out_field, "out"))
        ink_button.clicked.connect(lambda: self.browse_button_file(ink_field, "ink"))

        # Create main layout where everything will be added.
        main_layout = QVBoxLayout()

        # Spaces between fields.
        main_layout.setSpacing(5)

        # Create a sub-layout for the File Extension.
        extension_layout = QHBoxLayout()
        extension_layout.addWidget(extension_label)
        extension_layout.addWidget(dropdown_list)
        # Set content margins to get dropdown list closer to label.
        extension_layout.setContentsMargins(0, 0, 570, 0)
        main_layout.addLayout(extension_layout)

        # Create a sub-layout for the DPI.
        dpi_layout = QHBoxLayout()
        dpi_layout.addWidget(dpi_label)
        dpi_layout.addWidget(dpi_field)
        # Set content margins to get dpi field smaller.
        dpi_layout.setContentsMargins(0, 0, 665, 0)
        main_layout.addLayout(dpi_layout)

        # Create a sub-layout for the Text.
        text_layout = QHBoxLayout()
        text_layout.addWidget(text_label)
        text_layout.addWidget(text_field)
        # Set content margins to get text field smaller.
        text_layout.setContentsMargins(0, 0, 513, 0)
        main_layout.addLayout(text_layout)

        # Create a sub-layout for the svg.
        svg_layout = QHBoxLayout()
        svg_layout.addWidget(svg_label)
        svg_layout.addWidget(svg_field)
        svg_layout.addWidget(template_button)
        main_layout.addLayout(svg_layout)

        # Create a sub-layout for the xls.
        xlsx_layout = QHBoxLayout()
        xlsx_layout.addWidget(xlsx_label)
        xlsx_layout.addWidget(xlsx_field)
        xlsx_layout.addWidget(xlsx_button)
        main_layout.addLayout(xlsx_layout)

        # Create a sub-layout for the out_directory.
        out_directory_layout = QHBoxLayout()
        out_directory_layout.addWidget(out_label)
        out_directory_layout.addWidget(out_field)
        out_directory_layout.addWidget(out_button)
        main_layout.addLayout(out_directory_layout)

        # Create a sub-layout for the ink_directory.
        ink_layout = QHBoxLayout()
        ink_layout.addWidget(ink_label)
        ink_layout.addWidget(ink_field)
        ink_layout.addWidget(ink_button)
        main_layout.addLayout(ink_layout)

        # Add a horizontal spacer item to center the "Process" button horizontally.
        main_layout.addItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        # Add the "Process" button to the main layout.
        main_layout.addWidget(process_button, alignment=Qt.AlignCenter)
        # Add another horizontal spacer item to center the "Process" button horizontally.
        main_layout.addItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Set the main layout for the window.
        self.setLayout(main_layout)
        # Take values from json and replace them
        self.receive_values_from_json(CONFIG_FILE, svg_field, xlsx_field, out_field, ink_field,
                                      text_field, dropdown_list, dpi_field)

    def process_button_work(self):

        # Get the selected value from the dropdown list.
        extension_value = self.findChild(QComboBox, "extensions").currentText()

        # Get the paths to the selected files.
        text_value = self.findChild(QLineEdit, "text").text()
        template_path = self.findChild(QLineEdit, "template").text()
        xlsx_path = self.findChild(QLineEdit, "xlsx").text()
        out_path = self.findChild(QLineEdit, "out").text()
        ink_path = self.findChild(QLineEdit, "ink").text()
        dpi_value = self.findChild(QLineEdit, "dpi").text()

        if not text_value:
            QMessageBox.warning(None, "Text not found", "Please provide name of file.", QMessageBox.Ok)
            return
        # Only PNG extension need dpi value.
        if extension_value == "png":
            if not dpi_value:
                QMessageBox.warning(None, "Integer not found", "Please provide DPI value.", QMessageBox.Ok)
                return
            else:
                try:
                    dpi_value = int(dpi_value)
                except ValueError:
                    QMessageBox.warning(None, "Integer not found", "DPI value isn't number, please provide number "
                                                                   "value.", QMessageBox.Ok)
                    return
        # SVG doesn't need inkscape to works.
        elif extension_value == "svg":
            if not ink_path:
                ink_path = "nothing path"

        # Start making process.
        engine.process_starter(template_path, xlsx_path, text_value, out_path,
                               extension_value, ink_path, dpi_value)

        # Save fields to the json.
        self.save_values_to_json("template", self.findChild(QLineEdit, "template").text())
        self.save_values_to_json("xlsx", self.findChild(QLineEdit, "xlsx").text())
        self.save_values_to_json("out", self.findChild(QLineEdit, "out").text())
        self.save_values_to_json("ink", self.findChild(QLineEdit, "ink").text())
        self.save_values_to_json("text_value", text_value)
        self.save_values_to_json("extension_value", extension_value)
        self.save_values_to_json("dpi_value", dpi_value)

    def browse_button_file(self, file_field, path_key):

        # Open a file dialog to select a file.
        file_dialog = QFileDialog.getOpenFileName(self, "Select a file")

        if file_dialog[0]:
            # Set the selected file path to the file field.
            file_field.setText(file_dialog[0])
            self.save_values_to_json(path_key, file_dialog[0])

    def browse_button_directory(self, file_field, path_key):

        # Open a file dialog to select a file.
        selected_directory = QFileDialog.getExistingDirectory(self, "Select a directory")

        if selected_directory:
            # Set the selected directory to the directory field.
            file_field.setText(selected_directory)
            self.save_values_to_json(path_key, selected_directory)

    def receive_values_from_json(self, json_file, svg, xlsx, out, ink, text, extension, dpi):
        # Search for path to json settings if not found created new.
        if not os.path.exists(json_file):
            self.saved_paths = self.create_json()
        else:
            self.saved_paths = self.load_json()
            # Set values from json to fields.
            svg.setText(self.saved_paths.get("template", ""))
            xlsx.setText(self.saved_paths.get("xlsx", ""))
            out.setText(self.saved_paths.get("out", ""))
            ink.setText(self.saved_paths.get("ink", ""))
            text.setText(self.saved_paths.get("text_value", ""))
            extension.setCurrentText(self.saved_paths.get("extension_value", ""))
            dpi.setText(str(self.saved_paths.get("dpi_value", "")))

    def save_values_to_json(self, key, value):
        # Save a key-value pair to the internal state
        self.saved_paths[key] = value

    def save_json(self):
        # Save configuration settings to a file (JSON format)
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(self.saved_paths, config_file)

    def load_json(self):
        with open(CONFIG_FILE, 'r') as config_file:
            return json.load(config_file)

    def create_json(self):
        factory_json = {
            "template": "",
            "xlsx": "",
            "out": "",
            "ink": "",
            "text_value": "",
            "extension_value": "",
            "dpi_value": ""
        }
        return factory_json

    def closeEvent(self, event):

        # Save fields to the json.
        self.save_values_to_json("template", self.findChild(QLineEdit, "template").text())
        self.save_values_to_json("xlsx", self.findChild(QLineEdit, "xlsx").text())
        self.save_values_to_json("out", self.findChild(QLineEdit, "out").text())
        self.save_values_to_json("ink", self.findChild(QLineEdit, "ink").text())
        self.save_values_to_json("text_value", self.findChild(QLineEdit, "text").text())
        self.save_values_to_json("extension_value", self.findChild(QComboBox, "extensions").currentText())
        self.save_values_to_json("dpi_value", self.findChild(QLineEdit, "dpi").text())

        self.save_json()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FactoryGui()
    window.show()
    sys.exit(app.exec_())
