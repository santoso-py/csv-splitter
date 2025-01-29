import sys
import os
import duckdb
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QLabel, QLineEdit, QComboBox, QTextEdit, QProgressBar, QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor


class CSVSplitterApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Setup Window
        self.setWindowTitle("CSV Splitter")
        self.setGeometry(200, 200, 700, 500)

        # Central Widget & Layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        # File CSV Path
        self.csv_path_label = QLabel("File CSV Path:")
        self.csv_path_entry = QLineEdit()
        self.csv_path_entry.setReadOnly(True)
        self.browse_csv_button = QPushButton("Browse")
        self.browse_csv_button.clicked.connect(self.browse_csv_file)

        # Column Selection
        self.column_label = QLabel("Select Column to Split By:")
        self.column_combo = QComboBox()

        # Folder Path
        self.folder_path_label = QLabel("Folder Path for Results:")
        self.folder_path_entry = QLineEdit()
        self.folder_path_entry.setReadOnly(True)
        self.browse_folder_button = QPushButton("Browse")
        self.browse_folder_button.clicked.connect(self.browse_folder)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        # Log Text
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # Process Button
        self.process_button = QPushButton("Process Data")
        self.process_button.clicked.connect(self.process_data)

        # Add Widgets to Layout
        layout.addWidget(self.csv_path_label)
        layout.addWidget(self.csv_path_entry)
        layout.addWidget(self.browse_csv_button)
        layout.addWidget(self.column_label)
        layout.addWidget(self.column_combo)
        layout.addWidget(self.folder_path_label)
        layout.addWidget(self.folder_path_entry)
        layout.addWidget(self.browse_folder_button)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.log_text)
        layout.addWidget(self.process_button)

    def browse_csv_file(self):
        """Memilih file CSV dan membaca kolomnya."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.csv_path_entry.setText(file_path)
            self.load_csv_columns(file_path)

    def load_csv_columns(self, file_path):
        """Membaca kolom dari CSV dan menampilkan dalam dropdown."""
        try:
            con = duckdb.connect()
            df_sample = con.execute(f"SELECT * FROM read_csv_auto('{file_path}') LIMIT 1").fetchdf()
            con.close()
            self.column_combo.clear()
            self.column_combo.addItems(df_sample.columns.tolist())
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read CSV columns: {e}")

    def browse_folder(self):
        """Memilih folder untuk menyimpan hasil splitting."""
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path_entry.setText(folder_path)

    def process_data(self):
        """Proses splitting berdasarkan kolom yang dipilih."""
        file_csv_path = self.csv_path_entry.text()
        folder_path = self.folder_path_entry.text()
        column_name = self.column_combo.currentText()

        if not file_csv_path or not folder_path or not column_name:
            QMessageBox.warning(self, "Warning", "Please select a CSV file, column, and output folder!")
            return

        try:
            con = duckdb.connect()
            df = con.execute(f"SELECT * FROM read_csv_auto('{file_csv_path}')").fetchdf()
            con.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading CSV file: {e}")
            return

        if column_name not in df.columns:
            QMessageBox.critical(self, "Error", "Selected column does not exist in CSV!")
            return

        unique_values = df[column_name].unique()
        total_groups = len(unique_values)
        self.progress_bar.setMaximum(total_groups)
        self.progress_bar.setValue(0)

        self.log_text.clear()
        self.log_text.append("Starting the splitting process...\n")

        for i, value in enumerate(unique_values):
            group = df[df[column_name] == value]
            file_name = f"{value}.xlsx"
            file_path = os.path.join(folder_path, file_name)
            group.to_excel(file_path, index=False)
            self.log_text.append(f"Saved: {file_name}")
            self.progress_bar.setValue(i + 1)
            QApplication.processEvents()

        self.log_text.append("\n🎉 Splitting completed successfully!")
        QMessageBox.information(self, "Success", "All processes completed successfully!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CSVSplitterApp()
    window.show()
    sys.exit(app.exec())
