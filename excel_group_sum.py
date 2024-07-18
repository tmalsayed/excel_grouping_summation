import sys
from PyQt5 import QtWidgets
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from datetime import datetime

class ExcelSummarizerApp:
    def __init__(self):
        # Predefined defaults
        self.default_grouping = ['HS Code', 'HS Code Description', 'COO', 'Invoice Document#']
        self.default_summation = ['Qty', 'Shipment Value', 'Total Weight']
        self.df = None

    def load_excel_files(self):
        app = QtWidgets.QApplication(sys.argv)  # Ensure QApplication is initialized
        files, _ = QFileDialog.getOpenFileNames(None, "Load Excel Files", "", "Excel files (*.xlsx *.xls)")
        if files:
            for file in files:
                try:
                    self.df = pd.read_excel(file)
                    if self.validate_default_fields():
                        self.summarize_data(file)
                    else:
                        self.show_message("Error", f"Missing required fields in {file}")
                except Exception as e:
                    self.show_message("Error", f"Failed to load file {file}: {e}")
            self.show_message("Success", "All files have been processed.")
        else:
            self.show_message("Information", "No files were selected.")
        app.quit()  # Close the application after processing

    def validate_default_fields(self):
        missing_fields = [field for field in self.default_grouping + self.default_summation if field not in self.df.columns]
        if missing_fields:
            return False
        return True

    def summarize_data(self, input_file):
        try:
            grouped_df = self.df.groupby(self.default_grouping)[self.default_summation].sum().reset_index()
            output_file = f"{input_file}_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            grouped_df.to_excel(output_file, index=False)
        except Exception as e:
            self.show_message("Error", f"Error summarizing data for file {input_file}: {e}")

    def show_message(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec_()

if __name__ == "__main__":
    summarizer = ExcelSummarizerApp()
    summarizer.load_excel_files()
