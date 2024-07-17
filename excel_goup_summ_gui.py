import sys
from PyQt5 import QtWidgets, uic
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox


class ExcelSummarizerApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(ExcelSummarizerApp, self).__init__()
        uic.loadUi('gui.ui', self)

        # Predefined defaults
        self.default_grouping = ['HS Code', 'HS Code Description', 'COO', 'Invoice Document#']
        self.default_summation = ['Qty', 'Shipment Value', 'Total Weight']
        # Hide the vertical headers
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget_2.verticalHeader().setVisible(False)
        self.tableWidget_3.verticalHeader().setVisible(False)
        self.tableWidget_2.horizontalHeader().setVisible(False)
        self.tableWidget_3.horizontalHeader().setVisible(False)

        # Connect buttons
        self.pushButton_2.clicked.connect(self.load_excel_file)
        self.pushButton.clicked.connect(self.summarize_data)
        self.pushButton_4.clicked.connect(self.move_to_group)
        self.pushButton_3.clicked.connect(self.remove_from_group)
        self.pushButton_5.clicked.connect(self.move_to_summation)
        self.pushButton_6.clicked.connect(self.remove_from_summation)
        self.pushButton_7.clicked.connect(self.export_to_excel)

        # Data
        self.df = None

        # Show the GUI
        self.show()

    def load_excel_file(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load Excel File", "", "Excel files (*.xlsx *.xls)")
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.tableWidget.clear()  # Clear previous items
                self.tableWidget.setRowCount(len(self.df.columns))
                self.tableWidget.setColumnCount(1)  # Set to one column
                self.tableWidget.setHorizontalHeaderLabels(['Headers'])
                # Populate the table with headers
                for i, header in enumerate(self.df.columns):
                    self.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(header))

                if self.validate_default_fields():
                    self.populate_default_grouping()
                    self.populate_default_summation()
                    self.summarize_data()
                else:
                    QtWidgets.QMessageBox.warning(self, "Missing Fields", "The loaded file does not contain all the required fields.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load file: {e}")
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()
        self.tableWidget_4.resizeColumnsToContents()

    def validate_default_fields(self):
        # Check if all default grouping and summation fields are in the DataFrame
        missing_fields = [field for field in self.default_grouping + self.default_summation if field not in self.df.columns]
        if missing_fields:
            QtWidgets.QMessageBox.warning(self, "Missing Fields", f"The following fields are missing: {', '.join(missing_fields)}")
            return False
        return True

    def populate_default_grouping___(self):
        # Clear and then add defaults to tableWidget_2 if they exist in the DataFrame
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.setColumnCount(1)
        for header in self.default_grouping:
            if header in self.df.columns:
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
                self.tableWidget_2.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))

    def populate_default_summation___(self):
        # Clear and then add defaults to tableWidget_3 if they exist in the DataFrame
        self.tableWidget_3.setRowCount(0)
        self.tableWidget_3.setColumnCount(1)
        for header in self.default_summation:
            if header in self.df.columns:
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
                self.tableWidget_3.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))
    def populate_default_grouping___________(self):
        # Clear and then add defaults to tableWidget_2 if they exist in the DataFrame
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.setColumnCount(1)
        existing_items = [self.tableWidget_2.item(row, 0).text() for row in range(self.tableWidget_2.rowCount())]
        for header in self.default_grouping:
            if header in self.df.columns and header not in existing_items:
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
                self.tableWidget_2.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))

    def populate_default_grouping(self):
        # Clear and then add defaults to tableWidget_2 if they exist in the DataFrame
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.setColumnCount(1)

        # Collect existing items in tableWidget_2
        existing_items = [self.tableWidget_2.item(row, 0).text() for row in range(self.tableWidget_2.rowCount())]

        # Iterate through tableWidget to remove default grouping items
        for header in self.default_grouping:
            if header in self.df.columns and header not in existing_items:
                # Add item to tableWidget_2
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
                self.tableWidget_2.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))

                # Find and remove the item from tableWidget
                for row in range(self.tableWidget.rowCount()):
                    item = self.tableWidget.item(row, 0)
                    if item and item.text() == header:
                        self.tableWidget.removeRow(row)
                        break

        # Resize columns to fit the content
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()


    def populate_default_summation_____(self):
        # Clear and then add defaults to tableWidget_3 if they exist in the DataFrame
        self.tableWidget_3.setRowCount(0)
        self.tableWidget_3.setColumnCount(1)
        existing_items = [self.tableWidget_3.item(row, 0).text() for row in range(self.tableWidget_3.rowCount())]
        for header in self.default_summation:
            if header in self.df.columns and header not in existing_items:
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
                self.tableWidget_3.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))
                # Remove the row from tableWidget


    def populate_default_summation(self):
        # Clear and then add defaults to tableWidget_3 if they exist in the DataFrame
        self.tableWidget_3.setRowCount(0)
        self.tableWidget_3.setColumnCount(1)

        # Collect existing items in tableWidget_3
        existing_items = [self.tableWidget_3.item(row, 0).text() for row in range(self.tableWidget_3.rowCount())]

        # Iterate through tableWidget to remove default summation items
        for header in self.default_summation:
            if header in self.df.columns and header not in existing_items:
                # Add item to tableWidget_3
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
                self.tableWidget_3.setItem(row_position, 0, QtWidgets.QTableWidgetItem(header))

                # Find and remove the item from tableWidget
                for row in range(self.tableWidget.rowCount()):
                    item = self.tableWidget.item(row, 0)
                    if item and item.text() == header:
                        self.tableWidget.removeRow(row)
                        break

        # Resize columns to fit the content
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()



    def summarize_data(self):
        # Check if the DataFrame is loaded and not empty
        if self.df is None or self.df.empty:
            QtWidgets.QMessageBox.warning(self, "Data Not Loaded", "Please load an Excel file first.")
            return

        # Collect headers from tableWidget_2 (Group by fields)
        group_headers = [self.tableWidget_2.item(row, 0).text() for row in range(self.tableWidget_2.rowCount())]
        # Collect headers from tableWidget_3 (Summation fields)
        sum_headers = [self.tableWidget_3.item(row, 0).text() for row in range(self.tableWidget_3.rowCount())]

        if not group_headers or not sum_headers:
            QtWidgets.QMessageBox.warning(self, "Selection Required", "Please select fields for grouping and summation.")
            return

        try:
            # Grouping and summarizing data
            grouped_df = self.df.groupby(group_headers)[sum_headers].sum().reset_index()

            # Displaying results in tableWidget_4
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.setColumnCount(len(grouped_df.columns))
            self.tableWidget_4.setHorizontalHeaderLabels(grouped_df.columns.tolist())

            # Populate tableWidget_4 with formatted data
            for index, row in grouped_df.iterrows():
                row_position = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_position)
                for col_index, item in enumerate(row):
                    # Format item to string with three decimal places if it's a float
                    if isinstance(item, float):
                        formatted_item = f"{item:.3f}"
                    else:
                        formatted_item = str(item)
                    self.tableWidget_4.setItem(row_position, col_index, QtWidgets.QTableWidgetItem(formatted_item))

            # Resize columns to fit the content
            self.tableWidget_4.resizeColumnsToContents()

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error Summarizing Data", f"An error occurred: {e}")

    def move_to_group(self):
        # Get the selected items from tableWidget
        selected_items = self.tableWidget.selectedItems()

        if not selected_items:
            QtWidgets.QMessageBox.warning(self, "Selection Required", "Please select at least one header to move.")
            return

        # Collect unique rows to handle multiple selections per row if needed
        unique_rows = set(item.row() for item in selected_items)

        # Make sure tableWidget_2 is configured for a single column
        if self.tableWidget_2.columnCount() < 1:
            self.tableWidget_2.setColumnCount(1)  # Setup one column if not already set

        # Move selected items from tableWidget to tableWidget_2
        for row in sorted(unique_rows, reverse=True):  # Reverse to handle deletion correctly
            item = self.tableWidget.takeItem(row, 0)  # Take the item from the first column
            if item:
                # Inserting the item into tableWidget_2
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)
                self.tableWidget_2.setItem(row_position, 0, QtWidgets.QTableWidgetItem(item.text()))
                # Remove the row from tableWidget
                self.tableWidget.removeRow(row)

        # Optionally, if you want tableWidget_2 to have no headers
        self.tableWidget_2.horizontalHeader().setVisible(False)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()
        self.tableWidget_4.resizeColumnsToContents()

    def remove_from_group(self):
        # Get the selected items from tableWidget_2
        selected_items = self.tableWidget_2.selectedItems()

        if not selected_items:
            QtWidgets.QMessageBox.warning(self, "Selection Required", "Please select at least one header to remove.")
            return

        # Collect unique rows to handle multiple selections per row if needed
        unique_rows = set(item.row() for item in selected_items)

        # Prepare to add removed items back to tableWidget
        if self.tableWidget.columnCount() < 1:
            self.tableWidget.setColumnCount(1)  # Ensure there is at least one column

        # Remove selected items from tableWidget_2 and add them back to tableWidget
        for row in sorted(unique_rows, reverse=True):  # Reverse to handle deletion correctly
            item = self.tableWidget_2.takeItem(row, 0)  # Take the item from the first column
            if item:
                # Inserting the item back into tableWidget
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)
                self.tableWidget.setItem(row_position, 0, QtWidgets.QTableWidgetItem(item.text()))
                # Remove the row from tableWidget_2
                self.tableWidget_2.removeRow(row)

        # Update both tables to make sure all changes are visually represented
        self.tableWidget_2.viewport().update()
        self.tableWidget.viewport().update()
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()
        self.tableWidget_4.resizeColumnsToContents()

    def move_to_summation(self):
        # Get the selected items from tableWidget
        selected_items = self.tableWidget.selectedItems()

        if not selected_items:
            QtWidgets.QMessageBox.warning(self, "Selection Required", "Please select at least one header to move.")
            return

        # Collect unique rows to handle multiple selections per row if needed
        unique_rows = set(item.row() for item in selected_items)

        # Make sure tableWidget_3 is configured for a single column
        if self.tableWidget_3.columnCount() < 1:
            self.tableWidget_3.setColumnCount(1)  # Setup one column if not already set

        # Move selected items from tableWidget to tableWidget_3
        for row in sorted(unique_rows, reverse=True):  # Reverse to handle deletion correctly
            item = self.tableWidget.takeItem(row, 0)  # Take the item from the first column
            if item:
                # Inserting the item into tableWidget_3
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
                self.tableWidget_3.setItem(row_position, 0, QtWidgets.QTableWidgetItem(item.text()))
                # Remove the row from tableWidget
                self.tableWidget.removeRow(row)

        # Optionally, if you want tableWidget_3 to have no headers
        self.tableWidget_3.horizontalHeader().setVisible(False)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()
        self.tableWidget_4.resizeColumnsToContents()

    def remove_from_summation(self):
        # Get the selected items from tableWidget_3
        selected_items = self.tableWidget_3.selectedItems()

        if not selected_items:
            QtWidgets.QMessageBox.warning(self, "Selection Required", "Please select at least one header to remove.")
            return

        # Collect unique rows to handle multiple selections per row if needed
        unique_rows = set(item.row() for item in selected_items)

        # Prepare to add removed items back to tableWidget
        if self.tableWidget.columnCount() < 1:
            self.tableWidget.setColumnCount(1)  # Ensure there is at least one column

        # Remove selected items from tableWidget_3 and add them back to tableWidget
        for row in sorted(unique_rows, reverse=True):  # Reverse to handle deletion correctly
            item = self.tableWidget_3.takeItem(row, 0)  # Take the item from the first column
            if item:
                # Inserting the item back into tableWidget
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)
                self.tableWidget.setItem(row_position, 0, QtWidgets.QTableWidgetItem(item.text()))
                # Remove the row from tableWidget_3
                self.tableWidget_3.removeRow(row)

        # Update both tables to make sure all changes are visually represented
        self.tableWidget_3.viewport().update()
        self.tableWidget.viewport().update()
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.tableWidget_3.resizeColumnsToContents()
        self.tableWidget_4.resizeColumnsToContents()

    def export_to_excel(self):
        # Check if there are any rows to export
        if self.tableWidget_4.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "There is no data to export.")
            return

        # Prompt the user to select a file location and name for saving the Excel file
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)", options=options)
        
        if file_name:
            if not file_name.endswith('.xlsx'):
                file_name += '.xlsx'  # Ensure the file has the correct .xlsx extension
            
            # Create a DataFrame from the table data
            data = []
            headers = [self.tableWidget_4.horizontalHeaderItem(i).text() for i in range(self.tableWidget_4.columnCount())]
            
            for row in range(self.tableWidget_4.rowCount()):
                row_data = []
                for col in range(self.tableWidget_4.columnCount()):
                    item = self.tableWidget_4.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            
            df = pd.DataFrame(data, columns=headers)
            
            try:
                # Save the DataFrame to an Excel file
                df.to_excel(file_name, index=False)
                QMessageBox.information(self, "Success", "Data exported successfully to " + file_name)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = ExcelSummarizerApp()
    sys.exit(app.exec_())
