import sys
import os
import pandas as pd
import logging
from datetime import datetime
from slypGenarater import generate_payslip
from print_manager import PayslipPrintManager  # Changed to PayslipPrintManager
from PyQt5.QtWidgets import (
   QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
   QFileDialog, QLabel, QListWidget, QTableWidget, QTableWidgetItem, QSplitter,
   QMessageBox, QTextEdit, QDialog, QComboBox, QLineEdit
)
from PyQt5.QtCore import Qt
from config import PayslipConfig


# Simplified logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger('ExcelViewer')


class ExcelSheetViewer(QMainWindow):
   def __init__(self):
       super().__init__()
       self.setWindowTitle("Excel Sheet Viewer")
       self.setGeometry(100, 100, 1200, 800)  # Made window larger


       # Initialize printer
       self.printer = PayslipPrintManager(self)  # Use PayslipPrintManager instead of PayslipPDFGenerator


       # Main widget and layout
       self.central_widget = QWidget()
       self.setCentralWidget(self.central_widget)
       self.layout = QVBoxLayout()
       self.central_widget.setLayout(self.layout)


       # File selection
       file_container = QWidget()
       file_layout = QHBoxLayout()
       file_container.setLayout(file_layout)
      
       self.file_label = QLabel("No Excel file selected")
       file_layout.addWidget(self.file_label)


       self.select_file_button = QPushButton("Select Excel File")
       self.select_file_button.clicked.connect(self.select_file)
       file_layout.addWidget(self.select_file_button)
      
       self.layout.addWidget(file_container)


       # Create main vertical splitter
       self.main_splitter = QSplitter(Qt.Vertical)
       self.layout.addWidget(self.main_splitter)


       # Create top horizontal splitter
       self.top_splitter = QSplitter(Qt.Horizontal)
       self.main_splitter.addWidget(self.top_splitter)


       # Sheet list display (left side)
       self.sheet_list_container = QWidget()
       self.sheet_list_layout = QVBoxLayout()
       self.sheet_list_container.setLayout(self.sheet_list_layout)
      
       self.sheet_list_label = QLabel("Available Sheets:")
       self.sheet_list_layout.addWidget(self.sheet_list_label)
      
       self.sheet_list = QListWidget()
       self.sheet_list.itemClicked.connect(self.on_sheet_selected)
       self.sheet_list_layout.addWidget(self.sheet_list)
      
       # Add the sheet list to the left side of top splitter
       self.top_splitter.addWidget(self.sheet_list_container)


       # Payslip preview (right side)
       self.payslip_container = QWidget()
       self.payslip_layout = QVBoxLayout()
       self.payslip_container.setLayout(self.payslip_layout)
      
       self.payslip_label = QLabel("Payslip Preview:")
       self.payslip_layout.addWidget(self.payslip_label)
      
       self.payslip_preview = QTextEdit()
       self.payslip_preview.setReadOnly(True)
       self.payslip_preview.setFont(QApplication.font("Monospace"))
       self.payslip_layout.addWidget(self.payslip_preview)
      
       # Add payslip preview to the right side of top splitter
       self.top_splitter.addWidget(self.payslip_container)
      
       # Set initial top splitter sizes (30% for sheet list, 70% for payslip)
       self.top_splitter.setSizes([300, 700])


       # Data table display (bottom)
       self.table_container = QWidget()
       self.table_layout = QVBoxLayout()
       self.table_container.setLayout(self.table_layout)
      
       self.table_label = QLabel("Sheet data:")
       self.table_layout.addWidget(self.table_label)
      
       self.data_table = QTableWidget()
       self.table_layout.addWidget(self.data_table)
      
       # Add the table to the bottom of main splitter
       self.main_splitter.addWidget(self.table_container)
      
       # Set initial main splitter sizes (50% for top, 50% for bottom)
       self.main_splitter.setSizes([400, 400])
      
       # Status and buttons
       bottom_container = QWidget()
       bottom_layout = QHBoxLayout()
       bottom_container.setLayout(bottom_layout)
      
       self.status_label = QLabel("Status: Ready")
       bottom_layout.addWidget(self.status_label)


       button_container = QWidget()
       button_layout = QHBoxLayout()
       button_container.setLayout(button_layout)


       self.generate_payslip_button = QPushButton("Generate Payslip & Save PDF")
       self.generate_payslip_button.clicked.connect(self.generate_selected_payslip)
       self.generate_payslip_button.setEnabled(False)
       button_layout.addWidget(self.generate_payslip_button)


       self.print_selected_button = QPushButton("Print Selected Payslip")
       self.print_selected_button.clicked.connect(self.print_selected_payslip)
       self.print_selected_button.setEnabled(False)
       button_layout.addWidget(self.print_selected_button)


       self.generate_bulk_button = QPushButton("Generate All Payslips & PDFs")
       self.generate_bulk_button.clicked.connect(self.generate_bulk_payslips)
       self.generate_bulk_button.setEnabled(False)
       button_layout.addWidget(self.generate_bulk_button)


       self.print_bulk_button = QPushButton("Print All Payslips")
       self.print_bulk_button.clicked.connect(self.print_bulk_payslips)
       self.print_bulk_button.setEnabled(False)
       button_layout.addWidget(self.print_bulk_button)
      
       # Add Configure button
       self.configure_button = QPushButton("Configure Custom Fields")
       self.configure_button.clicked.connect(self.show_config_dialog)
       button_layout.addWidget(self.configure_button)
      
       bottom_layout.addWidget(button_container)
       self.layout.addWidget(bottom_container)


       self.selected_file = None
       self.current_df = None
       self.current_headers = []  # Add this line after self.current_df initialization


   def select_file(self):
       file_name, _ = QFileDialog.getOpenFileName(
           self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
       )
       if file_name:
           self.selected_file = file_name
           self.file_label.setText(f"Selected File: {os.path.basename(file_name)}")
           try:
               # Load Excel file to get sheet names
               xls = pd.ExcelFile(file_name)
               sheet_names = xls.sheet_names
              
               # Clear and update the sheet list
               self.sheet_list.clear()
               self.sheet_list.addItems(sheet_names)
              
               # Clear the table
               self.data_table.clear()
               self.data_table.setRowCount(0)
               self.data_table.setColumnCount(0)
               self.table_label.setText("Sheet data:")
              
               self.status_label.setText(
                   f"Status: Excel file loaded successfully. Found {len(sheet_names)} sheets."
               )
           except Exception as e:
               error_msg = f"Error loading Excel file: {str(e)}"
               logger.error(error_msg)
               self.status_label.setText(f"Status: {error_msg}")
               self.sheet_list.clear()
               self.data_table.clear()
       else:
           logger.info("No file selected by user")
           self.file_label.setText("No Excel file selected")
           self.status_label.setText("Status: No file selected.")
           self.sheet_list.clear()
           self.data_table.clear()
           self.table_label.setText("Sheet data:")


   def on_sheet_selected(self, item):
       sheet_name = item.text()
       logger.info(f"Sheet selected: {sheet_name}")
       self.show_sheet_data(str(sheet_name))


   def show_sheet_data(self, sheet_name):
       if not self.selected_file:
           return
      
       try:
           # Read Excel with parse_dates=False to keep original format
           df = pd.read_excel(self.selected_file, sheet_name=sheet_name, header=1)
           self.current_df = df
           self.current_headers = [str(h) for h in df.columns]  # Store headers
          
           # Known date columns (add any other date column names you have)
           date_columns = ['DOB', 'DOJ', 'D.O.B.', 'D.O.J', 'DATE']
          
           # Update table
           rows, cols = df.shape
           self.data_table.setRowCount(rows)
           self.data_table.setColumnCount(cols)
           self.data_table.setHorizontalHeaderLabels([str(h) for h in df.columns])
          
           # Fill data efficiently with proper date formatting
           for i in range(rows):
               for j in range(cols):
                   value = df.iloc[i, j]
                   col_name = str(df.columns[j])


                   # Check if column is a date column or value is datetime
                   is_date = (
                       col_name in date_columns or
                       pd.api.types.is_datetime64_any_dtype(df.iloc[:, j]) or
                       isinstance(value, pd.Timestamp)
                   )


                   if is_date and pd.notnull(value):
                       try:
                           # Convert to datetime if string
                           if isinstance(value, str):
                               value = pd.to_datetime(value)
                           # Only show the date part (remove timestamp)
                           formatted_value = value.strftime('%d/%m/%Y')
                       except Exception:
                           formatted_value = str(value)
                   elif isinstance(value, (int, float)) and pd.notnull(value):
                       # Format numbers to 2 decimal points
                       formatted_value = f"{value:.2f}"
                   else:
                       formatted_value = "" if pd.isna(value) else str(value)


                   self.data_table.setItem(i, j, QTableWidgetItem(formatted_value))
          
           self.data_table.resizeColumnsToContents()
           # Enable row selection
           self.data_table.setSelectionBehavior(QTableWidget.SelectRows)
           self.data_table.setSelectionMode(QTableWidget.SingleSelection)
           # Connect selection signal
           self.data_table.itemSelectionChanged.connect(self.on_row_selection_changed)
          
           # Enable bulk generation and print buttons when sheet is loaded
           self.generate_bulk_button.setEnabled(True)
           self.print_bulk_button.setEnabled(True)
          
           self.status_label.setText(f"Status: Loaded {rows} rows from '{sheet_name}'")
          
       except Exception as e:
           logger.error(f"Error loading sheet: {e}")
           self.status_label.setText(f"Status: Failed to load sheet")
           self.generate_bulk_button.setEnabled(False)
           self.print_bulk_button.setEnabled(False)


   def on_row_selection_changed(self):
       # Enable/disable payslip and print buttons based on row selection
       selected_rows = self.data_table.selectedItems()
       has_selection = len(selected_rows) > 0
       self.generate_payslip_button.setEnabled(has_selection)
       self.print_selected_button.setEnabled(has_selection)
       self.print_selected_button.setEnabled(len(selected_rows) > 0)


   def generate_selected_payslip(self):
       if self.current_df is not None:
           selected_items = self.data_table.selectedItems()
           if selected_items:
               row_index = selected_items[0].row()
               try:
                   # Get the row data as Series
                   row_data = self.current_df.iloc[row_index]
                  
                   # Generate payslip with current sheet name
                   current_sheet = self.sheet_list.currentItem().text()
                   payslip = generate_payslip(row_data, current_sheet)
                  
                   # Display payslip in preview area
                   self.payslip_preview.setText(payslip)
                  
                   # Also generate and save PDF with directory chooser
                   emp_name = str(row_data.get('NAME', f'Employee {row_index}'))
                   if self.printer.generate_single_pdf(emp_name, payslip, ask_directory=True):
                       pdf_dir = self.printer.output_directory
                       self.status_label.setText(f"Status: PDF generated for {emp_name} in {pdf_dir}")
                   else:
                       self.status_label.setText(f"Status: PDF generation cancelled for {emp_name}")
                  
               except Exception as e:
                   error_msg = f"Error generating payslip: {str(e)}"
                   logger.error(error_msg)
                   logger.exception("Detailed error:")
                   QMessageBox.critical(self, "Error", error_msg)
                   self.payslip_preview.clear()
           else:
               QMessageBox.warning(self, "Warning", "Please select a row first")
               self.payslip_preview.clear()
       else:
           QMessageBox.warning(self, "Warning", "Please load a sheet first")
           self.payslip_preview.clear()


   def generate_bulk_payslips(self):
       if self.current_df is None:
           QMessageBox.warning(self, "Warning", "Please load a sheet first")
           return


       try:
           current_sheet = self.sheet_list.currentItem().text()
           self._generate_bulk_pdfs(current_sheet)
       except Exception as e:
           error_msg = f"Error in bulk generation: {str(e)}"
           logger.error(error_msg)
           logger.exception("Detailed error:")
           QMessageBox.critical(self, "Error", error_msg)


   def _generate_bulk_pdfs(self, current_sheet):
       """Generate PDFs for bulk payslips with directory chooser"""
       progress = None
       try:
           # Prepare data for bulk PDF generation
           employees = []
           total_rows = len(self.current_df)
          
           # Show progress dialog
           progress = QMessageBox(self)
           progress.setWindowTitle("Preparing Data")
           progress.setText("Preparing data for PDF generation...")
           progress.setStandardButtons(QMessageBox.NoButton)
           progress.show()
           QApplication.processEvents()
          
           for index, row_data in self.current_df.iterrows():
               try:
                   # Get employee name
                   emp_name = str(row_data.get('NAME', f'Employee {index}'))
                  
                   # Generate payslip content
                   payslip = generate_payslip(row_data, current_sheet)
                  
                   # Show current payslip in UI
                   self.payslip_preview.setText(payslip)
                   QApplication.processEvents()  # Update UI
                  
                   # Add to the list of employees to process
                   employee_data = {'name': emp_name, 'data': row_data, 'sheet': current_sheet}
                   employees.append(employee_data)
                  
               except Exception as e:
                   logger.error(f"Error preparing data for row {index}: {str(e)}")
          
           # Close the preparing dialog
           if progress:
               progress.close()
               progress.deleteLater()
               progress = None
               QApplication.processEvents()
          
           # Define a content generator function
           def content_generator(employee):
               return generate_payslip(employee['data'], employee['sheet'])
          
           # Start PDF generation process with directory chooser
           if employees:
               self.printer.generate_bulk_pdfs(employees, content_generator, ask_directory=True)
               self.status_label.setText("Status: PDF generation started...")
           else:
               QMessageBox.warning(self, "Warning", "No data prepared for PDF generation")
              
       except Exception as e:
           # Ensure dialog is closed even if an error occurs
           if progress:
               progress.close()
               progress.deleteLater()
               progress = None
               QApplication.processEvents()
          
           error_msg = f"Error in PDF generation: {str(e)}"
           logger.error(error_msg)
           QMessageBox.critical(self, "Error", error_msg)


   def print_selected_payslip(self):
       """Prints the currently selected payslip with B4 paper size"""
       if self.current_df is not None:
           selected_items = self.data_table.selectedItems()
           if selected_items:
               row_index = selected_items[0].row()
               try:
                   # Get the row data as Series
                   row_data = self.current_df.iloc[row_index]
                  
                   # Get employee name for the print dialog
                   emp_name = str(row_data.get('NAME', f'Employee {row_index}'))
                  
                   # Generate payslip with current sheet name
                   current_sheet = self.sheet_list.currentItem().text()
                   payslip = generate_payslip(row_data, current_sheet)
                  
                   # Show the payslip in preview
                   self.payslip_preview.setText(payslip)
                  
                   # Ask user for print option (only printing options)
                   choice = QMessageBox.question(
                       self,
                       "Print Options",
                       f"How would you like to print the payslip for {emp_name}?\n\nYes = Print with Preview\nNo = Print Directly",
                       QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                       QMessageBox.Yes
                   )
                  
                   if choice == QMessageBox.Yes:
                       # Print with preview
                       if self.printer.print_with_preview(payslip, emp_name):
                           self.status_label.setText(f"Status: Payslip printed for {emp_name}")
                       else:
                           self.status_label.setText(f"Status: Print cancelled for {emp_name}")
                   elif choice == QMessageBox.No:
                       # Print directly
                       if self.printer.print_single_payslip(payslip, emp_name):
                           self.status_label.setText(f"Status: Payslip printed for {emp_name}")
                       else:
                           self.status_label.setText(f"Status: Print cancelled for {emp_name}")
                   else:
                       # User cancelled
                       self.status_label.setText(f"Status: Print cancelled for {emp_name}")
                  
               except Exception as e:
                   error_msg = f"Error printing payslip: {str(e)}"
                   logger.error(error_msg)
                   logger.exception("Detailed error:")
                   QMessageBox.critical(self, "Error", error_msg)
           else:
               QMessageBox.warning(self, "Warning", "Please select a row first")
       else:
           QMessageBox.warning(self, "Warning", "Please load a sheet first")
  
   def print_bulk_payslips(self):
       """Prints all payslips from the current sheet"""
       if self.current_df is None:
           QMessageBox.warning(self, "Warning", "Please load a sheet first")
           return
          
       try:
           # Ask for confirmation before printing many payslips
           total_rows = len(self.current_df)
           confirm = QMessageBox.question(
               self,
               "Bulk Print Confirmation",
               f"Are you sure you want to print {total_rows} payslips?\n\nThis will send all payslips to your printer one by one.",
               QMessageBox.Yes | QMessageBox.No,
               QMessageBox.No
           )
          
           if confirm == QMessageBox.No:
               return
          
           # Use the bulk printing method from print manager
           current_sheet = self.sheet_list.currentItem().text()
          
           # Prepare employees data
           employees = []
           for index, row_data in self.current_df.iterrows():
               emp_name = str(row_data.get('NAME', f'Employee {index}'))
               employee_data = {'name': emp_name, 'data': row_data, 'sheet': current_sheet}
               employees.append(employee_data)
          
           # Define content generator
           def content_generator(employee):
               return generate_payslip(employee['data'], employee['sheet'])
          
           # Use the print manager's bulk printing method (only printing, no PDF)
           if self.printer.print_bulk_payslips(employees, content_generator):
               self.status_label.setText("Status: Bulk printing completed successfully")
           else:
               self.status_label.setText("Status: Bulk printing was cancelled or failed")
              
       except Exception as e:
           error_msg = f"Error in bulk printing: {str(e)}"
           logger.error(error_msg)
           logger.exception("Detailed error:")
           QMessageBox.critical(self, "Error", error_msg)


   def show_config_dialog(self):
       dialog = ConfigDialog(self, self.current_headers)
       dialog.exec_()


class ConfigDialog(QDialog):
   def __init__(self, parent=None, headers=None):
       super().__init__(parent)
       self.config = PayslipConfig()
       self.headers = headers or []
       self.setup_ui()


   def setup_ui(self):
       self.setWindowTitle("Configure Custom Fields")
       layout = QVBoxLayout()


       # Sheet type selector
       self.sheet_type = QComboBox()
       self.sheet_type.addItems(['FIXED', 'FTC'])
       layout.addWidget(QLabel("Sheet Type:"))
       layout.addWidget(self.sheet_type)


       # Type selector
       self.mapping_type = QComboBox()
       self.mapping_type.addItems(['earnings', 'deductions'])
       layout.addWidget(QLabel("Mapping Type:"))
       layout.addWidget(self.mapping_type)


       # Display name input
       layout.addWidget(QLabel("Display Name:"))
       self.display_name = QLineEdit()
       self.display_name.setPlaceholderText("Display Name")
       layout.addWidget(self.display_name)


       # Excel header dropdown
       layout.addWidget(QLabel("Excel Column:"))
       self.excel_header = QComboBox()
       self.excel_header.addItems(self.headers)
       self.excel_header.setEditable(True)
       layout.addWidget(self.excel_header)


       # Add mapping button
       add_button = QPushButton("Add Mapping")
       add_button.clicked.connect(self.add_mapping)
       layout.addWidget(add_button)


       # Current mappings list
       layout.addWidget(QLabel("Current Mappings:"))
       self.mappings_list = QListWidget()
       self.update_mappings_list()
       layout.addWidget(self.mappings_list)


       # Remove mapping button
       remove_button = QPushButton("Remove Selected")
       remove_button.clicked.connect(self.remove_mapping)
       layout.addWidget(remove_button)


       self.setLayout(layout)


   def update_mappings_list(self):
       self.mappings_list.clear()
       sheet_type = self.sheet_type.currentText()
       mapping_type = self.mapping_type.currentText()
       mappings = self.config.get_mappings(sheet_type)[mapping_type]
      
       for display_name, excel_header in mappings.items():
           self.mappings_list.addItem(f"{display_name} → {excel_header}")


   def add_mapping(self):
       display_name = self.display_name.text().strip()
       excel_header = self.excel_header.currentText().strip()
      
       if display_name and excel_header:
           sheet_type = self.sheet_type.currentText()
           mapping_type = self.mapping_type.currentText()
          
           self.config.add_mapping(sheet_type, mapping_type, display_name, excel_header)
           self.update_mappings_list()
           self.display_name.clear()
           self.excel_header.clear()


   def remove_mapping(self):
       current_item = self.mappings_list.currentItem()
       if current_item:
           display_name = current_item.text().split(' → ')[0]
           sheet_type = self.sheet_type.currentText()
           mapping_type = self.mapping_type.currentText()
          
           self.config.remove_mapping(sheet_type, mapping_type, display_name)
           self.update_mappings_list()


def main():
   app = QApplication(sys.argv)
   window = ExcelSheetViewer()
   window.show()
   sys.exit(app.exec_())


if __name__ == "__main__":
   main()
