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
   QMessageBox, QTextEdit, QDialog, QComboBox, QLineEdit, QListWidgetItem
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette
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


       # Month selector (add after file selection buttons)
       month_container = QWidget()
       month_layout = QHBoxLayout()
       month_container.setLayout(month_layout)
        
       month_layout.addWidget(QLabel("Payslip Month:"))
       self.month_selector = QComboBox()
       current_year = datetime.now().year
       months = [f"{m.upper()} {y}" for y in [current_year-1, current_year, current_year+1] 
                for m in ["January", "February", "March", "April", "May", "June", 
                         "July", "August", "September", "October", "November", "December"]]
       self.month_selector.addItems(months)
        
       # Set current month
       current_month = datetime.now().strftime('%B %Y').upper()
       current_idx = self.month_selector.findText(current_month)
       if current_idx >= 0:
           self.month_selector.setCurrentIndex(current_idx)
            
       # Load saved month if exists
       config = PayslipConfig()
       saved_month = config.get_payslip_month()
       if saved_month:
           saved_idx = self.month_selector.findText(saved_month)
           if saved_idx >= 0:
               self.month_selector.setCurrentIndex(saved_idx)
        
       self.month_selector.currentTextChanged.connect(self.on_month_changed)
       month_layout.addWidget(self.month_selector)
        
       self.layout.addWidget(month_container)


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
       """Handle row selection and generate payslip preview"""
       selected_items = self.data_table.selectedItems()
       has_selection = len(selected_items) > 0
       self.generate_payslip_button.setEnabled(has_selection)
       self.print_selected_button.setEnabled(has_selection)

       # Generate preview when row is selected
       if has_selection and self.current_df is not None:
           try:
               # Get selected row index
               row_index = selected_items[0].row()
               # Get row data as Series
               row_data = self.current_df.iloc[row_index]
               # Get current sheet name
               current_sheet = self.sheet_list.currentItem().text()
               # Generate payslip
               payslip = generate_payslip(row_data, current_sheet)
               # Update preview
               self.payslip_preview.setText(payslip)
           except Exception as e:
               logger.error(f"Error generating preview: {str(e)}")
               self.payslip_preview.setText("Error generating preview")
       else:
           self.payslip_preview.clear()


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


   def on_month_changed(self, month):
        """Save selected month to config"""
        config = PayslipConfig()
        config.set_payslip_month(month)
        # Update any visible payslips
        self.on_row_selection_changed()


class ConfigDialog(QDialog):
    def __init__(self, parent=None, headers=None):
        super().__init__(parent)
        self.config = PayslipConfig()
        self.headers = headers or []
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle("Configure Custom Fields")
        self.setMinimumSize(500, 600)  # Set minimum size
        self.resize(600, 700)  # Set initial size
        layout = QVBoxLayout()


        # Sheet type selector
        self.sheet_type = QComboBox()
        self.sheet_type.addItems(['FIXED', 'FTC'])
        self.sheet_type.currentTextChanged.connect(self.update_mappings_list)
        layout.addWidget(QLabel("Sheet Type:"))
        layout.addWidget(self.sheet_type)


        # Type selector
        self.mapping_type = QComboBox()
        self.mapping_type.addItems(['earnings', 'deductions'])
        self.mapping_type.currentTextChanged.connect(self.update_mappings_list)
        layout.addWidget(QLabel("Mapping Type:"))
        layout.addWidget(self.mapping_type)


        # Mapping format selector
        self.mapping_format = QComboBox()
        self.mapping_format.addItems(['Single Column', 'Double Column'])
        layout.addWidget(QLabel("Mapping Format:"))
        layout.addWidget(self.mapping_format)
        self.mapping_format.currentIndexChanged.connect(self.on_mapping_format_changed)


        # Display name input (only one)
        self.display_name1_label = QLabel("Display Name:")
        layout.addWidget(self.display_name1_label)
        self.display_name1 = QLineEdit()
        self.display_name1.setPlaceholderText("Display Name")
        layout.addWidget(self.display_name1)

        # Remove Display Name 2
        # self.display_name2_label = QLabel("Display Name 2:")
        # layout.addWidget(self.display_name2_label)
        # self.display_name2 = QLineEdit()
        # self.display_name2.setPlaceholderText("Display Name 2")
        # layout.addWidget(self.display_name2)
        # self.display_name2_label.hide()
        # self.display_name2.hide()


        # Excel header dropdown(s)
        self.excel_header1_label = QLabel("Excel Column 1:")
        layout.addWidget(self.excel_header1_label)
        self.excel_header1 = QComboBox()
        self.excel_header1.addItems(self.headers)
        self.excel_header1.setEditable(True)
        layout.addWidget(self.excel_header1)

        self.excel_header2_label = QLabel("Excel Column 2:")
        layout.addWidget(self.excel_header2_label)
        self.excel_header2 = QComboBox()
        self.excel_header2.addItems(self.headers)
        self.excel_header2.setEditable(True)
        layout.addWidget(self.excel_header2)
        self.excel_header2_label.hide()
        self.excel_header2.hide()

        # Track editing state (initialize before update_mappings_list)
        self.editing_mapping = None  # Will store the original mapping being edited

        # Current mappings list
        self.mappings_status_label = QLabel()
        layout.addWidget(self.mappings_status_label)
        
        self.mappings_list = QListWidget()
        self.mappings_list.itemDoubleClicked.connect(self.edit_mapping)
        self.update_mappings_list()
        layout.addWidget(self.mappings_list)

        # Instruction label for editing
        edit_instruction = QLabel("💡 Tip: Double-click any mapping to edit it | Click and drag to reorder")
        edit_instruction.setStyleSheet("color: #666; font-style: italic; font-size: 11px;")
        layout.addWidget(edit_instruction)

        # Button container for add/update and remove
        button_container = QWidget()
        button_layout = QHBoxLayout()
        button_container.setLayout(button_layout)

        # Add/Update mapping button (dynamic text)
        self.add_update_button = QPushButton("Add Mapping")
        self.add_update_button.clicked.connect(self.add_or_update_mapping)
        button_layout.addWidget(self.add_update_button)

        # Cancel edit button (hidden by default)
        self.cancel_edit_button = QPushButton("Cancel Edit")
        self.cancel_edit_button.clicked.connect(self.cancel_edit)
        self.cancel_edit_button.hide()
        button_layout.addWidget(self.cancel_edit_button)

        # Remove mapping button
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self.remove_mapping)
        button_layout.addWidget(remove_button)

        layout.addWidget(button_container)

        self.setLayout(layout)


    def on_mapping_format_changed(self, idx):
        if self.mapping_format.currentText() == "Double Column":
            # Only show the second column input, not display name
            self.excel_header2_label.show()
            self.excel_header2.show()
        else:
            self.excel_header2_label.hide()
            self.excel_header2.hide()


    def update_mappings_list(self):
        """Update the mappings list based on current sheet type and mapping type selection"""
        # Cancel edit mode if changing sheet type or mapping type
        if self.editing_mapping:
            current_sheet = self.sheet_type.currentText()
            current_mapping = self.mapping_type.currentText()
            if (self.editing_mapping['sheet_type'] != current_sheet or 
                self.editing_mapping['mapping_type'] != current_mapping):
                self.cancel_edit()
        
        self.mappings_list.clear()
        sheet_type = self.sheet_type.currentText()
        mapping_type = self.mapping_type.currentText()
        
        # Update status label
        self.mappings_status_label.setText(f"Current Mappings ({sheet_type} - {mapping_type.title()}):")
        
        # Get mappings for the selected sheet type and mapping type
        mappings = self.config.get_mappings(sheet_type)[mapping_type]
        
        if not mappings:
            item_text = "No mappings configured for this combination"
            item = QListWidgetItem(item_text)
            item.setForeground(QApplication.palette().color(QPalette.Text))
            self.mappings_list.addItem(item)
            return
        
        for display_name, excel_header in mappings.items():
            # Format the display text based on whether it's single or double column
            if isinstance(excel_header, (list, tuple)) and len(excel_header) == 2:
                col1, col2 = excel_header
                item_text = f"{display_name} → [{col1}, {col2}]"
            else:
                item_text = f"{display_name} → {excel_header}"
            
            item = QListWidgetItem(item_text)
            
            # Highlight the item being edited
            if (self.editing_mapping and 
                self.editing_mapping['display_name'] == display_name and
                self.editing_mapping['sheet_type'] == sheet_type and
                self.editing_mapping['mapping_type'] == mapping_type):
                item.setBackground(QApplication.palette().color(QPalette.Highlight))
                item.setForeground(QApplication.palette().color(QPalette.HighlightedText))
            
            self.mappings_list.addItem(item)


    def remove_mapping(self):
        """Remove the selected mapping from the configuration"""
        current_item = self.mappings_list.currentItem()
        if current_item:
            text = current_item.text()
            
            # Check if this is the "No mappings" message
            if text.startswith("No mappings configured"):
                QMessageBox.information(self, "No Mappings", "There are no mappings to remove for this combination.")
                return
            
            # Parse the display text to extract the display name
            if ' → ' in text:
                display_name = text.split(' → ')[0].strip()
                
                sheet_type = self.sheet_type.currentText()
                mapping_type = self.mapping_type.currentText()
                
                # Check if we're trying to remove the mapping being edited
                if (self.editing_mapping and 
                    self.editing_mapping['display_name'] == display_name and
                    self.editing_mapping['sheet_type'] == sheet_type and
                    self.editing_mapping['mapping_type'] == mapping_type):
                    
                    reply = QMessageBox.question(self, "Remove Edited Mapping", 
                                               f"You are currently editing '{display_name}'.\n"
                                               f"Do you want to remove it instead?",
                                               QMessageBox.Yes | QMessageBox.No, 
                                               QMessageBox.No)
                    if reply == QMessageBox.No:
                        return
                    
                    # Cancel edit mode first
                    self.cancel_edit()
                else:
                    # Normal confirmation for non-edited mappings
                    reply = QMessageBox.question(self, "Confirm Removal", 
                                               f"Are you sure you want to remove the mapping '{display_name}'?",
                                               QMessageBox.Yes | QMessageBox.No, 
                                               QMessageBox.No)
                    
                    if reply == QMessageBox.No:
                        return
                
                # Remove the mapping using the display name as key
                self.config.remove_mapping(sheet_type, mapping_type, display_name)
                
                # Update the list to reflect the change
                self.update_mappings_list()
                
                # Show success message
                QMessageBox.information(self, "Success", f"Mapping '{display_name}' has been removed.")
            else:
                QMessageBox.warning(self, "Error", "Invalid mapping format selected.")
        else:
            QMessageBox.information(self, "No Selection", "Please select a mapping to remove.")


    def edit_mapping(self, item):
        """Handle double-click on a mapping item to edit it"""
        text = item.text()
        
        # Check if this is the "No mappings" message
        if text.startswith("No mappings configured"):
            return
        
        # Parse the display text to extract mapping information
        if ' → ' not in text:
            return
            
        display_name = text.split(' → ')[0].strip()
        excel_header_part = text.split(' → ')[1].strip()
        
        # Get the current sheet type and mapping type
        sheet_type = self.sheet_type.currentText()
        mapping_type = self.mapping_type.currentText()
        
        # Get the actual mapping from config
        mappings = self.config.get_mappings(sheet_type)[mapping_type]
        if display_name not in mappings:
            QMessageBox.warning(self, "Error", "Could not find mapping in configuration.")
            return
            
        excel_header = mappings[display_name]
        
        # Enter edit mode
        self.editing_mapping = {
            'display_name': display_name,
            'excel_header': excel_header,
            'sheet_type': sheet_type,
            'mapping_type': mapping_type
        }
        
        # Populate the input fields
        self.display_name1.setText(display_name)
        
        # Handle single vs double column
        if isinstance(excel_header, (list, tuple)) and len(excel_header) == 2:
            # Double column mapping
            self.mapping_format.setCurrentText("Double Column")
            self.excel_header1.setCurrentText(excel_header[0])
            self.excel_header2.setCurrentText(excel_header[1])
        else:
            # Single column mapping
            self.mapping_format.setCurrentText("Single Column")
            self.excel_header1.setCurrentText(str(excel_header))
            self.excel_header2.setCurrentText("")
        
        # Update UI for edit mode
        self.add_update_button.setText("Update Mapping")
        self.cancel_edit_button.show()
        
        # Highlight the item being edited
        item.setSelected(True)
        
        # Show info message
        QMessageBox.information(self, "Edit Mode", 
                              f"Now editing: {display_name}\n"
                              f"Make your changes and click 'Update Mapping' to save.")

    def cancel_edit(self):
        """Cancel the current edit operation"""
        self.editing_mapping = None
        self.add_update_button.setText("Add Mapping")
        self.cancel_edit_button.hide()
        
        # Clear input fields
        self.display_name1.clear()
        self.excel_header1.setCurrentIndex(0)
        self.excel_header2.setCurrentIndex(0)
        self.mapping_format.setCurrentText("Single Column")
        
        # Clear selection
        self.mappings_list.clearSelection()

    def add_or_update_mapping(self):
        """Add new mapping or update existing one based on current mode"""
        mapping_format = self.mapping_format.currentText()
        sheet_type = self.sheet_type.currentText()
        mapping_type = self.mapping_type.currentText()
        dn1 = self.display_name1.text().strip()
        col1 = self.excel_header1.currentText().strip()
        col2 = self.excel_header2.currentText().strip()
        
        # Validate input
        if mapping_format == "Double Column":
            if not (dn1 and col1 and col2):
                QMessageBox.warning(self, "Incomplete Data", 
                                  "Please fill in display name and both column names for double column mapping.")
                return
            excel_header = [col1, col2]
        else:
            if not (dn1 and col1):
                QMessageBox.warning(self, "Incomplete Data", 
                                  "Please fill in display name and column name for single column mapping.")
                return
            excel_header = col1
        
        if self.editing_mapping:
            # Update existing mapping
            old_display_name = self.editing_mapping['display_name']
            old_sheet_type = self.editing_mapping['sheet_type']
            old_mapping_type = self.editing_mapping['mapping_type']
            
            # If the key (display_name, sheet_type, or mapping_type) changed, remove old mapping
            if (old_display_name != dn1 or 
                old_sheet_type != sheet_type or 
                old_mapping_type != mapping_type):
                self.config.remove_mapping(old_sheet_type, old_mapping_type, old_display_name)
            
            # Add the updated mapping
            self.config.add_mapping(sheet_type, mapping_type, dn1, excel_header)
            
            # Exit edit mode
            self.cancel_edit()
            
            # Show success message
            QMessageBox.information(self, "Success", 
                                  f"Mapping '{dn1}' has been updated successfully!")
        else:
            # Add new mapping
            self.config.add_mapping(sheet_type, mapping_type, dn1, excel_header)
            
            # Clear input fields
            self.display_name1.clear()
            self.excel_header1.setCurrentIndex(0)
            self.excel_header2.setCurrentIndex(0)
            
            # Show success message
            QMessageBox.information(self, "Success", 
                                  f"Mapping '{dn1}' has been added to {sheet_type} {mapping_type}.")
        
        # Update the list to show changes
        self.update_mappings_list()

def main():
   app = QApplication(sys.argv)
   window = ExcelSheetViewer()
   window.show()
   sys.exit(app.exec_())


if __name__ == "__main__":
   main()
   main()
