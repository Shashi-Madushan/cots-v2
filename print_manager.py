import os
import platform
import tempfile
import logging
import subprocess
from datetime import datetime
from PyQt5.QtCore import QObject, pyqtSignal, QThread, QSizeF, Qt
from PyQt5.QtWidgets import QMessageBox, QDialog, QProgressBar, QLabel, QVBoxLayout, QPushButton, QHBoxLayout, QFileDialog, QApplication, QProgressDialog
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PyQt5.QtGui import QTextDocument, QFont
import uuid
import math

# Set up logging
logger = logging.getLogger('PrintManager')

def is_printer_available():
    """Check if any printers are available on the system"""
    # For PDF generation, we don't need an actual printer
    return True
    # Original printer detection code commented out
    '''
    try:
        printer = QPrinter()
        # If default printer name is empty, there may be no printers
        if not printer.printerName():
            if platform.system() == 'Windows':
                # On Windows, check for printers using wmic (requires admin)
                import subprocess
                result = subprocess.run(['wmic', 'printer', 'get', 'name'], 
                                      capture_output=True, text=True, check=False)
                # If any printers are listed after header line, return True
                lines = [line.strip() for line in result.stdout.split('\n') if line.strip()]
                return len(lines) > 1  # More than just the header
            elif platform.system() == 'Darwin':  # macOS
                result = subprocess.run(['lpstat', '-p'], 
                                      capture_output=True, text=True, check=False)
                return 'printer' in result.stdout.lower()
            else:  # Linux and other UNIX
                result = subprocess.run(['lpstat', '-p'], 
                                      capture_output=True, text=True, check=False)
                return result.returncode == 0 and result.stdout.strip() != ''
        return True
    except Exception as e:
        logger.error(f"Error checking for printers: {str(e)}")
        return False
    '''

class PDFGeneratorWorker(QThread):
    """Worker thread for generating payslip PDFs in the background"""
    progress_updated = pyqtSignal(int)
    process_finished = pyqtSignal(int, int)  # success_count, error_count
    single_pdf_complete = pyqtSignal(bool, str)  # success, message

    def __init__(self, employees, content_generator, output_directory=None):
        """
        Initialize PDF generator worker
        
        Args:
            employees: List of employee information (id, name, etc.)
            content_generator: Function that takes employee info and returns payslip content
            output_directory: Custom directory to save PDFs (optional)
        """
        super().__init__()
        self.employees = employees
        self.content_generator = content_generator
        self.cancelled = False
        self.output_directory = output_directory or os.path.join(os.path.expanduser("~"), "Payslips")
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory)

    def run(self):
        """Process all payslips for PDF generation, creating each PDF on demand"""
        success_count = 0
        error_count = 0
        total = len(self.employees)
        
        for i, employee in enumerate(self.employees):
            if self.cancelled:
                break
                
            try:
                # Generate payslip content on demand
                emp_name = employee.get('name', f"Employee {i+1}")
                
                # Get the payslip content from the generator function
                try:
                    payslip_content = self.content_generator(employee)
                except Exception as e:
                    logger.error(f"Error generating payslip for {emp_name}: {str(e)}")
                    error_count += 1
                    self.single_pdf_complete.emit(False, f"Failed to generate payslip for {emp_name}")
                    continue
                
                # Generate PDF using the common generator function
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # Include milliseconds
                pdf_path = generate_pdf(payslip_content, emp_name, self.output_directory, 
                                      timestamp)
                if pdf_path:
                    success_count += 1
                    self.single_pdf_complete.emit(True, f"Generated PDF for {emp_name} at {pdf_path}")
                else:
                    error_count += 1
                    self.single_pdf_complete.emit(False, f"Failed to generate PDF for {emp_name}")
                
            except Exception as e:
                logger.error(f"Error generating PDF: {str(e)}")
                error_count += 1
                self.single_pdf_complete.emit(False, f"Error: {str(e)}")
            
            # Update progress
            progress = int((i + 1) / total * 100)
            self.progress_updated.emit(progress)
        
        self.process_finished.emit(success_count, error_count)

    def cancel(self):
        """Cancel the PDF generation process"""
        self.cancelled = True


class PDFProgressDialog(QDialog):
    """Dialog displaying PDF generation progress"""
    def __init__(self, parent=None, total_count=0):
        super().__init__(parent)
        self.setWindowTitle("Generating PDF Payslips")
        self.setMinimumWidth(400)
        
        # Create layout
        layout = QVBoxLayout()
        
        # Status label
        self.status_label = QLabel(f"Generating 0/{total_count} PDF payslips...")
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Current job label
        self.current_job = QLabel("Starting...")
        layout.addWidget(self.current_job)
        
        # Cancel button
        button_layout = QHBoxLayout()
        self.cancel_button = QPushButton("Cancel")
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        self.total_count = total_count

    def update_progress(self, progress):
        """Update progress bar and status"""
        self.progress_bar.setValue(progress)
        completed = int(progress * self.total_count / 100)
        self.status_label.setText(f"Generating {completed}/{self.total_count} PDF payslips...")
    
    def update_current_job(self, success, message):
        """Update current job status"""
        self.current_job.setText(message)


class PageSettingsManager:
    """Centralized manager for page settings configuration using Singleton pattern"""
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(PageSettingsManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if not self._initialized:
            self._settings = {
                "paper_size": QPrinter.B5,
                "custom_size": QSizeF(220, 200),
                "margins": {"top": 0, "bottom": 0, "left": 0, "right": 0},
                "padding": 0,
                "orientation": QPrinter.Portrait,
                "font_family": "Courier New",
                "font_size": 10
            }
            self._initialized = True

    def configure_printer(self, printer):
        """Configure all printer settings in one place"""
        if self._settings.get("custom_size"):
            printer.setPaperSize(self._settings["custom_size"], QPrinter.Millimeter)
        else:
            printer.setPaperSize(self._settings.get("paper_size", QPrinter.B4))
        
        margins = self._settings.get("margins", {"top": 0, "bottom": 0, "left": 0, "right": 0})
        printer.setPageMargins(
            float(margins["left"]), 
            float(margins["top"]), 
            float(margins["right"]), 
            float(margins["bottom"]), 
            QPrinter.Millimeter
        )
        printer.setOrientation(self._settings.get("orientation", QPrinter.Portrait))
        printer.setFullPage(True)  # Enable full page printing

    def configure_document(self, doc):
        """Configure document settings"""
        font = QFont(
            self._settings.get("font_family", "Courier New"),
            self._settings.get("font_size", 10)
        )
        doc.setDefaultFont(font)

    def update(self, paper_size=None, custom_size=None, margins=None, padding=None, 
              orientation=None, font_family=None, font_size=None):
        """Update settings"""
        if paper_size:
            self._settings["paper_size"] = paper_size
            self._settings["custom_size"] = None
        if custom_size:
            self._settings["custom_size"] = QSizeF(*custom_size)
            self._settings["paper_size"] = None
        if margins:
            self._settings["margins"] = margins
        if padding is not None:
            self._settings["padding"] = padding
        if orientation is not None:
            self._settings["orientation"] = orientation
        if font_family:
            self._settings["font_family"] = font_family
        if font_size:
            self._settings["font_size"] = font_size

    @property
    def settings(self):
        """Get current settings"""
        return self._settings.copy()


class PayslipPDFGenerator:
    """Main class for handling payslip PDF generation"""
    def __init__(self, parent_widget):
        self.parent = parent_widget
        self.pdf_worker = None
        self.progress_dialog = None
        self.cancelled = False  # Add cancelled flag
        
        # For PDF generation, we don't need an actual printer
        self.printers_available = True
        
        # Create default output directory
        self.output_directory = os.path.join(os.path.expanduser("~"), "Payslips")
        if not os.path.exists(self.output_directory):
            try:
                os.makedirs(self.output_directory)
            except Exception as e:
                logger.error(f"Error creating output directory: {str(e)}")
                self.output_directory = os.path.expanduser("~")
        
        # Use singleton page settings manager
        self.page_settings_manager = PageSettingsManager()

    def set_page_settings(self, paper_size=None, custom_size=None, margins=None, padding=None):
        """Set custom page settings"""
        self.page_settings_manager.update(paper_size, custom_size, margins, padding)

    def get_output_directory(self, title="Select Output Folder"):
        """Get output directory from user via file dialog"""
        directory = QFileDialog.getExistingDirectory(
            self.parent,
            title,
            self.output_directory,
            QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
        )
        
        # If user canceled, return None
        if not directory:
            return None
            
        # Create directory if it doesn't exist
        if not os.path.exists(directory):
            try:
                os.makedirs(directory)
            except Exception as e:
                logger.error(f"Error creating selected directory: {str(e)}")
                QMessageBox.critical(self.parent, "Directory Error", 
                                    f"Error creating directory: {str(e)}")
                return None
                
        return directory

    def generate_single_pdf(self, employee_name, payslip_content, ask_directory=True):
        """
        Generate a single payslip PDF with optional directory selection
        
        Args:
            employee_name: Name of the employee for the payslip
            payslip_content: Content of the payslip
            ask_directory: Whether to ask user for directory selection
        """
        try:
            # Get output directory
            if ask_directory:
                output_dir = self.get_output_directory("Select Directory to Save PDF")
                if not output_dir:  # User cancelled
                    return False
            else:
                output_dir = self.output_directory
            
            # Ensure output directory exists
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Use timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Generate PDF using the common generator function
            pdf_path = generate_pdf(payslip_content, employee_name, output_dir, 
                                  timestamp)
            
            if pdf_path:
                # Update the output directory for status message
                self.output_directory = output_dir
                logger.info(f"PDF saved to: {pdf_path}")
                return True
            else:
                return False
            
        except Exception as e:
            logger.error(f"Error in PDF generation: {str(e)}")
            QMessageBox.critical(self.parent, "PDF Error", f"Error generating PDF: {str(e)}")
            return False

    def _cancel_generation(self):
        """Cancel the ongoing PDF generation process"""
        self.cancelled = True
        if self.pdf_worker:
            self.pdf_worker.cancel()
        if self.progress_dialog:
            self.progress_dialog.current_job.setText("Cancelling...")

    def generate_bulk_pdfs(self, employees, content_generator, ask_directory=True):
        """Generate multiple payslip PDFs with batch processing"""
        if not employees:
            QMessageBox.warning(self.parent, "No Data", "No employees selected for PDF generation.")
            return

        try:
            # Data validation - filter out invalid entries
            valid_employees = []
            for emp in employees:
                try:
                    # Validate required fields
                    if not emp or not isinstance(emp, dict):
                        continue
                    # Convert potential NaN values to None or default values
                    cleaned_emp = {}
                    for key, value in emp.items():
                        if isinstance(value, float) and math.isnan(value):
                            cleaned_emp[key] = None
                        else:
                            cleaned_emp[key] = value
                    valid_employees.append(cleaned_emp)
                except Exception as e:
                    logger.warning(f"Skipping invalid employee data: {str(e)}")
                    continue

            if not valid_employees:
                QMessageBox.warning(self.parent, "Invalid Data", 
                                  "No valid employee data found for PDF generation.")
                return

            # Get output directory
            output_dir = self.get_output_directory("Select Directory to Save PDF Payslips") if ask_directory else self.output_directory
            if not output_dir:
                return
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Create progress dialog
            self.progress_dialog = PDFProgressDialog(self.parent, len(valid_employees))
            self.progress_dialog.cancel_button.clicked.connect(self._cancel_generation)
            self.progress_dialog.show()
            
            # Initialize batch processor
            batch_processor = BatchProcessor(batch_size=10)
            
            def process_batch(batch):
                batch_success = 0
                batch_errors = 0
                
                for employee in batch:
                    if self.cancelled:
                        break
                        
                    try:
                        emp_name = employee.get('name', 'Employee')
                        payslip_content = content_generator(employee)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                        
                        pdf_path = generate_pdf(payslip_content, emp_name, output_dir, timestamp)
                        
                        if pdf_path:
                            batch_success += 1
                            self.progress_dialog.update_current_job(True, f"Generated PDF for {emp_name}")
                        else:
                            batch_errors += 1
                            self.progress_dialog.update_current_job(False, f"Failed to generate PDF for {emp_name}")
                            
                    except Exception as e:
                        batch_errors += 1
                        logger.error(f"Error processing {emp_name}: {str(e)}")
                        self.progress_dialog.update_current_job(False, f"Error: {str(e)}")
                    
                return batch_success, batch_errors
            
            # Process batches
            success_count, error_count = batch_processor.process_batches(
                valid_employees,
                process_batch,
                lambda progress: self.progress_dialog.update_progress(progress)
            )
            
            self._on_generation_finished(success_count, error_count, output_dir)
            
        except Exception as e:
            logger.error(f"Error in batch PDF generation: {str(e)}")
            if self.progress_dialog:
                self.progress_dialog.close()
            QMessageBox.critical(self.parent, "PDF Error", f"Error in batch processing: {str(e)}")

    def _on_generation_finished(self, success_count, error_count, output_directory):
        """Handle completion of PDF generation"""
        try:
            if self.progress_dialog:
                self.progress_dialog.close()
            
            message = (
                f"PDF generation completed.\n"
                f"Successfully generated: {success_count}\n"
                f"Failed: {error_count}\n\n"
                f"PDFs saved to: {output_directory}"
            )
            
            QMessageBox.information(self.parent, "PDF Generation Complete", message)
        except Exception as e:
            logger.error(f"Error showing completion message: {str(e)}")

    def generate_single_pdf_legacy(self, employee_name, payslip_content, ask_directory=True):
        """
        Generate a single payslip PDF (legacy method) - DEPRECATED
        
        Args:
            employee_name: Name of the employee for the payslip
            payslip_content: Content of the payslip
            ask_directory: Whether to ask user for directory selection
        """
        try:
            # Get output directory
            if ask_directory:
                output_dir = self.get_output_directory("Select Directory to Save PDF")
                if not output_dir:  # User cancelled
                    return False
            else:
                output_dir = self.output_directory
            
            # Ensure output directory exists
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Use timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Generate PDF using the common generator function
            pdf_path = generate_pdf(payslip_content, employee_name, output_dir, 
                                  timestamp)
            
            if pdf_path:
                # Update the output directory for status message
                self.output_directory = output_dir
                logger.info(f"PDF saved to: {pdf_path}")
                return True
            else:
                return False
            
        except Exception as e:
            logger.error(f"Error in PDF generation (legacy): {str(e)}")
            QMessageBox.critical(self.parent, "PDF Error", f"Error generating PDF (legacy): {str(e)}")
            return False

    def generate_bulk_pdfs_legacy(self, employees, content_generator, ask_directory=True):
        """Generate multiple payslip PDFs (legacy method) - DEPRECATED"""
        if not employees:
            QMessageBox.warning(self.parent, "No Data", "No employees selected for PDF generation.")
            return
        
        try:
            # Get output directory
            output_dir = self.get_output_directory("Select Directory to Save PDF Payslips") if ask_directory else self.output_directory
            if not output_dir:
                return
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Create progress dialog
            self.progress_dialog = PDFProgressDialog(self.parent, len(employees))
            self.progress_dialog.cancel_button.clicked.connect(self._cancel_generation)
            self.progress_dialog.show()
            
            # Process each employee sequentially (legacy method)
            for i, employee in enumerate(employees):
                if self.cancelled:
                    break
                
                try:
                    emp_name = employee.get('name', f"Employee {i+1}")
                    payslip_content = content_generator(employee)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                    
                    pdf_path = generate_pdf(payslip_content, emp_name, output_dir, timestamp)
                    
                    if pdf_path:
                        self.progress_dialog.update_current_job(True, f"Generated PDF for {emp_name} at {pdf_path}")
                    else:
                        self.progress_dialog.update_current_job(False, f"Failed to generate PDF for {emp_name}")
                
                except Exception as e:
                    logger.error(f"Error generating PDF for {emp_name}: {str(e)}")
                    self.progress_dialog.update_current_job(False, f"Error: {str(e)}")
                
                # Update progress
                progress = int((i + 1) / len(employees) * 100)
                self.progress_dialog.update_progress(progress)
            
            self.process_finished.emit(len(employees), 0)
            
        except Exception as e:
            logger.error(f"Error in bulk PDF generation (legacy): {str(e)}")
            if self.progress_dialog:
                self.progress_dialog.close()
            QMessageBox.critical(self.parent, "PDF Error", f"Error in bulk processing (legacy): {str(e)}")

    def cancel(self):
        """Cancel the PDF generation process"""
        self.cancelled = True


class BatchProcessor:
    """Handles batch processing operations with memory management"""
    
    def __init__(self, batch_size=10):
        self.batch_size = batch_size
    
    def process_batches(self, items, processor_func, progress_callback=None):
        """
        Process items in batches
        
        Args:
            items: List of items to process
            processor_func: Function to process each batch
            progress_callback: Optional callback for progress updates
        """
        total_items = len(items)
        processed = 0
        success_count = 0
        error_count = 0
        
        for i in range(0, total_items, self.batch_size):
            batch = items[i:i + self.batch_size]
            batch_success, batch_errors = processor_func(batch)
            
            success_count += batch_success
            error_count += batch_errors
            processed += len(batch)
            
            if progress_callback:
                progress = int((processed / total_items) * 100)
                progress_callback(progress)
                
        return success_count, error_count


class PayslipPrintManager:
    """Enhanced print manager that handles both printing and PDF generation for B4 paper size"""
    
    def __init__(self, parent_widget):
        self.parent = parent_widget
        self.pdf_generator = PayslipPDFGenerator(parent_widget)
    
    @property
    def output_directory(self):
        """Get the output directory from the PDF generator"""
        return self.pdf_generator.output_directory
        
    @output_directory.setter 
    def output_directory(self, value):
        """Set the output directory on the PDF generator"""
        self.pdf_generator.output_directory = value

    def print_single_payslip(self, payslip_content, employee_name="Employee"):
        """Print a single payslip directly to printer with appropriate page settings"""
        return self._print_payslips([{"content": payslip_content, "name": employee_name}], 
                                   title="Print Payslip")
    
    def print_with_preview(self, payslip_content, employee_name="Employee"):
        """Show print preview dialog before printing with appropriate page settings"""
        try:
            # Create a printer
            printer = QPrinter()
            
            # Apply the same page settings used for PDF generation
            self.pdf_generator.page_settings_manager.configure_printer(printer)
            printer.setOrientation(QPrinter.Portrait)
            
            # Create print preview dialog
            preview_dialog = QPrintPreviewDialog(printer, self.parent)
            preview_dialog.setWindowTitle(f"Print Preview - {employee_name}")
            
            # Connect preview signal
            def print_preview(printer):
                doc = QTextDocument()
                doc.setPlainText(payslip_content)
                
                # Configure document settings through page settings manager
                self.pdf_generator.page_settings_manager.configure_document(doc)
                doc.print_(printer)
            
            preview_dialog.paintRequested.connect(print_preview)
            
            # Show preview dialog
            if preview_dialog.exec_() == QPrintPreviewDialog.Accepted:
                logger.info(f"Payslip printed with preview for {employee_name}")
                return True
            else:
                logger.info(f"Print preview cancelled for {employee_name}")
                return False
                
        except Exception as e:
            logger.error(f"Error in print preview for {employee_name}: {str(e)}")
            QMessageBox.critical(self.parent, "Print Preview Error", f"Error in print preview: {str(e)}")
            return False
    
    def generate_pdf(self, payslip_content, employee_name="Employee", ask_directory=True):
        """Generate PDF using the existing PDF generator"""
        return self.pdf_generator.generate_single_pdf(employee_name, payslip_content, ask_directory)
    
    def generate_single_pdf(self, employee_name, payslip_content, ask_directory=True):
        """Generate a single PDF - delegate to PDF generator"""
        return self.pdf_generator.generate_single_pdf(employee_name, payslip_content, ask_directory)
    
    def generate_bulk_pdfs(self, employees, content_generator, ask_directory=True):
        """Generate bulk PDFs - delegate to PDF generator"""
        return self.pdf_generator.generate_bulk_pdfs(employees, content_generator, ask_directory)
    
    def print_bulk_payslips(self, employees, content_generator):
        """Print multiple payslips with improved batch processing"""
        if not employees:
            QMessageBox.warning(self.parent, "No Data", "No employees selected for printing.")
            return False

        try:
            printer = QPrinter()
            printer.setFullPage(True)
            self.pdf_generator.page_settings_manager.configure_printer(printer)
            printer.setOrientation(QPrinter.Portrait)

            print_dialog = QPrintDialog(printer, self.parent)
            print_dialog.setWindowTitle("Print Payslips")
            
            if print_dialog.exec_() != QPrintDialog.Accepted:
                logger.info("Print cancelled by user")
                return False

            progress = QProgressDialog("Printing payslips...", "Cancel", 0, len(employees), self.parent)
            progress.setWindowModality(Qt.WindowModal)
            
            batch_processor = BatchProcessor(batch_size=10)
            
            def process_print_batch(batch):
                batch_success = 0
                batch_errors = 0
                
                for employee in batch:
                    if progress.wasCanceled():
                        return batch_success, batch_errors
                        
                    try:
                        emp_name = employee.get('name', 'Employee')
                        payslip_content = content_generator(employee)
                        
                        doc = QTextDocument()
                        doc.setPlainText(payslip_content)
                        self.pdf_generator.page_settings_manager.configure_document(doc)
                        doc.print_(printer)
                        
                        batch_success += 1
                        logger.info(f"Printed payslip for {emp_name}")
                        
                    except Exception as e:
                        batch_errors += 1
                        logger.error(f"Error printing payslip for {emp_name}: {str(e)}")
                    
                    QApplication.processEvents()
                
                return batch_success, batch_errors
            
            success_count, error_count = batch_processor.process_batches(
                employees,
                process_print_batch,
                lambda p: progress.setValue(p * len(employees) // 100)
            )
            
            progress.close()
            
            message = (
                f"Printing completed.\n"
                f"Successfully printed: {success_count}\n"
                f"Failed: {error_count}"
            )
            QMessageBox.information(self.parent, "Print Complete", message)
            
            return success_count > 0
            
        except Exception as e:
            logger.error(f"Error in bulk printing: {str(e)}")
            QMessageBox.critical(self.parent, "Print Error", f"Error in printing: {str(e)}")
            return False

# For backwards compatibility, keep the original class name as an alias
PayslipPrinter = PayslipPDFGenerator


# Common PDF generation function used by both single and bulk operations
def generate_pdf(content, employee_name, output_directory, timestamp):
    """Generate a PDF from payslip content using singleton PageSettingsManager"""
    try:
        settings_manager = PageSettingsManager()
        
        safe_name = ''.join(c for c in employee_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        filename = f"Payslip_{safe_name}_{timestamp}.pdf"
        pdf_path = os.path.join(output_directory, filename)
        
        printer = QPrinter()
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(pdf_path)
        printer.setFullPage(True)  # Enable full page printing


        doc = QTextDocument()
        doc.setPlainText(content)
        
        settings_manager.configure_printer(printer)
        settings_manager.configure_document(doc)
        
        doc.print_(printer)
        
        return pdf_path
        
    except Exception as e:
        logger.error(f"PDF generation error for {employee_name}: {str(e)}")
        return None
