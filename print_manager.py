import os
import platform
import tempfile
import logging
import subprocess
from datetime import datetime
from PyQt5.QtCore import QObject, pyqtSignal, QThread, QSizeF
from PyQt5.QtWidgets import QMessageBox, QDialog, QProgressBar, QLabel, QVBoxLayout, QPushButton, QHBoxLayout, QFileDialog, QApplication
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog
from PyQt5.QtGui import QTextDocument, QFont
import uuid

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

    def __init__(self, employees, content_generator, output_directory=None, page_settings=None):
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
        
        # Use provided output directory or default to ~/Payslips
        self.output_directory = output_directory or os.path.join(os.path.expanduser("~"), "Payslips")
        
        # Create output directory if it doesn't exist
        if not os.path.exists(self.output_directory):
            os.makedirs(self.output_directory)
        
        # Page settings: paper size, margins, and padding
        self.page_settings = page_settings or {
            "paper_size": QPrinter.B4,
            "custom_size": QSizeF(220,200),  # Custom size as QSizeF(width, height) in millimeters
            "margins": {"top": 0, "bottom": 0, "left": 0, "right": 0},  # Default margins in mm
            "padding": 0  
        }

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
                                      timestamp, self.page_settings)
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


class PayslipPDFGenerator:
    """Main class for handling payslip PDF generation"""
    def __init__(self, parent_widget):
        self.parent = parent_widget
        self.pdf_worker = None
        self.progress_dialog = None
        
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
        
        # Page settings: paper size, margins, and padding
        self.page_settings = {
            "paper_size": QPrinter.B5,
            "custom_size": QSizeF(220,200),  # Custom size as QSizeF(width, height) in millimeters
            "margins": {"top": 0, "bottom": 0, "left": 0, "right": 0},  # Default margins in mm
            "padding":0   # Default padding in mm
        }

    def set_page_settings(self, paper_size=None, custom_size=None, margins=None, padding=None):
        """Set custom page settings"""
        if paper_size:
            self.page_settings["paper_size"] = paper_size
            self.page_settings["custom_size"] = None  # Reset custom size if paper size is set
        if custom_size:
            self.page_settings["custom_size"] = QSizeF(*custom_size)  # Expect tuple (width, height)
            self.page_settings["paper_size"] = None  # Reset paper size if custom size is set
        if margins:
            self.page_settings["margins"] = margins
        if padding is not None:
            self.page_settings["padding"] = padding

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
                                  timestamp, self.page_settings)
            
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

    def generate_bulk_pdfs(self, employees, content_generator, ask_directory=True):
        """
        Generate multiple payslip PDFs with optional directory selection
        
        Args:
            employees: List of employee information (id, name, etc.)
            content_generator: Function that takes employee info and returns payslip content
            ask_directory: Whether to ask user for directory selection
        """
        if not employees:
            QMessageBox.warning(self.parent, "No Data", "No employees selected for PDF generation.")
            return
        
        try:
            # Get output directory
            if ask_directory:
                output_dir = self.get_output_directory("Select Directory to Save PDF Payslips")
                if not output_dir:  # User cancelled
                    return
            else:
                output_dir = self.output_directory
            
            # Ensure output directory exists
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Create progress dialog
            self.progress_dialog = PDFProgressDialog(self.parent, len(employees))
            self.progress_dialog.cancel_button.clicked.connect(self._cancel_generation)
            self.progress_dialog.show()
            
            # Create worker thread with employees list, content generator, and output directory
            self.pdf_worker = PDFGeneratorWorker(employees, content_generator, output_dir, self.page_settings)
            self.pdf_worker.progress_updated.connect(self.progress_dialog.update_progress)
            self.pdf_worker.single_pdf_complete.connect(self.progress_dialog.update_current_job)
            self.pdf_worker.process_finished.connect(lambda success, error: 
                                                 self._on_generation_finished(success, error, output_dir))
            
            # Start PDF generation
            self.pdf_worker.start()
            
        except Exception as e:
            logger.error(f"Error setting up PDF generation: {str(e)}")
            if self.progress_dialog:
                self.progress_dialog.close()
            QMessageBox.critical(self.parent, "PDF Error", f"Error setting up PDF generation: {str(e)}")

    def _cancel_generation(self):
        """Cancel the ongoing PDF generation process"""
        if self.pdf_worker and self.pdf_worker.isRunning():
            self.pdf_worker.cancel()
            self.progress_dialog.current_job.setText("Cancelling...")

    def _on_generation_finished(self, success_count, error_count, output_directory):
        """Handle completion of PDF generation"""
        if self.progress_dialog:
            self.progress_dialog.close()
        
        message = (
            f"PDF generation completed.\n"
            f"Successfully generated: {success_count}\n"
            f"Failed: {error_count}\n\n"
            f"PDFs saved to: {output_directory}"
        )
        
        QMessageBox.information(self.parent, "PDF Generation Complete", message)

    def _setup_document_for_print(self, printer, content):
        """Set up the document for PDF generation with monospaced font and configured page settings"""
        from PyQt5.QtGui import QTextDocument, QFont
        
        # Apply the configured page settings
        apply_page_settings_to_printer(printer, self.page_settings)
        
        # Create a text document with the content
        doc = QTextDocument()
        font = QFont("Courier New", 10)  # Use monospaced font
        doc.setDefaultFont(font)
        doc.setPlainText(content)
        
        # Apply padding from page settings
        padding = self.page_settings.get("padding", 0)
        doc.setDocumentMargin(padding)
        
        # Generate the PDF
        doc.print_(printer)


# For backwards compatibility, keep the original class name as an alias
PayslipPrinter = PayslipPDFGenerator


class PayslipPrintManager:
    """Enhanced print manager that handles both printing and PDF generation for B4 paper size"""
    
    def __init__(self, parent_widget):
        self.parent = parent_widget
        self.pdf_generator = PayslipPDFGenerator(parent_widget)
    
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
            apply_page_settings_to_printer(printer, self.pdf_generator.page_settings)
            printer.setOrientation(QPrinter.Portrait)
            
            # Create print preview dialog
            preview_dialog = QPrintPreviewDialog(printer, self.parent)
            preview_dialog.setWindowTitle(f"Print Preview - {employee_name}")
            
            # Connect preview signal
            def print_preview(printer):
                doc = QTextDocument()
                font = QFont("Courier New", 10)
                doc.setDefaultFont(font)
                doc.setPlainText(payslip_content)
                
                # Apply padding from page settings
                padding = self.pdf_generator.page_settings.get("padding", 0)
                doc.setDocumentMargin(padding)
                
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
        """Print multiple payslips directly to printer with appropriate page settings"""
        if not employees:
            QMessageBox.warning(self.parent, "No Data", "No employees selected for printing.")
            return False
        
        # Generate all payslip contents
        payslips = []
        for employee in employees:
            try:
                emp_name = employee.get('name', 'Employee')
                payslip_content = content_generator(employee)
                payslips.append({
                    "content": payslip_content,
                    "name": emp_name
                })
            except Exception as e:
                logger.error(f"Error generating payslip content for {employee.get('name', 'Employee')}: {str(e)}")
                # Continue with other payslips even if one fails
        
        if not payslips:
            QMessageBox.warning(self.parent, "Generation Error", "Failed to generate any payslips for printing.")
            return False
            
        # Print all the generated payslips
        return self._print_payslips(payslips, title="Print All Payslips")
    
    def _print_payslips(self, payslips, title="Print Payslips"):
        """
        Common function to print one or more payslips
        
        Args:
            payslips: List of dictionaries with 'content' and 'name' keys for each payslip
            title: Title for the print dialog
            
        Returns:
            Boolean indicating success or failure
        """
        if not payslips:
            return False
            
        try:
            # Create a printer
            printer = QPrinter()
            
            # Apply the same page settings used for PDF generation
            apply_page_settings_to_printer(printer, self.pdf_generator.page_settings)
            printer.setOrientation(QPrinter.Portrait)
            
            # Show print dialog
            print_dialog = QPrintDialog(printer, self.parent)
            print_dialog.setWindowTitle(title)
            
            if print_dialog.exec_() != QPrintDialog.Accepted:
                logger.info("Print cancelled by user")
                return False
            
            success_count = 0
            error_count = 0
            
            for payslip in payslips:
                try:
                    # Set up the document for printing
                    doc = QTextDocument()
                    font = QFont("Courier New", 10)
                    doc.setDefaultFont(font)
                    doc.setPlainText(payslip["content"])
                    
                    # Apply padding from page settings
                    padding = self.pdf_generator.page_settings.get("padding", 0)
                    doc.setDocumentMargin(padding)
                    
                    # Print the document
                    doc.print_(printer)
                    
                    success_count += 1
                    logger.info(f"Printed payslip for {payslip['name']}")
                    
                except Exception as e:
                    error_count += 1
                    logger.error(f"Error printing payslip for {payslip['name']}: {str(e)}")
            
            # Show completion message if multiple payslips were printed
            if len(payslips) > 1:
                message = (
                    f"Printing completed.\n"
                    f"Successfully printed: {success_count}\n"
                    f"Failed: {error_count}"
                )
                QMessageBox.information(self.parent, "Print Complete", message)
                
            return success_count > 0
            
        except Exception as e:
            logger.error(f"Error in printing: {str(e)}")
            QMessageBox.critical(self.parent, "Print Error", f"Error in printing: {str(e)}")
            return False
    
    @property
    def output_directory(self):
        """Get the output directory from PDF generator"""
        return self.pdf_generator.output_directory


# Common function to apply page settings to printer
def apply_page_settings_to_printer(printer, page_settings):
    """
    Apply page settings from config to a printer instance
    
    Args:
        printer: QPrinter instance to configure
        page_settings: Dictionary containing paper size, margins, and padding settings
    """
    # Apply paper size
    if page_settings.get("custom_size"):
        custom_size = page_settings["custom_size"]
        printer.setPaperSize(custom_size, QPrinter.Millimeter)
    else:
        printer.setPaperSize(page_settings.get("paper_size", QPrinter.B4))
    
    # Apply margins with numeric defaults
    margins = page_settings.get("margins", {"top": 0, "bottom": 0, "left": 0, "right": 0})
    
    # Convert any "auto" or non-numeric values to 0
    margin_left = float(margins["left"]) if isinstance(margins["left"], (int, float)) else 0
    margin_top = float(margins["top"]) if isinstance(margins["top"], (int, float)) else 0
    margin_right = float(margins["right"]) if isinstance(margins["right"], (int, float)) else 0
    margin_bottom = float(margins["bottom"]) if isinstance(margins["bottom"], (int, float)) else 0
    
    printer.setPageMargins(margin_left, margin_top, margin_right, margin_bottom, QPrinter.Millimeter)


# Common PDF generation function used by both single and bulk operations
def generate_pdf(content, employee_name, output_directory, timestamp, page_settings):
    """
    Generate a PDF from payslip content with customizable page settings
    
    Args:
        content: Payslip content to include in PDF
        employee_name: Name of employee for filename generation
        output_directory: Directory to save the PDF
        timestamp: Timestamp to include in filename
        page_settings: Dictionary containing paper size, margins, and padding settings
    
    Returns:
        Path to generated PDF or None if generation failed
    """
    try:
        # Clean the employee name for use in filename
        safe_name = ''.join(c for c in employee_name if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_name = safe_name.replace(' ', '_')
        
        # Create filename
        filename = f"Payslip_{safe_name}_{timestamp}.pdf"
        pdf_path = os.path.join(output_directory, filename)
        
        # Create a printer set to PDF output
        printer = QPrinter()
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(pdf_path)

        # Apply page settings using the common function
        apply_page_settings_to_printer(printer, page_settings)
        
        # Create a text document with the content
        doc = QTextDocument()
        font = QFont("Courier New", 10)  # Use monospaced font
        doc.setDefaultFont(font)
        doc.setPlainText(content)
        
        # Apply padding (if needed, adjust content layout)
        padding = page_settings.get("padding", 0)
        doc.setDocumentMargin(padding)
        
        # Generate the PDF directly
        doc.print_(printer)
        
        return pdf_path
        
    except Exception as e:
        logger.error(f"PDF generation error for {employee_name}: {str(e)}")
        return None
