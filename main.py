import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import base64
from qr_generator import QRCodeGenerator
from qr_crypto import QRCodeCrypto
import threading
from datetime import datetime, timedelta
from report_generator import ExcelReportGenerator

# Import the Firestore database client
# This line executes the initialization code in firebase_client.py
from firebase_client import db
from data_importer import ImporterBuilder

class QRCodeGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SJLSHS Chronos QR Code Generator")

        # Make the Firestore client available to the app instance
        self.db = db
        
        # Encryption settings
        self.encryption_key = None
        self.encryption_enabled = False
        self.key_file = "encryption_key.key"
        self.root.geometry("800x700") # Increased size for better layout
        self.root.resizable(True, True)
        self.root.minsize(750, 600)

        if os.path.exists(self.key_file):
            with open(self.key_file, 'rb') as f:
                self.encryption_key = f.read()
        else:
            self.encryption_enabled = False
        
        # Initialize QR Code Generator
        self.qr_generator = QRCodeGenerator(encryption_key=self.encryption_key)
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=('Helvetica', 10))
        self.style.configure("TLabel", padding=6, font=('Helvetica', 10))
        self.style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        self.style.configure("Accent.TButton", background="#0078D7")

        # --- Main Application Structure ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # -- Tab 1: QR Code Generator --
        qr_generator_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(qr_generator_tab, text='   QR Code Generator   ')
        self._create_qr_generator_tab(qr_generator_tab)

        # -- Tab 2: Master List Import --
        master_list_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(master_list_tab, text='   Master List Import   ')
        self._create_master_list_tab(master_list_tab)

        # -- Tab 3: Attendance Reports --
        report_generator_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(report_generator_tab, text='   Attendance Reports   ')
        self._create_report_generator_tab(report_generator_tab)

    def _create_qr_generator_tab(self, tab):
        # Header
        header = ttk.Label(tab, text="Student QR Code Generator", style="Header.TLabel")
        header.pack(pady=(0, 20))
        
        # Excel File Selection
        excel_frame = ttk.LabelFrame(tab, text="1. Select Excel File for QR Codes", padding=10)
        excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_path = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(excel_frame, text="Browse...", command=self.browse_excel).pack(side=tk.RIGHT)
        
        # Output Folder Selection
        output_frame = ttk.LabelFrame(tab, text="2. Select Output Folder (Optional)", padding=10)
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), 'qr'))
        ttk.Entry(output_frame, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).pack(side=tk.RIGHT)
        
        # Encryption Settings
        encryption_frame = ttk.LabelFrame(tab, text="3. Encryption Settings", padding=10)
        encryption_frame.pack(fill=tk.X, pady=5)
        
        self.encryption_var = tk.BooleanVar(value=self.encryption_enabled)
        ttk.Checkbutton(
            encryption_frame, text="Enable Encryption", variable=self.encryption_var,
            command=self.toggle_encryption).pack(side=tk.LEFT, padx=5)
        
        self.key_status_var = tk.StringVar(value="Key: Not Set")
        ttk.Label(encryption_frame, textvariable=self.key_status_var, foreground="gray").pack(side=tk.LEFT, padx=5)
        ttk.Button(encryption_frame, text="Set Encryption Key", command=self.set_encryption_key).pack(side=tk.RIGHT, padx=5)
        
        # Progress Frame
        progress_frame = ttk.LabelFrame(tab, text="Progress", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to generate QR codes")
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # Log Frame
        log_frame = ttk.LabelFrame(progress_frame, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Generate Button
        self.generate_btn = ttk.Button(tab, text="Generate QR Codes", command=self.start_generation, style="Accent.TButton")
        self.generate_btn.pack(pady=10)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(tab, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

    def _create_master_list_tab(self, tab):
        # --- UI Elements for Master List Import ---
        self.master_list_excel_path = tk.StringVar()

        # Header
        header = ttk.Label(tab, text="Import Student Master List", style="Header.TLabel")
        header.pack(pady=(0, 20))

        # Instructions
        instructions = "Import an Excel file to populate the local student database. This database is used for validation."
        ttk.Label(tab, text=instructions, wraplength=500, justify=tk.LEFT).pack(fill=tk.X, pady=5)

        # Required Columns Info
        columns_frame = ttk.LabelFrame(tab, text="Required Excel Columns", padding=10)
        columns_frame.pack(fill=tk.X, pady=(10, 5))
        
        column_text = """
The Excel file must contain columns with the following headers (order does not matter):

• LRN
• LAST_NAME
• FIRST_NAME
• STUDENT_YEAR
• SECTION
• ADVISER
• GENDER"""
        ttk.Label(columns_frame, text=column_text, justify=tk.LEFT).pack(anchor=tk.W, padx=5, pady=5)

        # File Selection
        file_frame = ttk.LabelFrame(tab, text="1. Select Master List Excel File", padding=10)
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Entry(file_frame, textvariable=self.master_list_excel_path, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="Browse...", command=self._browse_master_list_file).pack(side=tk.RIGHT)

        # Import Button
        self.import_btn = ttk.Button(tab, text="Import Master List", command=self._start_master_list_import, style="Accent.TButton")
        self.import_btn.pack(pady=20)

        # Log Frame
        import_log_frame = ttk.LabelFrame(tab, text="Import Log", padding=10)
        import_log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.import_log_text = tk.Text(import_log_frame, height=10, wrap=tk.WORD, state='disabled')
        self.import_log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        import_scrollbar = ttk.Scrollbar(self.import_log_text, command=self.import_log_text.yview)
        import_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.import_log_text.config(yscrollcommand=import_scrollbar.set)

    def _create_report_generator_tab(self, tab):
        # --- UI Elements for Report Generation ---
        self.report_start_date = tk.StringVar()
        self.report_end_date = tk.StringVar()
        self.report_section = tk.StringVar()
        self.report_output_path = tk.StringVar()

        # Set default dates
        today = datetime.now()
        first_day_of_month = today.replace(day=1)
        self.report_start_date.set(first_day_of_month.strftime("%Y-%m-%d"))
        self.report_end_date.set(today.strftime("%Y-%m-%d"))

        # Header
        header = ttk.Label(tab, text="Generate Attendance Report", style="Header.TLabel")
        header.pack(pady=(0, 20))

        # 1. Parameters Frame
        params_frame = ttk.LabelFrame(tab, text="1. Report Parameters", padding=10)
        params_frame.pack(fill=tk.X, pady=5)

        # Date Range
        date_frame = ttk.Frame(params_frame)
        date_frame.pack(fill=tk.X, pady=5)
        ttk.Label(date_frame, text="Start Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(date_frame, textvariable=self.report_start_date, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(date_frame, text="End Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=(10, 5))
        ttk.Entry(date_frame, textvariable=self.report_end_date, width=15).pack(side=tk.LEFT, padx=5)

        # Section Filter
        section_frame = ttk.Frame(params_frame)
        section_frame.pack(fill=tk.X, pady=5)
        ttk.Label(section_frame, text="Section (Optional):").pack(side=tk.LEFT, padx=(0, 28))
        ttk.Entry(section_frame, textvariable=self.report_section, width=32).pack(side=tk.LEFT, padx=5)

        # 2. Output File
        output_frame = ttk.LabelFrame(tab, text="2. Output File", padding=10)
        output_frame.pack(fill=tk.X, pady=10)
        ttk.Entry(output_frame, textvariable=self.report_output_path, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_frame, text="Save As...", command=self._browse_report_output_file).pack(side=tk.RIGHT)

        # 3. Generate Button
        self.generate_report_btn = ttk.Button(tab, text="Generate Report", command=self._start_report_generation, style="Accent.TButton")
        self.generate_report_btn.pack(pady=20)

        # 4. Log Frame
        report_log_frame = ttk.LabelFrame(tab, text="Report Generation Log", padding=10)
        report_log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.report_log_text = tk.Text(report_log_frame, height=10, wrap=tk.WORD, state='disabled')
        self.report_log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        report_scrollbar = ttk.Scrollbar(self.report_log_text, command=self.report_log_text.yview)
        report_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.report_log_text.config(yscrollcommand=report_scrollbar.set)

    def _log_import(self, message):
        """Add a message to the import log."""
        self.import_log_text.config(state='normal')
        self.import_log_text.insert(tk.END, message + "\n")
        self.import_log_text.see(tk.END)
        self.import_log_text.config(state='disabled')
        self.root.update_idletasks()

    def _log_report_gen(self, message):
        """Add a message to the report generation log."""
        self.report_log_text.config(state='normal')
        self.report_log_text.insert(tk.END, message + "\n")
        self.report_log_text.see(tk.END)
        self.report_log_text.config(state='disabled')
        self.root.update_idletasks()

    def _browse_master_list_file(self):
        """Open file dialog to select the master list Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Master List Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.master_list_excel_path.set(file_path)

    def _browse_report_output_file(self):
        """Open file dialog to select the report output file path."""
        file_path = filedialog.asksaveasfilename(
            title="Save Report As",
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx",
            initialfile="Attendance Report.xlsx")
        if file_path:
            self.report_output_path.set(file_path)

    def _start_master_list_import(self):
        """Start the master list import process in a separate thread."""
        file_path = self.master_list_excel_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return

        self.import_btn.config(state='disabled')
        self._log_import("Starting import process...")

        import_thread = threading.Thread(
            target=self._run_master_list_import,
            args=(file_path,),
            daemon=True
        )
        import_thread.start()

    def _run_master_list_import(self, file_path):
        """The actual import logic that runs in a thread."""
        try:
            self._log_import(f"Reading data from: {os.path.basename(file_path)}")
            importer = ImporterBuilder(file_path).build()
            
            self._log_import("Parsing Excel file...")
            df = importer.parse_excel_file()
            self._log_import(f"Found {len(df)} records in the file.")

            self._log_import("Storing data into the local database... (This may take a moment)")
            importer.store_master_list(df)

            self._log_import("\nImport complete! The student master list has been updated.")
            messagebox.showinfo("Success", f"Successfully imported {len(df)} records into the master list.")

        except Exception as e:
            error_message = f"An error occurred: {e}"
            self._log_import(f"ERROR: {error_message}")
            messagebox.showerror("Import Failed", error_message)
        finally:
            self.root.after(0, lambda: self.import_btn.config(state='normal'))

    def _start_report_generation(self):
        """Validates inputs and starts the report generation in a thread."""
        # --- Input Validation ---
        try:
            start_date = datetime.strptime(self.report_start_date.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.report_end_date.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter dates in YYYY-MM-DD format.")
            return

        if start_date > end_date:
            messagebox.showerror("Invalid Date Range", "Start date cannot be after the end date.")
            return

        output_path = self.report_output_path.get()
        if not output_path:
            messagebox.showerror("Output Path Missing", "Please specify an output file path.")
            return

        section = self.report_section.get().strip()
        if not section:
            section = None # Pass None to the generator if the field is empty

        self.generate_report_btn.config(state='disabled')
        self._log_report_gen("Starting report generation...")

        # --- Run in Thread ---
        report_thread = threading.Thread(
            target=self._run_report_generation,
            args=(start_date, end_date, output_path, section),
            daemon=True
        )
        report_thread.start()

    def _run_report_generation(self, start_date, end_date, output_path, section):
        """The actual report generation logic that runs in a thread."""
        try:
            if not self.db:
                raise ConnectionError("Not connected to Firestore.")

            report_gen = ExcelReportGenerator(db_client=self.db)
            report_gen.generate_report(start_date, end_date, output_path, section)

            self._log_report_gen("\nSUCCESS: Report generated and saved to {output_path}")
            messagebox.showinfo("Success", "Attendance report has been generated successfully.")

        except Exception as e:
            error_message = f"An error occurred during report generation: {e}"
            self._log_report_gen(f"ERROR: {error_message}")
            messagebox.showerror("Report Generation Failed", error_message)
        finally:
            self.root.after(0, lambda: self.generate_report_btn.config(state='normal'))
    
    def log(self, message):
        """Add a message to the log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update the status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def browse_excel(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path.set(file_path)
            self.qr_generator.set_excel_path(file_path)
    
    def toggle_encryption(self):
        """Toggle encryption on/off"""
        self.encryption_enabled = self.encryption_var.get()
        if self.encryption_enabled and not self.encryption_key:
            if os.path.exists(self.key_file):
                try:
                    with open(self.key_file, 'rb') as f:
                        key_data = f.read()
                        self.encryption_key = base64.b64decode(key_data)
                        self.qr_generator = QRCodeGenerator(encryption_key=self.encryption_key)
                        self.encryption_enabled = True
                        self.encryption_var.set(True)
                        for widget in self.encryption_frame.winfo_children():
                            if isinstance(widget, (ttk.Checkbutton, ttk.Button)):
                                widget.config(state='disabled')
                        self.key_status_var.set("Key: Set (exists)")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load encryption key: {e}")
        
        self.update_ui_state()
    
    def set_encryption_key(self):
        """Set or generate an encryption key"""
        if os.path.exists(self.key_file):
            messagebox.showinfo("Key Exists", "Encryption key already exists and is in use.")
            return
            
        key_dialog = tk.Toplevel(self.root)
        key_dialog.title("Set Encryption Key")
        key_dialog.transient(self.root)
        key_dialog.grab_set()
        key_dialog.resizable(False, False)
        
        window_width, window_height = 400, 200
        screen_width, screen_height = key_dialog.winfo_screenwidth(), key_dialog.winfo_screenheight()
        x, y = (screen_width // 2) - (window_width // 2), (screen_height // 2) - (window_height // 2)
        key_dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        key_frame = ttk.Frame(key_dialog, padding=10)
        key_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(key_frame, text="Enter a 32-byte key (base64 encoded)").pack(pady=(0, 10))
        
        key_var = tk.StringVar()
        key_entry = ttk.Entry(key_frame, textvariable=key_var, width=50)
        key_entry.pack(pady=(0, 10), fill=tk.X)
        
        def on_generate():
            key = os.urandom(32)
            key_var.set(base64.b64encode(key).decode('utf-8'))
        
        def on_ok():
            try:
                key_str = key_var.get().strip()
                if not key_str:
                    messagebox.showerror("Error", "Key cannot be empty", parent=key_dialog)
                    return
                
                key = base64.b64decode(key_str)
                if len(key) != 32:
                    messagebox.showerror("Error", "Key must be 32 bytes (44 characters in base64)", parent=key_dialog)
                    return
                
                self.encryption_key = key
                self.qr_generator = QRCodeGenerator(encryption_key=key)
                self.qr_generator.crypto.save_key("encryption_key.key")
                self.key_status_var.set(f"Key: Set ({len(key)} bytes)")
                self.encryption_enabled = True
                self.encryption_var.set(True)
                key_dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Invalid key format: {str(e)}", parent=key_dialog)
        
        button_frame = ttk.Frame(key_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Generate Random Key", command=on_generate).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=key_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        key_entry.focus_set()
        key_dialog.bind('<Return>', lambda e: on_ok())
    
    def update_ui_state(self):
        if self.encryption_enabled and not self.encryption_key:
            self.encryption_var.set(False)
            self.encryption_enabled = False
        
        if self.encryption_key:
            key_preview = base64.b64encode(self.encryption_key[:4]).decode('utf-8')
            self.key_status_var.set(f"Key: Set ({key_preview}...)")
        else:
            self.key_status_var.set("Key: Not Set")
    
    def browse_output(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.output_path.set(dir_path)
            self.qr_generator.set_output_path(dir_path)
    
    def start_generation(self):
        if not self.excel_path.get():
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        if self.encryption_var.get() and not self.encryption_key:
            messagebox.showwarning("Warning", "Encryption is enabled but no key is set. Please set an encryption key first.")
            self.set_encryption_key()
            if not self.encryption_key:
                return
        
        self.generate_btn.config(state='disabled')
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        self.generation_thread = threading.Thread(target=self.generate_qr_codes, daemon=True)
        self.generation_thread.start()
        self.check_thread_status()
    
    def check_thread_status(self):
        if self.generation_thread.is_alive():
            self.root.after(100, self.check_thread_status)
        else:
            self.generate_btn.config(state='normal')
            self.update_status("QR Code generation completed!")
    
    def generate_qr_codes(self):
        try:
            self.update_status("Reading Excel file...")
            self.log("Reading Excel file...")
            
            try:
                df = self.qr_generator.read_excel()
                total_students = len(df)
                self.log(f"Found {total_students} students in the Excel file.")
            except Exception as e:
                self.log(f"Error reading Excel file: {str(e)}")
                messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
                return
            
            self.progress['maximum'] = total_students
            success_count = 0
            for index, row in df.iterrows():
                try:
                    student_id = str(row.get('Student ID', '')).strip()
                    if not student_id:
                        self.log(f"Skipping row {index + 2}: Missing Student ID")
                        continue
                    
                    student_name = row.get('Student Name', 'N/A').strip()
                    self.log(f"Generating QR code for {student_name} (ID: {student_id})")
                    
                    self.qr_generator.generate_qr_code(row)
                    success_count += 1
                    
                    self.progress['value'] = index + 1
                    self.update_status(f"Processed {index + 1}/{total_students} students")
                    
                except Exception as e:
                    self.log(f"Error processing student {student_id}: {str(e)}")
            
            if success_count > 0:
                self.log(f"\nSuccessfully generated {success_count} QR codes!")
                messagebox.showinfo("Success", f"Successfully generated {success_count} QR codes!")
            else:
                self.log("\nNo QR codes were generated. Please check the log for errors.")
                
        except Exception as e:
            self.log(f"An unexpected error occurred: {str(e)}")
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.update_status("Ready")


    def _create_qr_generator_tab(self, tab):
        # Header
        header = ttk.Label(tab, text="Student QR Code Generator", style="Header.TLabel")
        header.pack(pady=(0, 20))
        
        # Excel File Selection
        excel_frame = ttk.LabelFrame(tab, text="1. Select Excel File for QR Codes", padding=10)
        excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_path = tk.StringVar()
        ttk.Entry(excel_frame, textvariable=self.excel_path, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(excel_frame, text="Browse...", command=self.browse_excel).pack(side=tk.RIGHT)
        
        # Output Folder Selection
        output_frame = ttk.LabelFrame(tab, text="2. Select Output Folder (Optional)", padding=10)
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), 'qr'))
        ttk.Entry(output_frame, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).pack(side=tk.RIGHT)
        
        # Encryption Settings
        encryption_frame = ttk.LabelFrame(tab, text="3. Encryption Settings", padding=10)
        encryption_frame.pack(fill=tk.X, pady=5)
        
        self.encryption_var = tk.BooleanVar(value=self.encryption_enabled)
        ttk.Checkbutton(
            encryption_frame, text="Enable Encryption", variable=self.encryption_var,
            command=self.toggle_encryption).pack(side=tk.LEFT, padx=5)
        
        self.key_status_var = tk.StringVar(value="Key: Not Set")
        ttk.Label(encryption_frame, textvariable=self.key_status_var, foreground="gray").pack(side=tk.LEFT, padx=5)
        ttk.Button(encryption_frame, text="Set Encryption Key", command=self.set_encryption_key).pack(side=tk.RIGHT, padx=5)
        
        # Progress Frame
        progress_frame = ttk.LabelFrame(tab, text="Progress", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to generate QR codes")
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # Log Frame
        log_frame = ttk.LabelFrame(progress_frame, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Generate Button
        self.generate_btn = ttk.Button(tab, text="Generate QR Codes", command=self.start_generation, style="Accent.TButton")
        self.generate_btn.pack(pady=10)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(tab, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

    def _create_master_list_tab(self, tab):
        # --- UI Elements for Master List Import ---
        self.master_list_excel_path = tk.StringVar()

        # Header
        header = ttk.Label(tab, text="Import Student Master List", style="Header.TLabel")
        header.pack(pady=(0, 20))

        # Instructions
        instructions = "Import an Excel file to populate the local student database. This database is used for validation."
        ttk.Label(tab, text=instructions, wraplength=500, justify=tk.LEFT).pack(fill=tk.X, pady=5)

        # Required Columns Info
        columns_frame = ttk.LabelFrame(tab, text="Required Excel Columns", padding=10)
        columns_frame.pack(fill=tk.X, pady=(10, 5))
        
        column_text = """The Excel file must contain columns with the following headers (order does not matter):

• LRN
• LAST_NAME
• FIRST_NAME
• STUDENT_YEAR
• SECTION
• ADVISER
• GENDER"""
        ttk.Label(columns_frame, text=column_text, justify=tk.LEFT).pack(anchor=tk.W, padx=5, pady=5)

        # File Selection
        file_frame = ttk.LabelFrame(tab, text="1. Select Master List Excel File", padding=10)
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Entry(file_frame, textvariable=self.master_list_excel_path, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="Browse...", command=self._browse_master_list_file).pack(side=tk.RIGHT)

        # Import Button
        self.import_btn = ttk.Button(tab, text="Import Master List", command=self._start_master_list_import, style="Accent.TButton")
        self.import_btn.pack(pady=20)

        # Log Frame
        import_log_frame = ttk.LabelFrame(tab, text="Import Log", padding=10)
        import_log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.import_log_text = tk.Text(import_log_frame, height=10, wrap=tk.WORD, state='disabled')
        self.import_log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        import_scrollbar = ttk.Scrollbar(self.import_log_text, command=self.import_log_text.yview)
        import_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.import_log_text.config(yscrollcommand=import_scrollbar.set)

    def _create_report_generator_tab(self, tab):
        # --- UI Elements for Report Generation ---
        self.report_start_date = tk.StringVar()
        self.report_end_date = tk.StringVar()
        self.report_section = tk.StringVar()
        self.report_output_path = tk.StringVar()

        # Set default dates
        today = datetime.now()
        first_day_of_month = today.replace(day=1)
        self.report_start_date.set(first_day_of_month.strftime("%Y-%m-%d"))
        self.report_end_date.set(today.strftime("%Y-%m-%d"))

        # Header
        header = ttk.Label(tab, text="Generate Attendance Report", style="Header.TLabel")
        header.pack(pady=(0, 20))

        # 1. Parameters Frame
        params_frame = ttk.LabelFrame(tab, text="1. Report Parameters", padding=10)
        params_frame.pack(fill=tk.X, pady=5)

        # Date Range
        date_frame = ttk.Frame(params_frame)
        date_frame.pack(fill=tk.X, pady=5)
        ttk.Label(date_frame, text="Start Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(date_frame, textvariable=self.report_start_date, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(date_frame, text="End Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=(10, 5))
        ttk.Entry(date_frame, textvariable=self.report_end_date, width=15).pack(side=tk.LEFT, padx=5)

        # Section Filter
        section_frame = ttk.Frame(params_frame)
        section_frame.pack(fill=tk.X, pady=5)
        ttk.Label(section_frame, text="Section (Optional):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(section_frame, textvariable=self.report_section, width=32).pack(side=tk.LEFT, padx=5)

        # 2. Output File
        output_frame = ttk.LabelFrame(tab, text="2. Output File", padding=10)
        output_frame.pack(fill=tk.X, pady=10)
        ttk.Entry(output_frame, textvariable=self.report_output_path, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_frame, text="Save As...", command=self._browse_report_output_file).pack(side=tk.RIGHT)

        # 3. Generate Button
        self.generate_report_btn = ttk.Button(tab, text="Generate Report", command=self._start_report_generation, style="Accent.TButton")
        self.generate_report_btn.pack(pady=20)

        # 4. Log Frame
        report_log_frame = ttk.LabelFrame(tab, text="Report Generation Log", padding=10)
        report_log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.report_log_text = tk.Text(report_log_frame, height=10, wrap=tk.WORD, state='disabled')
        self.report_log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        report_scrollbar = ttk.Scrollbar(self.report_log_text, command=self.report_log_text.yview)
        report_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.report_log_text.config(yscrollcommand=report_scrollbar.set)

    def _log_import(self, message):
        """Add a message to the import log."""
        self.import_log_text.config(state='normal')
        self.import_log_text.insert(tk.END, message + "\n")
        self.import_log_text.see(tk.END)
        self.import_log_text.config(state='disabled')
        self.root.update_idletasks()

    def _log_report_gen(self, message):
        """Add a message to the report generation log."""
        self.report_log_text.config(state='normal')
        self.report_log_text.insert(tk.END, message + "\n")
        self.report_log_text.see(tk.END)
        self.report_log_text.config(state='disabled')
        self.root.update_idletasks()

    def _browse_master_list_file(self):
        """Open file dialog to select the master list Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Master List Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.master_list_excel_path.set(file_path)

    def _browse_report_output_file(self):
        """Open file dialog to select the report output file path."""
        file_path = filedialog.asksaveasfilename(
            title="Save Report As",
            filetypes=[("Excel files", "*.xlsx")],
            defaultextension=".xlsx",
            initialfile="Attendance Report.xlsx")
        if file_path:
            self.report_output_path.set(file_path)


    def _start_master_list_import(self):
        """Start the master list import process in a separate thread."""
        file_path = self.master_list_excel_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return

        self.import_btn.config(state='disabled')
        self._log_import("Starting import process...")

        import_thread = threading.Thread(
            target=self._run_master_list_import,
            args=(file_path,),
            daemon=True
        )
        import_thread.start()

    def _run_master_list_import(self, file_path):
        """The actual import logic that runs in a thread."""
        try:
            self._log_import(f"Reading data from: {os.path.basename(file_path)}")
            importer = ImporterBuilder(file_path).build()
            
            self._log_import("Parsing Excel file...")
            df = importer.parse_excel_file()
            self._log_import(f"Found {len(df)} records in the file.")

            self._log_import("Storing data into the local database... (This may take a moment)")
            importer.store_master_list(df)

            self._log_import("\nImport complete! The student master list has been updated.")
            messagebox.showinfo("Success", f"Successfully imported {len(df)} records into the master list.")

        except Exception as e:
            error_message = f"An error occurred: {e}"
            self._log_import(f"ERROR: {error_message}")
            messagebox.showerror("Import Failed", error_message)
        finally:
            self.root.after(0, lambda: self.import_btn.config(state='normal'))

    def _start_report_generation(self):
        """Validates inputs and starts the report generation in a thread."""
        # --- Input Validation ---
        try:
            start_date = datetime.strptime(self.report_start_date.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.report_end_date.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter dates in YYYY-MM-DD format.")
            return

        if start_date > end_date:
            messagebox.showerror("Invalid Date Range", "Start date cannot be after the end date.")
            return

        output_path = self.report_output_path.get()
        if not output_path:
            messagebox.showerror("Output Path Missing", "Please specify an output file path.")
            return

        section = self.report_section.get()
        if section == "All Sections":
            section = None # Pass None to the generator to get all sections

        self.generate_report_btn.config(state='disabled')
        self._log_report_gen("Starting report generation...")

        # --- Run in Thread ---
        report_thread = threading.Thread(
            target=self._run_report_generation,
            args=(start_date, end_date, output_path, section),
            daemon=True
        )
        report_thread.start()

    def _run_report_generation(self, start_date, end_date, output_path, section):
        """The actual report generation logic that runs in a thread."""
        try:
            if not self.db:
                raise ConnectionError("Not connected to Firestore.")

            report_gen = ExcelReportGenerator(db_client=self.db)
            report_gen.generate_report(start_date, end_date, output_path, section)

            self._log_report_gen(f"\nSUCCESS: Report generated and saved to {output_path}")
            messagebox.showinfo("Success", "Attendance report has been generated successfully.")

        except Exception as e:
            error_message = f"An error occurred during report generation: {e}"
            self._log_report_gen(f"ERROR: {error_message}")
            messagebox.showerror("Report Generation Failed", error_message)
        finally:
            self.root.after(0, lambda: self.generate_report_btn.config(state='normal'))
    
    def log(self, message):
        """Add a message to the log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update the status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def browse_excel(self):
        """Open file dialog to select Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path.set(file_path)
            self.qr_generator.set_excel_path(file_path)
    
    def toggle_encryption(self):
        """Toggle encryption on/off"""
        self.encryption_enabled = self.encryption_var.get()
        if self.encryption_enabled and not self.encryption_key:
            if os.path.exists(self.key_file):
                try:
                    with open(self.key_file, 'rb') as f:
                        key_data = f.read()
                        self.encryption_key = base64.b64decode(key_data)
                        self.qr_generator = QRCodeGenerator(encryption_key=self.encryption_key)
                        self.encryption_enabled = True
                        self.encryption_var.set(True)
                        for widget in self.encryption_frame.winfo_children():
                            if isinstance(widget, (ttk.Checkbutton, ttk.Button)):
                                widget.config(state='disabled')
                        self.key_status_var.set("Key: Set (exists)")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load encryption key: {e}")
        
        self.update_ui_state()
    
    def set_encryption_key(self):
        """Set or generate an encryption key"""
        if os.path.exists(self.key_file):
            messagebox.showinfo("Key Exists", "Encryption key already exists and is in use.")
            return
            
        key_dialog = tk.Toplevel(self.root)
        key_dialog.title("Set Encryption Key")
        key_dialog.transient(self.root)
        key_dialog.grab_set()
        key_dialog.resizable(False, False)
        
        window_width, window_height = 400, 200
        screen_width, screen_height = key_dialog.winfo_screenwidth(), key_dialog.winfo_screenheight()
        x, y = (screen_width // 2) - (window_width // 2), (screen_height // 2) - (window_height // 2)
        key_dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        key_frame = ttk.Frame(key_dialog, padding=10)
        key_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(key_frame, text="Enter a 32-byte key (base64 encoded)").pack(pady=(0, 10))
        
        key_var = tk.StringVar()
        key_entry = ttk.Entry(key_frame, textvariable=key_var, width=50)
        key_entry.pack(pady=(0, 10), fill=tk.X)
        
        def on_generate():
            key = os.urandom(32)
            key_var.set(base64.b64encode(key).decode('utf-8'))
        
        def on_ok():
            try:
                key_str = key_var.get().strip()
                if not key_str:
                    messagebox.showerror("Error", "Key cannot be empty", parent=key_dialog)
                    return
                
                key = base64.b64decode(key_str)
                if len(key) != 32:
                    messagebox.showerror("Error", "Key must be 32 bytes (44 characters in base64)", parent=key_dialog)
                    return
                
                self.encryption_key = key
                self.qr_generator = QRCodeGenerator(encryption_key=key)
                self.qr_generator.crypto.save_key("encryption_key.key")
                self.key_status_var.set(f"Key: Set ({len(key)} bytes)")
                self.encryption_enabled = True
                self.encryption_var.set(True)
                key_dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Invalid key format: {str(e)}", parent=key_dialog)
        
        button_frame = ttk.Frame(key_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Generate Random Key", command=on_generate).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=key_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        key_entry.focus_set()
        key_dialog.bind('<Return>', lambda e: on_ok())
    
    def update_ui_state(self):
        if self.encryption_enabled and not self.encryption_key:
            self.encryption_var.set(False)
            self.encryption_enabled = False
        
        if self.encryption_key:
            key_preview = base64.b64encode(self.encryption_key[:4]).decode('utf-8')
            self.key_status_var.set(f"Key: Set ({key_preview}...)")
        else:
            self.key_status_var.set("Key: Not Set")
    
    def browse_output(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.output_path.set(dir_path)
            self.qr_generator.set_output_path(dir_path)
    
    def start_generation(self):
        if not self.excel_path.get():
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        if self.encryption_var.get() and not self.encryption_key:
            messagebox.showwarning("Warning", "Encryption is enabled but no key is set. Please set an encryption key first.")
            self.set_encryption_key()
            if not self.encryption_key:
                return
        
        self.generate_btn.config(state='disabled')
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        self.generation_thread = threading.Thread(target=self.generate_qr_codes, daemon=True)
        self.generation_thread.start()
        self.check_thread_status()
    
    def check_thread_status(self):
        if self.generation_thread.is_alive():
            self.root.after(100, self.check_thread_status)
        else:
            self.generate_btn.config(state='normal')
            self.update_status("QR Code generation completed!")
    
    def generate_qr_codes(self):
        try:
            self.update_status("Reading Excel file...")
            self.log("Reading Excel file...")
            
            try:
                df = self.qr_generator.read_excel()
                total_students = len(df)
                self.log(f"Found {total_students} students in the Excel file.")
            except Exception as e:
                self.log(f"Error reading Excel file: {str(e)}")
                messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
                return
            
            self.progress['maximum'] = total_students
            success_count = 0
            for index, row in df.iterrows():
                try:
                    student_id = str(row.get('Student ID', '')).strip()
                    if not student_id:
                        self.log(f"Skipping row {index + 2}: Missing Student ID")
                        continue
                    
                    student_name = row.get('Student Name', 'N/A').strip()
                    self.log(f"Generating QR code for {student_name} (ID: {student_id})")
                    
                    self.qr_generator.generate_qr_code(row)
                    success_count += 1
                    
                    self.progress['value'] = index + 1
                    self.update_status(f"Processed {index + 1}/{total_students} students")
                    
                except Exception as e:
                    self.log(f"Error processing student {student_id}: {str(e)}")
            
            if success_count > 0:
                self.log(f"\nSuccessfully generated {success_count} QR codes!")
                messagebox.showinfo("Success", f"Successfully generated {success_count} QR codes!")
            else:
                self.log("\nNo QR codes were generated. Please check the log for errors.")
                
        except Exception as e:
            self.log(f"An unexpected error occurred: {str(e)}")
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.update_status("Ready")


def main():
    # Create the main window
    root = tk.Tk()
    
    # Set the theme (requires ttkthemes package)
    try:
        from ttkthemes import ThemedStyle
        style = ThemedStyle(root)
        style.set_theme("arc")
    except ImportError:
        # Fallback to default theme if ttkthemes is not available
        pass
    
    # Set application icon if available
    try:
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass
    
    # Create and run the application
    app = QRCodeGeneratorApp(root)
    
    # Make the window resizable
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    
    # Center the window on screen
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f'{window_width}x{window_height}+{x}+{y}')
    
    # Start the application
    root.mainloop()

if __name__ == "__main__":
    main()
