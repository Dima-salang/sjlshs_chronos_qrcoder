import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import base64
from qr_generator import QRCodeGenerator
from qr_crypto import QRCodeCrypto
import threading

# Import the Firestore database client
# This line executes the initialization code in firebase_client.py
from firebase_client import db
from data_importer import FirestoreDataImporter

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
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        self.root.minsize(600, 450)

        if os.path.exists(self.key_file):
            # read the file
            with open(self.key_file, 'rb') as f:
                self.encryption_key = f.read()
        else:
            self.encryption_enabled = False
        
        print(self.encryption_key)

        

        
        # Initialize QR Code Generator
        self.qr_generator = QRCodeGenerator(encryption_key=self.encryption_key)
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=('Helvetica', 10))
        self.style.configure("TLabel", padding=6, font=('Helvetica', 10))
        self.style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))
        
        # Main container
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Label(
            self.main_frame,
            text="Student QR Code Generator",
            style="Header.TLabel"
        )
        header.pack(pady=(0, 20))
        
        # Excel File Selection
        self.excel_frame = ttk.LabelFrame(self.main_frame, text="1. Select Excel File", padding=10)
        self.excel_frame.pack(fill=tk.X, pady=5)
        
        self.excel_path = tk.StringVar()
        ttk.Entry(self.excel_frame, textvariable=self.excel_path, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(
            self.excel_frame,
            text="Browse...",
            command=self.browse_excel
        ).pack(side=tk.RIGHT)
        
        # Output Folder Selection
        self.output_frame = ttk.LabelFrame(self.main_frame, text="3. Select Output Folder (Optional)", padding=10)
        self.output_frame.pack(fill=tk.X, pady=5)
        
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), 'qr'))
        ttk.Entry(self.output_frame, textvariable=self.output_path).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(
            self.output_frame,
            text="Browse...",
            command=self.browse_output
        ).pack(side=tk.RIGHT)
        
        # Encryption Settings
        self.encryption_frame = ttk.LabelFrame(self.main_frame, text="4. Encryption Settings", padding=10)
        self.encryption_frame.pack(fill=tk.X, pady=5)
        
        self.encryption_var = tk.BooleanVar(value=self.encryption_enabled)
        ttk.Checkbutton(
            self.encryption_frame,
            text="Enable Encryption",
            variable=self.encryption_var,
            command=self.toggle_encryption
        ).pack(side=tk.LEFT, padx=5)
        
        self.key_status_var = tk.StringVar(value="Key: Not Set")
        ttk.Label(
            self.encryption_frame,
            textvariable=self.key_status_var,
            foreground="gray"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            self.encryption_frame,
            text="Set Encryption Key",
            command=self.set_encryption_key
        ).pack(side=tk.RIGHT, padx=5)
        
        # Progress Frame
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="Progress", padding=10)
        self.progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.progress_label = ttk.Label(self.progress_frame, text="Ready to generate QR codes")
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(
            self.progress_frame,
            orient=tk.HORIZONTAL,
            length=100,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X, pady=5)
        
        # Log Frame
        self.log_frame = ttk.LabelFrame(self.progress_frame, text="Log")
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        
        self.log_text = tk.Text(self.log_frame, height=8, wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        scrollbar = ttk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        
        # Generate Button
        self.generate_btn = ttk.Button(
            self.main_frame,
            text="Generate QR Codes",
            command=self.start_generation,
            style="Accent.TButton"
        )
        self.generate_btn.pack(pady=10)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            self.main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # Configure the grid weights
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)
    
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
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.excel_path.set(file_path)
            self.qr_generator.set_excel_path(file_path)
    
    def toggle_encryption(self):
        """Toggle encryption on/off"""
        self.encryption_enabled = self.encryption_var.get()
        if self.encryption_enabled and not self.encryption_key:
            # Check if key file exists and load it
            if os.path.exists(self.key_file):
                try:
                    with open(self.key_file, 'rb') as f:
                        key_data = f.read()
                        self.encryption_key = base64.b64decode(key_data)
                        self.qr_generator = QRCodeGenerator(encryption_key=self.encryption_key)
                        print(self.encryption_key)
                        self.encryption_enabled = True
                        self.encryption_var.set(True)
                        # Disable encryption controls
                        for widget in self.encryption_frame.winfo_children():
                            if isinstance(widget, (ttk.Checkbutton, ttk.Button)):
                                widget.config(state='disabled')
                        self.key_status_var.set("Key: Set (exists)")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load encryption key: {e}")
        
        # Start with UI disabled until files are selected
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
        
        # Center the dialog
        window_width = 400
        window_height = 200
        screen_width = key_dialog.winfo_screenwidth()
        screen_height = key_dialog.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        key_dialog.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        # Key input frame
        key_frame = ttk.Frame(key_dialog, padding=10)
        key_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(key_frame, text="Enter a 32-byte key (base64 encoded)").pack(pady=(0, 10))
        
        key_var = tk.StringVar()
        key_entry = ttk.Entry(key_frame, textvariable=key_var, width=50)
        key_entry.pack(pady=(0, 10), fill=tk.X)
        
        def on_generate():
            """Generate a new random key"""
            key = os.urandom(32)
            key_var.set(base64.b64encode(key).decode('utf-8'))
        
        def on_ok():
            """Validate and set the key"""
            try:
                key_str = key_var.get().strip()
                if not key_str:
                    messagebox.showerror("Error", "Key cannot be empty")
                    return
                
                # Try to decode the key
                key = base64.b64decode(key_str)
                if len(key) != 32:
                    messagebox.showerror("Error", "Key must be 32 bytes (44 characters in base64)")
                    return
                
                self.encryption_key = key
                self.qr_generator = QRCodeGenerator(encryption_key=key)
                self.qr_generator.crypto.save_key("encryption_key.key")
                self.key_status_var.set(f"Key: Set ({len(key)} bytes)")
                self.encryption_enabled = True
                self.encryption_var.set(True)
                key_dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Invalid key format: {str(e)}")
        
        button_frame = ttk.Frame(key_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Generate Random Key", command=on_generate).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=key_dialog.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Set focus to the entry field
        key_entry.focus_set()
        key_dialog.bind('<Return>', lambda e: on_ok())
    
    def update_ui_state(self):
        """Update the UI state based on current settings"""
        if self.encryption_enabled and not self.encryption_key:
            self.encryption_var.set(False)
            self.encryption_enabled = False
        
        # Update key status
        if self.encryption_key:
            key_preview = base64.b64encode(self.encryption_key[:4]).decode('utf-8')
            encryption_key = base64.b64encode(self.encryption_key).decode('utf-8')
            print(encryption_key)
            self.key_status_var.set(f"Key: Set ({key_preview}...)")
        else:
            self.key_status_var.set("Key: Not Set")
    
    def browse_output(self):
        """Open directory dialog to select output folder"""
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.output_path.set(dir_path)
            self.qr_generator.set_output_path(dir_path)
    
    def start_generation(self):
        """Start the QR code generation in a separate thread"""
        if not self.excel_path.get():
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
            
        
        if self.encryption_var.get() and not self.encryption_key:
            messagebox.showwarning("Warning", "Encryption is enabled but no key is set. Please set an encryption key first.")
            self.set_encryption_key()
            if not self.encryption_key:
                return
        
        # Disable generate button during generation
        self.generate_btn.config(state='disabled')
        
        # Clear previous logs
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        
        # Start generation in a separate thread
        self.generation_thread = threading.Thread(
            target=self.generate_qr_codes,
            daemon=True
        )
        self.generation_thread.start()
        
        # Check the thread status periodically
        self.check_thread_status()
    
    def check_thread_status(self):
        """Check if the generation thread is still running"""
        if self.generation_thread.is_alive():
            self.root.after(100, self.check_thread_status)
        else:
            self.generate_btn.config(state='normal')
            self.update_status("QR Code generation completed!")
    
    def generate_qr_codes(self):
        """Generate QR codes based on the provided Excel file"""
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
            
            # Update progress bar
            self.progress['maximum'] = total_students
            
            # Generate QR codes
            success_count = 0
            for index, row in df.iterrows():
                try:
                    student_id = str(row.get('Student ID', '')).strip()
                    if not student_id:
                        self.log(f"Skipping row {index + 2}: Missing Student ID")
                        continue
                        
                    student_name = row.get('Student Name', 'N/A').strip()
                    year = str(row.get('Year', '')).strip()
                    section = str(row.get('Section', '')).strip()
                    
                    self.log(f"Generating QR code for {student_name} (ID: {student_id})")
                    
                    # Generate QR code
                    self.qr_generator.generate_qr_code(row)
                    success_count += 1
                    
                    # Update progress
                    self.progress['value'] = index + 1
                    self.update_status(f"Processed {index + 1}/{total_students} students")
                    
                except Exception as e:
                    self.log(f"Error processing student {student_id}: {str(e)}")
            
            # Show completion message
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
