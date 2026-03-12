import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
from PIL import Image, ImageTk
import pandas as pd
from qr_generator import QRCodeGenerator

BG_COLOR = "#0F111A"
CARD_COLOR = "#1B1E2B"
ACCENT_COLOR = "#3B82F6"
SUCCESS_COLOR = "#10B981"
ERROR_COLOR = "#EF4444"
TEXT_PRIMARY = "#F8FAFC"
TEXT_SECONDARY = "#94A3B8" 


class SaturnApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Saturn QR | Generator for Saturn")
        self.root.geometry("900x750")
        self.root.configure(bg=BG_COLOR)
        self.root.resizable(False, False)

        # Variables
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.path.join(os.getcwd(), "qr"))
        self.status_msg = tk.StringVar(value="System Ready")

        # Generator
        self.qr_generator = QRCodeGenerator(encryption_key=None)

        # UI Initialization
        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Modern Progress Bar
        style.configure(
            "Sleek.Horizontal.TProgressbar",
            thickness=6,
            background=ACCENT_COLOR,
            troughcolor=CARD_COLOR,
            bordercolor=CARD_COLOR,
        )

        # Scrollbar
        style.configure(
            "Sleek.Vertical.TScrollbar",
            gripcount=0,
            background=CARD_COLOR,
            troughcolor=BG_COLOR,
            bordercolor=BG_COLOR,
        )

    def _build_ui(self):
        sidebar = tk.Frame(self.root, bg=BG_COLOR, width=280)
        sidebar.pack(side="left", fill="y", padx=(20, 0), pady=40)
        sidebar.pack_propagate(False)

        # Branding
        tk.Label(
            sidebar,
            text="SATURN",
            font=("Inter", 24, "bold"),
            fg=ACCENT_COLOR,
            bg=BG_COLOR,
        ).pack(anchor="w")
        tk.Label(
            sidebar,
            text="QR GENERATOR",
            font=("Inter", 10),
            fg=TEXT_SECONDARY,
            bg=BG_COLOR,
        ).pack(anchor="w", pady=(0, 40))

        # Requirements Section
        req_frame = tk.Frame(sidebar, bg=BG_COLOR)
        req_frame.pack(fill="x", pady=20)

        tk.Label(
            req_frame,
            text="REQUIRED DATA",
            font=("Inter", 9, "bold"),
            fg=TEXT_SECONDARY,
            bg=BG_COLOR,
        ).pack(anchor="w", pady=10)

        def add_req(text):
            f = tk.Frame(req_frame, bg=BG_COLOR)
            f.pack(fill="x", pady=4)
            tk.Label(
                f, text="•", fg=ACCENT_COLOR, bg=BG_COLOR, font=("Inter", 12)
            ).pack(side="left")
            tk.Label(
                f, text=text, fg=TEXT_PRIMARY, bg=BG_COLOR, font=("Inter", 10)
            ).pack(side="left", padx=10)

        add_req("Student ID")
        add_req("Section")
        add_req("Full Name")

        # 2. Main Workspace
        workspace = tk.Frame(self.root, bg=BG_COLOR)
        workspace.pack(side="right", fill="both", expand=True, padx=40, pady=40)

        # Card: File Selection
        self._create_card(
            workspace,
            "Source Data",
            "Select the Excel file containing student records.",
            self.excel_path,
            self.browse_excel,
            "SELECT FILE",
        )

        self._create_card(
            workspace,
            "Output Location",
            "Choose where to save the generated QR codes.",
            self.output_dir,
            self.browse_output,
            "CHOOSE FOLDER",
        )

        process_frame = tk.Frame(workspace, bg=BG_COLOR)
        process_frame.pack(fill="x", side="bottom", pady=(20, 0))

        self.progress = ttk.Progressbar(
            process_frame, style="Sleek.Horizontal.TProgressbar", mode="determinate"
        )
        self.progress.pack(fill="x", pady=(0, 15))

        # Main Action
        self.run_btn = tk.Button(
            process_frame,
            text="GENERATE QR CODES",
            command=self.start_generation,
            bg=ACCENT_COLOR,
            fg=TEXT_PRIMARY,
            font=("Inter", 11, "bold"),
            relief="flat",
            padx=30,
            pady=12,
            activebackground="#2563EB",
            cursor="hand2",
        )
        self.run_btn.pack(side="right")

        tk.Label(
            process_frame,
            textvariable=self.status_msg,
            font=("Inter", 10),
            fg=TEXT_SECONDARY,
            bg=BG_COLOR,
        ).pack(side="left")

    def _create_card(self, parent, title, desc, variable, command, btn_text):
        card = tk.Frame(parent, bg=CARD_COLOR, padx=25, pady=25)
        card.pack(fill="x", pady=(0, 20))

        tk.Label(
            card,
            text=title.upper(),
            font=("Inter", 10, "bold"),
            fg=ACCENT_COLOR,
            bg=CARD_COLOR,
        ).pack(anchor="w")
        tk.Label(
            card, text=desc, font=("Inter", 10), fg=TEXT_SECONDARY, bg=CARD_COLOR
        ).pack(anchor="w", pady=(2, 12))

        input_row = tk.Frame(card, bg=CARD_COLOR)
        input_row.pack(fill="x")

        entry = tk.Entry(
            input_row,
            textvariable=variable,
            font=("Inter", 10),
            bg=BG_COLOR,
            fg=TEXT_PRIMARY,
            relief="flat",
            insertbackground=TEXT_PRIMARY,
            highlightthickness=1,
            highlightbackground="#2D3142",
        )
        entry.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))

        tk.Button(
            input_row,
            text=btn_text,
            command=command,
            bg="#2D3142",
            fg=TEXT_PRIMARY,
            font=("Inter", 8, "bold"),
            relief="flat",
            padx=15,
            pady=8,
            activebackground="#3D4259",
            cursor="hand2",
        ).pack(side="right")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_path.set(path)
            self.qr_generator.set_excel_path(path)
            self.status_msg.set("Excel file loaded")

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)
            self.qr_generator.set_output_path(path)

    def start_generation(self):
        if not self.excel_path.get():
            messagebox.showwarning(
                "Incomplete", "Please select a source Excel file first."
            )
            return

        self.run_btn.config(state="disabled", text="PROCESSING...")
        self.status_msg.set("Initialising...")
        threading.Thread(target=self._run_engine, daemon=True).start()

    def _run_engine(self):
        try:
            df = self.qr_generator.read_excel()
            total = len(df)
            self.progress["maximum"] = total

            success_count = 0
            for index, row in df.iterrows():
                success, _ = self.qr_generator.generate_qr_code(row)
                if success:
                    success_count += 1

                self.root.after(0, lambda v=index + 1: self._update_progress(v, total))

            messagebox.showinfo(
                "Complete", f"Successfully generated {success_count} QR codes."
            )
            self.status_msg.set("Ready")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_msg.set("Generation error")
        finally:
            self.root.after(
                0, lambda: self.run_btn.config(state="normal", text="GENERATE QR CODES")
            )

    def _update_progress(self, val, total):
        self.progress["value"] = val
        self.status_msg.set(f"Generating {val} of {total}...")


if __name__ == "__main__":
    root = tk.Tk()
    app = SaturnApp(root)
    root.mainloop()
