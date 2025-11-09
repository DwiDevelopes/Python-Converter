import os
import zipfile
import tarfile
import rarfile
import py7zr
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
from threading import Thread
from datetime import datetime
import webbrowser
import subprocess
import patoolib
import img2pdf
from fpdf import FPDF
import pytesseract
from pdf2image import convert_from_path

class FileExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Archive Manager - WinRAR Alternative")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 10))
        self.style.configure('TButton', font=('Segoe UI', 10))
        self.style.configure('TNotebook.Tab', font=('Segoe UI', 10))
        self.style.configure('Treeview', font=('Segoe UI', 10))
        self.style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))
        
        # Create main notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.create_extract_tab()
        self.create_compress_tab()
        self.create_image_converter_tab()
        self.create_advanced_tab()
        self.create_about_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize variables
        self.current_archive = ""
        self.current_files = []
        self.compression_level = tk.IntVar(value=6)
        self.compression_method = tk.StringVar(value="ZIP")
        self.password_var = tk.StringVar()
        self.split_size_var = tk.StringVar(value="0")
        self.image_format_var = tk.StringVar(value="JPEG")
        self.image_quality_var = tk.IntVar(value=85)
        
    def create_extract_tab(self):
        """Create the extraction tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Extract Files")
        
        # Top frame for archive selection
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(top_frame, text="Archive File:").pack(side=tk.LEFT)
        self.archive_path = tk.StringVar()
        self.archive_entry = ttk.Entry(top_frame, textvariable=self.archive_path, width=50)
        self.archive_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        browse_btn = ttk.Button(top_frame, text="Browse...", command=self.browse_archive)
        browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Middle frame for archive contents
        mid_frame = ttk.Frame(tab)
        mid_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(mid_frame, columns=('Size', 'Modified'), selectmode='extended')
        self.tree.heading('#0', text='Name')
        self.tree.heading('Size', text='Size')
        self.tree.heading('Modified', text='Modified')
        self.tree.column('Size', width=100, anchor='e')
        self.tree.column('Modified', width=120, anchor='e')
        
        vsb = ttk.Scrollbar(mid_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        mid_frame.grid_rowconfigure(0, weight=1)
        mid_frame.grid_columnconfigure(0, weight=1)
        
        # Bottom frame for extraction options
        bottom_frame = ttk.Frame(tab)
        bottom_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(bottom_frame, text="Extract to:").pack(side=tk.LEFT)
        self.extract_path = tk.StringVar()
        self.extract_entry = ttk.Entry(bottom_frame, textvariable=self.extract_path, width=50)
        self.extract_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        extract_browse_btn = ttk.Button(bottom_frame, text="Browse...", command=self.browse_extract_path)
        extract_browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Password frame
        pass_frame = ttk.Frame(bottom_frame)
        pass_frame.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(pass_frame, text="Password:").pack(side=tk.LEFT)
        self.password_entry = ttk.Entry(pass_frame, textvariable=self.password_var, show="*", width=15)
        self.password_entry.pack(side=tk.LEFT, padx=5)
        
        # Extract buttons
        extract_btn_frame = ttk.Frame(tab)
        extract_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        extract_btn = ttk.Button(extract_btn_frame, text="Extract Files", command=self.extract_files)
        extract_btn.pack(side=tk.LEFT, padx=5)
        
        extract_here_btn = ttk.Button(extract_btn_frame, text="Extract Here", command=lambda: self.extract_files(here=True))
        extract_here_btn.pack(side=tk.LEFT, padx=5)
        
        test_btn = ttk.Button(extract_btn_frame, text="Test Archive", command=self.test_archive)
        test_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(tab, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, padx=5, pady=5)
        
    def create_compress_tab(self):
        """Create the compression tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Create Archive")
        
        # Top frame for file selection
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(top_frame, text="Files to Compress:").pack(side=tk.LEFT)
        self.compress_path = tk.StringVar()
        self.compress_entry = ttk.Entry(top_frame, textvariable=self.compress_path, width=50)
        self.compress_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        compress_browse_btn = ttk.Button(top_frame, text="Browse...", command=self.browse_compress_files)
        compress_browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Middle frame for archive options
        mid_frame = ttk.LabelFrame(tab, text="Archive Options")
        mid_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Archive format
        format_frame = ttk.Frame(mid_frame)
        format_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(format_frame, text="Archive Format:").pack(side=tk.LEFT)
        formats = ["ZIP", "7Z", "TAR", "TAR.GZ", "TAR.BZ2", "TAR.XZ", "RAR"]
        format_menu = ttk.OptionMenu(format_frame, self.compression_method, *formats)
        format_menu.pack(side=tk.LEFT, padx=5)
        
        # Compression level
        level_frame = ttk.Frame(mid_frame)
        level_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(level_frame, text="Compression Level:").pack(side=tk.LEFT)
        ttk.Radiobutton(level_frame, text="Store", variable=self.compression_level, value=0).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(level_frame, text="Fast", variable=self.compression_level, value=3).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(level_frame, text="Normal", variable=self.compression_level, value=6).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(level_frame, text="Maximum", variable=self.compression_level, value=9).pack(side=tk.LEFT, padx=5)
        
        # Password protection
        pass_frame = ttk.Frame(mid_frame)
        pass_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(pass_frame, text="Password:").pack(side=tk.LEFT)
        self.compress_pass_entry = ttk.Entry(pass_frame, show="*", width=20)
        self.compress_pass_entry.pack(side=tk.LEFT, padx=5)
        
        # Split archive
        split_frame = ttk.Frame(mid_frame)
        split_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(split_frame, text="Split to Volumes (MB):").pack(side=tk.LEFT)
        ttk.Entry(split_frame, textvariable=self.split_size_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(split_frame, text="(0 for no split)").pack(side=tk.LEFT)
        
        # Archive name
        name_frame = ttk.Frame(tab)
        name_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(name_frame, text="Archive Name:").pack(side=tk.LEFT)
        self.archive_name = tk.StringVar()
        ttk.Entry(name_frame, textvariable=self.archive_name, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Compression button
        compress_btn_frame = ttk.Frame(tab)
        compress_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        compress_btn = ttk.Button(compress_btn_frame, text="Create Archive", command=self.create_archive)
        compress_btn.pack(side=tk.LEFT, padx=5)
        
    def create_image_converter_tab(self):
        """Create the image conversion tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Image Tools")
        
        # Top frame for image selection
        top_frame = ttk.Frame(tab)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(top_frame, text="Image File(s):").pack(side=tk.LEFT)
        self.image_path = tk.StringVar()
        self.image_entry = ttk.Entry(top_frame, textvariable=self.image_path, width=50)
        self.image_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        image_browse_btn = ttk.Button(top_frame, text="Browse...", command=self.browse_images)
        image_browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Middle frame for conversion options
        mid_frame = ttk.LabelFrame(tab, text="Conversion Options")
        mid_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Output format
        format_frame = ttk.Frame(mid_frame)
        format_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(format_frame, text="Output Format:").pack(side=tk.LEFT)
        formats = ["JPEG", "PNG", "BMP", "GIF", "TIFF", "WEBP", "PDF"]
        format_menu = ttk.OptionMenu(format_frame, self.image_format_var, *formats)
        format_menu.pack(side=tk.LEFT, padx=5)
        
        # Quality
        quality_frame = ttk.Frame(mid_frame)
        quality_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(quality_frame, text="Quality (1-100):").pack(side=tk.LEFT)
        ttk.Scale(quality_frame, from_=1, to=100, variable=self.image_quality_var, orient=tk.HORIZONTAL).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Label(quality_frame, textvariable=self.image_quality_var).pack(side=tk.LEFT, padx=5)
        
        # Output directory
        out_frame = ttk.Frame(tab)
        out_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(out_frame, text="Output Directory:").pack(side=tk.LEFT)
        self.image_out_path = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.image_out_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        out_browse_btn = ttk.Button(out_frame, text="Browse...", command=self.browse_image_output)
        out_browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Conversion buttons
        convert_btn_frame = ttk.Frame(tab)
        convert_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        convert_btn = ttk.Button(convert_btn_frame, text="Convert Images", command=self.convert_images)
        convert_btn.pack(side=tk.LEFT, padx=5)
        
        pdf_btn = ttk.Button(convert_btn_frame, text="Create PDF", command=self.create_pdf_from_images)
        pdf_btn.pack(side=tk.LEFT, padx=5)
        
        ocr_btn = ttk.Button(convert_btn_frame, text="Extract Text (OCR)", command=self.extract_text_from_images)
        ocr_btn.pack(side=tk.LEFT, padx=5)
        
        # Preview area
        preview_frame = ttk.LabelFrame(tab, text="Image Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.image_preview = ttk.Label(preview_frame)
        self.image_preview.pack(fill=tk.BOTH, expand=True)
        
    def create_advanced_tab(self):
        """Create the advanced tools tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Advanced Tools")
        
        # Repair archive
        repair_frame = ttk.LabelFrame(tab, text="Repair Archive")
        repair_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(repair_frame, text="Archive to Repair:").pack(side=tk.LEFT)
        self.repair_path = tk.StringVar()
        ttk.Entry(repair_frame, textvariable=self.repair_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        repair_browse_btn = ttk.Button(repair_frame, text="Browse...", command=self.browse_repair_file)
        repair_browse_btn.pack(side=tk.LEFT, padx=5)
        
        repair_btn = ttk.Button(repair_frame, text="Repair Archive", command=self.repair_archive)
        repair_btn.pack(side=tk.LEFT, padx=5)
        
        # Merge volumes
        merge_frame = ttk.LabelFrame(tab, text="Merge Split Archives")
        merge_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(merge_frame, text="First Volume:").pack(side=tk.LEFT)
        self.merge_path = tk.StringVar()
        ttk.Entry(merge_frame, textvariable=self.merge_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        merge_browse_btn = ttk.Button(merge_frame, text="Browse...", command=self.browse_merge_file)
        merge_browse_btn.pack(side=tk.LEFT, padx=5)
        
        merge_btn = ttk.Button(merge_frame, text="Merge Volumes", command=self.merge_volumes)
        merge_btn.pack(side=tk.LEFT, padx=5)
        
        # Batch processing
        batch_frame = ttk.LabelFrame(tab, text="Batch Processing")
        batch_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Batch extract
        batch_extract_frame = ttk.Frame(batch_frame)
        batch_extract_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(batch_extract_frame, text="Source Directory:").pack(side=tk.LEFT)
        self.batch_source = tk.StringVar()
        ttk.Entry(batch_extract_frame, textvariable=self.batch_source, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        batch_source_btn = ttk.Button(batch_extract_frame, text="Browse...", command=self.browse_batch_source)
        batch_source_btn.pack(side=tk.LEFT, padx=5)
        
        # Batch output
        batch_out_frame = ttk.Frame(batch_frame)
        batch_out_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Label(batch_out_frame, text="Output Directory:").pack(side=tk.LEFT)
        self.batch_output = tk.StringVar()
        ttk.Entry(batch_out_frame, textvariable=self.batch_output, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        batch_out_btn = ttk.Button(batch_out_frame, text="Browse...", command=self.browse_batch_output)
        batch_out_btn.pack(side=tk.LEFT, padx=5)
        
        # Batch buttons
        batch_btn_frame = ttk.Frame(batch_frame)
        batch_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        batch_extract_btn = ttk.Button(batch_btn_frame, text="Batch Extract", command=self.batch_extract)
        batch_extract_btn.pack(side=tk.LEFT, padx=5)
        
        batch_compress_btn = ttk.Button(batch_btn_frame, text="Batch Compress", command=self.batch_compress)
        batch_compress_btn.pack(side=tk.LEFT, padx=5)
        
        batch_convert_btn = ttk.Button(batch_btn_frame, text="Batch Convert Images", command=self.batch_convert_images)
        batch_convert_btn.pack(side=tk.LEFT, padx=5)
        
        # Log area
        log_frame = ttk.LabelFrame(tab, text="Operation Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def create_about_tab(self):
        """Create the about tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="About")
        
        # Logo and info
        info_frame = ttk.Frame(tab)
        info_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Placeholder for logo
        logo_label = ttk.Label(info_frame, text="Advanced Archive Manager", font=('Segoe UI', 16, 'bold'))
        logo_label.pack(pady=10)
        
        # Version info
        version_label = ttk.Label(info_frame, text="Version 1.0", font=('Segoe UI', 10))
        version_label.pack(pady=5)
        
        # Description
        desc_frame = ttk.Frame(tab)
        desc_frame.pack(fill=tk.X, padx=20, pady=10)
        
        desc_text = """
        Advanced Archive Manager is a comprehensive file compression and extraction tool 
        with support for multiple archive formats and advanced features including:
        
        - Support for ZIP, RAR, 7Z, TAR, and more
        - Password protection and encryption
        - Split archive creation
        - Image conversion and PDF creation
        - Batch processing
        - Archive repair and testing
        
        This application is developed in Python as an alternative to WinRAR with 
        additional functionality and a modern interface.
        """
        desc_label = ttk.Label(desc_frame, text=desc_text, justify=tk.LEFT)
        desc_label.pack(anchor=tk.W)
        
        # Credits
        credits_frame = ttk.Frame(tab)
        credits_frame.pack(fill=tk.X, padx=20, pady=10)
        
        credits_text = """
        Credits:
        - Python 3
        - tkinter for GUI
        - Pillow for image processing
        - py7zr, rarfile, patoolib for archive handling
        - pytesseract for OCR
        """
        credits_label = ttk.Label(credits_frame, text=credits_text, justify=tk.LEFT)
        credits_label.pack(anchor=tk.W)
        
        # Buttons
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill=tk.X, padx=20, pady=10)
        
        website_btn = ttk.Button(btn_frame, text="Visit Website", command=self.open_website)
        website_btn.pack(side=tk.LEFT, padx=5)
        
        docs_btn = ttk.Button(btn_frame, text="Documentation", command=self.open_docs)
        docs_btn.pack(side=tk.LEFT, padx=5)
        
        license_btn = ttk.Button(btn_frame, text="License", command=self.show_license)
        license_btn.pack(side=tk.LEFT, padx=5)
        
    # File browsing methods
    def browse_archive(self):
        filetypes = (
            ("Archive Files", "*.zip *.rar *.7z *.tar *.tar.gz *.tar.bz2 *.tar.xz"),
            ("All Files", "*.*")
        )
        filename = filedialog.askopenfilename(title="Select Archive File", filetypes=filetypes)
        if filename:
            self.archive_path.set(filename)
            self.load_archive_contents(filename)
    
    def browse_extract_path(self):
        path = filedialog.askdirectory(title="Select Extraction Directory")
        if path:
            self.extract_path.set(path)
    
    def browse_compress_files(self):
        files = filedialog.askopenfilenames(title="Select Files to Compress")
        if files:
            self.compress_path.set("; ".join(files))
    
    def browse_images(self):
        filetypes = (
            ("Image Files", "*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff *.webp"),
            ("All Files", "*.*")
        )
        files = filedialog.askopenfilenames(title="Select Image Files", filetypes=filetypes)
        if files:
            self.image_path.set("; ".join(files))
            self.show_image_preview(files[0])
    
    def browse_image_output(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.image_out_path.set(path)
    
    def browse_repair_file(self):
        filetypes = (
            ("Archive Files", "*.zip *.rar *.7z"),
            ("All Files", "*.*")
        )
        filename = filedialog.askopenfilename(title="Select Archive to Repair", filetypes=filetypes)
        if filename:
            self.repair_path.set(filename)
    
    def browse_merge_file(self):
        filetypes = (
            ("Archive Files", "*.part1.rar *.zip.001 *.7z.001"),
            ("All Files", "*.*")
        )
        filename = filedialog.askopenfilename(title="Select First Volume", filetypes=filetypes)
        if filename:
            self.merge_path.set(filename)
    
    def browse_batch_source(self):
        path = filedialog.askdirectory(title="Select Source Directory")
        if path:
            self.batch_source.set(path)
    
    def browse_batch_output(self):
        path = filedialog.askdirectory(title="Select Output Directory")
        if path:
            self.batch_output.set(path)
    
    # Archive handling methods
    def load_archive_contents(self, archive_path):
        """Load and display the contents of the selected archive"""
        self.tree.delete(*self.tree.get_children())
        self.current_archive = archive_path
        self.current_files = []
        
        try:
            if archive_path.lower().endswith('.zip'):
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        self.tree.insert('', 'end', text=file_info.filename, 
                                       values=(self.format_size(file_info.file_size), 
                                              self.format_time(file_info.date_time)))
                        self.current_files.append(file_info.filename)
            
            elif archive_path.lower().endswith('.rar'):
                with rarfile.RarFile(archive_path, 'r') as rar_ref:
                    for file_info in rar_ref.infolist():
                        self.tree.insert('', 'end', text=file_info.filename, 
                                       values=(self.format_size(file_info.file_size), 
                                              self.format_time(file_info.date_time)))
                        self.current_files.append(file_info.filename)
            
            elif archive_path.lower().endswith(('.7z', '.7zip')):
                with py7zr.SevenZipFile(archive_path, 'r') as sevenz_ref:
                    for file_info in sevenz_ref.list():
                        self.tree.insert('', 'end', text=file_info.filename, 
                                       values=(self.format_size(file_info.uncompressed), 
                                              self.format_time(file_info.creationtime)))
                        self.current_files.append(file_info.filename)
            
            elif archive_path.lower().endswith(('.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                mode = 'r'
                if archive_path.endswith('.gz'):
                    mode += ':gz'
                elif archive_path.endswith('.bz2'):
                    mode += ':bz2'
                elif archive_path.endswith('.xz'):
                    mode += ':xz'
                
                with tarfile.open(archive_path, mode) as tar_ref:
                    for file_info in tar_ref.getmembers():
                        self.tree.insert('', 'end', text=file_info.name, 
                                       values=(self.format_size(file_info.size), 
                                              self.format_time(file_info.mtime)))
                        self.current_files.append(file_info.name)
            
            self.status_var.set(f"Loaded {len(self.current_files)} files from archive")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open archive:\n{str(e)}")
            self.status_var.set("Error loading archive")
    
    def extract_files(self, here=False):
        """Extract files from the archive"""
        if not self.current_archive:
            messagebox.showwarning("Warning", "No archive selected")
            return
        
        password = self.password_var.get() or None
        extract_path = self.extract_path.get()
        
        if here:
            extract_path = os.path.dirname(self.current_archive)
            self.extract_path.set(extract_path)
        
        if not extract_path:
            messagebox.showwarning("Warning", "Please select an extraction directory")
            return
        
        try:
            selected_items = self.tree.selection()
            files_to_extract = None
            
            if selected_items:
                files_to_extract = [self.tree.item(item, 'text') for item in selected_items]
            
            def extraction_thread():
                try:
                    self.progress['value'] = 0
                    self.status_var.set("Extracting files...")
                    
                    if self.current_archive.lower().endswith('.zip'):
                        with zipfile.ZipFile(self.current_archive, 'r') as zip_ref:
                            if files_to_extract:
                                total = len(files_to_extract)
                                for i, filename in enumerate(files_to_extract):
                                    zip_ref.extract(filename, extract_path, pwd=password.encode() if password else None)
                                    self.progress['value'] = (i + 1) / total * 100
                                    self.root.update_idletasks()
                            else:
                                zip_ref.extractall(extract_path, pwd=password.encode() if password else None)
                                self.progress['value'] = 100
                    
                    elif self.current_archive.lower().endswith('.rar'):
                        with rarfile.RarFile(self.current_archive, 'r') as rar_ref:
                            if files_to_extract:
                                total = len(files_to_extract)
                                for i, filename in enumerate(files_to_extract):
                                    rar_ref.extract(filename, extract_path, pwd=password)
                                    self.progress['value'] = (i + 1) / total * 100
                                    self.root.update_idletasks()
                            else:
                                rar_ref.extractall(extract_path, pwd=password)
                                self.progress['value'] = 100
                    
                    elif self.current_archive.lower().endswith(('.7z', '.7zip')):
                        with py7zr.SevenZipFile(self.current_archive, 'r', password=password) as sevenz_ref:
                            if files_to_extract:
                                all_files = sevenz_ref.getnames()
                                filtered = [f for f in files_to_extract if f in all_files]
                                sevenz_ref.extract(targets=filtered, path=extract_path)
                            else:
                                sevenz_ref.extractall(path=extract_path)
                            self.progress['value'] = 100
                    
                    elif self.current_archive.lower().endswith(('.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                        mode = 'r'
                        if self.current_archive.endswith('.gz'):
                            mode += ':gz'
                        elif self.current_archive.endswith('.bz2'):
                            mode += ':bz2'
                        elif self.current_archive.endswith('.xz'):
                            mode += ':xz'
                        
                        with tarfile.open(self.current_archive, mode) as tar_ref:
                            if files_to_extract:
                                members = [m for m in tar_ref.getmembers() if m.name in files_to_extract]
                                total = len(members)
                                for i, member in enumerate(members):
                                    tar_ref.extract(member, extract_path)
                                    self.progress['value'] = (i + 1) / total * 100
                                    self.root.update_idletasks()
                            else:
                                tar_ref.extractall(extract_path)
                                self.progress['value'] = 100
                    
                    self.status_var.set(f"Extraction complete to {extract_path}")
                    messagebox.showinfo("Success", "Files extracted successfully")
                    self.log_operation(f"Extracted files from {os.path.basename(self.current_archive)} to {extract_path}")
                
                except Exception as e:
                    self.status_var.set("Extraction failed")
                    messagebox.showerror("Error", f"Failed to extract files:\n{str(e)}")
                    self.log_operation(f"Error extracting {os.path.basename(self.current_archive)}: {str(e)}")
                
                finally:
                    self.progress['value'] = 0
            
            Thread(target=extraction_thread, daemon=True).start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract files:\n{str(e)}")
            self.status_var.set("Extraction failed")
    
    def test_archive(self):
        """Test the integrity of the archive"""
        if not self.current_archive:
            messagebox.showwarning("Warning", "No archive selected")
            return
        
        password = self.password_var.get() or None
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Testing archive...")
            
            if self.current_archive.lower().endswith('.zip'):
                with zipfile.ZipFile(self.current_archive, 'r') as zip_ref:
                    bad_file = zip_ref.testzip()
                    if bad_file:
                        messagebox.showerror("Error", f"Archive is corrupt. Bad file: {bad_file}")
                        self.status_var.set("Archive test failed")
                    else:
                        messagebox.showinfo("Success", "Archive is OK")
                        self.status_var.set("Archive test passed")
            
            elif self.current_archive.lower().endswith('.rar'):
                with rarfile.RarFile(self.current_archive, 'r') as rar_ref:
                    rar_ref.testrar(pwd=password)
                    messagebox.showinfo("Success", "Archive is OK")
                    self.status_var.set("Archive test passed")
            
            elif self.current_archive.lower().endswith(('.7z', '.7zip')):
                with py7zr.SevenZipFile(self.current_archive, 'r', password=password) as sevenz_ref:
                    if sevenz_ref.test():
                        messagebox.showinfo("Success", "Archive is OK")
                        self.status_var.set("Archive test passed")
                    else:
                        messagebox.showerror("Error", "Archive is corrupt")
                        self.status_var.set("Archive test failed")
            
            elif self.current_archive.lower().endswith(('.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                mode = 'r'
                if self.current_archive.endswith('.gz'):
                    mode += ':gz'
                elif self.current_archive.endswith('.bz2'):
                    mode += ':bz2'
                elif self.current_archive.endswith('.xz'):
                    mode += ':xz'
                
                with tarfile.open(self.current_archive, mode) as tar_ref:
                    # Tar files don't have a test method, so we'll just try to read the first file
                    first_member = tar_ref.next()
                    if first_member:
                        messagebox.showinfo("Success", "Archive appears to be OK")
                        self.status_var.set("Archive test passed")
                    else:
                        messagebox.showwarning("Warning", "Archive is empty")
                        self.status_var.set("Archive is empty")
            
            self.progress['value'] = 100
            self.log_operation(f"Tested archive {os.path.basename(self.current_archive)}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to test archive:\n{str(e)}")
            self.status_var.set("Archive test failed")
            self.log_operation(f"Error testing {os.path.basename(self.current_archive)}: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    def create_archive(self):
        """Create a new archive"""
        files = self.compress_path.get().split("; ")
        if not files or not files[0]:
            messagebox.showwarning("Warning", "No files selected for compression")
            return
        
        archive_name = self.archive_name.get().strip()
        if not archive_name:
            messagebox.showwarning("Warning", "Please enter an archive name")
            return
        
        # Add extension if not present
        ext = self.compression_method.get().lower()
        if ext == 'tar.gz':
            ext = 'tar.gz'
        elif ext == 'tar.bz2':
            ext = 'tar.bz2'
        elif ext == 'tar.xz':
            ext = 'tar.xz'
        
        if not archive_name.lower().endswith(f'.{ext}'):
            archive_name += f'.{ext}'
        
        output_path = filedialog.asksaveasfilename(
            title="Save Archive As",
            initialfile=archive_name,
            defaultextension=f'.{ext}',
            filetypes=[(f"{ext.upper()} Archive", f"*.{ext}")]
        )
        
        if not output_path:
            return  # User canceled
        
        password = self.compress_pass_entry.get() or None
        level = self.compression_level.get()
        split_size = int(self.split_size_var.get()) * 1024 * 1024  # Convert MB to bytes
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Creating archive...")
            
            def compression_thread():
                try:
                    if output_path.lower().endswith('.zip'):
                        with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=level) as zip_ref:
                            total = len(files)
                            for i, file in enumerate(files):
                                arcname = os.path.basename(file)
                                zip_ref.write(file, arcname, compresslevel=level)
                                if password:
                                    # ZIP password protection is weak, but we'll add it
                                    zip_ref.setpassword(password.encode())
                                self.progress['value'] = (i + 1) / total * 100
                                self.root.update_idletasks()
                    
                    elif output_path.lower().endswith('.7z'):
                        filters = [{'id': py7zr.FILTER_DEFLATE, 'level': level}]
                        with py7zr.SevenZipFile(output_path, 'w', password=password, filters=filters) as sevenz_ref:
                            total = len(files)
                            for i, file in enumerate(files):
                                sevenz_ref.write(file, os.path.basename(file))
                                self.progress['value'] = (i + 1) / total * 100
                                self.root.update_idletasks()
                    
                    elif output_path.lower().endswith(('.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                        mode = 'w'
                        if output_path.endswith('.gz'):
                            mode += ':gz'
                        elif output_path.endswith('.bz2'):
                            mode += ':bz2'
                        elif output_path.endswith('.xz'):
                            mode += ':xz'
                        
                        with tarfile.open(output_path, mode) as tar_ref:
                            total = len(files)
                            for i, file in enumerate(files):
                                tar_ref.add(file, arcname=os.path.basename(file))
                                self.progress['value'] = (i + 1) / total * 100
                                self.root.update_idletasks()
                    
                    elif output_path.lower().endswith('.rar'):
                        # RAR creation requires external tools
                        cmd = ['rar', 'a', '-r', f'-m{level}']
                        if password:
                            cmd.append(f'-p{password}')
                        if split_size > 0:
                            cmd.append(f'-v{split_size}b')
                        cmd.append(output_path)
                        cmd.extend(files)
                        
                        result = subprocess.run(cmd, capture_output=True, text=True)
                        if result.returncode != 0:
                            raise Exception(f"RAR error: {result.stderr}")
                        
                        self.progress['value'] = 100
                    
                    self.status_var.set(f"Archive created: {os.path.basename(output_path)}")
                    messagebox.showinfo("Success", "Archive created successfully")
                    self.log_operation(f"Created archive {os.path.basename(output_path)}")
                
                except Exception as e:
                    self.status_var.set("Archive creation failed")
                    messagebox.showerror("Error", f"Failed to create archive:\n{str(e)}")
                    self.log_operation(f"Error creating archive: {str(e)}")
                
                finally:
                    self.progress['value'] = 0
            
            Thread(target=compression_thread, daemon=True).start()
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create archive:\n{str(e)}")
            self.status_var.set("Archive creation failed")
    
    def repair_archive(self):
        """Attempt to repair a damaged archive"""
        archive_path = self.repair_path.get()
        if not archive_path:
            messagebox.showwarning("Warning", "No archive selected for repair")
            return
        
        try:
            self.progress['value'] = 50
            self.status_var.set("Attempting to repair archive...")
            
            if archive_path.lower().endswith('.zip'):
                # ZIP repair is limited - we can try to extract what we can
                temp_dir = os.path.join(os.path.dirname(archive_path), "recovered")
                os.makedirs(temp_dir, exist_ok=True)
                
                with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Create new archive
                new_path = os.path.splitext(archive_path)[0] + "_repaired.zip"
                with zipfile.ZipFile(new_path, 'w') as new_zip:
                    for root, _, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            new_zip.write(file_path, arcname)
                
                shutil.rmtree(temp_dir)
                self.progress['value'] = 100
                messagebox.showinfo("Success", f"Repaired archive saved as:\n{new_path}")
                self.status_var.set("Archive repair completed")
                self.log_operation(f"Repaired archive {os.path.basename(archive_path)}")
            
            elif archive_path.lower().endswith('.rar'):
                # Use RAR's built-in repair function
                cmd = ['rar', 'r', archive_path]
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode == 0:
                    self.progress['value'] = 100
                    messagebox.showinfo("Success", "RAR repair completed")
                    self.status_var.set("Archive repair completed")
                    self.log_operation(f"Repaired RAR archive {os.path.basename(archive_path)}")
                else:
                    raise Exception(f"RAR repair failed: {result.stderr}")
            
            else:
                messagebox.showwarning("Warning", "Repair not supported for this archive type")
                self.status_var.set("Repair not supported")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to repair archive:\n{str(e)}")
            self.status_var.set("Archive repair failed")
            self.log_operation(f"Error repairing {os.path.basename(archive_path)}: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    def merge_volumes(self):
        """Merge split archive volumes"""
        first_volume = self.merge_path.get()
        if not first_volume:
            messagebox.showwarning("Warning", "No volume selected")
            return
        
        try:
            self.progress['value'] = 10
            self.status_var.set("Merging volumes...")
            
            # Determine the output file name
            if '.part' in first_volume.lower() or '.rar' in first_volume.lower():
                # RAR volumes
                output_path = first_volume.split('.part')[0] + '.rar'
                cmd = ['rar', 'r', output_path]
            else:
                # Assume ZIP or 7Z volumes (001, 002, etc.)
                base = first_volume[:-3]
                output_path = base + '.7z' if first_volume.endswith('.7z.001') else base + '.zip'
                
                # Find all volumes
                volumes = []
                i = 1
                while True:
                    vol = f"{base}.{i:03d}"
                    if not os.path.exists(vol):
                        break
                    volumes.append(vol)
                    i += 1
                
                if not volumes:
                    raise Exception("No volumes found to merge")
                
                # Combine the files
                with open(output_path, 'wb') as outfile:
                    for vol in volumes:
                        with open(vol, 'rb') as infile:
                            shutil.copyfileobj(infile, outfile)
                        self.progress['value'] = 10 + (90 * (volumes.index(vol) + 1) / len(volumes))
                        self.root.update_idletasks()
            
            self.progress['value'] = 100
            messagebox.showinfo("Success", f"Merged volumes to:\n{output_path}")
            self.status_var.set("Volume merge completed")
            self.log_operation(f"Merged volumes to {os.path.basename(output_path)}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to merge volumes:\n{str(e)}")
            self.status_var.set("Volume merge failed")
            self.log_operation(f"Error merging volumes: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    # Image conversion methods
    def show_image_preview(self, image_path):
        """Show a preview of the selected image"""
        try:
            img = Image.open(image_path)
            img.thumbnail((300, 300))
            photo = ImageTk.PhotoImage(img)
            self.image_preview.config(image=photo)
            self.image_preview.image = photo  # Keep a reference
        except Exception as e:
            self.image_preview.config(image=None, text="Preview not available")
    
    def convert_images(self):
        """Convert images to the selected format"""
        image_paths = self.image_path.get().split("; ")
        if not image_paths or not image_paths[0]:
            messagebox.showwarning("Warning", "No images selected for conversion")
            return
        
        output_dir = self.image_out_path.get()
        if not output_dir:
            output_dir = os.path.dirname(image_paths[0])
            self.image_out_path.set(output_dir)
        
        output_format = self.image_format_var.get().lower()
        quality = self.image_quality_var.get()
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Converting images...")
            
            for i, img_path in enumerate(image_paths):
                try:
                    with Image.open(img_path) as img:
                        # Convert to RGB if saving as JPEG
                        if output_format in ['jpg', 'jpeg'] and img.mode != 'RGB':
                            img = img.convert('RGB')
                        
                        # Create output filename
                        base = os.path.splitext(os.path.basename(img_path))[0]
                        output_path = os.path.join(output_dir, f"{base}.{output_format}")
                        
                        # Handle different formats
                        if output_format in ['jpg', 'jpeg']:
                            img.save(output_path, 'JPEG', quality=quality)
                        elif output_format == 'png':
                            img.save(output_path, 'PNG', compress_level=9 - (quality // 12))
                        elif output_format == 'bmp':
                            img.save(output_path, 'BMP')
                        elif output_format == 'gif':
                            img.save(output_path, 'GIF')
                        elif output_format == 'tiff':
                            img.save(output_path, 'TIFF', compression='tiff_lzw')
                        elif output_format == 'webp':
                            img.save(output_path, 'WEBP', quality=quality)
                        
                        self.progress['value'] = (i + 1) / len(image_paths) * 100
                        self.root.update_idletasks()
                
                except Exception as e:
                    self.log_operation(f"Error converting {os.path.basename(img_path)}: {str(e)}")
            
            self.status_var.set(f"Converted {len(image_paths)} images to {output_format.upper()}")
            messagebox.showinfo("Success", "Image conversion completed")
            self.log_operation(f"Converted {len(image_paths)} images to {output_format.upper()}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert images:\n{str(e)}")
            self.status_var.set("Image conversion failed")
        
        finally:
            self.progress['value'] = 0
    
    def create_pdf_from_images(self):
        """Create a PDF from the selected images"""
        image_paths = self.image_path.get().split("; ")
        if not image_paths or not image_paths[0]:
            messagebox.showwarning("Warning", "No images selected for PDF creation")
            return
        
        output_dir = self.image_out_path.get()
        if not output_dir:
            output_dir = os.path.dirname(image_paths[0])
            self.image_out_path.set(output_dir)
        
        # Ask for PDF filename
        pdf_path = filedialog.asksaveasfilename(
            initialdir=output_dir,
            initialfile="images.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")]
        )
        
        if not pdf_path:
            return  # User canceled
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Creating PDF...")
            
            # Convert images to PDF
            with open(pdf_path, "wb") as f:
                f.write(img2pdf.convert(image_paths))
            
            self.progress['value'] = 100
            self.status_var.set(f"PDF created: {os.path.basename(pdf_path)}")
            messagebox.showinfo("Success", "PDF created successfully")
            self.log_operation(f"Created PDF from {len(image_paths)} images")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF:\n{str(e)}")
            self.status_var.set("PDF creation failed")
            self.log_operation(f"Error creating PDF: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    def extract_text_from_images(self):
        """Extract text from images using OCR"""
        image_paths = self.image_path.get().split("; ")
        if not image_paths or not image_paths[0]:
            messagebox.showwarning("Warning", "No images selected for text extraction")
            return
        
        output_dir = self.image_out_path.get()
        if not output_dir:
            output_dir = os.path.dirname(image_paths[0])
            self.image_out_path.set(output_dir)
        
        # Ask for text filename
        text_path = filedialog.asksaveasfilename(
            initialdir=output_dir,
            initialfile="extracted_text.txt",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )
        
        if not text_path:
            return  # User canceled
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Extracting text...")
            
            # Check if Tesseract is installed
            try:
                pytesseract.get_tesseract_version()
            except EnvironmentError:
                messagebox.showerror("Error", "Tesseract OCR is not installed or not in your PATH")
                return
            
            # Process each image
            with open(text_path, 'w', encoding='utf-8') as f:
                for i, img_path in enumerate(image_paths):
                    try:
                        text = pytesseract.image_to_string(Image.open(img_path))
                        f.write(f"=== Text from {os.path.basename(img_path)} ===\n")
                        f.write(text)
                        f.write("\n\n")
                        
                        self.progress['value'] = (i + 1) / len(image_paths) * 100
                        self.root.update_idletasks()
                    
                    except Exception as e:
                        self.log_operation(f"Error processing {os.path.basename(img_path)}: {str(e)}")
            
            self.progress['value'] = 100
            self.status_var.set(f"Text extracted to {os.path.basename(text_path)}")
            messagebox.showinfo("Success", "Text extraction completed")
            self.log_operation(f"Extracted text from {len(image_paths)} images")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text:\n{str(e)}")
            self.status_var.set("Text extraction failed")
            self.log_operation(f"Error extracting text: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    # Batch processing methods
    def batch_extract(self):
        """Batch extract archives from a directory"""
        source_dir = self.batch_source.get()
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showwarning("Warning", "Invalid source directory")
            return
        
        output_dir = self.batch_output.get()
        if not output_dir:
            output_dir = os.path.join(source_dir, "extracted")
            os.makedirs(output_dir, exist_ok=True)
            self.batch_output.set(output_dir)
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Starting batch extraction...")
            
            # Find all archive files
            archives = []
            for root, _, files in os.walk(source_dir):
                for file in files:
                    if file.lower().endswith(('.zip', '.rar', '.7z', '.tar', '.tar.gz', '.tar.bz2', '.tar.xz')):
                        archives.append(os.path.join(root, file))
            
            if not archives:
                messagebox.showinfo("Info", "No archive files found in the source directory")
                return
            
            # Process each archive
            for i, archive in enumerate(archives):
                try:
                    self.status_var.set(f"Extracting {i+1}/{len(archives)}: {os.path.basename(archive)}")
                    
                    # Create output subdirectory
                    base = os.path.splitext(os.path.basename(archive))[0]
                    if base.endswith('.tar'):
                        base = os.path.splitext(base)[0]
                    archive_out = os.path.join(output_dir, base)
                    os.makedirs(archive_out, exist_ok=True)
                    
                    # Extract the archive
                    patoolib.extract_archive(archive, outdir=archive_out)
                    
                    self.progress['value'] = (i + 1) / len(archives) * 100
                    self.root.update_idletasks()
                    self.log_operation(f"Extracted {os.path.basename(archive)} to {archive_out}")
                
                except Exception as e:
                    self.log_operation(f"Error extracting {os.path.basename(archive)}: {str(e)}")
            
            self.status_var.set(f"Batch extraction complete: {len(archives)} archives processed")
            messagebox.showinfo("Success", f"Batch extraction completed\n{len(archives)} archives processed")
            self.log_operation(f"Batch extraction completed - {len(archives)} archives processed")
        
        except Exception as e:
            messagebox.showerror("Error", f"Batch extraction failed:\n{str(e)}")
            self.status_var.set("Batch extraction failed")
            self.log_operation(f"Batch extraction error: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    def batch_compress(self):
        """Batch compress directories to archives"""
        source_dir = self.batch_source.get()
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showwarning("Warning", "Invalid source directory")
            return
        
        output_dir = self.batch_output.get()
        if not output_dir:
            output_dir = os.path.join(source_dir, "archives")
            os.makedirs(output_dir, exist_ok=True)
            self.batch_output.set(output_dir)
        
        format = self.compression_method.get().lower()
        level = self.compression_level.get()
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Starting batch compression...")
            
            # Find all directories (excluding the output directory)
            dirs = []
            for root, directories, _ in os.walk(source_dir):
                for directory in directories:
                    dir_path = os.path.join(root, directory)
                    if not dir_path.startswith(output_dir):
                        dirs.append(dir_path)
            
            if not dirs:
                messagebox.showinfo("Info", "No directories found to compress")
                return
            
            # Process each directory
            for i, dir_path in enumerate(dirs):
                try:
                    dir_name = os.path.basename(dir_path)
                    self.status_var.set(f"Compressing {i+1}/{len(dirs)}: {dir_name}")
                    
                    # Create archive path
                    archive_path = os.path.join(output_dir, f"{dir_name}.{format}")
                    
                    # Compress the directory
                    if format == 'zip':
                        with zipfile.ZipFile(archive_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=level) as zipf:
                            for root, _, files in os.walk(dir_path):
                                for file in files:
                                    file_path = os.path.join(root, file)
                                    arcname = os.path.relpath(file_path, dir_path)
                                    zipf.write(file_path, arcname)
                    
                    elif format == '7z':
                        with py7zr.SevenZipFile(archive_path, 'w', filters=[{'id': py7zr.FILTER_DEFLATE, 'level': level}]) as sevenzf:
                            sevenzf.writeall(dir_path, os.path.basename(dir_path))
                    
                    elif format in ['tar', 'tar.gz', 'tar.bz2', 'tar.xz']:
                        mode = 'w'
                        if format.endswith('.gz'):
                            mode += ':gz'
                        elif format.endswith('.bz2'):
                            mode += ':bz2'
                        elif format.endswith('.xz'):
                            mode += ':xz'
                        
                        with tarfile.open(archive_path, mode) as tarf:
                            tarf.add(dir_path, arcname=os.path.basename(dir_path))
                    
                    self.progress['value'] = (i + 1) / len(dirs) * 100
                    self.root.update_idletasks()
                    self.log_operation(f"Compressed {dir_name} to {os.path.basename(archive_path)}")
                
                except Exception as e:
                    self.log_operation(f"Error compressing {dir_name}: {str(e)}")
            
            self.status_var.set(f"Batch compression complete: {len(dirs)} directories processed")
            messagebox.showinfo("Success", f"Batch compression completed\n{len(dirs)} directories processed")
            self.log_operation(f"Batch compression completed - {len(dirs)} directories processed")
        
        except Exception as e:
            messagebox.showerror("Error", f"Batch compression failed:\n{str(e)}")
            self.status_var.set("Batch compression failed")
            self.log_operation(f"Batch compression error: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    def batch_convert_images(self):
        """Batch convert images in a directory"""
        source_dir = self.batch_source.get()
        if not source_dir or not os.path.isdir(source_dir):
            messagebox.showwarning("Warning", "Invalid source directory")
            return
        
        output_dir = self.batch_output.get()
        if not output_dir:
            output_dir = os.path.join(source_dir, "converted")
            os.makedirs(output_dir, exist_ok=True)
            self.batch_output.set(output_dir)
        
        format = self.image_format_var.get().lower()
        quality = self.image_quality_var.get()
        
        try:
            self.progress['value'] = 0
            self.status_var.set("Starting batch image conversion...")
            
            # Find all image files
            images = []
            for root, _, files in os.walk(source_dir):
                for file in files:
                    if file.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tif', '.tiff', '.webp')):
                        images.append(os.path.join(root, file))
            
            if not images:
                messagebox.showinfo("Info", "No image files found in the source directory")
                return
            
            # Process each image
            for i, img_path in enumerate(images):
                try:
                    img_name = os.path.basename(img_path)
                    self.status_var.set(f"Converting {i+1}/{len(images)}: {img_name}")
                    
                    # Create output filename
                    base = os.path.splitext(img_name)[0]
                    output_path = os.path.join(output_dir, f"{base}.{format}")
                    
                    # Convert the image
                    with Image.open(img_path) as img:
                        # Convert to RGB if saving as JPEG
                        if format in ['jpg', 'jpeg'] and img.mode != 'RGB':
                            img = img.convert('RGB')
                        
                        # Save in the new format
                        if format in ['jpg', 'jpeg']:
                            img.save(output_path, 'JPEG', quality=quality)
                        elif format == 'png':
                            img.save(output_path, 'PNG', compress_level=9 - (quality // 12))
                        elif format == 'bmp':
                            img.save(output_path, 'BMP')
                        elif format == 'gif':
                            img.save(output_path, 'GIF')
                        elif format == 'tiff':
                            img.save(output_path, 'TIFF', compression='tiff_lzw')
                        elif format == 'webp':
                            img.save(output_path, 'WEBP', quality=quality)
                    
                    self.progress['value'] = (i + 1) / len(images) * 100
                    self.root.update_idletasks()
                    self.log_operation(f"Converted {img_name} to {format.upper()}")
                
                except Exception as e:
                    self.log_operation(f"Error converting {img_name}: {str(e)}")
            
            self.status_var.set(f"Batch conversion complete: {len(images)} images processed")
            messagebox.showinfo("Success", f"Batch image conversion completed\n{len(images)} images processed")
            self.log_operation(f"Batch image conversion completed - {len(images)} images processed")
        
        except Exception as e:
            messagebox.showerror("Error", f"Batch conversion failed:\n{str(e)}")
            self.status_var.set("Batch conversion failed")
            self.log_operation(f"Batch conversion error: {str(e)}")
        
        finally:
            self.progress['value'] = 0
    
    # Utility methods
    def format_size(self, size):
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"
    
    def format_time(self, time_tuple):
        """Format time tuple to readable string"""
        if isinstance(time_tuple, (int, float)):
            return datetime.fromtimestamp(time_tuple).strftime('%Y-%m-%d %H:%M:%S')
        return datetime(*time_tuple[:6]).strftime('%Y-%m-%d %H:%M:%S')
    
    def log_operation(self, message):
        """Add a message to the operation log"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.update()
    
    def open_website(self):
        """Open the application website"""
        webbrowser.open("https://example.com/archive-manager")
    
    def open_docs(self):
        """Open the documentation"""
        webbrowser.open("https://example.com/archive-manager/docs")
    
    def show_license(self):
        """Show the license information"""
        license_text = """
        Advanced Archive Manager - License Agreement
        
        This software is provided 'as-is', without any express or implied warranty.
        The author will not be held liable for any damages arising from the use of this software.
        
        Permission is granted to use this software for any purpose, including commercial applications.
        Redistribution is permitted with attribution to the original author.
        """
        messagebox.showinfo("License Information", license_text)

if __name__ == "__main__":
    root = tk.Tk()
    app = FileExtractorApp(root)
    root.mainloop()