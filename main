import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText
from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph
from pptx import Presentation
import os
import threading

class PDFShrinker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Document Text Extractor")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")
        
        # Style configuration
        style = ttk.Style()
        style.configure("TButton", padding=10, font=('Helvetica', 10))
        style.configure("TLabel", font=('Helvetica', 11))
        
        self.total_original_size = 0
        self.total_new_size = 0
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Input Files", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Input files selection
        self.input_entry = ttk.Entry(file_frame)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_btn = ttk.Button(
            file_frame,
            text="Browse Files",
            command=self.select_input_files
        )
        browse_btn.pack(side=tk.RIGHT)
        
        # Output folder frame
        output_frame = ttk.LabelFrame(main_frame, text="Output Location", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.output_entry = ttk.Entry(output_frame)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        output_btn = ttk.Button(
            output_frame,
            text="Select Folder",
            command=self.select_output_folder
        )
        output_btn.pack(side=tk.RIGHT)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = ScrolledText(
            progress_frame,
            height=10,
            wrap=tk.WORD,
            font=('Courier', 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # Convert button
        self.convert_btn = ttk.Button(
            main_frame,
            text="Convert Documents",
            command=self.start_conversion,
            style="TButton"
        )
        self.convert_btn.pack(pady=10)
        
    def select_input_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Supported Files", "*.pdf *.ppt *.pptx"), ("PDF Files", "*.pdf"), ("PowerPoint Files", "*.ppt *.pptx")]
        )
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, ", ".join(files))
        
    def select_output_folder(self):
        folder_path = filedialog.askdirectory()
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, folder_path)
        
    def log_message(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        
    def extract_text_from_pdf(self, input_file):
        pdf_reader = PdfReader(input_file)
        text_content = ""
        for page in pdf_reader.pages:
            text_content += page.extract_text() + "\n"
        return text_content
    
    def extract_text_from_pptx(self, input_file):
        presentation = Presentation(input_file)
        text_content = ""
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_content += shape.text + "\n"
        return text_content
    
    def create_pdf_from_text(self, text_content, output_file):
        doc = SimpleDocTemplate(
            output_file,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        styles = getSampleStyleSheet()
        style = ParagraphStyle(
            'CustomStyle',
            parent=styles['Normal'],
            fontSize=10,
            leading=14
        )
        
        story = []
        paragraphs = text_content.split('\n')
        for para in paragraphs:
            if para.strip():
                story.append(Paragraph(para, style))
        
        doc.build(story)
    
    def get_file_size(self, file_path):
        return os.path.getsize(file_path) / (1024 * 1024)  # Size in MB
    
    def convert_file(self, input_file, output_folder):
        try:
            base_name = os.path.basename(input_file)
            file_name, file_ext = os.path.splitext(base_name)
            output_file = os.path.join(output_folder, f"{file_name}_text_only.pdf")
            
            original_size = self.get_file_size(input_file)
            self.total_original_size += original_size
            
            self.log_message(f"Processing: {base_name} (Original Size: {original_size:.2f} MB)")
            
            if file_ext.lower() in ['.ppt', '.pptx']:
                text_content = self.extract_text_from_pptx(input_file)
            elif file_ext.lower() == '.pdf':
                text_content = self.extract_text_from_pdf(input_file)
            else:
                self.log_message(f"Unsupported file type: {file_ext}")
                return
            
            self.create_pdf_from_text(text_content, output_file)
            
            new_size = self.get_file_size(output_file)
            self.total_new_size += new_size
            
            self.log_message(f"Successfully converted: {base_name} (New Size: {new_size:.2f} MB)")
            
        except Exception as e:
            self.log_message(f"Error processing {base_name}: {str(e)}")
    
    def start_conversion(self):
        self.clear_log()  # Clear the log at the start of each conversion
        input_files = self.input_entry.get().split(", ")
        output_folder = self.output_entry.get()
        
        if not input_files or not output_folder:
            self.log_message("Please select both input files and output folder.")
            return
        
        self.convert_btn.config(state='disabled')
        self.progress_var.set(0)
        
        def conversion_thread():
            self.total_original_size = 0  # Reset total sizes
            self.total_new_size = 0
            
            total_files = len(input_files)
            for i, input_file in enumerate(input_files):
                self.convert_file(input_file, output_folder)
                progress = ((i + 1) / total_files) * 100
                self.progress_var.set(progress)
            
            self.log_message(f"\nTotal Original Size: {self.total_original_size:.2f} MB")
            self.log_message(f"Total New Size: {self.total_new_size:.2f} MB")
            self.log_message(f"Size Reduction: {self.total_original_size - self.total_new_size:.2f} MB")
            
            self.convert_btn.config(state='normal')
        
        threading.Thread(target=conversion_thread, daemon=True).start()
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = PDFShrinker()
    app.run()
