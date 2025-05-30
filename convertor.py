# -*- coding: utf-8 -*-
"""
Created on Thu May 29 12:14:46 2025

@author: Aditya
"""

# --- Imports ---
import os
from PIL import Image  # For image processing
Image.MAX_IMAGE_PIXELS = None  # Disable DecompressionBombWarning for large images

import img2pdf  # For converting images to PDF
from pdf2image import convert_from_path  # For converting PDF to images
from docx2pdf import convert as docx2pdf_convert  # For DOCX to PDF conversion
import pandas as pd  # For Excel to CSV conversion
import ttkbootstrap as tb  # For modern Tkinter GUI
from ttkbootstrap.constants import *  # Bootstrap constants for styling

from tkinter import filedialog, messagebox  # For file dialogs and popups
from pdf2docx import Converter as PDF2DocxConverter  # For PDF to DOCX conversion
import tabula  # For PDF to XLSX conversion (extracts tables)

# --- Conversion Functions ---

def convert_image(input_path, output_path, output_format):
    """
    Convert an image from one format to another (JPG, PNG, JPEG).
    """
    img = Image.open(input_path)
    if output_format.lower() == 'jpg':
        output_format = 'JPEG'
    img.save(output_path, output_format.upper())

def image_to_pdf(input_path, output_path):
    """
    Convert a single image to a PDF file.
    """
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(input_path))

def pdf_to_images(input_path, output_folder, output_format):
    """
    Convert each page of a PDF to separate image files.
    Returns a list of output image paths.
    """
    images = convert_from_path(input_path)
    paths = []
    for i, img in enumerate(images):
        out_path = os.path.join(output_folder, f"page_{i+1}.{output_format}")
        img.save(out_path, output_format.upper())
        paths.append(out_path)
    return paths

def docx_to_pdf(input_path, output_path):
    """
    Convert a DOCX file to PDF.
    """
    docx2pdf_convert(input_path, output_path)

def excel_to_csv(input_path, output_path):
    """
    Convert an Excel (XLSX) file to CSV.
    """
    df = pd.read_excel(input_path)
    df.to_csv(output_path, index=False)

def pptx_to_pdf(input_path, output_path):
    """
    Convert a PowerPoint (PPTX) file to PDF.
    Only works on Windows with Microsoft Office installed.
    """
    import comtypes.client
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    ppt = powerpoint.Presentations.Open(input_path)
    ppt.SaveAs(output_path, 32)  # 32 = PDF format
    ppt.Close()
    powerpoint.Quit()

def pdf_to_docx(input_path, output_path):
    """
    Convert a PDF file to DOCX.
    """
    cv = PDF2DocxConverter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def pdf_to_xlsx(input_path, output_path):
    """
    Convert a PDF file to XLSX by extracting tables.
    """
    dfs = tabula.read_pdf(input_path, pages='all', multiple_tables=True)
    if dfs:
        with pd.ExcelWriter(output_path) as writer:
            for idx, df in enumerate(dfs):
                sheet_name = f"Sheet{idx+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        raise Exception("No tables found in PDF.")

# --- Supported Formats and Conversion Map ---

SUPPORTED_FORMATS = ['pdf', 'jpg', 'jpeg', 'png', 'docx', 'xlsx', 'pptx']

# Map input formats to possible output formats for the dropdown
CONVERSION_MAP = {
    'jpg':    ['png', 'jpeg', 'pdf'],
    'jpeg':   ['jpg', 'png', 'pdf'],
    'png':    ['jpg', 'jpeg', 'pdf'],
    'pdf':    ['jpg', 'jpeg', 'png', 'docx', 'xlsx'],
    'docx':   ['pdf'],
    'xlsx':   ['csv'],
    'pptx':   ['pdf'],
}

# --- GUI Class ---

class ConverterGUI:
    """
    Main GUI class for the Universal File Converter.
    Handles user interaction and conversion logic.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Universal File Converter")
        self.file_path = tb.StringVar()
        self.output_path = tb.StringVar()
        self.output_format = tb.StringVar()

        # --- File path selection ---
        file_frame = tb.Frame(root)
        file_frame.pack(pady=5, fill='x')
        tb.Label(file_frame, text="📂 Select Input File:").pack(side='left', padx=5)
        tb.Entry(file_frame, textvariable=self.file_path, width=40).pack(side='left', padx=5)
        tb.Button(file_frame, text="Browse", command=self.browse_file, bootstyle="secondary").pack(side='left', padx=5)

        # --- Output path display ---
        out_frame = tb.Frame(root)
        out_frame.pack(pady=5, fill='x')
        tb.Label(out_frame, text="💾 Converted File Will Be Saved As:").pack(side='left', padx=5)
        tb.Entry(out_frame, textvariable=self.output_path, width=40).pack(side='left', padx=5)

        # --- Output format selection ---
        convert_frame = tb.LabelFrame(root, text="")
        convert_frame.pack(pady=5, fill='x')
        tb.Label(convert_frame, text="Convert To:").pack(side='left', padx=5)
        self.format_box = tb.Combobox(
            convert_frame,
            textvariable=self.output_format,
            values=[],
            width=8,
            state='readonly'
        )
        self.format_box.pack(side='left', padx=5)
        self.format_box.bind("<<ComboboxSelected>>", lambda e: self.on_format_change())

        # --- Convert button ---
        tb.Button(
            root,
            text="Convert",
            command=self.run_conversion,
            bootstyle="success-outline",
            width=40
        ).pack(pady=10)

        # --- Open file/folder buttons ---
        btn_frame = tb.Frame(root)
        btn_frame.pack(pady=2)
        tb.Button(
            btn_frame,
            text="Open Converted File",
            command=self.open_converted_file,
            bootstyle="info"
        ).pack(side='left', padx=10)
        tb.Button(
            btn_frame,
            text="Open Output Folder",
            command=self.open_output_folder,
            bootstyle="secondary"
        ).pack(side='left', padx=10)

        # --- Info label with supported conversions ---
        info_text = (
            "Supported Conversions:\n"
            "- Images ↔ Images (JPG, PNG, JPEG)\n"
            "- Images → PDF\n"
            "- PDF → Images\n"
            "- PDF → DOCX\n"
            "- PDF → XLSX\n"
            "- DOCX → PDF\n"
            "- XLSX → CSV\n"
            "- PPTX → PDF (Windows only)"
        )
        tb.Label(root, text=info_text, justify='left', foreground='gray').pack(pady=5)

    def browse_file(self):
        """
        Open file dialog for user to select input file.
        Updates format dropdown and output path.
        """
        path = filedialog.askopenfilename()
        if path:
            path = os.path.normpath(path)  # Normalize to OS default separator
            self.file_path.set(path)
            self.update_format_box()
            self.update_output_path()

    def update_format_box(self):
        """
        Update the output format dropdown based on selected input file type.
        """
        input_path = self.file_path.get()
        if not input_path:
            self.format_box['values'] = []
            self.output_format.set('')
            return
        input_format = os.path.splitext(input_path)[1][1:].lower()
        valid_outputs = CONVERSION_MAP.get(input_format, [])
        self.format_box['values'] = valid_outputs
        if valid_outputs:
            self.output_format.set(valid_outputs[0])
            self.update_output_path()  # Ensure output path updates when format box is set
        else:
            self.output_format.set('')

    def on_format_change(self):
        """
        Called when the output format dropdown selection changes.
        Updates the output path to match the new extension.
        """
        self.update_output_path()

    def update_output_path(self):
        """
        Suggest or update the output file path based on input file and selected output format.
        If the user has edited the output path, only update the extension.
        Ensures consistent use of OS path separators.
        """
        input_path = self.file_path.get()
        if input_path:
            input_path = os.path.normpath(input_path)  # Normalize input path
            self.file_path.set(input_path)
        output_format = self.output_format.get()
        current_output = self.output_path.get()
        if input_path and output_format:
            base = os.path.splitext(os.path.basename(input_path))[0]
            default_dir = os.path.dirname(input_path)
            default_name = f"{base}.{output_format}"
            default_path = os.path.join(default_dir, default_name)
            default_path = os.path.normpath(default_path)

            if not current_output or os.path.normpath(current_output) == default_path:
                # Suggest default
                self.output_path.set(default_path)
            else:
                # User has edited output path; update only the extension
                user_dir = os.path.dirname(current_output)
                user_base = os.path.splitext(os.path.basename(current_output))[0]
                new_path = os.path.join(user_dir, f"{user_base}.{output_format}")
                new_path = os.path.normpath(new_path)
                self.output_path.set(new_path)

    def run_conversion(self):
        """
        Run the appropriate conversion based on user selections.
        Handles errors and shows success/info dialogs.
        """
        input_path = self.file_path.get()
        output_path = self.output_path.get()
        output_format = self.output_format.get()

        # --- Input validation ---
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid file.")
            return

        input_format = os.path.splitext(input_path)[1][1:].lower()
        if input_format not in SUPPORTED_FORMATS:
            messagebox.showerror("Unsupported Format", f"The file type '.{input_format}' is not supported.")
            return

        try:
            # --- Image to image conversion ---
            if input_format in ['jpg', 'jpeg', 'png'] and output_format in ['jpg', 'jpeg', 'png']:
                convert_image(input_path, output_path, output_format)
            # --- Image to PDF ---
            elif input_format in ['jpg', 'jpeg', 'png'] and output_format == 'pdf':
                image_to_pdf(input_path, output_path)
            # --- PDF to images ---
            elif input_format == 'pdf' and output_format in ['jpg', 'jpeg', 'png']:
                pdf_to_images(input_path, os.path.dirname(output_path), output_format)
                messagebox.showinfo("Success", f"PDF converted to images in folder:\n{os.path.dirname(output_path)}")
                return
            # --- PDF to DOCX ---
            elif input_format == 'pdf' and output_format == 'docx':
                pdf_to_docx(input_path, output_path)
            # --- PDF to XLSX ---
            elif input_format == 'pdf' and output_format == 'xlsx':
                pdf_to_xlsx(input_path, output_path)
            # --- DOCX to PDF ---
            elif input_format == 'docx' and output_format == 'pdf':
                docx_to_pdf(input_path, output_path)
            # --- XLSX to CSV ---
            elif input_format == 'xlsx' and output_format == 'csv':
                excel_to_csv(input_path, output_path)
            # --- PPTX to PDF ---
            elif input_format == 'pptx' and output_format == 'pdf':
                pptx_to_pdf(input_path, output_path)
            else:
                messagebox.showinfo("Info", "Conversion not supported for this format pair.")
                return

            messagebox.showinfo("Success", f"File converted and saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_converted_file(self):
        """
        Open the converted file with the default application.
        """
        path = self.output_path.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Converted file does not exist.")
            return
        try:
            if os.name == 'nt':
                os.startfile(path)
            elif os.name == 'posix':
                import subprocess
                subprocess.Popen(['xdg-open', path])
            else:
                messagebox.showinfo("Info", f"Please open the file manually:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    def open_output_folder(self):
        """
        Open the folder containing the converted file.
        """
        path = self.output_path.get()
        folder = os.path.dirname(path)
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Error", "Output folder does not exist.")
            return
        try:
            if os.name == 'nt':
                os.startfile(folder)
            elif os.name == 'posix':
                import subprocess
                subprocess.Popen(['xdg-open', folder])
            else:
                messagebox.showinfo("Info", f"Please open the folder manually:\n{folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

# --- Main Application Entry Point ---

def main():
    """
    Launch the Universal File Converter GUI.
    """
    app = tb.Window(themename="superhero")  # Try: flatly, lumen, morph, journal, "superhero", "cosmo", "solar", "cyborg"
    ConverterGUI(app)
    app.mainloop()

if __name__ == "__main__":
    main()