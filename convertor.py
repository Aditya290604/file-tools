import os
from PIL import Image
Image.MAX_IMAGE_PIXELS = None  # Allow very large images

import img2pdf
from pdf2image import convert_from_path
from docx2pdf import convert as docx2pdf_convert
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def convert_image(input_path, output_path, output_format):
    img = Image.open(input_path)
    if output_format.lower() == 'jpg':
        output_format = 'JPEG'
    img.save(output_path, output_format.upper())

def image_to_pdf(input_path, output_path):
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(input_path))

def pdf_to_images(input_path, output_folder, output_format):
    images = convert_from_path(input_path)
    paths = []
    for i, img in enumerate(images):
        out_path = os.path.join(output_folder, f"page_{i+1}.{output_format}")
        img.save(out_path, output_format.upper())
        paths.append(out_path)
    return paths

def docx_to_pdf(input_path, output_path):
    docx2pdf_convert(input_path, output_path)

def excel_to_csv(input_path, output_path):
    df = pd.read_excel(input_path)
    df.to_csv(output_path, index=False)

def pptx_to_pdf(input_path, output_path):
    import comtypes.client
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    ppt = powerpoint.Presentations.Open(input_path)
    ppt.SaveAs(output_path, 32)
    ppt.Close()
    powerpoint.Quit()

SUPPORTED_FORMATS = ['pdf', 'jpg', 'jpeg', 'png', 'docx', 'xlsx', 'pptx']

class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal File Converter")
        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_format = tk.StringVar(value=SUPPORTED_FORMATS[0])

        # Input file path
        file_frame = tk.Frame(root)
        file_frame.pack(pady=5, fill='x')
        tk.Label(file_frame, text="📂 Select Input File:").pack(side='left', padx=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=40).pack(side='left', padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_file).pack(side='left', padx=5)

        # Output file path
        out_frame = tk.Frame(root)
        out_frame.pack(pady=5, fill='x')
        tk.Label(out_frame, text="💾 Converted File Will Be Saved As:").pack(side='left', padx=5)
        tk.Entry(out_frame, textvariable=self.output_path, width=40).pack(side='left', padx=5)

        # Format selection
        convert_frame = tk.LabelFrame(root, text="")
        convert_frame.pack(pady=5, fill='x')
        tk.Label(convert_frame, text="Convert To:").pack(side='left', padx=5)
        format_box = ttk.Combobox(convert_frame, textvariable=self.output_format, values=SUPPORTED_FORMATS, width=8, state='readonly')
        format_box.pack(side='left', padx=5)
        format_box.bind("<<ComboboxSelected>>", lambda e: self.update_output_path())

        # Convert button
        tk.Button(root, text="Convert", command=self.run_conversion, bg='#4CAF50', fg='white', height=2).pack(pady=10, fill='x')

    def browse_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.file_path.set(path)
            self.update_output_path()

    def update_output_path(self):
        input_path = self.file_path.get()
        output_format = self.output_format.get()
        manual_output = self.output_path.get()
        if input_path and output_format and (not manual_output or manual_output.endswith(f".{self.output_format.get()}")):
            base = os.path.splitext(os.path.basename(input_path))[0]
            out_path = os.path.join(os.path.dirname(input_path), f"{base}.{output_format}")
            self.output_path.set(out_path)

    def run_conversion(self):
        input_path = self.file_path.get()
        output_path = self.output_path.get()
        output_format = self.output_format.get()

        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid file.")
            return

        input_format = os.path.splitext(input_path)[1][1:].lower()
        if input_format not in SUPPORTED_FORMATS:
            messagebox.showerror("Unsupported Format", f"The file type '.{input_format}' is not supported.")
            return

        try:
            if input_format in ['jpg', 'jpeg', 'png'] and output_format in ['jpg', 'jpeg', 'png']:
                convert_image(input_path, output_path, output_format)
            elif input_format in ['jpg', 'jpeg', 'png'] and output_format == 'pdf':
                image_to_pdf(input_path, output_path)
            elif input_format == 'pdf' and output_format in ['jpg', 'jpeg', 'png']:
                pdf_to_images(input_path, os.path.dirname(output_path), output_format)
                messagebox.showinfo("Success", f"PDF converted to images in folder:\n{os.path.dirname(output_path)}")
                return
            elif input_format == 'docx' and output_format == 'pdf':
                docx_to_pdf(input_path, output_path)
            elif input_format == 'xlsx' and output_format == 'csv':
                excel_to_csv(input_path, output_path)
            elif input_format == 'pptx' and output_format == 'pdf':
                pptx_to_pdf(input_path, output_path)
            else:
                messagebox.showinfo("Info", "Conversion not supported for this format pair.")
                return

            messagebox.showinfo("Success", f"File converted and saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk()
    ConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
