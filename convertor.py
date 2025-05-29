import os
from PIL import Image
import img2pdf
from pdf2image import convert_from_path
from docx2pdf import convert as docx2pdf_convert
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Conversion helper functions
def convert_image(input_path, output_path, output_format):
    img = Image.open(input_path)
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
    ppt.SaveAs(output_path, 32)  # 32 for PDF
    ppt.Close()
    powerpoint.Quit()

SUPPORTED_FORMATS = ['pdf', 'jpg', 'jpeg', 'png', 'docx', 'xlsx', 'pptx']

class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Universal File Converter")
        self.file_path = tk.StringVar()
        self.input_format = tk.StringVar(value=SUPPORTED_FORMATS[0])
        self.output_format = tk.StringVar(value=SUPPORTED_FORMATS[0])

        # File selection section
        file_frame = tk.Frame(root)
        file_frame.pack(pady=10, fill='x')
        tk.Label(file_frame, text="File Path:").pack(side='left', padx=5)
        tk.Entry(file_frame, textvariable=self.file_path, width=40).pack(side='left', padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_file).pack(side='left', padx=5)

        # Conversion section
        convert_frame = tk.LabelFrame(root, text="Convert")
        convert_frame.pack(pady=10, fill='x')
        tk.Label(convert_frame, text="Current Type:").pack(side='left', padx=5)
        ttk.Combobox(convert_frame, textvariable=self.input_format, values=SUPPORTED_FORMATS, width=8).pack(side='left', padx=5)
        tk.Label(convert_frame, text="Convert To:").pack(side='left', padx=5)
        ttk.Combobox(convert_frame, textvariable=self.output_format, values=SUPPORTED_FORMATS, width=8).pack(side='left', padx=5)

        # Convert button
        tk.Button(root, text="Convert", command=self.run_conversion, bg='#4CAF50', fg='white', height=2).pack(pady=15, fill='x')

    def browse_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.file_path.set(path)
            ext = os.path.splitext(path)[1][1:].lower()
            if ext in SUPPORTED_FORMATS:
                self.input_format.set(ext)

    def run_conversion(self):
        input_path = self.file_path.get()
        input_format = self.input_format.get()
        output_format = self.output_format.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid file.")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=f'.{output_format}', filetypes=[(output_format.upper(), f'*.{output_format}')])
        if not output_path:
            return
        try:
            if input_format in ['jpg', 'jpeg', 'png'] and output_format in ['jpg', 'jpeg', 'png']:
                convert_image(input_path, output_path, output_format)
            elif input_format in ['jpg', 'jpeg', 'png'] and output_format == 'pdf':
                image_to_pdf(input_path, output_path)
            elif input_format == 'pdf' and output_format in ['jpg', 'jpeg', 'png']:
                pdf_to_images(input_path, os.path.dirname(output_path), output_format)
            elif input_format == 'docx' and output_format == 'pdf':
                docx_to_pdf(input_path, output_path)
            elif input_format == 'xlsx' and output_format == 'csv':
                excel_to_csv(input_path, output_path)
            elif input_format == 'pptx' and output_format == 'pdf':
                pptx_to_pdf(input_path, output_path)
            else:
                messagebox.showinfo("Info", "Conversion not supported yet.")
                return
            messagebox.showinfo("Success", f"File converted and saved to {output_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk()
    ConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
