import os
import PyPDF2
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

def slice_pdf(input_pdf, start_page, end_page, output_pdf):
    """
    Slices a PDF from start_page to end_page (inclusive) and saves it as output_pdf.
    """
    try:
        with open(input_pdf, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            writer = PyPDF2.PdfWriter()
            # Adjust for 0-based indexing in PyPDF2
            for page_num in range(start_page - 1, end_page):
                writer.add_page(reader.pages[page_num])
            os.makedirs(os.path.dirname(output_pdf), exist_ok=True)
            with open(output_pdf, 'wb') as output_file:
                writer.write(output_file)
        return True, f"PDF sliced successfully and saved as:\n{output_pdf}"
    except Exception as e:
        return False, f"An error occurred: {e}"

class PDFSliceGUI:
    """
    GUI for slicing a PDF file by page range.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Slicer")
        self.pdf_path = tb.StringVar()
        self.start_page = tb.StringVar()
        self.end_page = tb.StringVar()

        # --- PDF file selection ---
        file_frame = tb.Frame(root)
        file_frame.pack(pady=5, fill='x')
        tb.Label(file_frame, text="📄 Select PDF File:").pack(side='left', padx=5)
        tb.Entry(file_frame, textvariable=self.pdf_path, width=40).pack(side='left', padx=5)
        tb.Button(file_frame, text="Browse", command=self.browse_pdf, bootstyle="secondary").pack(side='left', padx=5)

        # --- Start page input ---
        start_frame = tb.Frame(root)
        start_frame.pack(pady=5, fill='x')
        tb.Label(start_frame, text="Start Page:").pack(side='left', padx=5)
        tb.Entry(start_frame, textvariable=self.start_page, width=10).pack(side='left', padx=5)

        # --- End page input ---
        end_frame = tb.Frame(root)
        end_frame.pack(pady=5, fill='x')
        tb.Label(end_frame, text="End Page:").pack(side='left', padx=5)
        tb.Entry(end_frame, textvariable=self.end_page, width=10).pack(side='left', padx=5)

        # --- Slice button ---
        tb.Button(root, text="Slice PDF", command=self.run_slice, bootstyle="success-outline", width=30).pack(pady=10)

        # --- Info label ---
        info_text = (
            "Select a PDF and enter the start and end page numbers (1-based).\n"
            "The sliced PDF will be saved in the same folder as the original."
        )
        tb.Label(root, text=info_text, justify='left', foreground='gray').pack(pady=5)

    def browse_pdf(self):
        """
        Open file dialog to select a PDF file.
        """
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path.set(path)

    def run_slice(self):
        """
        Perform the PDF slicing operation based on user input.
        """
        pdf_path = self.pdf_path.get()
        start = self.start_page.get()
        end = self.end_page.get()

        # --- Input validation ---
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file.")
            return
        try:
            start_page = int(start)
            end_page = int(end)
        except ValueError:
            messagebox.showerror("Error", "Start and end page must be integers.")
            return
        if start_page < 1 or end_page < start_page:
            messagebox.showerror("Error", "Invalid page range.")
            return

        # --- Output path ---
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_folder = os.path.dirname(pdf_path)
        out_path = os.path.join(out_folder, f"{base}_pages_{start_page}_to_{end_page}.pdf")

        # --- Slice PDF ---
        success, msg = slice_pdf(pdf_path, start_page, end_page, out_path)
        if success:
            messagebox.showinfo("Success", msg)
        else:
            messagebox.showerror("Error", msg)

def main():
    """
    Launch the PDF Slicer GUI.
    """
    app = tb.Window(themename="superhero")
    PDFSliceGUI(app)
    app.mainloop()

if __name__ == "__main__":
    main()





