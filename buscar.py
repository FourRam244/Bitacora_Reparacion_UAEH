import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import fitz

class PDFViewerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Visor de PDF")

        self.main_frame = tk.Frame(self.master)
        self.main_frame.pack(pady=10)

        self.open_button = tk.Button(self.main_frame, text="Abrir PDF", command=self.open_pdf)
        self.open_button.pack()

    def open_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

        if not file_path:
            return

        pdf_window = tk.Toplevel(self.master)
        pdf_window.title("PDF Viewer")
        pdf_window.geometry("650x650")  # Cambia el tamaño de la ventana a 800x600 píxeles

        pdf_canvas = tk.Canvas(pdf_window)
        pdf_canvas.pack(fill=tk.BOTH, expand=True)

        pdf_document = fitz.open(file_path)

        for page_number in range(pdf_document.page_count):
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap()
            width, height = pix.width, pix.height

            img = Image.frombytes("RGB", [width, height], pix.samples)
            tk_img = ImageTk.PhotoImage(img)

            pdf_canvas.create_image(0, 0, anchor=tk.NW, image=tk_img)
            pdf_canvas.image = tk_img

def main():
    root = tk.Tk()
    app = PDFViewerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
