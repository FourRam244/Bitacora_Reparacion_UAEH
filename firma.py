import tkinter as tk
from PIL import Image, ImageDraw

class Firma:
    def __init__(self, root):
        self.root = root
        self.root.title("Firma con el Mouse")

        self.canvas = tk.Canvas(self.root, width=400, height=200, bg="white")
        self.canvas.pack()

        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(fill=tk.X)

        self.clear_button = tk.Button(self.button_frame, text="Limpiar", command=self.limpiar_lienzo)
        self.clear_button.pack(side=tk.LEFT)

        self.save_button = tk.Button(self.button_frame, text="Guardar", command=self.guardar_firma)
        self.save_button.pack(side=tk.LEFT)

        self.canvas.bind("<B1-Motion>", self.pintar)
        self.image = Image.new("RGB", (400, 200), "white")
        self.pintar = ImageDraw.Draw(self.image)

    def pintar(self, event):
        x1, y1 = (event.x - 1), (event.y - 1)
        x2, y2 = (event.x + 1), (event.y + 1)
        self.canvas.create_oval(x1, y1, x2, y2, fill="black", width=2)
        self.pintar.line([x1, y1, x2, y2], fill="black", width=2)

    def limpiar_lienzo(self):
        self.canvas.delete("all")
        self.image = Image.new("RGB", (400, 200), "white")
        self.draw = ImageDraw.Draw(self.image)

    def guardar_firma(self):
        self.image.save("signature.png")
        print("Firma guardada como signature.png")

def open_signature_app():
    signature_window = tk.Toplevel(root)
    app = Firma(signature_window)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Ventana Principal")

    open_button = tk.Button(root, text="Abrir ventana de firma", command=open_signature_app)
    open_button.pack(pady=20)

    root.mainloop()
