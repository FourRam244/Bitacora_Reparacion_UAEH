import tkinter as tk
from PIL import ImageTk, Image
import subprocess


def alta():
    subprocess.run(["python", "Bitacora1.py"])

def consulta():
    # Coloca aquí el código para la consulta
    subprocess.run(["python", "Bitacora2.py"])
    pass

def finalizar():
    # Coloca aquí el código para finalizar la aplicación
    subprocess.run(["python", "Bitacora.py"])

# Crear la ventana principal
root = tk.Tk()
root.title("Menu Bitacora")

# Cargar las imágenes y redimensionarlas
image1 = Image.open("./img/logo1.png")
image1 = image1.resize((200, 100))
image1 = ImageTk.PhotoImage(image1)

image2 = Image.open("./img/logo2.png")
image2 = image2.resize((200, 100))
image2 = ImageTk.PhotoImage(image2)

# Crear los widgets para las imágenes
imagen1_label = tk.Label(root, image=image1)
imagen1_label.pack(side=tk.LEFT, padx=10)

imagen2_label = tk.Label(root, image=image2)
imagen2_label.pack(side=tk.RIGHT, padx=10)

# Función para crear botones
def crear_boton(texto, comando):
    return tk.Button(root, text=texto, command=comando, width=20)

# Crear los botones
boton_alta = crear_boton("Alta", alta)
boton_alta.pack(pady=10)
boton_consulta = crear_boton("Consulta", consulta)
boton_consulta.pack(pady=10)
boton_finalizar = crear_boton("Finalizar", finalizar)
boton_finalizar.pack(pady=10)



# Ejecutar la aplicación
root.mainloop()
