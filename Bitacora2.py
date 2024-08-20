# Importación de módulos para el manejo de imágenes
from PIL import Image, ImageTk  # Para trabajar con imágenes en Tkinter

# Importación del módulo tkinter para la creación de la interfaz gráfica
import tkinter as tk
from tkinter import ttk
from tkinter import Scrollbar  # Importación específica de la barra de desplazamiento
from tkinter import messagebox  # Para mostrar mensajes de cuadro de diálogo

from tkinter import simpledialog

from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
import win32print
import win32ui
from PIL import Image, ImageWin


# Importación del widget de calendario para Tkinter
from tkcalendar import DateEntry
#Importancion de pandas para manipulacion de info de excel

# Importación del módulo 'os' para operaciones del sistema




import json

class BitacoraMantenimiento:
    def __init__(self, root):   
        
        self.root = root
        self.root.title("Bitácora de Mantenimiento")
        
        
        # Bloquear la opción de maximizar la ventana principal
        self.root.resizable(False, False)
      
        # Crear un lienzo para colocar el frame con la barra de desplazamiento
        canvas = tk.Canvas(root)
        canvas.pack(side='left', fill='both', expand=True)
        
        # Crear un frame dentro del lienzo
        self.frame = tk.Frame(canvas)
        self.frame.pack(padx=10, pady=10)
        
        # Añadir la barra de desplazamiento vertical
        scrollbar = tk.Scrollbar(root, orient='vertical', command=canvas.yview)
        scrollbar.pack(side='right', fill='y')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configurar el lienzo para usar la barra de desplazamiento
        canvas.create_window((0, 0), window=self.frame, anchor='nw')
        
        # Vincular la barra de desplazamiento con el movimiento del ratón
        self.frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # Enlazar el desplazamiento del mouse al lienzo
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        # Agregar contenido al frame
        self.add_content()
        

    def add_content(self):
        
        #self.contador = tk.IntVar(value=2)
        # Intentar cargar el número guardado desde el archivo
        self.numero_guardado = self.cargar_numero()

        # Crear una etiqueta para mostrar el contador
        self.contador_label = tk.Label(root, text=f"{self.numero_guardado}")
        self.contador_label.place(x=850, y=10)  # Colocar la etiqueta en la esquina superior derecha
        
                  # Cargar las imágenes
        self.logo1_image = Image.open("./img/logo1.png")
        self.logo1_image = self.logo1_image.resize((180, 100), )  # Redimensionar la imagen
        self.logo1_photo = ImageTk.PhotoImage(self.logo1_image)
        
        self.logo2_image = Image.open("./img/logo2.png")
        self.logo2_image = self.logo2_image.resize((180, 100), )  # Redimensionar la imagen
        self.logo2_photo = ImageTk.PhotoImage(self.logo2_image)
        
        # Etiqueta para la primera imagen
        self.logo1_label = tk.Label(self.frame, image=self.logo1_photo)
        self.logo1_label.grid(row=0, column=0, padx=10, pady=10)
        
        # Etiqueta para la segunda imagen
        self.logo2_label = tk.Label(self.frame, image=self.logo2_photo)
        self.logo2_label.grid(row=0, column=3, padx=10, pady=10)
    
        # Título para la sección de Información del Equipo
        titulo_informacion_equipo = tk.Label(self.frame, text="Bitacora de Mantenimiento", font=("Arial", 14, "bold"), pady=10)
        titulo_informacion_equipo.grid(row=0, column=0, columnspan=4)
        titulo_informacion_equipo.config(justify="center")
        
        # Campos de entrada
        tk.Label(self.frame, text="Fecha de Recepción:").grid(row=1, column=0, sticky="e")
        self.fecha_recepcion_entry = DateEntry(self.frame, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.fecha_recepcion_entry.grid(row=1, column=1)
        
        tk.Label(self.frame, text="Fecha de Entrega:").grid(row=1, column=2, sticky="e")
        self.fecha_entrega_entry = DateEntry(self.frame, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.fecha_entrega_entry.grid(row=1, column=3)
        
        # Título 
        titulo_datos_equipo = tk.Label(self.frame, text="Datos del Responsable del Equipo", font=("Helvetica", 12, "bold"), pady=10)
        titulo_datos_equipo.grid(row=2, column=0, columnspan=4)
        titulo_datos_equipo.config(justify="center")
        
        tk.Label(self.frame, text="Nombre:").grid(row=3, column=0, sticky="e")
        self.nombre_responsable_entry = tk.Entry(self.frame)
        self.nombre_responsable_entry.grid(row=3, column=1)
        
        tk.Label(self.frame, text="Teléfono:").grid(row=4, column=0, sticky="e")
        self.telefono_responsable_entry = tk.Entry(self.frame)
        self.telefono_responsable_entry.grid(row=4, column=1)
        
        tk.Label(self.frame, text="Correo:").grid(row=3, column=2, sticky="e")
        self.correo_entry = tk.Entry(self.frame)
        self.correo_entry.grid(row=3, column=3)
        
        tk.Label(self.frame, text="Área:").grid(row=4, column=2, sticky="e")
        self.area_responsable_entry = tk.Entry(self.frame)
        self.area_responsable_entry.grid(row=4, column=3)
        
        # Título 
        titulo_descripcion_equipo = tk.Label(self.frame, text="Descripcion de Equipo", font=("Helvetica", 12, "bold"), pady=10)
        titulo_descripcion_equipo.grid(row=6, column=0, columnspan=4)
        titulo_descripcion_equipo.config(justify="center")
        # Crear el Label "Equipo:"
        tk.Label(self.frame, text="Equipo:").grid(row=7, column=0, sticky="e")
        
        # Definir opciones para el Combobox
        equipo_options = ["Autoclave", "Balanza digital", "Balanza Granataria", "Basculas", "Bocina", "Bomba de agua", "Bomba de vacío", 
                          "Campana", "Cámaras", "Cargador", "Centrifugadora", "Cortadora", "Compresora", "Computadora", "Control Remoto", "Cronometro", 
                          "Dispensador de agua", "Diseño 3D", "Eliminador", "Esmeril", "Espectrofotómetro", "Estufa", "Fuentes de poder", "Generador de funciones", 
                          "Horno", "Impresión 3D", "Impresora", "Incubadora", "Lámpara", "Licuadora", "Lavadoras", "Mano factura de piezas", "Manta de calentamiento", 
                          "Microscopio", "Microscopio Estereoscópico", "Monitor", "Motor", "Mufla", "Multímetro", "Multi-contactos", "No break", "Osciloscopio","Pantalla", 
                          "Parilla con agitación", "Parrilla de calentamiento", "Parrilla de calentamiento y agitación", "Pipetas", "Proyector", "Pulidora", "Purificador", 
                          "Radio", "Refrigerador", "Regulador de Voltaje", "Sensores", "Termómetro digital", "Taladro", "Torneo y fresado", "Tornos","Torniquete","UPS","Otros" ]
        
        # Crear el Combobox y posicionarlo a la derecha del Label
        self.equipo_combobox = ttk.Combobox(self.frame, values=equipo_options)
        self.equipo_combobox.grid(row=7, column=1, sticky="w")
        self.equipo_combobox.bind("<<ComboboxSelected>>", self.toggle_equipo)
        
        tk.Label(self.frame, text="Otros:").grid(row=7, column=2, sticky="e")
        self.otro_equipo_entry = tk.Entry(self.frame, state="disabled")
        self.otro_equipo_entry.grid(row=7, column=3)
        
        tk.Label(self.frame, text="Modelo:").grid(row=8, column=0, sticky="e")
        self.modelo_equipo_entry = tk.Entry(self.frame)
        self.modelo_equipo_entry.grid(row=8, column=1)
        
        tk.Label(self.frame, text="Marca:").grid(row=8, column=2, sticky="e")
        self.marca_equipo_entry = tk.Entry(self.frame)
        self.marca_equipo_entry.grid(row=8, column=3)
        
        tk.Label(self.frame, text="No. Serie").grid(row=10, column=0, sticky="e")
        self.no_serie_var = tk.StringVar(value="N/A")
        self.no_serie_entry = tk.Entry(self.frame, textvariable=self.no_serie_var)
        self.no_serie_entry.grid(row=10, column=1)
        
        tk.Label(self.frame, text="No. Inventario").grid(row=10, column=2, sticky="e")
        self.no_inventario_var = tk.StringVar(value="N/A")
        self.no_inventario_entry = tk.Entry(self.frame, textvariable=self.no_inventario_var)
        self.no_inventario_entry.grid(row=10, column=3)
        
        # Título 
        titulo_estado_equipo = tk.Label(self.frame, text="Estado del Equipo", font=("Helvetica", 12, "bold"), pady=10)
        titulo_estado_equipo.grid(row=11, column=0, columnspan=4)
        titulo_estado_equipo.config(justify="center")
        
        # Sección de Estado del Equipo
        tk.Label(self.frame, text="Estado del Equipo:").grid(row=12, column=0, sticky="e")
        self.estado_enciende = tk.BooleanVar()
        self.enciende_checkbutton = tk.Checkbutton(self.frame, text="Enciende", variable=self.estado_enciende)
        self.enciende_checkbutton.grid(row=12, column=1, sticky="w")
        
        self.estado_cable = tk.BooleanVar()
        self.cable_checkbutton = tk.Checkbutton(self.frame, text="Cable de Alimentación", variable=self.estado_cable)
        self.cable_checkbutton.grid(row=12, column=2, sticky="w")
        
        self.estado_componente = tk.BooleanVar()
        self.componente_checkbutton = tk.Checkbutton(self.frame, text="Falta algún componente", variable=self.estado_componente,command=self.toggle_falta)
        self.componente_checkbutton.grid(row=13, column=1, sticky="w")
        
        
        self.faltante_equipo_entry = tk.Entry(self.frame, state="disabled")
        self.faltante_equipo_entry.grid(row=13, column=3)
        
        
        self.estado_dano_botones = tk.BooleanVar()
        self.dano_botones_checkbutton = tk.Checkbutton(self.frame, text="Daño en botones/perillas", variable=self.estado_dano_botones)
        self.dano_botones_checkbutton.grid(row=13, column=2, sticky="w")
        
        self.estado_corrosion = tk.BooleanVar()
        self.corrosion_checkbutton = tk.Checkbutton(self.frame, text="Presenta corrosión/oxidación", variable=self.estado_corrosion)
        self.corrosion_checkbutton.grid(row=14, column=1, sticky="w")
        
        self.estado_dano_carcasa = tk.BooleanVar()
        self.dano_carcasa_checkbutton = tk.Checkbutton(self.frame, text="Daño en Carcasa", variable=self.estado_dano_carcasa)
        self.dano_carcasa_checkbutton.grid(row=14, column=2, sticky="w")
        
        # Título 
        titulo_estado_equipo = tk.Label(self.frame, text="Descripción Detallada del Equipo", font=("Helvetica", 12, "bold"), pady=10)
        titulo_estado_equipo.grid(row=15, column=0, columnspan=4)
        titulo_estado_equipo.config(justify="center")
        
        # Otros campos de información
        tk.Label(self.frame, text="Descripción Detallada del Equipo:").grid(row=16, column=0, sticky="ne")
        self.descripcion_detallada_entry = tk.Text(self.frame, height=5, width=80, wrap=tk.WORD)
        self.descripcion_detallada_entry.grid(row=16, column=1, columnspan=3, sticky="w")
        
        self.impr = tk.Button(self.frame, text="Imprimir Etiqueta", command=lambda: self.imprimir())
        self.impr.grid(row=37, column=1, columnspan=1, pady=10, sticky="w")
        
    
        self.cargar_datos()
            
    def toggle_equipo(self, event):
        selected_item = self.equipo_combobox.get()  # Obtener el elemento seleccionado
        if selected_item == "Otros":
            self.otro_equipo_entry.config(state="normal")  # Activar la caja de texto "otro_equipo"
        else:
            self.otro_equipo_entry.config(state="disabled")  # Desactivar la caja de texto "otro_equipo"
        
    def guardar_datos(self, event=None):
        # Obtener el equipo seleccionado en el combobox
        descripcion_equipo = self.equipo_combobox.get()
        
        # Obtener el estado de los checkboxes
        estado_enciende = self.estado_enciende.get()
        estado_cable = self.estado_cable.get()
        estado_componente = self.estado_componente.get()
        estado_dano_botones = self.estado_dano_botones.get()
        estado_corrosion = self.estado_corrosion.get()
        estado_dano_carcasa = self.estado_dano_carcasa.get()
     
        # Si la descripción del equipo es "Otros", usar el valor de la entrada de texto
        if descripcion_equipo == "Otros":
            descripcion_equipo = self.otro_equipo_entry.get()
        
        datos = {
            "fecha_recepcion": self.fecha_recepcion_entry.get(),
            "fecha_entrega": self.fecha_entrega_entry.get(),
            "nombre_responsable": self.nombre_responsable_entry.get(),
            "telefono_responsable": self.telefono_responsable_entry.get(),
            "area_responsable": self.area_responsable_entry.get(),
            "correo": self.correo_entry.get(),
            "descripcion_equipo": descripcion_equipo,
            "otro_equipo": self.otro_equipo_entry.get(),
            "falta_equipo": self.faltante_equipo_entry.get(),
            "modelo_equipo": self.modelo_equipo_entry.get(),
            "marca_equipo": self.marca_equipo_entry.get(),
            "no_serie": self.no_serie_entry.get(),
            "no_inventario": self.no_inventario_entry.get(),
            "descripcion_detallada": self.descripcion_detallada_entry.get("1.0", tk.END).strip(),
            "estado_enciende": estado_enciende,
            "estado_cable": estado_cable,
            "estado_componente": estado_componente,
            "estado_dano_botones": estado_dano_botones,
            "estado_corrosion": estado_corrosion,
            "estado_dano_carcasa": estado_dano_carcasa,
            # Obtener los tipos de mantenimiento seleccionados
        }
        # Generar el ticket
        numero = self.numero_guardado
        with open(f"./Archivos/Progresos/{numero}.json", "w") as f:
            json.dump(datos, f)
            messagebox.showinfo("Guardar Progreso", "Progreso Guardado Correctamente")
    
    def imprimir(self):
        
        # Dimensiones de la imagen (2x1 pulgadas)
        ancho = (144 * 2)*12
        alto = (144 * 1)*12
        
        # Crear una nueva imagen en blanco
        imagen = Image.new("RGB", (ancho, alto), "white")
        draw = ImageDraw.Draw(imagen)
        
        # Cargar los logotipos
        logo_izquierdo = Image.open("./img/l1.png").resize((int(ancho * 0.25), int(alto * 0.21)))
        logo_derecho = Image.open("./img/logo2.jpg").resize((int(ancho * 0.25), int(alto * 0.21)))
        
        # Pegar los logotipos en la parte superior
        imagen.paste(logo_izquierdo, (int(ancho * 0.05), int(alto * 0.05)))
        imagen.paste(logo_derecho, (ancho - int(ancho * 0.25), int(alto * 0.05)))
        # Agregar texto
        texto1 = "Preventivo ( )   Correctivo ( )"
        fecha=self.fecha_recepcion_entry.get()
        texto2 = f"Fecha: {fecha} "
        fuente = ImageFont.truetype("arial.ttf", int(alto * 0.1))

        # Calcular el tamaño del texto1
        texto1_rectangulo = draw.textbbox((0, 0), texto1, font=fuente)
        texto1_ancho = texto1_rectangulo[2] - texto1_rectangulo[0]

        # Calcular el tamaño del texto2
        texto2_rectangulo = draw.textbbox((0, 0), texto2, font=fuente)
        texto2_ancho = texto2_rectangulo[2] - texto2_rectangulo[0]

        # Calcular las coordenadas para centrar el texto horizontalmente
        texto1_x = (ancho - texto1_ancho) // 2
        texto2_x = (ancho - texto2_ancho) // 2

        # Pegar el texto en la imagen
        draw.text((texto1_x, int(alto * 0.3)), texto1, fill="black", font=fuente)
        draw.text((texto2_x, int(alto * 0.45)), texto2, fill="black", font=fuente)
        
        # Generar código de barras
        numero = self.contador_label.cget("text")
        codigo = barcode.get_barcode_class('code128')
        codigo_imagen = codigo(numero, writer=ImageWriter())
        codigo_imagen.save('codigo_barras')
        
        # Cargar y pegar el código de barras en la parte inferior (centrado horizontalmente)
        codigo_barras = Image.open('codigo_barras.png').resize((int(ancho * 0.6), int(alto * 0.4)))
        posicion_x = (ancho - codigo_barras.width) // 2
        posicion_y = alto - int(alto * 0.4)
        imagen.paste(codigo_barras, (posicion_x, posicion_y))
        
        # Guardar la imagen generada
        imagen.save(f"./Archivos/Tickets/{numero}.jpg")
        #
        # HORZRES / VERTRES = área imprimible
        #
        HORZRES = 8
        VERTRES = 10
        #
        # LOGPIXELS = puntos por pulgada
        #
        LOGPIXELSX = 88
        LOGPIXELSY = 90
        #
        # PHYSICALWIDTH/HEIGHT = area total
        #
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111
        #
        # PHYSICALOFFSETX/Y = margen izquierdo/superior
        #
        PHYSICALOFFSETX = 112
        PHYSICALOFFSETY = 113

        printer_name = win32print.GetDefaultPrinter()
        file_name = f"./Archivos/Tickets/{numero}.jpg" #nota la imagenes se ven mejor en monocromatico .bmp
        #
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
        printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
        printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)
        #
        #
        bmp = Image.open(file_name)
        if bmp.size[0] > bmp.size[1]:
            bmp = bmp.rotate(0)  # Rotar 

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min(ratios)

        #
        # Inicie el trabajo de impresión y dibuje el mapa de bits para
        #  el dispositivo de impresión en el tamaño escalado.
        #
        hDC.StartDoc(file_name)
        hDC.StartPage()

        dib = ImageWin.Dib(bmp)
        scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
        x1 = int((printer_size[0] - scaled_width) / 2)
        y1 = int((printer_size[1] - scaled_height) / 2)
        x2 = x1 + scaled_width
        y2 = y1 + scaled_height
        dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()
        
    def cargar_datos(self):
        while True:
            try:
                folio = simpledialog.askstring("Cargar Datos", "Ingrese el folio del archivo JSON a cargar:")
                if folio is None:  # El usuario canceló la entrada
                    break
                with open(f"./Archivos/Progresos/{folio}.json", "r") as f:
                    datos = json.load(f)
                    self.fecha_recepcion_entry.delete(0, tk.END)
                    self.fecha_entrega_entry.delete(0, tk.END)
                    self.no_serie_entry.delete(0, tk.END)
                    self.no_inventario_entry.delete(0, tk.END)
                    # Asignar los datos cargados a los campos correspondientes
                    self.fecha_recepcion_entry.insert(0, datos.get("fecha_recepcion", ""))
                    self.fecha_entrega_entry.insert(0, datos.get("fecha_entrega", ""))
                    self.nombre_responsable_entry.insert(0, datos.get("nombre_responsable", ""))
                    self.telefono_responsable_entry.insert(0, datos.get("telefono_responsable", ""))
                    self.area_responsable_entry.insert(0, datos.get("area_responsable", ""))
                    self.correo_entry.insert(0, datos.get("correo", ""))
                    self.otro_equipo_entry.insert(0, datos.get("otro_equipo", ""))
        
                    # Verificar si el campo "estado_equipo" es True y el campo "falta_equipo" es "Falta algun componente"
                    estado_componente = datos.get("estado_componente", False)
                    
                    if estado_componente:
                        self.faltante_equipo_entry.config(state="normal")
                        self.faltante_equipo_entry.insert(0, datos.get("falta_equipo", ""))
                    else:
                        self.faltante_equipo_entry.config(state="disabled")
        
                    # Cargar los estados de los Checkbuttons
                    self.modelo_equipo_entry.insert(0, datos.get("modelo_equipo", ""))
                    self.marca_equipo_entry.insert(0, datos.get("marca_equipo", ""))
                    self.no_serie_entry.insert(0, datos.get("no_serie", ""))
                    self.no_inventario_entry.insert(0, datos.get("no_inventario", ""))
                    self.descripcion_detallada_entry.insert("1.0", datos.get("descripcion_detallada", ""))
                    # Si la descripción del equipo es "Otros", establecer el valor en el combobox y habilitar la entrada de texto
                    descripcion_equipo = datos.get("descripcion_equipo", "")
                    if descripcion_equipo == "Otros":
                        self.equipo_combobox.set("Otros")
                        self.otro_equipo_entry.config(state="normal")
                    else:
                        self.equipo_combobox.set(descripcion_equipo)
                        self.otro_equipo_entry.config(state="disabled")
                    # Cargar los estados de los Checkbuttons
                    self.estado_enciende.set(datos.get("estado_enciende", False))
                    self.estado_cable.set(datos.get("estado_cable", False))
                    self.estado_componente.set(datos.get("estado_componente", False))
                    self.estado_dano_botones.set(datos.get("estado_dano_botones", False))
                    self.estado_corrosion.set(datos.get("estado_corrosion", False))
                    self.estado_dano_carcasa.set(datos.get("estado_dano_carcasa", False))
                    messagebox.showinfo("Cargar Progreso", "Progreso Cargado Correctamente")
                    self.contador_label.config(text=folio)
                    break  # Salir del bucle después de cargar correctamente
            except FileNotFoundError:
                messagebox.showinfo("Cargar Progreso", "Progreso NO Cargado Correctamente")
    def cargar_numero(self):
    # Intentar cargar el número guardado desde el archivo
    
        try:
            with open("./Folio/numero_guardado.txt", "r") as file:
                numero = int(file.read())
                
            return numero
        except FileNotFoundError:
            return 0
        except ValueError:
            messagebox.showwarning("Advertencia", "El archivo de número guardado está dañado.")
            return 0
        
    def guardar_numero(self, numero):
        # Guardar el número en el archivo
        with open("./Folio/numero_guardado.txt", "w") as file:
            file.write(str(numero))
        
    
    def toggle_otros(self):
        # Habilitar o deshabilitar la caja de texto "Otros" según el estado del checkbox "Otros"
        if self.otro_var.get():
            self.otros_materiales_entry.config(state="normal")
        else:
            self.otros_materiales_entry.config(state="disabled")
            
    def toggle_falta(self):
        # Habilitar o deshabilitar la caja de texto "falta" según el estado del checkbox "Otros"
        if self.estado_componente.get():
            self.faltante_equipo_entry.config(state="normal")
        else:
            self.faltante_equipo_entry.config(state="disabled")

   
    



# Inicializar la aplicación Tkinter
root = tk.Tk()
root.geometry("900x610")
app = BitacoraMantenimiento(root)



# Ejecutar la aplicación
root.mainloop()