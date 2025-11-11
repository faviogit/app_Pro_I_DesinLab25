import tkinter as tk                                     # importamos tkinter para la interfaz gráfica
from tkinter import filedialog, messagebox, scrolledtext # importamos componentes específicos de tkinter
from PIL import Image, ImageTk                           # Pillow: para manipular imágenes (iconos si se requieren usar PNG/ICO)
import pandas as pd                                      # pandas: para leer Excel
from datetime import datetime                            # para obtener la fecha actual
from docxtpl import DocxTemplate                         # docxtpl: para renderizar plantillas .docx con variables
import os                                                # operaciones con paths y carpetas
import sys                                               # para detectar si estamos en un exe (PyInstaller)

# --- Helper: obtener ruta de recursos (soporta PyInstaller) ---
def resource_path(relative_path):                        # definimos una función que construye la ruta correcta
    """
    Devuelve la ruta absoluta de un archivo, funcionando igual
    en desarrollo o cuando el script está empaquetado con PyInstaller.
    """
    if getattr(sys, 'frozen', False):                    # si estamos en un ejecutable (PyInstaller)...
        base_path = sys._MEIPASS                         # PyInstaller extrae los archivos a _MEIPASS
    else:
        base_path = os.path.abspath(".")                 # en desarrollo usamos la carpeta actual
    return os.path.join(base_path, relative_path)        # unimos base + ruta relativa y la retornamos

# --- Funciones para seleccionar archivos con diálogo ---
def seleccionar_excel():                                 # abre el explorador para elegir un archivo Excel
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    entrada_excel.delete(0, tk.END)                      # limpiamos la entrada de texto
    entrada_excel.insert(0, ruta)                        # ponemos la ruta seleccionada en la entrada

def seleccionar_plantilla():                             # abre el explorador para elegir la plantilla .docx
    ruta = filedialog.askopenfilename(
        title="Selecciona la plantilla Word",
        filetypes=[("Plantilla Word", "*.docx")]
    )
    entrada_plantilla.delete(0, tk.END)                  # limpiamos la entrada de texto
    entrada_plantilla.insert(0, ruta)                    # ponemos la ruta seleccionada en la entrada

# --- Función principal que genera los documentos ---
def generar_documentos():
    try:
        excel_path = entrada_excel.get().strip()         # leemos la ruta del Excel desde la entrada
        doc_path = entrada_plantilla.get().strip()       # leemos la ruta de la plantilla desde la entrada

        if not excel_path or not doc_path:               # validamos que ambos archivos hayan sido seleccionados
            messagebox.showwarning("Advertencia", "Selecciona el Excel y la plantilla antes de continuar.")
            return

        # aseguramos que la carpeta de salida exista
        os.makedirs("prueba", exist_ok=True)

        # constantes que se inyectarán en la plantilla
        nombre = "favio candia"
        telefono = "61508169"
        correo = "favioalvarez26@gmail.com"
        fecha = datetime.today().strftime("%d/%m/%Y")   # formateamos la fecha a DD/MM/YYYY

        constantes = {
            'nombre': nombre,
            'telefono': telefono,
            'correo': correo,
            'fecha': fecha
        }

        # --- Lectura del Excel ---
        df = pd.read_excel(excel_path)                   # leemos el archivo Excel con pandas

        # Validación de columnas (evita KeyError y ayuda al usuario)
        required_cols = ['Nombre del Alumno', 'Mat', 'Fis', 'Qui']
        for col in required_cols:
            if col not in df.columns:
                messagebox.showerror("Error", f"Falta la columna '{col}' en el Excel. Asegúrate del nombre exacto.")
                return

        # --- Generación de documentos ---
        for indice, fila in df.iterrows():                # iteramos fila por fila del DataFrame
            doc = DocxTemplate(doc_path)                  # abrimos la plantilla .docx

            contenido = {
                'nombre_alumno': fila['Nombre del Alumno'], # valores tomados de las columnas del Excel
                'nota_mat': fila['Mat'],
                'nota_fis': fila['Fis'],
                'nota_qui': fila['Qui']
            }

            contenido.update(constantes)                 # añadimos las constantes al diccionario de contexto
            nombre_archivo = f"prueba/Notas_de_{fila['Nombre del Alumno']}.docx" # nombre de salida
            doc.render(contenido)                         # renderizamos la plantilla con los datos
            doc.save(nombre_archivo)                      # guardamos el archivo .docx generado

            log_text.insert(tk.END, f" Generado: {nombre_archivo}\n")   # escribimos en el log de la UI
            log_text.see(tk.END)                         # hacemos scroll automático al final

        messagebox.showinfo("Éxito", "Documentos generados correctamente en la carpeta 'prueba'")

    except Exception as e:
        messagebox.showerror("Error", str(e))           # mostramos cualquier excepción al usuario en un popup

# --- Ventana principal (UI) ---
ventana = tk.Tk()                                      # creamos la ventana principal de la aplicación
ventana.title("Generador de Documentos - TECBA")       # título de la ventana
ventana.geometry("700x500")                            # tamaño inicial de la ventana

# Intentamos cargar un icono .ico si existe (usado en Windows)
try:
    icon_path = resource_path("icono/images.ico")     # obtenemos la ruta usando resource_path
    ventana.iconbitmap(icon_path)                     # establecemos el icono (solo .ico funciona con iconbitmap)
except Exception:
    # Si falla, intentamos un PNG vía iconphoto (más flexible)
    try:
        png_icon = resource_path("icono/images.png")
        img = Image.open(png_icon)
        photo = ImageTk.PhotoImage(img)
        ventana.iconphoto(True, photo)                # iconphoto acepta PNG
    except Exception:
        print("No se pudo cargar el ícono (ni .ico ni .png). Continuando sin icono.")

# --- Widgets: título ---
titulo = tk.Label(ventana, text="Generador de Documentos", font=("Arial", 16, "bold"))
titulo.pack(pady=10)

# --- Frame: selección de Excel ---
frame_excel = tk.Frame(ventana)
frame_excel.pack(pady=5, fill="x", padx=10)

tk.Label(frame_excel, text="Archivo Excel:", font=("Arial", 11)).pack(side="left")  # etiqueta
entrada_excel = tk.Entry(frame_excel, width=50)               # entrada de texto para la ruta del Excel
entrada_excel.pack(side="left", padx=5)
tk.Button(frame_excel, text="Buscar", command=seleccionar_excel).pack(side="left") # botón para abrir diálogo

# --- Frame: selección de plantilla Word ---
frame_doc = tk.Frame(ventana)
frame_doc.pack(pady=5, fill="x", padx=10)

tk.Label(frame_doc, text="Plantilla Word:", font=("Arial", 11)).pack(side="left")  # etiqueta
entrada_plantilla = tk.Entry(frame_doc, width=50)         # entrada de texto para la ruta de la plantilla
entrada_plantilla.pack(side="left", padx=5)
tk.Button(frame_doc, text="Buscar", command=seleccionar_plantilla).pack(side="left") # botón para abrir diálogo

# --- Botón principal para generar documentos ---
btn_generar = tk.Button(ventana, text="Generar Documentos", font=("Arial", 12, "bold"),
                        bg="#2e7d32", fg="white", command=generar_documentos)
btn_generar.pack(pady=15)

# --- Área de log (salida) ---
log_text = scrolledtext.ScrolledText(ventana, width=80, height=15, font=("Consolas", 10))
log_text.pack(padx=10, pady=10)

# arrancamos el bucle principal de la interfaz gráfica
ventana.mainloop()

