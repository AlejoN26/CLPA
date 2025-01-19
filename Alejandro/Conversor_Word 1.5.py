import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import os

def seleccionar_archivo_word():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Word", "*.docx")])
    if archivo:
        entry_archivo.config(state="normal")  # Habilitar para actualizar la ruta
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)
        entry_archivo.config(state="readonly")  # Volver a solo lectura

def seleccionar_archivo_pdf():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    if archivo:
        entry_archivo.config(state="normal")  # Habilitar para actualizar la ruta
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)
        entry_archivo.config(state="readonly")  # Volver a solo lectura

def convertir_a_pdf():
    archivo_word = entry_archivo.get()
    if archivo_word == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo Word.")
        return

    if archivo_word.lower().endswith(".pdf"):
        messagebox.showerror("Error", "El archivo ya está en formato PDF.")
        return

    # Mostrar mensaje de progreso
    label_progreso.config(text="Convirtiendo a PDF...")
    ventana.update_idletasks()

    word = None
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        archivo_word_absoluto = os.path.abspath(archivo_word)

        if not os.path.exists(archivo_word_absoluto):
            raise FileNotFoundError("El archivo seleccionado no existe.")

        doc = word.Documents.Open(archivo_word_absoluto, ReadOnly=True)

        archivo_pdf_absoluto = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile=os.path.splitext(os.path.basename(archivo_word))[0]
        )

        if not archivo_pdf_absoluto:
            label_progreso.config(text="Operación cancelada por el usuario.")
            return

        # Normalizar ruta para evitar codificaciones
        archivo_pdf_absoluto = os.path.normpath(archivo_pdf_absoluto)

        # Guardar como PDF usando la opción de mantener formato
        doc.SaveAs(archivo_pdf_absoluto, FileFormat=17, EmbedTrueTypeFonts=True)
        label_progreso.config(text="Conversión completada con éxito.")
        messagebox.showinfo("Éxito", f"Archivo convertido a PDF: {archivo_pdf_absoluto}")
    except FileNotFoundError as e:
        label_progreso.config(text="Error: Archivo no encontrado.")
        messagebox.showerror("Error", str(e))
    except Exception as e:
        label_progreso.config(text="Error durante la conversión.")
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        label_progreso.config(text="")  # Limpiar el mensaje de progreso

def convertir_a_word():
    archivo_pdf = entry_archivo.get()
    if archivo_pdf == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo PDF.")
        return

    # Mostrar mensaje de progreso
    label_progreso.config(text="Convirtiendo a Word...")
    ventana.update_idletasks()

    word = None
    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        archivo_pdf_absoluto = os.path.abspath(archivo_pdf)

        if not os.path.exists(archivo_pdf_absoluto):
            raise FileNotFoundError("El archivo seleccionado no existe.")

        doc = word.Documents.Open(archivo_pdf_absoluto)

        archivo_word_absoluto = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Archivos Word", "*.docx")],
            initialfile=os.path.splitext(os.path.basename(archivo_pdf))[0]
        )

        if not archivo_word_absoluto:
            label_progreso.config(text="Operación cancelada por el usuario.")
            return

        # Normalizar ruta para evitar codificaciones
        archivo_word_absoluto = os.path.normpath(archivo_word_absoluto)

        # Guardar como Word usando formato original
        doc.SaveAs(archivo_word_absoluto, FileFormat=16)
        label_progreso.config(text="Conversión completada con éxito.")
        messagebox.showinfo("Éxito", f"Archivo convertido a Word: {archivo_word_absoluto}")
    except FileNotFoundError as e:
        label_progreso.config(text="Error: Archivo no encontrado.")
        messagebox.showerror("Error", str(e))
    except Exception as e:
        label_progreso.config(text="Error durante la conversión.")
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        label_progreso.config(text="")  # Limpiar el mensaje de progreso

# Interfaz gráfica con Tkinter
ventana = tk.Tk()
ventana.title("Conversor de Word a PDF / PDF a Word")

# Establecer tamaño inicial de la ventana
ventana.geometry("600x400")  # Ajustar el tamaño según sea necesario

# Crear el campo de entrada en estado de solo lectura
entry_archivo = tk.Entry(ventana, width=50, state="readonly")
entry_archivo.pack(pady=10)

boton_seleccionar_word = tk.Button(ventana, text="Seleccionar archivo Word", command=seleccionar_archivo_word)
boton_seleccionar_word.pack(pady=5)

boton_seleccionar_pdf = tk.Button(ventana, text="Seleccionar archivo PDF", command=seleccionar_archivo_pdf)
boton_seleccionar_pdf.pack(pady=5)

boton_convertir_pdf = tk.Button(ventana, text="Convertir a PDF", command=convertir_a_pdf)
boton_convertir_pdf.pack(pady=5)

boton_convertir_word = tk.Button(ventana, text="Convertir a Word", command=convertir_a_word)
boton_convertir_word.pack(pady=5)

# Etiqueta para mostrar el progreso
label_progreso = tk.Label(ventana, text="", fg="blue")
label_progreso.pack(pady=10)

ventana.mainloop()