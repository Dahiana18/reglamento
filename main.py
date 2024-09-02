import tkinter as tk
from tkinter import messagebox
from docx import Document
from docx.shared import Pt  # Para establecer el tamaño de fuente
from docx.oxml import OxmlElement  # Para estilos de texto (negrita)

def reemplazar_datos_en_plantilla(nombre, municipio, departamento,horarios):
    # Cargar el documento de Word
    doc = Document('template.docx')
    
    # Convertir 'nombre' a mayúsculas
    nombre = nombre.upper()

    # Recorrer todos los párrafos y reemplazar las palabras clave
    for p in doc.paragraphs:
        for run in p.runs:
            if "NOMBRE" in run.text:
                # Reemplazar "NOMBRE" con el nombre ingresado en mayúsculas
                run.text = run.text.replace('|NOMBRE|', nombre)
                run.bold = True  # Aplicar negrita al run modificado

            if "MUNICIPIO" in run.text:
                run.text = run.text.replace('|MUNICIPIO|', municipio)
            
            if "DEPARTAMENTO" in run.text:
                run.text = run.text.replace('|DEPARTAMENTO|', departamento)
                
            if "HORARIO" in run.text:
                run.text = run.text.replace('|HORARIO|', horarios)

    # Guardar el documento modificado
    doc.save('documento_completado.docx')
    doc = Document('documento_completado.docx')
    for p in doc.paragraphs:
        for run in p.runs:
            if nombre.upper() in run.text:
                # Aplicar diferentes estilos
                run.text = run.text.replace(nombre.upper(), nombre.upper())  # Cambiar a mayúsculas
                run.bold = True   # Aplicar negrita
                run.italic = False  # Aplicar cursiva
                run.underline = False  # Aplicar subrayado
                run.font.size = Pt(11)  # Cambiar tamaño de fuente a 14 puntos

    # Guardar el documento modificado
    nuevo_nombre_archivo = 'documento_completado.docx'
    doc.save(nuevo_nombre_archivo)
    
    messagebox.showinfo("Éxito", "El documento ha sido generado exitosamente.")

def crear_formulario():
    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.title("Formulario de Datos")
    
    # Etiquetas y campos de entrada para Nombre, Municipio, Departamento
    tk.Label(ventana, text="Nombre:").grid(row=0, column=0)
    nombre_entry = tk.Entry(ventana)
    nombre_entry.grid(row=0, column=1)

    tk.Label(ventana, text="Municipio:").grid(row=1, column=0)
    municipio_entry = tk.Entry(ventana)
    municipio_entry.grid(row=1, column=1)

    tk.Label(ventana, text="Departamento:").grid(row=2, column=0)
    departamento_entry = tk.Entry(ventana)
    departamento_entry.grid(row=2, column=1)
    
    # Checkbuttons para seleccionar horario de trabajo
    tk.Label(ventana, text="Horario de trabajo:").grid(row=3, column=0)
    
    operativo_var = tk.IntVar()
    administrativo_var = tk.IntVar()

    operativo_cb = tk.Checkbutton(ventana, text="Horario de trabajo personal operativo", variable=operativo_var)
    operativo_cb.grid(row=3, column=1, sticky="w")

    administrativo_cb = tk.Checkbutton(ventana, text="Horario de trabajo personal administrativo", variable=administrativo_var)
    administrativo_cb.grid(row=4, column=1, sticky="w")

    def on_submit():
        nombre = nombre_entry.get()
        municipio = municipio_entry.get()
        departamento = departamento_entry.get()
        
        # Construir la cadena de horarios seleccionados
        horarios = []
        if operativo_var.get():
            horarios.append("Horario de trabajo personal operativo")
        if administrativo_var.get():
            horarios.append("Horario de trabajo personal administrativo")
        
        horarios_str = ', '.join(horarios) if horarios else "Ninguno"

        reemplazar_datos_en_plantilla(nombre, municipio, departamento, horarios_str)

    # Botón para enviar el formulario
    submit_button = tk.Button(ventana, text="Generar Documento", command=on_submit)
    submit_button.grid(row=5, columnspan=2)

    # Iniciar el bucle de la aplicación Tkinter
    ventana.mainloop()

if __name__ == "__main__":
    crear_formulario()
