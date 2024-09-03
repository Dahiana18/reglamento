import tkinter as tk
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def insertar_tabla(doc, paragraph, horarios):
    # Añadir una tabla con 2 columnas y tantas filas como elementos en horarios
    table = doc.add_table(rows=1, cols=2)
    
    # Añadir encabezados
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Tipo de Horario'
    hdr_cells[1].text = 'Horario'
    
    # Añadir filas con datos
    for horario in horarios:
        row_cells = table.add_row().cells
        row_cells[0].text = horario
        row_cells[1].text = ''
    
    # Mover la tabla al lugar correcto
    tbl = table._tbl
    paragraph._element.addnext(tbl)

    # Añadir bordes a la tabla
    tbl = table._tbl
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tbl.tblPr.append(tblBorders)

def reemplazar_datos_en_plantilla(nombre, municipio, departamento, horarios):
    # Cargar el documento de Word
    doc = Document('template.docx')
    
    # Convertir 'nombre' a mayúsculas
    nombre = nombre.upper()

    # Recorrer todos los párrafos y reemplazar las palabras clave
    for p in doc.paragraphs:
        if "|NOMBRE|" in p.text:
            p.text = p.text.replace('|NOMBRE|', nombre)
            for run in p.runs:
                if nombre in run.text:
                    run.bold = True  # Aplicar negrita al run modificado

        if "|MUNICIPIO|" in p.text:
            p.text = p.text.replace('|MUNICIPIO|', municipio)
        
        if "|DEPARTAMENTO|" in p.text:
            p.text = p.text.replace('|DEPARTAMENTO|', departamento)
            
        if "|HORARIO|" in p.text:
            p.text = p.text.replace('|HORARIO|', "")
            insertar_tabla(doc, p, horarios)
            break  # Salir del bucle después de insertar la tabla

    # Guardar el documento modificado
    doc.save('documento_completado.docx')

def generar_tabla():
    global table_frame  # Asegúrate de que table_frame esté accesible
    for widget in table_frame.winfo_children():
        widget.destroy()
    
    row = 0
    tk.Label(table_frame, text="Tipo de Horario").grid(row=row, column=0)
    tk.Label(table_frame, text="Horario").grid(row=row, column=1)
    row += 1
    
    if operativo_var.get():
        tk.Label(table_frame, text="Horario de trabajo personal operativo").grid(row=row, column=0)
        tk.Label(table_frame, text="").grid(row=row, column=1)
        row += 1
    
    if administrativo_var.get():
        tk.Label(table_frame, text="Horario de trabajo personal administrativo").grid(row=row, column=0)
        tk.Label(table_frame, text="").grid(row=row, column=1)
        row += 1

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
    
    global operativo_var, administrativo_var, table_frame
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
        
        # Generar la tabla automáticamente
        generar_tabla()

        reemplazar_datos_en_plantilla(nombre, municipio, departamento, horarios)

    # Botón para enviar el formulario
    submit_button = tk.Button(ventana, text="Generar Documento", command=on_submit)
    submit_button.grid(row=5, columnspan=2)

    # Crear un frame para la tabla
    global table_frame  # Asegúrate de que table_frame esté accesible
    table_frame = tk.Frame(ventana)
    table_frame.grid(row=6, column=0, columnspan=2)

    # Iniciar el bucle de la aplicación Tkinter
    ventana.mainloop()

# Llamar a la función para crear el formulario
crear_formulario()