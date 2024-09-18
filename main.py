import tkinter as tk
from tkinter import ttk
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import RGBColor
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

def reemplazar_datos_en_plantilla(nombre, municipio, departamento, objeto_social, fecha_pago, horarios):
    # Cargar el documento de Word
    doc = Document('template.docx')
    
    # Convertir 'nombre' a mayúsculas
    nombre = nombre.upper()

    print(f"Valor de fecha_pago: {fecha_pago}")

    # Recorrer todos los párrafos y reemplazar las palabras clave
    for p in doc.paragraphs:
        for run in p.runs:
            if "|NOMBRE|" in run.text:
                run.text = run.text.replace('|NOMBRE|', nombre)
                run.bold = True  # Aplicar negrita al run modificado

            if "|MUNICIPIO|" in run.text:
                run.text = run.text.replace('|MUNICIPIO|', municipio)
            
            if "|DEPARTAMENTO|" in run.text:
                run.text = run.text.replace('|DEPARTAMENTO|', departamento)

            if "|FECHA_PAGO|" in run.text:
                print(f"Encontrado |FECHA_PAGO| en: {run.text}")  # Mensaje de depuración
                run.text = run.text.replace('|FECHA_PAGO|', fecha_pago)
                print(f"Reemplazado por: {run.text}")  # Mensaje de depuración  

            if "|OBJETO_SOCIAL|" in run.text and isinstance(objeto_social, str):
                run.text = run.text.replace('|OBJETO_SOCIAL|', objeto_social) 

                    
            

    for p in doc.paragraphs:

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

    global nombre_entry, municipio_entry, departamento_entry, fecha_pago_entry, operativo_var, administrativo_var, objeto_social_entry, table_frame
   

    font_style = ("Helvetica", 24, "italic")
    bg_color = '#b0d4ec'    

    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.title("Formulario de Datos")
    ventana.configure(bg= bg_color)

    # Frame para los datos personales (Nombre, Municipio, Departamento)
    frame_datos = tk.Frame(ventana, bg=bg_color)
    frame_datos.pack(padx=10, pady=10)

    tk.Label(frame_datos, text="Nombre Empresa:", font=font_style, bg=bg_color).grid(row=0, column=0)
    nombre_entry = tk.Entry(frame_datos, font=font_style)
    nombre_entry.grid(row=0, column=1, padx=(0, 20))

    tk.Label(frame_datos, text="Departamento:", font=font_style, bg=bg_color).grid(row=0, column=2)
    departamento_entry = tk.Entry(frame_datos, font=font_style)
    departamento_entry.grid(row=0, column=3, padx=(0, 20))

    tk.Label(frame_datos, text="Municipio:", font=font_style, bg=bg_color).grid(row=0, column=4)
    municipio_entry = tk.Entry(frame_datos, font=font_style)
    municipio_entry.grid(row=0, column=5, pady=20)   


    tk.Label(frame_datos, text="Objeto Social:", font=font_style, bg=bg_color).grid(row=1, column=0)
    objeto_social_entry = tk.Entry(frame_datos, font=font_style)
    objeto_social_entry.grid(row=1, column=1, columnspan=5, sticky="we", pady=20)

    tk.Label(frame_datos, text="Fecha de Pago:", font=font_style, bg=bg_color).grid(row=2, column=0)
    opciones_pago = ["los días 30 de cada mes", "los días 15 y 30 de cada mes", "catorcenales", "semanales"]
    fecha_pago_entry = ttk.Combobox(frame_datos, values=opciones_pago, font=font_style, state="readonly")
    fecha_pago_entry.grid(row=2, column=1,columnspan=5, sticky="we", pady=20)
    fecha_pago_entry.set("Seleccione una opción")

    # Configurar la fuente del menú desplegable
    ventana.option_add('*TCombobox*Listbox.font', font_style)   
    

    # Frame para los checkbuttons (Horario de trabajo)
    frame_horarios = tk.Frame(ventana)
    frame_horarios.pack(padx=10, pady=10)

    tk.Label(frame_horarios, text="Horario de trabajo:").grid(row=0, column=0)

    operativo_var = tk.IntVar()
    administrativo_var = tk.IntVar()

    operativo_cb = tk.Checkbutton(frame_horarios, text="Horario de trabajo personal operativo", variable=operativo_var)
    operativo_cb.grid(row=1, column=0, sticky="w")

    administrativo_cb = tk.Checkbutton(frame_horarios, text="Horario de trabajo personal administrativo", variable=administrativo_var)
    administrativo_cb.grid(row=2, column=0, sticky="w")

    

    # Función para manejar el envío de datos
    def on_submit():        

        nombre = nombre_entry.get()
        municipio = municipio_entry.get()
        departamento = departamento_entry.get()
        objeto_social = objeto_social_entry.get()
        fecha_pago = fecha_pago_entry.get()

        # Construir la cadena de horarios seleccionados
        horarios = []
        if operativo_var.get():
            horarios.append("Horario de trabajo personal operativo")
        if administrativo_var.get():
            horarios.append("Horario de trabajo personal administrativo")

        # Generar la tabla automáticamente
        generar_tabla()

        # Reemplazar datos en la plantilla
        reemplazar_datos_en_plantilla(nombre, municipio, departamento, objeto_social, fecha_pago, horarios)

    # Frame para el botón de enviar
    frame_botones = tk.Frame(ventana)
    frame_botones.pack(padx=10, pady=10)

    # Botón para enviar el formulario
    submit_button = tk.Button(frame_botones, text="Generar Documento", command=on_submit)
    submit_button.grid(row=0, column=0)
    
    # Frame para la tabla
    table_frame = tk.Frame(ventana)
    table_frame.pack(padx=10, pady=10)

    # Iniciar el bucle de la aplicación Tkinter
    ventana.mainloop()

# Llamar a la función para crear el formulario
crear_formulario()
