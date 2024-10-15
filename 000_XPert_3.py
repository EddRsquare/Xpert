# -*- coding: utf-8 -*-
"""
Created on Thu Sep  5 16:29:12 2024

@author: Edd_Rsquare
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import openpyxl

# Variables globales
datos = None
datos_nuevos = pd.DataFrame(columns=["CONTRACT ID", "A.ERC_AssetManger", "A.ERC_Lawyer",
                                     "B.ERC_StartDate_AssetManager", "B.ERC_Meses_estimados_recobro_AssetManager",
                                     "C.ERC_StartDate_Lawyer", "C.ERC_Meses_estimados_recobro_Lawyer",
                                     "D.AssetComments", "D.LawyerComments", "Fecha", "Contrato",
                                     "Movimientos", "Demanda", "Despacho de Ejecución", "Av patrimonial",
                                     "Ultimo Hito Procesal", "Informe AC", "Informe Convenio",
                                     "Informe Liquidación", "Adhesion al convenio", "Comunicación de créditos",
                                     "Comunicación cuenta al convenio", "Textos definitivos", "Certificado de defunción",
                                     "Notas simples", "Burofax/otros", "CompanyStatus"])

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Xpert 2.0")
ventana.geometry("1200x800")

# Añadir un título al inicio de la ventana principal
titulo_label = tk.Label(ventana, text="XPERT 2.0 By Edgar Aguilar - SMEs Manual valuation ", font=("Helvetica", 16))
titulo_label.grid(row=0, column=0, pady=10, sticky="w", padx=(10, 0))

# Etiqueta Documentos
doc_label = tk.Label(ventana, text="Available documentation, check the documentation available in the VDR:", font=("Helvetica", 12))
doc_label.grid(row=2, column=0, pady=10, sticky="w", padx=(10, 0))

# Etiqueta ERC
doc_label = tk.Label(ventana, text="Manual ERCs, enter the estimated value, approximate date of first payment, number of months and comments:", font=("Helvetica", 12))
doc_label.grid(row=13, column=0, pady=10, sticky="w", padx=(10, 0))

# Crear la segunda ventana para mostrar datos_nuevos
ventana_nuevos_datos = tk.Toplevel(ventana)
ventana_nuevos_datos.title("Nuevos Datos")
ventana_nuevos_datos.geometry("1200x400")
ventana_nuevos_datos.geometry(f"+{ventana.winfo_x()}+{ventana.winfo_y() + 820}")

# Evitar que la ventana "Nuevos Datos" se cierre accidentalmente
def on_closing_nuevos_datos():
    # Hacer nada al intentar cerrar la ventana
    pass

# Asignar la función para evitar el cierre
ventana_nuevos_datos.protocol("WM_DELETE_WINDOW", on_closing_nuevos_datos)

# Crear la tabla para mostrar los datos_nuevos
columns_nuevos_datos = ["CONTRACT ID", "A.ERC_AssetManger", "A.ERC_Lawyer",
                        "B.ERC_StartDate_AssetManager", "B.ERC_Meses_estimados_recobro_AssetManager",
                        "C.ERC_StartDate_Lawyer", "C.ERC_Meses_estimados_recobro_Lawyer",
                        "D.AssetComments", "D.LawyerComments", "Fecha", "Contrato",
                        "Movimientos", "Demanda", "Despacho de Ejecución", "Av patrimonial",
                        "Ultimo Hito Procesal", "Informe AC", "Informe Convenio",
                        "Informe Liquidación", "Adhesion al convenio", "Comunicación de créditos",
                        "Comunicación cuenta al convenio", "Textos definitivos", "Certificado de defunción",
                        "Notas simples", "Burofax/otros","CompanyStatus"]

tree_nuevos_datos = ttk.Treeview(ventana_nuevos_datos, columns=columns_nuevos_datos, show="headings")
for col in columns_nuevos_datos:
    tree_nuevos_datos.heading(col, text=col)
    tree_nuevos_datos.column(col, width=100)
tree_nuevos_datos.pack(fill=tk.BOTH, expand=True)


# Función para mostrar los datos de un CIF
def mostrar_datos_cif(event):
    cif_seleccionado = lista_cif.get()
    datos_cif = datos[datos["CIF"] == cif_seleccionado]
   
    # Llenar el Combobox de Contract ID con los valores correspondientes
    campo_contract_id['values'] = list(datos_cif["CONTRACT ID"].unique())
    campo_contract_id.current(0)
   
    for i in tree.get_children():
        tree.delete(i)
    for index, row in datos_cif.iterrows():
        tree.insert("", tk.END, values=row.tolist())

# Crear un DataFrame global para almacenar los registros de la ventana Calendarios
df_calendarios = pd.DataFrame(columns=["Contrato"] + [f"%FY{i}" for i in range(1, 16)])

# Variable global para rastrear el modo de edición
edit_mode = False
selected_index = None



def abrir_ventana_calendarios():
    ventana_calendarios = tk.Toplevel(ventana)
    ventana_calendarios.title("Calendarios")
    ventana_calendarios.geometry("800x400")  # Ajuste del tamaño de la ventana para dar espacio a la tabla horizontal

    global edit_mode, selected_index  # Acceder a las variables globales de edición

    # ComboBox para seleccionar el "CIF"
    tk.Label(ventana_calendarios, text="CIF").grid(row=0, column=0, pady=5, sticky="e")
    cif_combobox = ttk.Combobox(ventana_calendarios)
    cif_combobox.grid(row=0, column=1, pady=5, sticky="w")
    cif_combobox['values'] = list(datos["CIF"].unique())  # Poblar el ComboBox con los CIF únicos de los datos
    cif_combobox.current(0)

    # ComboBox para seleccionar "NumContratoCalendar" filtrado por el "CIF" seleccionado en la ventana "Calendarios"
    tk.Label(ventana_calendarios, text="NumContratoCalendar").grid(row=1, column=0, pady=5, sticky="e")
    contract_id_combobox = ttk.Combobox(ventana_calendarios)
    contract_id_combobox.grid(row=1, column=1, pady=5, sticky="w")

    # Actualizar el combobox de "CONTRACT ID" cuando se seleccione un CIF
    def actualizar_contract_id_combobox():
        cif_seleccionado = cif_combobox.get()
        if cif_seleccionado:
            datos_cif = datos[datos["CIF"] == cif_seleccionado]
            contract_id_combobox['values'] = list(datos_cif["CONTRACT ID"].unique())
            if contract_id_combobox['values']:
                contract_id_combobox.current(0)

    # Llamar a la función para actualizar los contratos al inicio
    actualizar_contract_id_combobox()

    # Vincular el cambio de CIF a la actualización del ComboBox de contratos
    cif_combobox.bind("<<ComboboxSelected>>", lambda event: actualizar_contract_id_combobox())

    # Crear entradas para "%FY1" a "%FY15"
    fy_entries = []
    for i in range(1, 16):
        tk.Label(ventana_calendarios, text=f"%FY{i}").grid(row=i + 1, column=0, pady=5, sticky="e")
        entry = tk.Entry(ventana_calendarios)
        entry.insert(0, "0")  # Valor por defecto de 0
        entry.grid(row=i + 1, column=1, pady=5, sticky="w")
        fy_entries.append(entry)

    # Etiqueta para mostrar la suma de los porcentajes
    suma_label = tk.Label(ventana_calendarios, text="Suma: 0%")
    suma_label.grid(row=17, column=0, columnspan=2, pady=10)

    # Crear la tabla (Treeview) para mostrar los valores capturados horizontalmente
    columns_calendarios = [f"%FY{i}" for i in range(1, 16)]
    columns_calendarios.insert(0, "Contrato")  # Añadir columna para contrato

    tree_calendarios = ttk.Treeview(ventana_calendarios, columns=columns_calendarios, show="headings", height=5)
    for col in columns_calendarios:
        tree_calendarios.heading(col, text=col)
        tree_calendarios.column(col, width=60)

    tree_calendarios.grid(row=1, column=3, rowspan=16, padx=10, pady=5, sticky="nsew")

    # Función para actualizar la suma de los valores y actualizar la tabla
    def actualizar_suma(event=None):
        try:
            suma = sum(float(entry.get()) for entry in fy_entries)
            suma_label.config(text=f"Suma: {suma:.2f}%")
        except ValueError:
            suma_label.config(text="Suma: Error")

    # Vincular la función actualizar_suma a las entradas
    for entry in fy_entries:
        entry.bind("<KeyRelease>", actualizar_suma)  # Actualizar suma al cambiar valor

    # Función para guardar los valores y añadir una nueva fila a la tabla y al DataFrame
    def guardar_valores():
        global df_calendarios, edit_mode, selected_index
        
        valores = [contract_id_combobox.get()] + [entry.get() for entry in fy_entries]

        if edit_mode:
            # Sobrescribir el registro existente en el DataFrame
            df_calendarios.iloc[selected_index] = valores
            # Actualizar el Treeview con el registro editado
            for i, col in enumerate(columns_calendarios):
                tree_calendarios.set(tree_calendarios.get_children()[selected_index], column=col, value=valores[i])

            edit_mode = False  # Salir del modo de edición
            selected_index = None
            messagebox.showinfo("Editar", "Registro editado correctamente", parent=ventana_calendarios)
        else:
            # Añadir al DataFrame global
            df_calendarios = pd.concat([df_calendarios, pd.DataFrame([valores], columns=df_calendarios.columns)], ignore_index=True)
            # Actualizar el Treeview con el nuevo registro
            tree_calendarios.insert("", "end", values=valores)
            messagebox.showinfo("Guardar", "Datos guardados", parent=ventana_calendarios)

        # Limpiar todos los campos después de guardar o editar
        contract_id_combobox.set('')  # Limpiar el ComboBox de Contract ID
        for entry in fy_entries:
            entry.delete(0, tk.END)
            entry.insert(0, "0")

    # Función para editar un registro seleccionado
    def editar_valores():
        global edit_mode, selected_index

        selected_item = tree_calendarios.selection()
        if selected_item:
            item = selected_item[0]
            selected_index = tree_calendarios.index(item)
            values = tree_calendarios.item(item, "values")
            
            # Rellenar las entradas con los valores seleccionados
            contract_id_combobox.set(values[0])
            for i, entry in enumerate(fy_entries):
                entry.delete(0, tk.END)
                entry.insert(0, values[i + 1])
            
            edit_mode = True  # Entrar en modo de edición
        else:
            messagebox.showwarning("Advertencia", "Seleccione un registro para editar.", parent=ventana_calendarios)

    # Función para borrar un registro seleccionado
    def borrar_valores():
        global df_calendarios
        selected_item = tree_calendarios.selection()
        if selected_item:
            item = selected_item[0]
            selected_index = tree_calendarios.index(item)

            # Confirmar antes de borrar
            confirmacion = messagebox.askyesno("Confirmar Borrado", "¿Está seguro de que desea borrar el registro seleccionado?", parent=ventana_calendarios)
            if confirmacion:
                # Eliminar del DataFrame
                df_calendarios.drop(index=selected_index, inplace=True)
                df_calendarios.reset_index(drop=True, inplace=True)  # Reiniciar los índices del DataFrame
                # Eliminar del Treeview
                tree_calendarios.delete(item)
                ventana_calendarios.focus_set()  # Mantener la ventana "Calendarios" enfocada
                messagebox.showinfo("Borrar", "Registro borrado.", parent=ventana_calendarios)
        else:
            messagebox.showwarning("Advertencia", "Seleccione un registro para borrar.", parent=ventana_calendarios)

    # Función para exportar los registros acumulados a un archivo Excel
    def exportar_valores():
        global df_calendarios
        # Abrir el cuadro de diálogo para guardar el archivo
        nombre_archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Exportar archivo Excel",
            parent=ventana_calendarios  # Mantener la ventana "Calendarios" visible
        )
        if nombre_archivo:
            # Exportar el DataFrame a un archivo Excel
            df_calendarios.to_excel(nombre_archivo, index=False)
            ventana_calendarios.focus_set()  # Mantener la ventana "Calendarios" enfocada
            messagebox.showinfo("Exportar", "Datos exportados correctamente.", parent=ventana_calendarios)

    # Crear un Frame para los botones
    frame_botones = tk.Frame(ventana_calendarios)
    frame_botones.grid(row=18, column=0, columnspan=4, pady=10, sticky="ew")

    # Definir el ancho de los botones para que sean del mismo tamaño
    ancho_boton = 20

    # Botones de acción con el mismo ancho
    boton_guardar = tk.Button(frame_botones, text="Guardar", command=guardar_valores, width=ancho_boton)
    boton_guardar.grid(row=0, column=0, padx=5, pady=5)

    boton_editar = tk.Button(frame_botones, text="Editar", command=editar_valores, width=ancho_boton)
    boton_editar.grid(row=0, column=1, padx=5, pady=5)

    boton_borrar = tk.Button(frame_botones, text="Borrar", command=borrar_valores, width=ancho_boton)
    boton_borrar.grid(row=0, column=2, padx=5, pady=5)

    boton_exportar = tk.Button(frame_botones, text="Exportar", command=exportar_valores, width=ancho_boton)
    boton_exportar.grid(row=0, column=3, padx=5, pady=5)

  
def cargar_archivo():
    global datos
    try:
        archivo_excel = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx; *.xls")],
        )
        if archivo_excel:  # Asegúrate de que se seleccionó un archivo
            print(f"Cargando archivo: {archivo_excel}")  # Para depuración
            datos = pd.read_excel(archivo_excel)
            print("Datos cargados correctamente")
            lista_cif['values'] = list(datos["CIF"].unique())
            lista_cif.current(0)
            mostrar_datos_cif(None)
        else:
            print("No se seleccionó ningún archivo.")
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")


# Función para guardar los datos nuevos o modificados
def guardar_datos():
    global datos_nuevos
    # Obtener los datos ingresados por el usuario
    cif = lista_cif.get()
    contrato = campo_contract_id.get()
    a_erc_asset_manager = campo_a_erc_asset_manager.get()
    a_erc_lawyer = campo_a_erc_lawyer.get()
    b_erc_start_date_asset_manager = campo_b_erc_start_date_asset_manager.get()
    b_erc_meses_estimados_recobro_asset_manager = campo_b_erc_meses_estimados_recobro_asset_manager.get()
    c_erc_start_date_lawyer = campo_c_erc_start_date_lawyer.get()
    c_erc_meses_estimados_recobro_lawyer = campo_c_erc_meses_estimados_recobro_lawyer.get()
    d_asset_comments = campo_d_asset_comments.get("1.0", tk.END)
    d_lawyer_comments = campo_d_lawyer_comments.get("1.0", tk.END)
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    checkbox_values = {
        "Contrato": contrato_check.get(),
        "Movimientos": movimientos_check.get(),
        "Demanda": demanda_check.get(),
        "Despacho de Ejecución": despacho_check.get(),
        "Av patrimonial": av_patrimonial_check.get(),
        "Ultimo Hito Procesal": ultimo_hito_check.get(),
        "Informe AC": informe_ac_check.get(),
        "Informe Convenio": informe_convenio_check.get(),
        "Informe Liquidación": informe_liquidacion_check.get(),
        "Adhesion al convenio": adhesion_convenio_check.get(),
        "Comunicación de créditos": comunicacion_creditos_check.get(),
        "Comunicación cuenta al convenio": comunicacion_cuenta_check.get(),
        "Textos definitivos": textos_definitivos_check.get(),
        "Certificado de defunción": certificado_defuncion_check.get(),
        "Notas simples": notas_simples_check.get(),
        "Burofax/otros": burofax_check.get()
    }
    company_status = campo_company_status.get()

    # Verificar si estamos editando un registro existente o agregando uno nuevo
    if tree_nuevos_datos.selection():
        selected_item = tree_nuevos_datos.selection()[0]
        index = tree_nuevos_datos.index(selected_item)
        datos_nuevos.at[index, "CONTRACT ID"] = contrato
        datos_nuevos.at[index, "A.ERC_AssetManger"] = a_erc_asset_manager
        datos_nuevos.at[index, "A.ERC_Lawyer"] = a_erc_lawyer
        datos_nuevos.at[index, "B.ERC_StartDate_AssetManager"] = b_erc_start_date_asset_manager
        datos_nuevos.at[index, "B.ERC_Meses_estimados_recobro_AssetManager"] = b_erc_meses_estimados_recobro_asset_manager
        datos_nuevos.at[index, "C.ERC_StartDate_Lawyer"] = c_erc_start_date_lawyer
        datos_nuevos.at[index, "C.ERC_Meses_estimados_recobro_Lawyer"] = c_erc_meses_estimados_recobro_lawyer
        datos_nuevos.at[index, "D.AssetComments"] = d_asset_comments
        datos_nuevos.at[index, "D.LawyerComments"] = d_lawyer_comments
        datos_nuevos.at[index, "CompanyStatus"] = company_status
        for key, value in checkbox_values.items():
            datos_nuevos.at[index, key] = value
        tree_nuevos_datos.item(selected_item, values=(
            contrato, a_erc_asset_manager, a_erc_lawyer, b_erc_start_date_asset_manager,
            b_erc_meses_estimados_recobro_asset_manager, c_erc_start_date_lawyer,
            c_erc_meses_estimados_recobro_lawyer, d_asset_comments, d_lawyer_comments,
            fecha_actual, *checkbox_values.values(), company_status
        ))
    else:
        nuevo_dato = pd.DataFrame({
            "CONTRACT ID": [contrato],
            "A.ERC_AssetManger": [a_erc_asset_manager],
            "A.ERC_Lawyer": [a_erc_lawyer],
            "B.ERC_StartDate_AssetManager": [b_erc_start_date_asset_manager],
            "B.ERC_Meses_estimados_recobro_AssetManager": [b_erc_meses_estimados_recobro_asset_manager],
            "C.ERC_StartDate_Lawyer": [c_erc_start_date_lawyer],
            "C.ERC_Meses_estimados_recobro_Lawyer": [c_erc_meses_estimados_recobro_lawyer],
            "D.AssetComments": [d_asset_comments],
            "D.LawyerComments": [d_lawyer_comments],
            "Fecha": [fecha_actual],
            **checkbox_values,
            "CompanyStatus": [company_status]
        })

        # Concatenar los nuevos datos al DataFrame global
        datos_nuevos = pd.concat([datos_nuevos, nuevo_dato], ignore_index=True)
        tree_nuevos_datos.insert("", tk.END, values=(
            contrato, a_erc_asset_manager, a_erc_lawyer, b_erc_start_date_asset_manager,
            b_erc_meses_estimados_recobro_asset_manager, c_erc_start_date_lawyer,
            c_erc_meses_estimados_recobro_lawyer, d_asset_comments, d_lawyer_comments,
            fecha_actual, *checkbox_values.values(), company_status
        ))

    # Limpiar los campos de entrada
    campo_contract_id.delete(0, tk.END)
    campo_a_erc_asset_manager.delete(0, tk.END)
    campo_a_erc_lawyer.delete(0, tk.END)
    campo_b_erc_start_date_asset_manager.delete(0, tk.END)
    campo_b_erc_meses_estimados_recobro_asset_manager.delete(0, tk.END)
    campo_c_erc_start_date_lawyer.delete(0, tk.END)
    campo_c_erc_meses_estimados_recobro_lawyer.delete(0, tk.END)
    campo_d_asset_comments.delete("1.0", tk.END)
    campo_d_lawyer_comments.delete("1.0", tk.END)
    contrato_check.set(False)
    movimientos_check.set(False)
    demanda_check.set(False)
    despacho_check.set(False)
    av_patrimonial_check.set(False)
    ultimo_hito_check.set(False)
    informe_ac_check.set(False)
    informe_convenio_check.set(False)
    informe_liquidacion_check.set(False)
    adhesion_convenio_check.set(False)
    comunicacion_creditos_check.set(False)
    comunicacion_cuenta_check.set(False)
    textos_definitivos_check.set(False)
    certificado_defuncion_check.set(False)
    notas_simples_check.set(False)
    burofax_check.set(False)
    campo_company_status.delete(0, tk.END)

    messagebox.showinfo("Registro guardado", "Registro guardado correctamente")

# Función para actualizar el Treeview de nuevos datos
def actualizar_treeview_nuevos_datos():
    if not ventana_nuevos_datos.winfo_ismapped():
        ventana_nuevos_datos.deiconify()
    for i in tree_nuevos_datos.get_children():
        tree_nuevos_datos.delete(i)
    for index, row in datos_nuevos.iterrows():
        tree_nuevos_datos.insert("", tk.END, values=row.tolist())

# Función para exportar todos los datos a un archivo Excel
def exportar_datos():
    global datos_nuevos
    nombre_archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")],
        title="Exportar archivo Excel"
    )
    if nombre_archivo:
        datos_nuevos.to_excel(nombre_archivo, index=False)
        messagebox.showinfo("Datos exportados", "Datos exportados correctamente")

# Función para manejar el doble clic en el Treeview y habilitar los campos para edición
def on_double_click(event):
    item = tree_nuevos_datos.selection()[0]
    values = tree_nuevos_datos.item(item, "values")

    # Llenar los campos con los valores seleccionados
    campo_contract_id.delete(0, tk.END)
    campo_contract_id.insert(0, values[0])

    campo_a_erc_asset_manager.delete(0, tk.END)
    campo_a_erc_asset_manager.insert(0, values[1])

    campo_a_erc_lawyer.delete(0, tk.END)
    campo_a_erc_lawyer.insert(0, values[2])

    campo_b_erc_start_date_asset_manager.delete(0, tk.END)
    campo_b_erc_start_date_asset_manager.insert(0, values[3])

    campo_b_erc_meses_estimados_recobro_asset_manager.delete(0, tk.END)
    campo_b_erc_meses_estimados_recobro_asset_manager.insert(0, values[4])

    campo_c_erc_start_date_lawyer.delete(0, tk.END)
    campo_c_erc_start_date_lawyer.insert(0, values[5])

    campo_c_erc_meses_estimados_recobro_lawyer.delete(0, tk.END)
    campo_c_erc_meses_estimados_recobro_lawyer.insert(0, values[6])

    campo_d_asset_comments.delete("1.0", tk.END)
    campo_d_asset_comments.insert("1.0", values[7])

    campo_d_lawyer_comments.delete("1.0", tk.END)
    campo_d_lawyer_comments.insert("1.0", values[8])

    # Actualizar los valores de las checkboxes
    contrato_check.set(values[10] == 'True')
    movimientos_check.set(values[11] == 'True')
    demanda_check.set(values[12] == 'True')
    despacho_check.set(values[13] == 'True')
    av_patrimonial_check.set(values[14] == 'True')
    ultimo_hito_check.set(values[15] == 'True')
    informe_ac_check.set(values[16] == 'True')
    informe_convenio_check.set(values[17] == 'True')
    informe_liquidacion_check.set(values[18] == 'True')
    adhesion_convenio_check.set(values[19] == 'True')
    comunicacion_creditos_check.set(values[20] == 'True')
    comunicacion_cuenta_check.set(values[21] == 'True')
    textos_definitivos_check.set(values[22] == 'True')
    certificado_defuncion_check.set(values[23] == 'True')
    notas_simples_check.set(values[24] == 'True')
    burofax_check.set(values[25] == 'True')

    campo_company_status.set(values[26])

# Función para borrar el registro seleccionado
def borrar_registro():
    if not tree_nuevos_datos.selection():
        messagebox.showwarning("Seleccionar registro", "Seleccione un registro para borrar.")
        return
    selected_item = tree_nuevos_datos.selection()[0]
    index = tree_nuevos_datos.index(selected_item)
    confirm = messagebox.askyesno("Confirmar borrado", "¿Está seguro de que desea borrar este registro?")
    if confirm:
        datos_nuevos.drop(index=index, inplace=True)
        tree_nuevos_datos.delete(selected_item)
        messagebox.showinfo("Registro borrado", "Registro borrado correctamente")

# Bind the double click event
tree_nuevos_datos.bind("<Double-1>", on_double_click)

# Interfaz gráfica
# Campos para mostrar datos
campo_contract_id = ttk.Combobox(ventana)  # Convertir a Combobox
campo_company_status = ttk.Combobox(ventana, values=["1_BK_Agreement", "2_BK_Common_phase", "3_BK_Liquidation", "4_Abierta", "5_Cerrada", "6_Otros"])

# Campos para ingresar nuevos datos
campo_a_erc_asset_manager = tk.Entry(ventana)
campo_a_erc_lawyer = tk.Entry(ventana)
campo_b_erc_start_date_asset_manager = tk.Entry(ventana)
campo_b_erc_meses_estimados_recobro_asset_manager = tk.Entry(ventana)
campo_c_erc_start_date_lawyer = tk.Entry(ventana)
campo_c_erc_meses_estimados_recobro_lawyer = tk.Entry(ventana)
campo_d_asset_comments = tk.Text(ventana, height=4)
campo_d_lawyer_comments = tk.Text(ventana, height=4)

# Variables de las checkboxes
contrato_check = tk.BooleanVar()
movimientos_check = tk.BooleanVar()
demanda_check = tk.BooleanVar()
despacho_check = tk.BooleanVar()
av_patrimonial_check = tk.BooleanVar()
ultimo_hito_check = tk.BooleanVar()
informe_ac_check = tk.BooleanVar()
informe_convenio_check = tk.BooleanVar()
informe_liquidacion_check = tk.BooleanVar()
adhesion_convenio_check = tk.BooleanVar()
comunicacion_creditos_check = tk.BooleanVar()
comunicacion_cuenta_check = tk.BooleanVar()
textos_definitivos_check = tk.BooleanVar()
certificado_defuncion_check = tk.BooleanVar()
notas_simples_check = tk.BooleanVar()
burofax_check = tk.BooleanVar()

# Checkboxes para contrato y otros elementos
check_contrato = tk.Checkbutton(ventana, text="Contrato", variable=contrato_check)
check_movimientos = tk.Checkbutton(ventana, text="Movimientos", variable=movimientos_check)
check_demanda = tk.Checkbutton(ventana, text="Demanda", variable=demanda_check)
check_despacho = tk.Checkbutton(ventana, text="Despacho de Ejecución", variable=despacho_check)
check_av_patrimonial = tk.Checkbutton(ventana, text="Av patrimonial", variable=av_patrimonial_check)
check_ultimo_hito = tk.Checkbutton(ventana, text="Ultimo Hito Procesal", variable=ultimo_hito_check)
check_informe_ac = tk.Checkbutton(ventana, text="Informe AC", variable=informe_ac_check)
check_informe_convenio = tk.Checkbutton(ventana, text="Informe Convenio", variable=informe_convenio_check)
check_informe_liquidacion = tk.Checkbutton(ventana, text="Informe Liquidación", variable=informe_liquidacion_check)
check_adhesion_convenio = tk.Checkbutton(ventana, text="Adhesion al convenio", variable=adhesion_convenio_check)
check_comunicacion_creditos = tk.Checkbutton(ventana, text="Comunicación de créditos", variable=comunicacion_creditos_check)
check_comunicacion_cuenta = tk.Checkbutton(ventana, text="Comunicación cuenta al convenio", variable=comunicacion_cuenta_check)
check_textos_definitivos = tk.Checkbutton(ventana, text="Textos definitivos", variable=textos_definitivos_check)
check_certificado_defuncion = tk.Checkbutton(ventana, text="Certificado de defunción", variable=certificado_defuncion_check)
check_notas_simples = tk.Checkbutton(ventana, text="Notas simples", variable=notas_simples_check)
check_burofax = tk.Checkbutton(ventana, text="Burofax/otros", variable=burofax_check)

# Botones de acción
boton_cargar_archivo = tk.Button(ventana, text="Cargar archivo Excel", command=cargar_archivo)
boton_guardar_datos = tk.Button(ventana, text="Guardar datos", command=guardar_datos)
boton_exportar_datos = tk.Button(ventana, text="Exportar información", command=exportar_datos)
boton_borrar_registro = tk.Button(ventana, text="Borrar registro", command=borrar_registro)
boton_calendarios = tk.Button(ventana, text="Calendarios", command=abrir_ventana_calendarios)  # Vinculando la función

# Lista desplegable para seleccionar CIF
tk.Label(ventana, text="CIF").grid(row=1, column=0, pady=5, sticky="e", padx=(10, 0))
lista_cif = ttk.Combobox(ventana)
lista_cif.bind("<<ComboboxSelected>>", mostrar_datos_cif)
lista_cif.grid(row=1, column=1, pady=5, sticky="w", padx=(0, 10))
boton_cargar_archivo.grid(row=1, column=2, pady=5, sticky="w", padx=(0, 10))

# Posicionar widgets en la ventana
boton_cargar_archivo.grid(row=0, column=2)

# Posicionar los widgets en la ventana principal con espacio a los lados
tk.Label(ventana, text="CIF").grid(row=0, column=0, pady=5, sticky="e", padx=(10, 0))
lista_cif.grid(row=0, column=1, pady=5, sticky="w", padx=(0, 10))
boton_cargar_archivo.grid(row=0, column=2, pady=5, sticky="w", padx=(0, 10))

# Posicionar el Combobox de Contract ID en la ventana principal
tk.Label(ventana, text="Contract ID").grid(row=1, column=0, pady=5, sticky="e", padx=(10, 0))
campo_contract_id.grid(row=1, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="Company Status").grid(row=2, column=0, pady=5, sticky="e", padx=(10, 0))
campo_company_status.grid(row=2, column=1, pady=5, sticky="w", padx=(0, 10))

# Posicionar los checkboxes
check_contrato.grid(row=5, column=0, pady=5, sticky="w", padx=(10, 0))
check_movimientos.grid(row=5, column=1, pady=5, sticky="w", padx=(10, 0))
check_demanda.grid(row=5, column=2, pady=5, sticky="w", padx=(10, 0))
check_despacho.grid(row=6, column=2, pady=5, sticky="w", padx=(10, 0))
check_av_patrimonial.grid(row=6, column=0, pady=5, sticky="w", padx=(10, 0))
check_ultimo_hito.grid(row=6, column=1, pady=5, sticky="w", padx=(10, 0))
check_informe_ac.grid(row=7, column=2, pady=5, sticky="w", padx=(10, 0))
check_informe_convenio.grid(row=7, column=0, pady=5, sticky="w", padx=(10, 0))
check_informe_liquidacion.grid(row=7, column=1, pady=5, sticky="w", padx=(10, 0))
check_adhesion_convenio.grid(row=8, column=2, pady=5, sticky="w", padx=(10, 0))
check_comunicacion_creditos.grid(row=8, column=0, pady=5, sticky="w", padx=(10, 0))
check_comunicacion_cuenta.grid(row=8, column=1, pady=5, sticky="w", padx=(10, 0))
check_textos_definitivos.grid(row=9, column=2, pady=5, sticky="w", padx=(10, 0))
check_certificado_defuncion.grid(row=9, column=0, pady=5, sticky="w", padx=(10, 0))
check_notas_simples.grid(row=9, column=1, pady=5, sticky="w", padx=(10, 0))
check_burofax.grid(row=10, column=2, pady=5, sticky="w", padx=(10, 0))

tk.Label(ventana, text="A. ERC Asset Manager").grid(row=14, column=0, pady=5, sticky="e", padx=(10, 0))
campo_a_erc_asset_manager.grid(row=14, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="A. ERC Lawyer").grid(row=15, column=0, pady=5, sticky="e", padx=(10, 0))
campo_a_erc_lawyer.grid(row=15, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="B. ERC Start Date Asset Manager").grid(row=16, column=0, pady=5, sticky="e", padx=(10, 0))
campo_b_erc_start_date_asset_manager.grid(row=16, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="B. ERC Meses Estimados Recobro Asset Manager").grid(row=17, column=0, pady=5, sticky="e", padx=(10, 0))
campo_b_erc_meses_estimados_recobro_asset_manager.grid(row=17, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="C. ERC Start Date Lawyer").grid(row=18, column=0, pady=5, sticky="e", padx=(10, 0))
campo_c_erc_start_date_lawyer.grid(row=18, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="C. ERC Meses Estimados Recobro Lawyer").grid(row=19, column=0, pady=5, sticky="e", padx=(10, 0))
campo_c_erc_meses_estimados_recobro_lawyer.grid(row=19, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="D. Asset Comments").grid(row=20, column=0, pady=5, sticky="e", padx=(10, 0))
campo_d_asset_comments.grid(row=20, column=1, pady=5, sticky="w", padx=(0, 10))

tk.Label(ventana, text="D. Lawyer Comments").grid(row=21, column=0, pady=5, sticky="e", padx=(10, 0))
campo_d_lawyer_comments.grid(row=21, column=1, pady=5, sticky="w", padx=(0, 10))

# Crear un Frame para contener los botones
frame_botones = tk.Frame(ventana)
frame_botones.grid(row=0, column=0, columnspan=4, pady=10)  # Posicionar el Frame en la parte superior

# Definir el ancho de los botones para que sean del mismo tamaño
ancho_boton = 20  # Ajusta este valor según sea necesario para que todos los botones tengan el mismo tamaño

# Posicionar los botones en el Frame usando grid y establecer el mismo ancho
boton_guardar_datos.grid(row=14, column=4, padx=1, pady=5)
boton_guardar_datos.config(width=ancho_boton)  # Establecer ancho del botón

boton_exportar_datos.grid(row=15, column=4, padx=1, pady=5)
boton_exportar_datos.config(width=ancho_boton)  # Establecer ancho del botón

boton_borrar_registro.grid(row=16, column=4, padx=1, pady=5)
boton_borrar_registro.config(width=ancho_boton)  # Establecer ancho del botón

boton_calendarios.grid(row=17, column=4, padx=1, pady=5)
boton_calendarios.config(width=ancho_boton)  # Establecer ancho del botón

# Posicionar otros widgets en la ventana usando grid
boton_cargar_archivo.grid(row=1, column=2, pady=5, sticky="w", padx=(0, 10))

# Crear la tabla para mostrar los datos
columns = ["CONTRACT ID", "SegmentI", "CIF", "Name", "Default_Date", "Product", "Individual", "NumIntervinientes", "Procedimiento1", "Class_Litigation_New", "InitialAmount", "TOTAL_ERCS_NBO", "DD_FLAG"]
tree = ttk.Treeview(ventana, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.grid(row=35, column=0, columnspan=3)

ventana.mainloop()