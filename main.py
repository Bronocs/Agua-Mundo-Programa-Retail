import gspread
from gspread_formatting import CellFormat, NumberFormat, format_cell_range
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from tkinter import filedialog, messagebox
import unicodedata
import io
import PyPDF2
import unicodedata
import os
import re
import os.path
import sys
import webbrowser
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import re
from collections import defaultdict
from tkinter import *
from tkcalendar import DateEntry
from datetime import datetime
from pathlib import Path


# 🔹 1️⃣ Autenticación con Google Drive
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# 🔹 3️⃣ Configuración de autenticación con OAuth
CREDENTIALS_FILE = "credentials.json"  # Ruta del archivo JSON descargado

APP_NAME = "Aguamundo"

appdata_auth_dir = Path(os.getenv("APPDATA")) / APP_NAME / "auth"
appdata_auth_dir.mkdir(parents=True, exist_ok=True)
token_path = appdata_auth_dir / "token.json"

temp_dir = Path(os.getenv("APPDATA")) / "Aguamundo" / "temp"
temp_dir.mkdir(parents=True, exist_ok=True)

# Si ya se autenticó antes, usa el token guardado
if os.path.exists(token_path):
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
else:
    # Inicia el flujo de autenticación OAuth 2.0
    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)  # Abre una ventana para autorizar
    # Guarda las credenciales para futuras ejecuciones
    with open(token_path, "w") as token:
        token.write(creds.to_json())

drive_service = build("drive", "v3", credentials=creds)

def actualizar_data():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    CREDENTIALS_FILE = "credentials.json"

    # 🔹 3️⃣ Configuración de autenticación con OAuth
    CREDENTIALS_FILE = "credentials.json"  # Ruta del archivo JSON descargado

    # Si ya se autenticó antes, usa el token guardado
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        # Inicia el flujo de autenticación OAuth 2.0
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)  # Abre una ventana para autorizar
        # Guarda las credenciales para futuras ejecuciones
        with open(token_path, "w") as token:
            token.write(creds.to_json())    
    service = build("sheets", "v4", credentials=creds)

    SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"  # Reemplaza con el ID de tu hoja de cálculo
    RANGE = "Registro ventas 2.0"  

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    data = result.get("values", [])

    max_columnas = max(len(fila) for fila in data)
    data_normalizada = [fila + [""] * (max_columnas - len(fila)) for fila in data]
    data = data_normalizada

    return(data)

def actualizar_data_productos():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    CREDENTIALS_FILE = "credentials.json"

    # 🔹 3️⃣ Configuración de autenticación con OAuth
    CREDENTIALS_FILE = "credentials.json"  # Ruta del archivo JSON descargado

    # Si ya se autenticó antes, usa el token guardado
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        # Inicia el flujo de autenticación OAuth 2.0
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)  # Abre una ventana para autorizar
        # Guarda las credenciales para futuras ejecuciones
        with open(token_path, "w") as token:
            token.write(creds.to_json())    
    service = build("sheets", "v4", credentials=creds)

    SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"  # Reemplaza con el ID de tu hoja de cálculo
    RANGE = "Ofi_skus"  

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    data = result.get("values", [])

    max_columnas = max(len(fila) for fila in data)
    data_normalizada = [fila + [""] * (max_columnas - len(fila)) for fila in data]
    data = data_normalizada

    return(data)

def actualizar_data_clientes():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    CREDENTIALS_FILE = "credentials.json"

    # 🔹 3️⃣ Configuración de autenticación con OAuth
    CREDENTIALS_FILE = "credentials.json"  # Ruta del archivo JSON descargado

    # Si ya se autenticó antes, usa el token guardado
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        # Inicia el flujo de autenticación OAuth 2.0
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)  # Abre una ventana para autorizar
        # Guarda las credenciales para futuras ejecuciones
        with open(token_path, "w") as token:
            token.write(creds.to_json())    
    service = build("sheets", "v4", credentials=creds)

    SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"  # Reemplaza con el ID de tu hoja de cálculo
    RANGE = "Clon_clientes_proves"  

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    data = result.get("values", [])

    max_columnas = max(len(fila) for fila in data)
    data_normalizada = [fila + [""] * (max_columnas - len(fila)) for fila in data]
    data = data_normalizada

    return(data)

def actualizar_data_ventas():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    CREDENTIALS_FILE = "credentials.json"

    # 🔹 3️⃣ Configuración de autenticación con OAuth
    CREDENTIALS_FILE = "credentials.json"  # Ruta del archivo JSON descargado

    # Si ya se autenticó antes, usa el token guardado
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        # Inicia el flujo de autenticación OAuth 2.0
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)  # Abre una ventana para autorizar
        # Guarda las credenciales para futuras ejecuciones
        with open(token_path, "w") as token:
            token.write(creds.to_json())    
    service = build("sheets", "v4", credentials=creds)

    SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"  # Reemplaza con el ID de tu hoja de cálculo
    RANGE = "Venta producto 2.0"  

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
    data = result.get("values", [])

    max_columnas = max(len(fila) for fila in data)
    data_normalizada = [fila + [""] * (max_columnas - len(fila)) for fila in data]
    data = data_normalizada

    return(data)

def coordenada_a_rango(nombre_hoja, fila, columna):
    # Convierte número de columna (ej: 2) a letra (ej: B)
    letra_columna = ''
    while columna > 0:
        columna, resto = divmod(columna - 1, 26)
        letra_columna = chr(65 + resto) + letra_columna
    return f"{nombre_hoja}!{letra_columna}{fila}"

def descargar_pdf(file_id, nombre_salida):
    request = drive_service.files().get_media(fileId=file_id)
    file = io.BytesIO()
    downloader = MediaIoBaseDownload(file, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    
    with open(nombre_salida, "wb") as f:
        f.write(file.getvalue())
    
    print(f"Archivo descargado: {nombre_salida}")

def abrir_enlace(event):
    item = tree.selection()  # Obtener el item seleccionado
    if item:
        enlace = tree.item(item, "values")[2]  # Obtener el valor de la columna Link
        if enlace:
            webbrowser.open(enlace)  # Abrir el enlace en el navegador

def extraer_id_drive(url):
    # Expresión regular para extraer el ID de un enlace de Google Drive
    patron = r"(?:/d/|id=)([a-zA-Z0-9_-]{10,})"
    coincidencia = re.search(patron, url)
    
    if coincidencia:
        return coincidencia.group(1)
    else:
        return None  # No se encontró un ID válido


# 🔹 1️⃣ Función para quitar tildes
def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

# 🔹 2️⃣ Función para buscar coincidencias
def buscar_coincidencias(lista, patron):
    patron = quitar_tildes(patron.lower())  # Convertir patrón a minúsculas y eliminar tildes
    return [(i, elemento) for i, elemento in enumerate(lista) if patron in quitar_tildes(elemento.lower())]

# 🔹 5️⃣ Obtener el enlace de descarga del archivo subido
def obtener_link(file_id):
    file = drive_service.files().get(fileId=file_id, fields="webViewLink").execute()
    return file.get("webViewLink")

# 🔹 3️⃣ Unir PDFs con PyPDF2
def combinar_pdfs(archivos, salida):
    pdf_merger = PyPDF2.PdfMerger()
    for archivo in archivos:
        pdf_merger.append(archivo)
    pdf_merger.write(salida)
    pdf_merger.close()
    print(f"PDF combinado: {salida}")



# 🔹 4️⃣ Subir el PDF combinado a Google Drive
def subir_pdf(nombre_archivo, carpeta_id=None):
    file_metadata = {"name": nombre_archivo}
    if carpeta_id:
        file_metadata["parents"] = [carpeta_id]
    
    media = MediaFileUpload(nombre_archivo, mimetype="application/pdf")
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    print(f"Archivo subido a Drive con ID: {file.get('id')}")
    return file.get("id")

def mostrar_lista():
    # Ocultar el menú principal
    frame_inicio.pack_forget()

    # Mostrar el frame con la lista
    frame_general_lista.pack(fill="both", expand=True, padx=20, pady=10)

def mostrar_Registro_OC(orden_lista):
    # Ocultar el menú principal
    frame_inicio.pack_forget()

    # Mostrar el frame con la lista
    Registro_OC.pack(fill="both", expand=True, padx=20, pady=10)
    data_orden = actualizar_data()
    orden = int(data_orden[len(data_orden)-1][0]) + 1
    orden_lista.append(orden)

def volver_al_inicio():
    # Ocultar la lista
    frame_general_lista.pack_forget()

    # Mostrar el menú principal
    frame_inicio.pack(fill="both", expand=True)

def volver_al_inicio_Registro_OC():
    # Ocultar la lista
    Registro_OC.pack_forget()

    # Mostrar el menú principal
    frame_inicio.pack(fill="both", expand=True)

def refrescar_datos(data, tree):
    # Limpiar el tree actual
    for fila in tree.get_children():
        tree.delete(fila)
    # Volver a cargar los datos
    # Agrega elementos (simulados)
    for i in range(len(data)):
        if data[i][6] == 'pend':
            tree.insert("", "end", values=(data[i][10], data[i][7], data[i][30]))

def obtener_links_consolidados_por_oc(
    datos,
    columna_oc=10,
    columna_verificar=11,
    columna_links=21,
    columna_link_oc=14
):
    """
    Busca patrones tipo XdeY en la columna de verificación, y extrae links únicos por OC.
    También combina todos los links asociados con la OC en una sola lista.
    
    - columna_link_oc: si se especifica, es la columna que contiene el "link de la OC principal".
    """

    patron = re.compile(r'(\d)de(\d)')
    oc_dict = defaultdict(set)
    links_oc_principales = {}

    for fila in datos:
        if len(fila) <= max(columna_oc, columna_verificar, columna_links):
            continue

        oc = str(fila[columna_oc]).strip()
        campo_verificar = str(fila[columna_verificar]).lower()

        # Guardar el link de la OC si corresponde
        if columna_link_oc is not None and len(fila) > columna_link_oc:
            link_oc = str(fila[columna_link_oc]).strip()
            if link_oc:
                links_oc_principales[oc] = link_oc

        # Buscar coincidencias tipo XdeY
        matches = patron.findall(campo_verificar)
        if matches:
            link = str(fila[columna_links]).strip()
            if link:
                oc_dict[oc].add(link)

    # Construir resultado con link de OC + los links encontrados
    resultado = {}
    for oc, links in oc_dict.items():
        lista_final = []
        link_oc = links_oc_principales.get(oc)
        if link_oc:
            lista_final.append(link_oc)
        lista_final.extend(sorted(links))  # opcional: ordena para mantener consistencia
        resultado[oc] = lista_final

    archivos = []
    file_IDs = []
    excluir = []

    for i in resultado:        
        archivos = []
        file_IDs = []
        excluir = []
        for j in range(len(datos)):
            if i == datos[j][10] and (len(datos[j][14]) > 10) and (len(datos[j][21]) > 10) and (j not in excluir) and (len(datos[j][31]) == 0)and (datos[j][8] != "ANULADO") and (datos[j][columna_links] != "pend"):
                print(i)
                print(datos[j][10])
                excluir.append(j)
                archivos = []
                file_IDs = []
                for k in range(len(resultado[i])):
                    archivo_temp = temp_dir / f"archivo{k+1}.pdf"
                    archivos.append(str(archivo_temp))
                    file_IDs.append(extraer_id_drive(resultado[i][k]))
                    print(resultado[i][k])
                    descargar_pdf(extraer_id_drive(resultado[i][k]), str(archivo_temp))
                
                pdf_final = temp_dir / f"{datos[j][23]}.pdf"
                print("pdffinal", datos[j][23] + ".pdf")
                combinar_pdfs(archivos, str(pdf_final))

                ID_CARPETA_DESTINO = "1yAXsWHBok2K7yBAyz27Ec4bhSok_yGn2"
                id_pdf_subido = subir_pdf(str(pdf_final), ID_CARPETA_DESTINO)

                link_descarga = obtener_link(id_pdf_subido)
                print(f"✅ PDF combinado disponible en: {link_descarga}")

                service = build("sheets", "v4", credentials=creds)
                sheet = service.spreadsheets()
                SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"

                # Modificar la celda
                result = sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=coordenada_a_rango("Registro ventas 2.0", j+1, 32),
                    valueInputOption="RAW",
                    body={"values": [[link_descarga]]}
                ).execute()
                
    return resultado


def combinar_guias(creds):
    data = actualizar_data()
    filas = len(data)

    for i in range(1, filas):
        if len(data[i][30]) > 0:
            continue
        
        # IDs de los archivos en Drive (reemplaza con los tuyos)
        FILE_ID_1 = extraer_id_drive(data[i][14])
        FILE_ID_2 = extraer_id_drive(data[i][16])

        if data[i][4] != "SDX":
            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            # Modificar la celda
            result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=coordenada_a_rango("Registro ventas 2.0", i+1, 31),
            valueInputOption="RAW",
            body={"values": [[data[i][14]]]}
            ).execute()
            print("e")

            continue
        else:
            if (FILE_ID_1) == None:
                continue

            if (FILE_ID_2) == None:
                continue

            # Nombres temporales para los archivos descargados
            archivo1 = temp_dir / "archivo1.pdf"
            archivo2 = temp_dir / "archivo2.pdf"

            descargar_pdf(FILE_ID_1, str(archivo1))
            descargar_pdf(FILE_ID_2, str(archivo2))

            valor_modificado = data[i][10].replace("/", " ") if len(data[i]) > 10 else "resultado"
            pdf_final = temp_dir / f"{valor_modificado}.pdf"


            combinar_pdfs([str(archivo1), str(archivo2)], str(pdf_final))

            ID_CARPETA_DESTINO = "1yAXsWHBok2K7yBAyz27Ec4bhSok_yGn2"
            id_pdf_subido = subir_pdf(str(pdf_final), ID_CARPETA_DESTINO)

            link_descarga = obtener_link(id_pdf_subido)
            print(f"✅ PDF combinado disponible en: {link_descarga}")

            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()

            # Modificar la celda
            result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=coordenada_a_rango("Registro ventas 2.0", i+1, 31),
            valueInputOption="RAW",
            body={"values": [[link_descarga]]}
            ).execute() 

            print(f"✅ Celda {range} actualizada con éxito: {link_descarga[0][0]}")
            print(data[i][0])

def agregar_producto(productos_agregados, nombres_agregados, productos_lista, tree, seleccion):
    producto = seleccion.get()
    
    if producto:  # Verifica que haya algo seleccionado
        for i in range(len(productos_lista)):
            if productos_lista[i][1] == producto:
                productos_agregados.append(productos_lista[i][:])
                actualizar_lista_productos(productos_agregados, tree)

def eliminar_producto(productos_agregados, tree, nombres_agregados):
    seleccionado = list(tree.item(tree.selection(), "values"))
    for i in range(len(productos_agregados)):
        #print(seleccionado[0])
        #print(productos_agregados[i][0])
        print(productos_agregados)
        print(i)
        if (seleccionado[0] == productos_agregados[i][0]) and (seleccionado[1] == productos_agregados[i][1]):
            print(productos_agregados)
            print(seleccionado)
            productos_agregados.remove(seleccionado)
            actualizar_lista_productos(productos_agregados, tree)

def actualizar_lista_productos(productos_agregados,tree):
    print(productos_agregados)
    for fila in tree.get_children():
        tree.delete(fila)
    for i in range(len(productos_agregados)):
        productos_agregados[i][0] = str(i+1)
    # Luego inserta en el tree
    for i in range(len(productos_agregados)):
        tree.insert("", "end", values=tuple(productos_agregados[i]))

def debugin(productos_agregados):
    print(productos_agregados)

def mostrar_valor_spinbox(spinbox):
    cantidad = int(spinbox.get())
    print(cantidad)

def agregar_nOD(opcion_OD, tree_productos_agregados, productos_agregados):
    OD_seleccionado = opcion_OD.get()
    producto_seleccionado = list(tree_productos_agregados.item(tree_productos_agregados.selection(), "values"))
    print(opcion_OD)
    print(producto_seleccionado)
    print(productos_agregados)
    if producto_seleccionado: 
        print("a34") # Verifica que haya algo seleccionado
        for i in range(len(productos_agregados)):
            if (productos_agregados[i][0] == producto_seleccionado[0]) and (productos_agregados[i][1] == producto_seleccionado[1]):
                print("entro34")
                productos_agregados[i][5] = OD_seleccionado
                print(productos_agregados)
                actualizar_lista_productos(productos_agregados, tree_productos_agregados)

def agregar_cantidad(spinbox_cantidad, tree_productos_agregados, productos_agregados):
    OD_seleccionado = spinbox_cantidad.get()
    producto_seleccionado = list(tree_productos_agregados.item(tree_productos_agregados.selection(), "values"))
    print(OD_seleccionado)
    print(producto_seleccionado)
    print(productos_agregados)
    if producto_seleccionado: 
        print("a34") # Verifica que haya algo seleccionado
        for i in range(len(productos_agregados)):
            if (productos_agregados[i][0] == producto_seleccionado[0]) and (productos_agregados[i][1] == producto_seleccionado[1]):
                print("entro34")
                productos_agregados[i][4] = OD_seleccionado
                print(productos_agregados)
                actualizar_lista_productos(productos_agregados, tree_productos_agregados)

def agregar_precio(precio, tree_productos_agregados, productos_agregados):
    precio_seleccionado = precio.get()
    producto_seleccionado = list(tree_productos_agregados.item(tree_productos_agregados.selection(), "values"))
    print(precio_seleccionado)
    print(producto_seleccionado)
    print(productos_agregados)
    if producto_seleccionado: 
        print("a34") # Verifica que haya algo seleccionado
        for i in range(len(productos_agregados)):
            if (productos_agregados[i][0] == producto_seleccionado[0]) and (productos_agregados[i][1] == producto_seleccionado[1]):
                print("entro34")
                productos_agregados[i][6] = precio_seleccionado
                print(productos_agregados)
                actualizar_lista_productos(productos_agregados, tree_productos_agregados)

# 🔹 4️⃣ Conectarse a Google Sheets
service = build("sheets", "v4", credentials=creds)

SPREADSHEET_ID = "13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ"  # Reemplaza con el ID de tu hoja de cálculo
RANGE = "Registro ventas 2.0"  

sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
data = result.get("values", [])

max_columnas = max(len(fila) for fila in data)
data_normalizada = [fila + [""] * (max_columnas - len(fila)) for fila in data]
data = data_normalizada

nombres_agregados = []

# 2. Autenticación con OAuth2
def autenticar():
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        try:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
            with open(token_path, "w") as token:
                token.write(creds.to_json())
        except Exception as e:
            messagebox.showerror("Error de autenticación", str(e))
            sys.exit(1)

    return creds

# 3. Subir archivo a Drive
def subir_archivo_od(creds, folder_id, orden, nro_Orden_OC, Nro_despacho, total_despacho, abreviacion):
    # Seleccionar archivo local
    ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo" + "de OD N°" + Nro_despacho)
    if not ruta_archivo:
        return

    try:
        service = build("drive", "v3", credentials=creds)

        archivo_metadata = {
            "name": orden + ". OD " + nro_Orden_OC + "-" + Nro_despacho + "DE" + total_despacho + "-" + abreviacion,
            "parents": [folder_id]  # ID de la carpeta destino
        }

        media = MediaFileUpload(ruta_archivo, resumable=True)
        archivo = service.files().create(
            body=archivo_metadata,
            media_body=media,
            fields="id"
        ).execute()

        messagebox.showinfo("Éxito", f"Archivo subido con ID: {archivo.get('id')}")
        return f"https://drive.google.com/file/d/{archivo.get('id')}/view"

    except Exception as e:
        messagebox.showerror("Error al subir", str(e))

def subir_archivo_gr(creds, folder_id, orden, nro_Orden_OC, Nro_despacho, total_despacho, abreviacion):
    # Seleccionar archivo local
    ruta_archivo = filedialog.askopenfilename(title="Seleccionar archivo" + "de GR N°" + Nro_despacho)
    if not ruta_archivo:
        return

    try:
        service = build("drive", "v3", credentials=creds)

        archivo_metadata = {
            "name": orden + ". GR  " + nro_Orden_OC + "-" + Nro_despacho + "DE" + total_despacho + "-" + abreviacion,
            "parents": [folder_id]  # ID de la carpeta destino
        }

        media = MediaFileUpload(ruta_archivo, resumable=True)
        archivo = service.files().create(
            body=archivo_metadata,
            media_body=media,
            fields="id"
        ).execute()

        messagebox.showinfo("Éxito", f"Archivo subido con ID: {archivo.get('id')}")
        return f"https://drive.google.com/file/d/{archivo.get('id')}/view"

    except Exception as e:
        messagebox.showerror("Error al subir", str(e))        

def subir_archivo_oc(creds, folder_id, orden, nro_Orden_OC, abreviacion):
    # Seleccionar archivo local
    ruta_archivo = filedialog.askopenfilename(title="Seleccionar OC")
    if not ruta_archivo:
        return

    nombre_archivo = os.path.basename(ruta_archivo)

    try:
        service = build("drive", "v3", credentials=creds)

        archivo_metadata = {
            "name": str(orden) + ". OC " + str(nro_Orden_OC)+ "-" + str(abreviacion),
            "parents": [folder_id]  # ID de la carpeta destino
        }

        media = MediaFileUpload(ruta_archivo, resumable=True)
        archivo = service.files().create(
            body=archivo_metadata,
            media_body=media,
            fields="id"
        ).execute()

        messagebox.showinfo("Éxito", f"Archivo subido con ID: {archivo.get('id')}")
        return f"https://drive.google.com/file/d/{archivo.get('id')}/view"

    except Exception as e:
        messagebox.showerror("Error al subir", str(e))        

# 4. Interfaz básica con tkinter
def iniciar_app_oc(orden, nro_Orden_OC, cliente, data_cliente):
    creds = autenticar()
    abreviacion = ""
    for i in range(len(data_cliente)):
        if cliente == data_cliente[i][1]:
            abreviacion = data_cliente[i][3]

    # Reemplaza este ID con el de tu carpeta en Google Drive
    if cliente == "SERGIO JHONATAN INOX SAC":
        folder_id = "1ukvcaN6UyUDKpKqEePdqwYaHLVLXev0M"
    if cliente == "SERVICIOS PETROLEROS Y CONSTRUCCIONES SEPCON SA":
        folder_id = "14_QzsrNvEYSoX7WvJZ9LVsCzCVdttZps"
    if cliente == "INLAND ENERGY SAC":
        folder_id = "1vrx-z4t6hqMXdRQI9XvVtjEftSIsUswD"
    if cliente == "HIDROLED S.A.C.":
        folder_id = "1ohQKkGwnZP-FFDkW9NC4RaDo4L1D8KGB"
    if cliente == "PESQUERA HAYDUK SA":
        folder_id = "14ml5NB_uSFLaCASJW9FR_xihcyMQBa65"
    if cliente == "FERYMAR S.A.C":
        folder_id = "1vkflpVvqME3lpSg6HjnaBpBp4Qs6xfdk"
    if cliente == "ENERGOTEC SAC":
        folder_id = "13bAVH_qi9_JQx_6F7WXwR26GxGNEVxOB"
    if cliente == "CITRICOS PERUANOS S.A.":
        folder_id = "1S2N4TXXadf3DWGMSfGbVxFqfmIhrKxN6"
    if cliente == "AQUA EXPEDITIONS S.A.C.":
        folder_id = "1lDJzu62S3x6h9HKMTlVKkSVhm9oaNOyr"
    if cliente == "AJEPER DEL ORIENTE S.A":
        folder_id = "1mG_aCah6RrsZE_sAMYcKATJWmhFrsIxF"
    if cliente == "SODEXO PERU S.A.C.":
        folder_id = "1cu4OhoqhI6lQBNDPEOxFBISJxXYbdjje"
    else:
        folder_id = "1025e3tfwHBW_Q01qxNl1dNaBLBgcD-Ia"
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    link = subir_archivo_oc(creds, folder_id, orden, nro_Orden_OC, abreviacion)
    return link

def iniciar_app_od(orden, nro_Orden_OC, Nro_despacho, total_despacho, cliente, data_cliente):
    creds = autenticar()
    abreviacion = ""
    for i in range(len(data_cliente)):
        if cliente == data_cliente[i][1]:
            abreviacion = data_cliente[i][3]

    # Reemplaza este ID con el de tu carpeta en Google Drive
    if cliente == "SERGIO JHONATAN INOX SAC":
        folder_id = "1ukvcaN6UyUDKpKqEePdqwYaHLVLXev0M"
    if cliente == "SERVICIOS PETROLEROS Y CONSTRUCCIONES SEPCON SA":
        folder_id = "14_QzsrNvEYSoX7WvJZ9LVsCzCVdttZps"
    if cliente == "INLAND ENERGY SAC":
        folder_id = "1vrx-z4t6hqMXdRQI9XvVtjEftSIsUswD"
    if cliente == "HIDROLED S.A.C.":
        folder_id = "1ohQKkGwnZP-FFDkW9NC4RaDo4L1D8KGB"
    if cliente == "PESQUERA HAYDUK SA":
        folder_id = "14ml5NB_uSFLaCASJW9FR_xihcyMQBa65"
    if cliente == "FERYMAR S.A.C":
        folder_id = "1vkflpVvqME3lpSg6HjnaBpBp4Qs6xfdk"
    if cliente == "ENERGOTEC SAC":
        folder_id = "13bAVH_qi9_JQx_6F7WXwR26GxGNEVxOB"
    if cliente == "CITRICOS PERUANOS S.A.":
        folder_id = "1S2N4TXXadf3DWGMSfGbVxFqfmIhrKxN6"
    if cliente == "AQUA EXPEDITIONS S.A.C.":
        folder_id = "1lDJzu62S3x6h9HKMTlVKkSVhm9oaNOyr"
    if cliente == "AJEPER DEL ORIENTE S.A":
        folder_id = "1mG_aCah6RrsZE_sAMYcKATJWmhFrsIxF"
    if cliente == "SODEXO PERU S.A.C.":
        folder_id = "1cu4OhoqhI6lQBNDPEOxFBISJxXYbdjje"
    else:
        folder_id = "1025e3tfwHBW_Q01qxNl1dNaBLBgcD-Ia"

    # Reemplaza este ID con el de tu carpeta en Google Drive
    folder_id = "1025e3tfwHBW_Q01qxNl1dNaBLBgcD-Ia"

    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    link = subir_archivo_od(creds, folder_id, orden, nro_Orden_OC, Nro_despacho, total_despacho, abreviacion)
    return link

def iniciar_app_gr(orden, nro_Orden_OC, Nro_despacho, total_despacho, cliente, data_cliente):
    creds = autenticar()
    abreviacion = ""
    for i in range(len(data_cliente)):
        if cliente == data_cliente[i][1]:
            abreviacion = data_cliente[i][3]

    # Reemplaza este ID con el de tu carpeta en Google Drive
    folder_id = "13srmOQYQJy03UbhKc67X7KVp28NRcUSR"

    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    link = subir_archivo_gr(creds, folder_id, orden, nro_Orden_OC, Nro_despacho, total_despacho, abreviacion)
    return link

def mostrar_cliente(seleccion_cliente):
    print(seleccion_cliente.get())

def actualizar_desplegable_ods(contenedor_desplegable_ods, cantidad_ods, opcion_od):
    for widget in contenedor_desplegable_ods.winfo_children():
        widget.destroy()

    total = cantidad_ods.get()
    opciones = [str(num) for num in range(1, total + 1)]

    if opciones:
        opcion_od.set(opciones[0])
        label_menu = tk.Label(contenedor_desplegable_ods, text="Seleccionar N° de OD")
        label_menu.pack(side="left")
        menu = ttk.OptionMenu(contenedor_desplegable_ods, opcion_od, opciones[0], *opciones)
        menu.pack(side="left")
        #botón añadir nOD
        boton_nOD = ttk.Button(contenedor_desplegable_ods, text="Seleccionar n° OD", command=lambda:agregar_nOD(opcion_od, tree_productos_agregados, productos_agregados))
        boton_nOD.pack(pady=5)

# Botón para aplicar la cantidad de OD's
def aplicar_cantidad_ods(contenedor_desplegable_ods, ctd_ods, opcion_od):
    actualizar_desplegable_ods(contenedor_desplegable_ods, ctd_ods, opcion_od)

def abrir_enlace_od(url):
    import webbrowser
    webbrowser.open_new_tab(url)

def agregar_link_oc(orden, etiqueta, link_oc, nro_Orden_OC, cliente, data_cliente):
    #falta añadir link oc al texto que se enviará a la base de datos

    link = iniciar_app_oc(orden, nro_Orden_OC, cliente, data_cliente)
    etiqueta.config(text=link)
    etiqueta.bind("<Button-1>", lambda e, url=link: abrir_enlace_od(url))
    link_oc.clear
    link_oc.append(link)

def mostrar_links(orden, frame_links, cantidad_ods, links_od, nro_Orden_OC, total_despacho, cliente, data_cliente):

    for widget in frame_links.winfo_children():
        widget.destroy()


    enlaces = []
    links_od.clear()
    # Lista de enlaces (puedes cambiarla dinámicamente)
    for i in range(cantidad_ods.get()):
        enlaces.append(iniciar_app_od(str(orden), str(nro_Orden_OC), str(i+1), str(total_despacho), cliente, data_cliente))
        
    links_od.extend(enlaces)

    for i, enlace in enumerate(enlaces, start=1):
        fila = tk.Frame(frame_links)
        fila.pack(anchor="w", pady=2)

        # Número de orden
        numero = tk.Label(fila, text=f"{i}.", width=4, anchor="w")
        numero.pack(side="left")

        # Enlace como texto clickable
        etiqueta = tk.Label(fila, text=enlace, fg="blue", cursor="hand2", relief="solid", padx=5, pady=2)
        etiqueta.pack(side="left")
        etiqueta.bind("<Button-1>", lambda e, url=enlace: webbrowser.open_new_tab(url))

def mostrar_links_2(orden, frame_links, cantidad_ods, links_od, nro_Orden_OC, total_despacho, cliente, data_cliente):
    for widget in frame_links.winfo_children():
        widget.destroy()
    enlaces = []
    links_od.clear()

    # Lista de enlaces (puedes cambiarla dinámicamente)
    for i in range(cantidad_ods.get()):
        enlaces.append(iniciar_app_gr(str(orden), str(nro_Orden_OC), str(i+1), str(total_despacho), cliente, data_cliente))

    links_od.extend(enlaces)

    for i, enlace in enumerate(enlaces, start=1):
        fila = tk.Frame(frame_links)
        fila.pack(anchor="w", pady=2)

        # Número de orden
        numero = tk.Label(fila, text=f"{i}.", width=4, anchor="w")
        numero.pack(side="left")

        # Enlace como texto clickable
        etiqueta = tk.Label(fila, text=enlace, fg="blue", cursor="hand2", relief="solid", padx=5, pady=2)
        etiqueta.pack(side="left")
        etiqueta.bind("<Button-1>", lambda e, url=enlace: webbrowser.open_new_tab(url))

def col_num_a_letra(n):
    letra = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letra = chr(65 + rem) + letra
    return letra

def enviar_OC(data, fecha_registro, data_cliente, cliente, numero_oc, ods_totales, lista_tree, direccion_entrega, link_oc, links_od, links_gr, moneda):

    campos = {
        "data": data,
        "fecha_registro": fecha_registro,
        "data_cliente": data_cliente,
        "Cliente": cliente,
        "Número de OC": numero_oc,
        "ods_totales": ods_totales,
        "Lista de productos": lista_tree,
        "Dirección de entrega": direccion_entrega,
        "Link de OC": link_oc,
        "Links de OD": links_od,
        "Links de GR": links_gr
    }

    # Lista para guardar los campos que están vacíos
    campos_vacios = [nombre for nombre, valor in campos.items() if not valor]

    if campos_vacios:
        mensaje = "Los siguientes campos están vacíos o no fueron proporcionados:\n\n"
        mensaje += "\n".join(f"- {campo}" for campo in campos_vacios)
        messagebox.showerror("Error", mensaje)
        return  # Sale de la función

    for i in range(len(data_cliente)):
        if cliente == data_cliente[i][1]:
            abreviacion = data_cliente[i][3]
            ruc = data_cliente[i][0]

    fecha_hoy = datetime.today().strftime('%d/%m/%Y')
    orden = int(data[len(data) - 1][0]) + 1
    fila_elegida = len(data) + 1

    for i in range(int(ods_totales)):
        valores_personalizados = {
            1: int(orden),
            2: fecha_registro,
            3: fecha_hoy,
            4: cliente,
            5: abreviacion,
            6: ruc,
            11: numero_oc,
            12: str(i+1) + "DE" + str(ods_totales),
            13: direccion_entrega,
            15: link_oc[0],
            17: links_od[i],
            20: links_gr[i],
        }
        for col, valor in valores_personalizados.items():
            celda = f"Registro ventas 2.0!{col_num_a_letra(col)}{fila_elegida}"
            sheet.values().update(
                spreadsheetId="13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ",
                range=celda,
                valueInputOption="USER_ENTERED",
                body={"values": [[valor]]}
            ).execute()
        fila_elegida += 1

        #combinar_guias(creds)

    fila_elegida = len(actualizar_data_ventas()) + 1

    for i in range(len(lista_tree)):
        valores_personalizados = {
            1: str(orden) + ". OC " + numero_oc+"-"+lista_tree[i][5]+"DE"+str(ods_totales)+"-"+abreviacion,
            3: link_oc[0],
            5: lista_tree[i][1],
            6: lista_tree[i][4],
            7: lista_tree[i][3],
            8: lista_tree[i][6],
            10: str(moneda),
        }
        for col, valor in valores_personalizados.items():
            celda = f"Venta producto 2.0!{col_num_a_letra(col)}{fila_elegida}"
            sheet.values().update(
                spreadsheetId="13ariJ1CCsqLFUIqHgaf7tuJ6Y5WM2IoMvt8EOcFWZNQ",
                range=celda,
                valueInputOption="USER_ENTERED",
                body={"values": [[valor]]}
            ).execute()
        fila_elegida += 1

    """

    ### Luego editar Venta producto 2.0
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    # Modificar la celda
    result = sheet.values().update(
    spreadsheetId=SPREADSHEET_ID,
    range=coordenada_a_rango("Registro ventas 2.0", i+1, 31),
    valueInputOption="RAW",
    body={"values": [[link_descarga]]}
    ).execute()"""

# 🔹 5️⃣ Procesar los datos
claves = data[0]  # Encabezados (primera fila)
filas = len(data)  # Datos (resto de las filas)

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Agua Mundo SAC")
ventana.state("zoomed")

########### Frame de inicio

frame_inicio = tk.Frame(ventana)

orden_lista = []

# Etiqueta
tk.Label(frame_inicio, text="Aguamundo", font=("Helvetica", 16)).pack(pady=20)

# Primer botón
tk.Button(frame_inicio, text="Ver Órdenes pendientes", command=mostrar_lista).pack(pady=(0, 20))  # Espacio debajo del primer botón

tk.Button(frame_inicio, text="Combinar OC y OD", command=lambda: combinar_guias(creds)).pack(pady=(20, 0))  # Espacio encima del segundo botón

# Segundo botón
tk.Button(frame_inicio, text="Combinar Orden y Guía sellada", command=lambda: obtener_links_consolidados_por_oc(actualizar_data())).pack(pady=(20, 0))  # Espacio encima del segundo botón

tk.Button(frame_inicio, text="Agregar OC",command=lambda: mostrar_Registro_OC(orden_lista)).pack(pady=(20, 0))  # Espacio encima del segundo botón


frame_inicio.pack(fill="both", expand=True)

############# Frame lista

frame_general_lista = tk.Frame(ventana)

# 🟨 Título
titulo = tk.Label(frame_general_lista, text="Órdenes de compra pendientes", font=("Helvetica", 16, "bold"), bg="lightgray")
titulo.pack(pady=(10, 5), side = "top")

texto = tk.Label(frame_general_lista, text = "Haz doble click en la orden para acceder al pdf", font=("Helvetica", 10, "bold"), bg="lightgray")
texto.pack(pady=(10, 5), side = "top")

# 🟦 Frame contenedor del Listbox que se expande
frame_lista = tk.Frame(frame_general_lista)
frame_lista.pack(fill = "both", expand=True)

# 🔷 Estilo para Treeview con líneas
style = ttk.Style()
style.configure("Treeview", 
    rowheight=25,
    font=("Helvetica", 10),
    borderwidth=1,
    relief="solid")
style.configure("Treeview.Heading", 
    font=("Helvetica", 10, "bold"),
    borderwidth=1,
    relief="solid")
style.layout("Treeview", [
    ('Treeview.treearea', {'sticky': 'nswe'})
])

tree = ttk.Treeview(frame_lista, columns=("Número de orden", "Prioridad", "Link al pdf"), show="headings")
tree.pack(pady = 20,fill = "both", expand=True)

# Definir las columnas
tree.heading("Número de orden", text="Número de orden")
tree.heading("Prioridad", text="Prioridad")
tree.heading("Link al pdf", text="Link al pdf")

# Configurar las columnas
tree.column("Número de orden", width=150)
tree.column("Prioridad", width=150)
tree.column("Link al pdf", width=150)

# 🔽 Scrollbar
# Crear una barra de desplazamiento vinculada al Listbox
scrollbar = tk.Scrollbar(tree, orient="vertical", command=tree.yview)
scrollbar.pack(side="right", fill="y")
tree.config(yscrollcommand=scrollbar.set)

# Agrega elementos (simulados)
for i in range(len(data)):
    if data[i][6] == 'pend':
        tree.insert("", "end", values=(data[i][10], data[i][7], data[i][30]))

Botones = tk.Frame(frame_general_lista) 
Botones.pack(fill="x", pady=10, anchor="center")
tk.Button(Botones, text="Actualizar", command=lambda: refrescar_datos(actualizar_data(), tree)).pack(side = "left", fill = "x")
tk.Button(Botones, text="Volver", command=volver_al_inicio).pack(side = "right", fill = "x")

# Asociar el evento de clic
tree.bind("<Double-1>", abrir_enlace)  # Al hacer doble clic en una fila se abrirá el enlace

############ Frame de registro de OC's

Registro_OC = tk.Frame(ventana)

data_orden = actualizar_data()
orden = int(data_orden[len(data_orden)-1][0]) + 1

productos = []
Clientes = []

data_clientes = actualizar_data_clientes()
data_productos = actualizar_data_productos()

for i in range(len(data_productos)):
    if i > 2:
        productos.append(["", data_productos[i][5], data_productos[i][9], data_productos[i][2], "", "", ""])

for i in range(len(data_clientes)):
    if (i > 4) and (data_clientes[i][1] != ""):
        Clientes.append(data_clientes[i][1])


links_od = []
links_gr = []

# Título
titulo_Registro_OC = tk.Label(Registro_OC, text="Agregar orden de compra", font=("Helvetica", 16, "bold"), bg="lightgray")
titulo_Registro_OC.pack(pady=(10, 5), side = "top")

botones_arriba = tk.Frame(Registro_OC)
botones_arriba.pack(side="bottom")

#### Frame primera fila Registro de OC's

primera_fila = tk.Frame(Registro_OC)
primera_fila.pack(side="top")

extra_fila = tk.Frame(Registro_OC)
extra_fila.pack(side="top")

segunda_fila = tk.Frame(Registro_OC)
segunda_fila.pack(side="top")

tercera_fila = tk.Frame(Registro_OC)
tercera_fila.pack(side="top")

## Frame contenedor  fecha
frame_fecha_registro = tk.Frame(primera_fila)
frame_fecha_registro.pack(pady=10, side="left", padx=10)

# Etiqueta a la izquierda
tk.Label(frame_fecha_registro, text="Fecha de registro:").pack(side="left", padx=(0, 10))

# Campo de entrada a la derecha
fecha_seleccionada = DateEntry(frame_fecha_registro, date_pattern='dd/mm/yyyy')  # Puedes usar 'yyyy-mm-dd' si prefieres
fecha_seleccionada.pack(pady=5)

#Para leer el valor aplico fecha_seleccionada.get()

## fin frame contenedor fecha

## Frame contenedor horizontal NRO OC
frame_oc = tk.Frame(primera_fila)
frame_oc.pack(pady=10, side="left", padx=10)

# Etiqueta a la izquierda
tk.Label(frame_oc, text="Número de OC:").pack(side="left", padx=(0, 10))

# Campo de entrada a la derecha
entrada_oc = tk.Entry(frame_oc, width=30)
entrada_oc.pack(side="left")
#Para leer el valor aplico entrada_oc.get()

## fin frame contenedor horizontal NRO OC

##### frame para cantidad de ods

# Variable para la cantidad total de OD's
cantidad_ods = tk.IntVar(value=1)

frame_cantidad_ods = tk.Frame(primera_fila)
frame_cantidad_ods.pack(pady=10, side="left", padx=10)

tk.Label(frame_cantidad_ods, text="Cantidad de OD's:").pack(pady=(10, 0), side="left")
selector_cantidad = tk.Spinbox(frame_cantidad_ods, from_=1, to=100, textvariable=cantidad_ods, width=5)
selector_cantidad.pack(side="left")

# Variable para la opción seleccionada del desplegable
opcion_od = tk.StringVar()

# Frame contenedor del desplegable
contenedor_desplegable_ods = tk.Frame(segunda_fila)
contenedor_desplegable_ods.pack(pady=10, side="left")

tk.Button(frame_cantidad_ods, text="Agregar cantidad de OD's", command=lambda:aplicar_cantidad_ods(contenedor_desplegable_ods, cantidad_ods, opcion_od)).pack(pady=5, side="left")

#opcion_od es la cantidad seleccionada, debemos usar esta variable para añadir este valor a la columna n°OD
#cantidad ods es la cantidad total

##### fin de frame para cantidad de OD's

## Frame contenedor horizontal direccion
frame_direccion = tk.Frame(primera_fila)
frame_direccion.pack(pady=10, side="left", padx=10)

# Etiqueta a la izquierda
tk.Label(frame_direccion, text="Dirección de entrega:").pack(side="left")

# Campo de entrada a la derecha
entrada_direccion = tk.Entry(frame_direccion, width=30)
entrada_direccion.pack(side="left")

## fin contenedor horizontal direccion

#### fin primera fila

#### Frame extra fila

## Desplegable clientes

frame_clientes = tk.Frame(extra_fila)
frame_clientes.pack(pady=10, side="left", padx=10)

# Etiqueta a la izquierda
tk.Label(frame_clientes, text="Cliente:").pack(side="left", padx=(0, 10))

# Variable de selección del desplegable
seleccion_cliente = tk.StringVar()

# Combobox editable
combobox_clientes = ttk.Combobox(frame_clientes, textvariable=seleccion_cliente)
combobox_clientes['values'] = Clientes
combobox_clientes.config(width=60)  
combobox_clientes.pack(side="left", pady=10)

# Habilita la escritura (editable)
combobox_clientes.configure(state="normal")

# Función de filtrado sin perder el foco
def autocompletar(event):
    texto = seleccion_cliente.get().lower()
    coincidencias = [c for c in Clientes if texto in c.lower()]

    # Evita que se borren todas si no hay coincidencia
    combobox_clientes['values'] = coincidencias if coincidencias else ["(sin coincidencias)"]
    
    # No despliega el menú automáticamente (evita perder el foco)

# Asocia a cada pulsación de tecla
combobox_clientes.bind('<KeyRelease>', autocompletar)

## fin desplegable clientes

### Frame contenedor seleccion de productos

frame_seleccion_producto = tk.Frame(extra_fila)
frame_seleccion_producto.pack(pady=10, side="left", padx=10)

# Variable de selección del desplegable
seleccion = tk.StringVar()

productos_nombre = []

for i in range(len(productos)):
    productos_nombre.append(productos[i][1])

# titulo seleccionar producto
titulo_seleccionar_producto = tk.Label(frame_seleccion_producto, text="Selecciona el producto: ")
titulo_seleccionar_producto.pack(side="left", pady=10, padx=10)

# Combobox editable
combobox_productos = ttk.Combobox(frame_seleccion_producto, textvariable=seleccion)
combobox_productos['values'] = productos_nombre
combobox_productos.config(width=100)
combobox_productos.pack(side="left", pady=10, padx=10)

# Permitir edición
combobox_productos.configure(state="normal")

# Función de autocompletado
def autocompletar_productos(event):
    texto = seleccion.get().lower()
    coincidencias = [nombre for nombre in productos_nombre if texto in nombre.lower()]
    combobox_productos['values'] = coincidencias if coincidencias else ["(sin coincidencias)"]

# Asociar al evento de escritura
combobox_productos.bind('<KeyRelease>', autocompletar_productos)

# Botón para agregar producto
boton_agregar_producto = ttk.Button(frame_seleccion_producto, text="Agregar producto", command=lambda:agregar_producto(productos_agregados, nombres_agregados, productos, tree_productos_agregados, seleccion))
boton_agregar_producto.pack(pady=10, padx= 10, side="left")

### Fin Frame seleccion de producto


#### fin extra fila

#### Frame segunda fila

### Frame agregar cantidad a producto

frame_agregar_cantidad = tk.Frame(segunda_fila)
frame_agregar_cantidad.pack(pady=10, side="left", padx=10)

### titulo añadir cantidad
titulo_agregar_cantidad = tk.Label(frame_agregar_cantidad, text="Selecciona la cantidad: ")
titulo_agregar_cantidad.pack(side="left", pady=10, padx=10)

### spinbox cantidad de productos
spinbox_cantidad = tk.Spinbox(frame_agregar_cantidad, from_=1, to=100, width=5)
spinbox_cantidad.pack(side= "left", pady=10, padx=10)

### boton añadir cantidad
boton_cantidad = ttk.Button(frame_agregar_cantidad, text="Añadir cantidad", command=lambda:agregar_cantidad(spinbox_cantidad, tree_productos_agregados, productos_agregados))
boton_cantidad.pack(side="left", pady = 10, padx = 10)

### Fin frame agregar cantidad a producto

#Boton para eliminar producto
boton_eliminar_producto = ttk.Button(segunda_fila, text="Eliminar producto", command=lambda:eliminar_producto(productos_agregados, tree_productos_agregados, nombres_agregados))
boton_eliminar_producto.pack(side="left",pady=10, padx= 10)

### Frame link OC

frame_agregar_oc = tk.Frame(segunda_fila)
frame_agregar_oc.pack(pady=10, side="left", padx=10)

link_oc = []

# Etiqueta link OC
titulo_agregar_oc = tk.Label(frame_agregar_oc, text="Link OC: ")
titulo_agregar_oc.pack(side="left", pady=10, padx=10)

# recuadro link OC
etiqueta_oc = tk.Label(frame_agregar_oc, text="", relief="solid",fg="blue", cursor="hand2", width=30, height=1)
etiqueta_oc.pack(pady=10, padx = 10, side="left")

# boton link OC
tk.Button(frame_agregar_oc, text="Subir OC", command=lambda:agregar_link_oc(orden, etiqueta_oc, link_oc, entrada_oc.get(), seleccion_cliente.get(), data_clientes)).pack(pady=10, padx = 10, side="left")

#### Fin segunda fila

#### Frame tercera fila

### Frame tipo de moneda:

frame_moneda = tk.Frame(tercera_fila)
frame_moneda.pack(pady=10, side="left", padx=10)

# Variable para almacenar la opción seleccionada
opcion_moneda = tk.StringVar()
opcion_moneda.set("d")  # Valor predeterminado

titulo_agregar_od = tk.Label(frame_moneda, text="Tipo de moneda: ")
titulo_agregar_od.pack(side="left", pady=10, padx=10)

# Crear el desplegable
desplegable = tk.OptionMenu(frame_moneda, opcion_moneda, "d", "s")
desplegable.pack()

### Fin frame tipo de moneda

### Frame agregar precio

frame_agregar_precio = tk.Frame(tercera_fila)
frame_agregar_precio.pack(pady=10, side="left", padx=10)

### titulo agregar precio
titulo_agregar_precio = tk.Label(frame_agregar_precio, text="Indica el precio: ")
titulo_agregar_precio.pack(side="left", pady=10, padx=10)

# Campo de entrada precio
precio = tk.Entry(frame_agregar_precio, width=7)
precio.pack(side="left")

### boton añadir cantidad
boton_precio = ttk.Button(frame_agregar_precio, text="Añadir precio", command=lambda:agregar_precio(precio, tree_productos_agregados, productos_agregados))
boton_precio.pack(side="left", pady = 10, padx = 10)

### Fin frame agregar precio

### Frame agregar OD
frame_agregar_od = tk.Frame(tercera_fila)
frame_agregar_od.pack(pady=10, side="left", padx=10)

### Frame titulo y boton + links
frame_agregar_od_arriba = tk.Frame(frame_agregar_od)
frame_agregar_od_arriba.pack(pady=10, side="top", padx=10)

# Etiqueta links OD
titulo_agregar_od = tk.Label(frame_agregar_od_arriba, text="Links OD: ")
titulo_agregar_od.pack(side="left", pady=10, padx=10)

# Marco donde se colocarán los recuadros de links OD
frame_links = tk.Frame(frame_agregar_od)
frame_links.pack(side="top")

# Boton agregar OD
tk.Button(frame_agregar_od_arriba, text="Subir OD", command=lambda:mostrar_links(orden, frame_links, cantidad_ods, links_od, entrada_oc.get(), cantidad_ods.get(),seleccion_cliente.get(), data_clientes)).pack(pady=10, padx = 10, side="left")

# Marco donde se colocarán los recuadros de links OD
frame_links = tk.Frame(frame_agregar_od)
frame_links.pack(side="top")

### Frame agregar GR
frame_agregar_gr = tk.Frame(tercera_fila)
frame_agregar_gr.pack(pady=10, side="left", padx=10)

### Frame titulo y boton + links
frame_agregar_gr_arriba = tk.Frame(frame_agregar_gr)
frame_agregar_gr_arriba.pack(pady=10, side="top", padx=10)

# Etiqueta links GR
titulo_agregar_gr = tk.Label(frame_agregar_gr_arriba, text="Links GR: ")
titulo_agregar_gr.pack(side="left", pady=10, padx=10)

# Marco donde se colocarán los recuadros de links GR
frame_links_2 = tk.Frame(frame_agregar_gr)
frame_links_2.pack(side="top")

# Boton agregar GR
tk.Button(frame_agregar_gr_arriba, text="Subir GR", command=lambda:mostrar_links_2(orden, frame_links_2, cantidad_ods, links_gr, entrada_oc.get(), cantidad_ods.get(),seleccion_cliente.get(), data_clientes)).pack(pady=5)

#### fin frame tercera fila

# 🟦 Frame contenedor del Listbox que se expande
frame_productos_agregados = tk.Frame(Registro_OC)
frame_productos_agregados.pack(fill = "both", expand=True, side = "top")

# 🔷 Estilo para Treeview con líneas
style_productos_agregados = ttk.Style()
style_productos_agregados.configure("Treeview", 
    rowheight=25,
    font=("Helvetica", 10),
    borderwidth=1,
    relief="solid")
style_productos_agregados.configure("Treeview.Heading", 
    font=("Helvetica", 10, "bold"),
    borderwidth=1,
    relief="solid")
style_productos_agregados.layout("Treeview", [
    ('Treeview.treearea', {'sticky': 'nswe'})  
])

tree_productos_agregados = ttk.Treeview(frame_productos_agregados, columns=("N°", "Nombre", "Codigo", "Unidad de Medición", "Cantidad", "n° OD", "Precio"), show="headings")
tree_productos_agregados.pack(pady = 20,fill = "both", expand=True)

# Definir las columnas
tree_productos_agregados.heading("N°", text="N°")
tree_productos_agregados.heading("Nombre", text="Nombre")
tree_productos_agregados.heading("Codigo", text="Codigo")
tree_productos_agregados.heading("Unidad de Medición", text="Unidad de Medición")
tree_productos_agregados.heading("Cantidad", text="Cantidad")
tree_productos_agregados.heading("n° OD", text="n° OD")
tree_productos_agregados.heading("Precio", text="Precio")

# Configurar las columnas
tree_productos_agregados.column("N°", width=80, stretch=False)
tree_productos_agregados.column("Nombre", width=10, stretch=True)
tree_productos_agregados.column("Codigo", width=200, stretch=False)
tree_productos_agregados.column("Unidad de Medición", width=150, stretch=False)
tree_productos_agregados.column("Cantidad", width=100, stretch=False)
tree_productos_agregados.column("n° OD", width=100, stretch=False)
tree_productos_agregados.column("Precio", width=200, stretch=False)


# 🔽 Scrollbar
# Crear una barra de desplazamiento vinculada al Listbox
scrollbar_productos_agregados = tk.Scrollbar(tree_productos_agregados, orient="vertical", command=tree_productos_agregados.yview)
scrollbar_productos_agregados.pack(side="right", fill="y")
tree_productos_agregados.config(yscrollcommand=scrollbar_productos_agregados.set)

productos_agregados = []

# Agrega elementos (simulados)
for i in range(len(productos_agregados)):
    tree_productos_agregados.insert("", "end", values=(productos_agregados[i][0], productos_agregados[i][1], productos_agregados[i][2],))

### Frame contenedor botones abajo

### Frame agregar GR
frame_botones_derecha= tk.Frame(Registro_OC)
frame_botones_derecha.pack(pady=10, side="bottom")

tk.Button(frame_botones_derecha, text="Volver", command=volver_al_inicio_Registro_OC).pack(side = "left", fill = "x", padx = 10)

tk.Button(frame_botones_derecha, text="Enviar", command=lambda:enviar_OC(actualizar_data(), fecha_seleccionada.get(), data_clientes,seleccion_cliente.get(), entrada_oc.get(), cantidad_ods.get(), productos_agregados, entrada_direccion.get(), link_oc, links_od, links_gr, opcion_moneda.get())).pack(side = "left", fill = "x", padx = 10)

#botón print productos agregados
boton_print = ttk.Button(frame_botones_derecha, text="print lista a enviar", command=lambda:debugin(productos_agregados))
boton_print.pack(padx = 10, side="left")

frame_inicio.pack(fill="both", expand=True)

ventana.mainloop()
