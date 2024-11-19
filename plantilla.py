import pandas as pd 
from openpyxl import load_workbook
import subprocess
import os

archivo_origen = "clon_info.xlsx"
df = pd.read_excel(archivo_origen)

archivo_plantilla  = "plantilla_ajustada2.xlsx"

os.makedirs("Excel_Generados", exist_ok=True)
os.makedirs("PDF_Generados", exist_ok=True)

#itera por columnas
for idx, proveedor in df.iterrows():
    ejercicio = proveedor['ejercicio']  
    relacion = proveedor['relacion']
    seccion = proveedor['seccion']
    nSeccion = proveedor['nombreSeccion']
    categoria = proveedor['categoria']
    nCategoria = proveedor['nombreCategoria']  
    prove = proveedor['noProveedor']
    nProve = proveedor['nombreProveedor']
    descnota=proveedor['concepto']
    parcialidad=proveedor['parcialidad']
    subtotal = proveedor['subtotal']
    ieps = proveedor['IEPS']
    iva = proveedor['IVA']
    total = proveedor['total']
    
    #wb es para cargar la hoja de la biblio openpyxl
    wb = load_workbook(archivo_plantilla)
    hoja = wb.active
    
    hoja["M2"] = ejercicio
    hoja["M3"] = relacion
    hoja["C12"] = seccion
    hoja["E12"] = nSeccion
    hoja["C13"] = categoria
    hoja["E13"] = nCategoria
    hoja["C14"] = prove
    hoja["E14"] = nProve
    hoja["B24"]=descnota
    hoja["B26"]=parcialidad
    hoja["M22"] = subtotal
    hoja["M23"] = ieps
    hoja["M24"] = iva
    hoja["M25"] = total
    
    
    excels = f"Excel_Generados/{prove}_{nProve}_{ejercicio}.xlsx" 
    wb.save(excels)
    print(f"Archivo Excel creado: {excels}")
    
    
    pdf = f"PDF_Generados/{prove}_{nProve}_{ejercicio}.pdf"
    subprocess.run([
        "libreoffice", 
        "--headless", 
        "--convert-to", "pdf", 
        "--outdir", "PDF_Generados", 
        excels
    ])
    print(f"Archivo PDF creado: {pdf}")
