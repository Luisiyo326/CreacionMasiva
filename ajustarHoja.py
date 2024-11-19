from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins


archivo_plantilla  = "plantilla.xlsx"
#wb es para cargar la hoja de la biblio openpyxl
wb = load_workbook(archivo_plantilla)
hoja = wb.active
    
    #definamos la area de impresion
hoja.print_area = "A1:N38"
    #margenes
hoja.page_margins = PageMargins(
    left=0,   
    right=0,  
    top=0,    
    bottom=0
    )
    #orientacion de la hoja
hoja.page_setup.orientation = "landscape"
    #tipo de hoja
hoja.page_setup.paperSize = hoja.PAPERSIZE_A4
    #ajusta el ancho
hoja.page_setup.fitToWidth = 0
    #Ajusta la altura
hoja.page_setup.fitToHeight = 0
archivo_salida = "plantilla_ajustada2.xlsx"
wb.save(archivo_salida)