from openpyxl import load_workbook

archivo = "plantilla.xlsx"  
wb = load_workbook(archivo)
hoja = wb.active


celdas_interes = ["M3","M4", "D13","F13","D14", "F14","D15","F15","N23","N24","N25","N26"]

for celda in celdas_interes:
    pertenece_a_rango = False
    for rango in hoja.merged_cells.ranges:
        if celda in rango:
            pertenece_a_rango = True
            celda_sup_izq = rango.coord.split(':')[0]
            print(f"La celda {celda} pertenece al rango fusionado: {rango}")
            print(f"La celda superior izquierda del rango es: {celda_sup_izq}")
            break
    if not pertenece_a_rango:
        print(f"La celda {celda} no pertenece a un rango fusionado")

wb.close()
