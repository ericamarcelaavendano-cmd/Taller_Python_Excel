
#  RETO 1: Mostrar únicamente Name y Email

import openpyxl
from tabulate import tabulate

# Abrimos el Excel y tomamos la hoja activa (igual que en el taller)
libro = openpyxl.load_workbook("personas.xlsx")
hoja = libro.active

data = []

# Recorremos las filas desde la 2 (la 1 es el encabezado).
# values_only=True nos da directamente los valores, sin objetos "celda".
for fila in hoja.iter_rows(min_row=2, values_only=True):
    # Cada fila trae: (id, name, company, email, mac_address)
    #                   0    1      2       3        4
    nombre = fila[1]
    email = fila[3]
    data.append([nombre, email])

print(tabulate(data, headers=["Name", "Email"], tablefmt="fancy_grid", showindex=range(1, len(data) + 1)))
