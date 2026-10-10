
#  RETO 3: Filtrar personas cuya empresa contenga la palabra "Group"

import openpyxl
from tabulate import tabulate

libro = openpyxl.load_workbook("personas.xlsx")
hoja = libro.active

palabra = "Group"
resultado = []

for fila in hoja.iter_rows(min_row=2, values_only=True):
    empresa = fila[2]
    # .lower() pasa todo a minúsculas, así encuentra "Group", "group" o "GROUP"
    if palabra.lower() in empresa.lower():
        resultado.append([fila[1], empresa, fila[3]])

print(f"\nPersonas cuya empresa contiene '{palabra}':\n")

if resultado:
    print(tabulate(resultado, headers=["Name", "Company", "Email"], tablefmt="fancy_grid"))
    print(f"\nSe encontraron {len(resultado)} personas.")
else:
    print("No se encontró ninguna empresa con esa palabra.")
