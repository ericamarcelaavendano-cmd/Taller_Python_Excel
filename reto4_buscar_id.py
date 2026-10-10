
#  RETO 4: Pedir un ID con input() y mostrar solo esa persona
#  Como los Id son muy largos (ej: 05936c10-1f8e-4a4c-...),
#  se puede buscar de dos formas:
#    - Por el número de fila (#), por ejemplo: 3
#    - Por el Id completo o por su inicio, por ejemplo: 05936c10

import openpyxl
from tabulate import tabulate

libro = openpyxl.load_workbook("personas.xlsx")
hoja = libro.active
personas = list(hoja.iter_rows(min_row=2, values_only=True))

busqueda = input("Escribe el # de la persona o su Id: ").strip()   # strip() quita espacios

encontrada = None

if busqueda == "":
    print("No escribiste nada. Intenta de nuevo.")

elif busqueda.isdigit():                       # ¿Escribió solo números? -> buscar por #
    numero = int(busqueda)
    if 1 <= numero <= len(personas):
        encontrada = personas[numero - 1]      # las listas empiezan en 0
    else:
        print(f"El número debe estar entre 1 y {len(personas)}.")

else:                                          # Si no, buscar por Id
    for persona in personas:
        if persona[0].startswith(busqueda.lower()):
            encontrada = persona
            break                              # ya la encontramos, dejamos de buscar
    if encontrada is None:
        print("No existe ninguna persona con ese Id.")

if encontrada:
    headers = ["Id", "Name", "Company", "Email", "MAC Address"]
    print(tabulate([encontrada], headers=headers, tablefmt="fancy_grid"))
