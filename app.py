import pandas as pd
import random
import openpyxl

# Leo los datos del Excel
def LeerDatos():
    df = pd.read_excel("coords_provincias.xlsx", sheet_name = "MatrizAdyacencia", header = None, index_col = 0)
    control = pd.read_excel("coords_provincias.xlsx", sheet_name = "Control", index_col = 0)

    return df, control


# Accedo en modo escritura al control de las comunidades
wb = openpyxl.load_workbook("coords_provincias.xlsx")
sheet = wb["Control"]

# Array con todas las comunidades
comunidades = ["Andalucía", "Aragón", "Asturias", "Islas Baleares", "Islas Canarias", "Cantabria", "Castilla y León", "Castilla-La Mancha", "Cataluña", "Extremadura", "Galicia", "Madrid", "Murcia", "Navarra", "País Vasco", "La Rioja", "Comunidad Valenciana", "Ceuta", "Melilla", "Islas Azores", "Kiev"]

# Creamos una variable global que indique qué nodos se han visitado ya
nodosProbados = list()

# Elige una comunidad aleatoria que atacará
def ElegirAtacante():
    return random.randint(0, 20)

# Elige la comunidad que defenderá
def ElegirDefensor(comAtaque):

    # Indica que se ha comprobado el nodo actual
    global nodosProbados
    nodosProbados.append(comAtaque)

    # Guarda el valor del equipo que controla la comunidad
    equAtaque = control.iloc[comAtaque, 0]
    
    # Saca la lista de objetivos de la comunidad atacante y la baraja
    objetivos = list(df.iloc[comAtaque].values)
    random.shuffle(objetivos)

    # Pasa por la lista de objetivos hasta que encuentre uno para atacar
    for objetivo in objetivos:

        # Si un objetivo no es controlado por el mismo equipo que el atacante, se selecciona como objetivo
        if objetivo != -1 and control.iloc[objetivo, 0] != equAtaque:

            print (comunidades[comAtaque] + " ha atacado a " + comunidades[objetivo])

            # Actualizamos el valor en el excel
            celda = "B" + str(objetivo + 2)
            sheet[celda] = equAtaque

            nodosProbados = list()

            return objetivo

    # Si no ha encontrado objetivo es que todas sus comunidades colindantes son controladas por el mismo, por lo que hacemos un bucle que use sus territorios para buscar nuevos objetivos
    for objetivo in objetivos:

        if objetivo != -1 and objetivo not in nodosProbados:

            ataque = ElegirDefensor(objetivo)

            if ataque != -1:

                return ataque


    return -1

while True:
    df, control = LeerDatos()

    accion = ElegirDefensor(ElegirAtacante())

    if accion == -1:
        break

    wb.save("coords_provincias.xlsx")
            









