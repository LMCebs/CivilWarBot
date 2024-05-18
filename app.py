import pandas as pd
import random
import openpyxl

# Número de provincias que se capturan por cada ataque (1 -> 55 - 75 días, 2 -> 19 - 27 días, 3 -> 15 - 24 días)
provxAtaque = 30

# Accedo en modo escritura al control de las comunidades
wb = openpyxl.load_workbook("coords_provincias.xlsx")
sheet = wb["Control"]

# Array con todas las comunidades
comunidades = ["Andalucía", "Aragón", "Asturias", "Islas Baleares", "Islas Canarias", "Cantabria", "Castilla y León", "Castilla-La Mancha", "Cataluña", "Extremadura", "Galicia", "Madrid", "Murcia", "Navarra", "País Vasco", "La Rioja", "Comunidad Valenciana", "Ceuta", "Melilla", "Islas Azores", "Kiev"]

# Creamos una variable global que indique qué nodos se han visitado ya
nodosProbados = list()

# Leo los datos del Excel
def LeerDatos():
    df = pd.read_excel("coords_provincias.xlsx", sheet_name = "MatrizAdyacencia", header = None, index_col = 0)
    control = pd.read_excel("coords_provincias.xlsx", sheet_name = "Control", index_col = 0)

    print(control)

    return df, control

# Bucle de reseteo del Excel
def Reset():
    cont = 2
    
    while(cont < 23):
        if cont < 5:
            sheet["B" + str(cont)] = "Al-Lagam"
        
        elif cont < 8:
            sheet["B" + str(cont)] = "Recreativo de Juerga"

        elif cont < 11:
            sheet["B" + str(cont)] = "Real Matriz"

        elif cont < 14:
            sheet["B" + str(cont)] = "Real Club de Parados"

        elif cont < 17:
            sheet["B" + str(cont)] = "Pombo F.C."

        elif cont < 20:
            sheet["B" + str(cont)] = "Minabo de Kiev"

        else:
            sheet["B" + str(cont)] = "Gambote del Norte S.A.D."

        cont+=1

    wb.save("coords_provincias.xlsx")

# Elige una comunidad aleatoria que atacará
def Day():
    ComunidadSeleccionada = random.randint(0,20)

    if random.randint(1,12) == 1 and ComunidadSeleccionada != control.iloc[ComunidadSeleccionada, 5]:

        # En una probabilidad entre 12, la provincia atacante se rebelará y volverá a su control inicial
        celda = "B" + str(ComunidadSeleccionada + 2)
        sheet[celda] = control.iloc[ComunidadSeleccionada, 5]

        print("Tras una rebelión en " + comunidades[ComunidadSeleccionada] + ", la comunidad ha vuelto al control del " + control.iloc[ComunidadSeleccionada, 5])

    else:
        accion = Ataque(ComunidadSeleccionada)

        if accion == -1:
            print ("Ha ganado el " + control.iloc[0,0])

# Elige provxAtaque - 1 comunidades colindantes al ataque y las conquista también 
def CambioCelda(comDefensa, equAtaque):

    adyacentes = list(df.iloc[comDefensa].values)
    random.shuffle(adyacentes)

    global provxAtaque
    numprovConq = 0

    for adyacente in adyacentes:

        # Si ya se han conquistado todas las provincias indicadas por la variable global, se para el bucle
        if numprovConq >= provxAtaque - 1:
            break

        # Seleccionamos otras provincias a su alrededor que sean del mismo propietario
        if adyacente != -1 and control.iloc[adyacente, 0] == control.iloc[comDefensa, 0]:
            celda = "B" + str(adyacente + 2)
            sheet[celda] = equAtaque
            
    # Finalmente cambia la celda original por su conquistador        
    celda = "B" + str(comDefensa + 2)
    sheet[celda] = equAtaque
    

# Elige la comunidad que defenderá
def Ataque(comAtaque):

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

        equDefensor = control.iloc[objetivo, 0]

        # Si un objetivo no es controlado por el mismo equipo que el atacante, se selecciona como objetivo
        if objetivo != -1 and equDefensor != equAtaque:

            print (equAtaque + " ha atacado desde su base en " + comunidades[comAtaque] + " a la base de " + equDefensor + " en " + comunidades[objetivo])

            # Actualizamos el valor en el excel
            CambioCelda(objetivo, equAtaque)

            nodosProbados = list()

            return objetivo

    # Si no ha encontrado objetivo es que todas sus comunidades colindantes son controladas por el mismo, por lo que hacemos un bucle que use sus territorios para buscar nuevos objetivos
    for objetivo in objetivos:

        if objetivo != -1 and objetivo not in nodosProbados:

            ataque = Ataque(objetivo)

            if ataque != -1:

                return ataque


    return -1

df, control = LeerDatos()

Day()
            









