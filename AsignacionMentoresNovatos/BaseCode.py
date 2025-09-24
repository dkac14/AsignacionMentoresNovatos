from openpyxl import *
from collections import defaultdict
from itertools import cycle

wbMentores = load_workbook(r"C:\Users\Lenovo\Downloads\PAO II 2025 Inscripcion Mentores.xlsx")
wsMentores = wbMentores.active

wbNovatos = load_workbook(r"C:\Users\Lenovo\Downloads\PAO II 2025 Inscripcion Novatos.xlsx")
wsNovatos = wbNovatos.active

encM = [c.value for c in wsMentores[1]]
encN = [c.value for c in wsNovatos[1]]

# Mentores
COL_M_NOMBRE  = "Nombre"
COL_M_CARRERA = "CARRERA2"
COL_M_CORREO  = "Correo electrónico"
COL_M_TEL     = "Número Telefónico"

idx_encabezados_Mentores = {i: encM.index(i) for i in [COL_M_NOMBRE, COL_M_CARRERA, COL_M_CORREO, COL_M_TEL]}

# Novatos
COL_N_NOMBRE  = "Nombre Completo"
COL_N_CARRERA = "CARRERA2"
COL_N_CORREO  = "Correo de Espol, como estudiante de Espol te debieron asignar un correo, si todavia no sabes cual es, puedes dejarlo en blanco"
COL_N_TEL     = "Número Telefónico"

idx_encabezados_Novatos = {i: encN.index(i) for i in [COL_N_NOMBRE, COL_N_CARRERA, COL_N_CORREO, COL_N_TEL]}

# Emparejamiento 
emparejamiento = Workbook()
s_emparejamiento = emparejamiento.active
s_emparejamiento.title = "Asignación"

s_emparejamiento.append([COL_N_NOMBRE, COL_N_CARRERA, COL_N_CORREO, COL_N_TEL, "Mentor Asignado", COL_M_CARRERA, COL_M_CORREO, COL_M_TEL])

def norm(s):
    return "" if s is None else str(s).strip().upper()

# Agrupación mentores por carrera 
mentores_por_carrera = {}
punteros = {}

for rowM in wsMentores.iter_rows(min_row=2, values_only=True):
    carreraM = norm(rowM[idx_encabezados_Mentores[COL_M_CARRERA]])
    if carreraM == "": 
        continue
    mentores_por_carrera.setdefault(carreraM, []).append(rowM)

for car in mentores_por_carrera:
    punteros[car] = 0  

for rowN in wsNovatos.iter_rows(min_row=2, values_only=True):
    novato_data = [
        rowN[idx_encabezados_Novatos[COL_N_NOMBRE]],
        rowN[idx_encabezados_Novatos[COL_N_CARRERA]],
        rowN[idx_encabezados_Novatos[COL_N_CORREO]],
        rowN[idx_encabezados_Novatos[COL_N_TEL]],
    ]

    carrera_novato_norm = norm(novato_data[1])
    lista_ment = mentores_por_carrera.get(carrera_novato_norm, [])

    if not lista_ment:
        s_emparejamiento.append(novato_data + ["SIN MENTOR", "", "", ""])
    else:
        p = punteros[carrera_novato_norm]
        rowM = lista_ment[p]
        punteros[carrera_novato_norm] = (p + 1) % len(lista_ment)

        mentor_data = [
            rowM[idx_encabezados_Mentores[COL_M_NOMBRE]],
            rowM[idx_encabezados_Mentores[COL_M_CARRERA]],
            rowM[idx_encabezados_Mentores[COL_M_CORREO]],
            rowM[idx_encabezados_Mentores[COL_M_TEL]],
        ]

        s_emparejamiento.append(novato_data + mentor_data)


emparejamiento.save("Emparejamiento_balanceado.xlsx")
print("OK -> Emparejamiento_balanceado.xlsx")

