from openpyxl import load_workbook
from collections import defaultdict
import json

ruta = r"C:\Users\Gateway\Downloads\PAO I 2026 Mentores & Novatos(1-33).xlsx"

wb = load_workbook(ruta)
ws = wb.active

if ws is None:
    raise RuntimeError("Workbook has no active worksheet")

encabezados = [c.value for c in ws[1]]


# =========================
# 2. Funciones auxiliares
# =========================

def norm(texto):  # sourcery skip: assign-if-exp, reintroduce-else
    """
    Normaliza texto para comparar:
    - Convierte None en ""
    - Quita espacios al inicio y final
    - Convierte a mayúsculas
    """
    if texto is None:
        return ""
    return str(texto).strip().upper()


def limpiar(texto):  # sourcery skip: assign-if-exp, reintroduce-else
    """
    Limpia texto para guardar en el JSON.
    No lo convierte a mayúsculas porque queremos conservar el formato original.
    """
    if texto is None:
        return ""
    return str(texto).strip()


def obtener_indice(nombre_columna):
    """
    Busca una columna por su nombre exacto.
    Si no la encuentra, lanza un error claro.
    """
    if nombre_columna not in encabezados:
        raise ValueError(f"No se encontró la columna: {nombre_columna}")
    return encabezados.index(nombre_columna)


def obtener_carrera(row, indices_carreras):
    """
    Revisa las columnas CARRERA, CARRERA2, CARRERA3...
    y devuelve la primera carrera que no esté vacía.
    """
    for idx in indices_carreras:
        valor = row[idx]
        if limpiar(valor) != "":
            return limpiar(valor)
    return ""


# =========================
# 3. Columnas importantes
# =========================

# sourcery skip: for-append-to-extend, use-next
COL_NOMBRE = "Nombre"
COL_CORREO = "Correo electrónico"
COL_TELEFONO = "Número Telefónico"

# Esta columna tiene un nombre largo.
# Para evitar copiar todo el texto exacto, la buscamos por una parte del nombre.
col_nivel = None

for encabezado in encabezados:
    if encabezado is not None and "Nivel de malla" in str(encabezado):
        col_nivel = encabezado
        break

if col_nivel is None:
    raise ValueError("No se encontró la columna de Nivel de malla")

idx_nombre = obtener_indice(COL_NOMBRE)
idx_correo = obtener_indice(COL_CORREO)
idx_telefono = obtener_indice(COL_TELEFONO)
idx_nivel = obtener_indice(col_nivel)

# Columnas posibles de carrera
columnas_carreras = [
    "CARRERA",
    "CARRERA2",
    "CARRERA3",
    "CARRERA4",
    "CARRERA5",
    "CARRERA6",
    "CARRERA7",
    "CARRERA8"
]

indices_carreras = []

for col in columnas_carreras:
    if col in encabezados:
        indices_carreras.append(encabezados.index(col))


# =========================
# 4. Separar mentores y novatos
# =========================

mentores_por_carrera = defaultdict(list)
novatos = []

for row in ws.iter_rows(min_row=2, values_only=True):
    nombre = limpiar(row[idx_nombre])
    correo = limpiar(row[idx_correo])
    telefono = limpiar(row[idx_telefono])
    nivel = norm(row[idx_nivel])
    carrera = obtener_carrera(row, indices_carreras)

    persona = {
        "nombre": nombre,
        "correo": correo,
        "telefono": telefono,
        "carrera": carrera
    }

    # Si no tiene carrera, no se puede emparejar bien
    if carrera == "":
        continue

    carrera_norm = norm(carrera)

    if nivel == "100 I":
        novatos.append(persona)
    else:
        mentores_por_carrera[carrera_norm].append(persona)


# =========================
# 5. Preparar estructura de salida
# =========================

resultado = []

# Esta estructura tendrá un bloque por cada mentor
for carrera_norm, lista_mentores in mentores_por_carrera.items():
    for mentor in lista_mentores:
        resultado.append({
            "mentor": mentor,
            "novatos": []
        })


# Para encontrar rápidamente los bloques de mentores por carrera
bloques_por_carrera = defaultdict(list)

for bloque in resultado:
    carrera_mentor_norm = norm(bloque["mentor"]["carrera"])
    bloques_por_carrera[carrera_mentor_norm].append(bloque)


# Punteros para asignación balanceada
punteros = defaultdict(int)


# =========================
# 6. Asignar novatos a mentores
# =========================

novatos_sin_mentor = []

for novato in novatos:
    carrera_novato_norm = norm(novato["carrera"])

    bloques_mentores = bloques_por_carrera.get(carrera_novato_norm, [])

    if not bloques_mentores:
        novatos_sin_mentor.append(novato)
    else:
        posicion = punteros[carrera_novato_norm]

        bloque_mentor = bloques_mentores[posicion]
        bloque_mentor["novatos"].append(novato)

        punteros[carrera_novato_norm] = (posicion + 1) % len(bloques_mentores)


# =========================
# 7. Agregar novatos sin mentor
# =========================

if novatos_sin_mentor:
    resultado.append({
        "mentor": {
            "nombre": "SIN MENTOR",
            "correo": "",
            "telefono": "",
            "carrera": ""
        },
        "novatos": novatos_sin_mentor
    })


# =========================
# 8. Guardar JSON en archivo .txt
# =========================

with open("asignacion_mentores.txt", "w", encoding="utf-8") as archivo:
    json.dump(resultado, archivo, ensure_ascii=False, indent=2)

print("OK -> asignacion_mentores.txt")