from openpyxl import load_workbook


base_incompleta = load_workbook("PAO I 2025 Mentores Novatos(1-67).xlsx") #load_workbook carga el libro que le falta info
hoja_incompleta = base_incompleta.active # .active activa las hojas dentro del libro

base_anterior = load_workbook("PAO II 2024 Mentores Novatos(1-634).xlsx") #carga el libro de la base anterior
hoja_full = base_anterior.active


data_full = {}
encabezados = [cell.value for cell in hoja_full[1]]
id_col_idx = encabezados.index("Correo electrónico") + 1  #match con correo electrónico


for row in hoja_full.iter_rows(min_row=2, values_only=False):
    id_value = row[id_col_idx - 1].value
    if id_value:
        data_full[id_value] = [cell.value for cell in row]


for row in hoja_incompleta.iter_rows(min_row=2):
    id_value = row[id_col_idx - 1].value
    if id_value in data_full:
        full_row = data_full[id_value]
        for i, cell in enumerate(row):
            if cell.value is None and i < len(full_row):
                cell.value = full_row[i]


base_incompleta.save("CopiaCompleta_BaseDeRegistro PAO I 2025 Mentores Novatos.xlsx")
print("Archivo actualizado guardado como 'CopiaCompleta_BaseDeRegistro PAO I 2025 Mentores Novatos.xlsx'")
