import os
import sys

import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from tqdm import tqdm


def load_excel(file):
    file_extension = os.path.splitext(file)[1]
    if file_extension == '.xls':
        return pd.read_excel(file, engine='xlrd')
    else:
        return pd.read_excel(file, engine='openpyxl')


def search_match(name, name_list):
    match, score = process.extractOne(name, name_list, scorer=fuzz.token_set_ratio)
    return match, score


def data_row():
    print("Introduce 1 para usar columnas por defecto y 2 para columnas personalizadas")
    print("\t- Columnas por defecto: Usar columna NOMBRE CRM, Coincidir con Organización y obtener valor de Cód. cliente")
    option = input()

    if option == '1':
        search_row = "NOMBRE CRM"
        match_row = "Organización"
        identifier_row = "Cód. cliente"
    else:
        print("Introduce el encabezado de la fila que necesitamos usar (Nombre, ...): ")
        search_row = input().strip()
        print("\nAhora introduce el encabezado de la fila con la que queremos realizar las coincidencias (Organización, ...): ")
        match_row = input().strip()
        print(
            "\nDel resultado, escriba el nombre de la columna donde se encuentra el valor con el que te quieres quedar (Cód cliente, ...): ")
        identifier_row = input().strip()

    return search_row, match_row, identifier_row


def process_problem(file_in, file_out):
    df = load_excel(file_in)
    results = []

    search_row, match_row, identifier_row = data_row()

    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="Procesando filas"):
        name_column_b = row[search_row.upper()]
        name_column_c = df[match_row].tolist()

        match, score = search_match(name_column_b, name_column_c)
        row_identify = df[df[match_row] == match]
        if not row_identify.empty:
            identifier = row_identify[identifier_row].values[0]
        else:
            identifier = 'No coincidences'

        results.append([name_column_b, identifier, match, score])

    df_results = pd.DataFrame(results, columns=['Nombre', 'Identificador', 'Coincidencia', 'Porcentaje'])
    df_results.to_excel(file_out, index=False)


def main():
    if len(sys.argv) > 1:
        file_in = sys.argv[1]
    else:
        file_in = "old_file.xlsx"
    file_out = "file_trained.xlsx"

    print("Comenzando el procesamiento...")
    process_problem(file_in, file_out)
    print("Procesamiento completado.")


if __name__ == '__main__':
    try:
        print("############################################")
        print("############################################")
        print("######### EXCEL MATCH PROGRAMME ############")
        print("############################################")
        print("############################################\n")
        main()
    except KeyboardInterrupt:
        print("Keyboard Interrupt")
