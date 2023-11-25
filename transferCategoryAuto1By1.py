import pandas as pd
import re

while True:
    # Carica i file
    leaf_node_corretti = pd.read_excel("Leaf node Corretti - v2.xlsx")
    fabbrica_dei_segna = pd.read_excel("FABBRICA DEI SEGNI - IL MELOGRANO IDQ SCORE OTT 2023.xlsx")
    classificazioni = pd.read_excel("Classificazioni codici THEMA.xlsx")

    row_with_empty_node_id = leaf_node_corretti[leaf_node_corretti["Node ID 1"].isnull()]

    if not row_with_empty_node_id.empty:
        ean_to_search = row_with_empty_node_id.iloc[0]["EAN"]

        matching_row = fabbrica_dei_segna[fabbrica_dei_segna["ean"] == ean_to_search]

        if not matching_row.empty:
            sottocategoria = matching_row.iloc[0]["Sottocategoria"]
            print(f"Il Node ID 1 vuoto corrisponde a {sottocategoria}")

            numbers = re.findall(r'\b\d+\b', sottocategoria)
            print("Numeri trovati:", numbers)

            for idx, number in enumerate(numbers, start=1):
                column_name = f"Node ID {idx}"
                # Converti in stringa e imposta la colonna come stringa
                leaf_node_corretti[column_name] = leaf_node_corretti[column_name].astype(str)
                leaf_node_corretti.loc[row_with_empty_node_id.index[0], column_name] = str(number)

                matching_row = classificazioni[classificazioni.iloc[:, 0].astype(str) == str(number)]
                if not matching_row.empty:
                    node_path = matching_row.iloc[0]["Node Path"]
                    tema_classification = matching_row.iloc[0]["Thema Classification"]
                    # Calcola l'indice delle colonne in cui inserire i dati
                    column_index = leaf_node_corretti.columns.get_loc(f"Node ID {idx}") + 1
                    # Inserisci le informazioni nei posti corretti
                    leaf_node_corretti.iloc[row_with_empty_node_id.index[0], column_index] = node_path
                    leaf_node_corretti.iloc[row_with_empty_node_id.index[0], column_index + 1] = tema_classification

            leaf_node_corretti.to_excel("Leaf node Corretti - v2.xlsx", index=False)
        else:
            print("Nessuna corrispondenza trovata nel file FABBRICA DEI SEGNI - IL MELOGRANO IDQ SCORE OTT 2023.xlsx")
    else:
        print("Processo completato, nessun Node ID 1 vuoto trovato nel file Leaf node Corretti - v2.xlsx")
        break  # Interrompi il ciclo while quando non ci sono più righe vuote
