import os
import pandas as pd

# Richiesta all'utente del percorso della cartella
folder_path = input("Inserisci il percorso della cartella: ")

# Costruzione dei percorsi dei file Excel
path_melograno = os.path.join(folder_path, 'FABBRICA DEI SEGNI - IL MELOGRANO IDQ SCORE OTT 2023.xlsx')
path_leaf_node = os.path.join(folder_path, 'Leaf node Corretti.xlsx')
path_thema = os.path.join(folder_path, 'Classificazioni codici THEMA.xlsx')

# Leggi i dati dalle tabelle Excel
df_melograno = pd.read_excel(path_melograno)
df_leaf_node = pd.read_excel(path_leaf_node)
df_thema = pd.read_excel(path_thema)

# Permetti all'utente di inserire l'EAN desiderato
ean_scelto = input("Inserisci l'EAN desiderato: ")

# Trova il titolo corrispondente all'EAN scelto
titolo = df_melograno.loc[df_melograno['EAN'] == int(ean_scelto), 'Title'].values[0]

# Trova la prima riga vuota disponibile nella tabella Leaf node Corretti
riga_vuota = df_leaf_node[df_leaf_node.isnull().all(axis=1)].index[0]

# Copia i campi ASIN, EAN, Title nella riga vuota della tabella Leaf node Corretti
df_leaf_node.loc[riga_vuota, ['ASIN', 'EAN', 'Title']] = df_melograno.loc[df_melograno['EAN'] == int(ean_scelto), ['ASIN', 'EAN', 'Title']].values[0]

# Consenti all'utente di inserire un numero
numero_inserito = int(input("Inserisci un numero: "))

# Cerca la riga corrispondente al numero nella tabella Classificazioni codici THEMA
riga_thema = df_thema[df_thema['Numero'] == numero_inserito].iloc[:, :3]

# Incolla le informazioni nella prima colonna disponibile dell'EAN scelto nella tabella Leaf node Corretti
df_leaf_node.loc[riga_vuota, 'Nuova_Colonna'] = riga_thema.values.flatten()

# Scrivi i dati aggiornati nei rispettivi file Excel
with pd.ExcelWriter('Leaf node Corretti.xlsx') as writer:
    df_leaf_node.to_excel(writer, index=False)
