import pandas as pd
import random

# Chemins des fichiers (à adapter)
input_file_path = "Data/export-tombola-de-l-institut-de-myologie-2024-la-myocoop-25_11_2024-14_01_2025.xlsx"
output_file_path = "ProcessData/expanded_tombola_data.xlsx"

# Charger les données depuis le fichier Excel
excel_data = pd.ExcelFile(input_file_path)
df = excel_data.parse('Feuille 1')

# Sélectionner les colonnes nécessaires et les renommer
columns_needed = ['Prénom participant', 'Nom participant', 'Numéro de billet', 'Tarif', 'Email payeur']
df_selected = df[columns_needed].rename(columns={
    'Prénom participant': 'Prénom',
    'Nom participant': 'Nom',
    'Numéro de billet': 'Numéro du billet original',
    'Tarif': 'Nombre de tickets',
    'Email payeur': 'Adresse e-mail'
})

# Calculer le nombre total de tickets
total_tickets = sum([int(row['Nombre de tickets'].split()[0]) for _, row in df_selected.iterrows()])

# Calculer le nombre de participants uniques (basé sur prénom et nom)
# Créez une copie des colonnes transformées en majuscules pour détecter les doublons
df_temp = df_selected[['Prénom', 'Nom']].apply(lambda x: x.str.upper())

# Identifiez les lignes uniques en ignorant la casse
unique_participants = df_selected.loc[df_temp.drop_duplicates().index, ['Prénom', 'Nom']]

num_unique_participants = len(unique_participants)

# Imprimer le nombre unique de participants
print(f"Nombre unique de participants : {num_unique_participants}")
print(f"Nombre total de ticket : {total_tickets}")

# Générer une liste de numéros uniques et mélanger
random.seed(42) 
random_numbers = list(range(1, total_tickets + 1))
random.shuffle(random_numbers, )

# Créer une liste pour stocker les données étendues
expanded_data = []

# Assigner les numéros uniques aléatoires
current_random_index = 0

# Parcourir chaque ligne et dupliquer selon le nombre de tickets
for index, row in df_selected.iterrows():
    num_tickets = int(row['Nombre de tickets'].split()[0])  # Extraire le nombre de tickets
    for i in range(1, num_tickets + 1):
        expanded_data.append({
            'Numéro d\'index': len(expanded_data) + 1,
            'Prénom': row['Prénom'],
            'Nom': row['Nom'],
            'Adresse e-mail': row['Adresse e-mail'],  # Ajouter l'adresse e-mail
            'Numéro du billet original': f"{row['Numéro du billet original']}-{i}",
            'Nombre unique': random_numbers[current_random_index]  # Assigner un numéro unique aléatoire
        })
        current_random_index += 1

# Créer le DataFrame final
df_expanded = pd.DataFrame(expanded_data)

# Sauvegarder le DataFrame dans un nouveau fichier Excel
df_expanded.to_excel(output_file_path, index=False)

print(f"Le fichier étendu a été créé : {output_file_path}")