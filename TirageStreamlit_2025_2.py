# -*- coding: utf-8 -*-
"""
Application de tirage au sort - Tombola
Gestion des lots et des tickets avec interface graphique Streamlit.
"""

import streamlit as st
import pandas as pd
import os

# === Chemins des fichiers ===
tickets_file_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\ProcessData\expanded_tombola_data.xlsx"
lots_file_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\Data\Lots25.xlsx"
output_file_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\ProcessData\tirage_gagnants.xlsx"
export_file_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\ProcessData\tirage_gagnants_export.xlsx"

logo_afm_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\Data\AFM_Telethon.png"
logo_institut_path = r"C:\Users\m.jacoupy\OneDrive - Institut\Documents\3 - Developpements informatiques\Tombola\Data\institut_de_myologie_couleur_francais_fond_transparent.png"

# === Fonctions utilitaires ===

@st.cache_data
def load_data():
    """Charge les données des tickets et des lots depuis les fichiers Excel."""
    tickets_df = pd.read_excel(tickets_file_path)
    lots_df = pd.read_excel(lots_file_path)
    return tickets_df, lots_df

def load_existing_results():
    """Charge les résultats enregistrés s'ils existent."""
    try:
        return pd.read_excel(output_file_path).to_dict('records')
    except FileNotFoundError:
        return []

def save_results(results):
    """Enregistre les résultats dans un fichier Excel."""
    pd.DataFrame(results).to_excel(output_file_path, index=False)

def export_results(results):
    """Crée un fichier d'export avec Prénom, initiale du nom de famille, ticket, offert par, email, et numéro du lot."""
    export_data = []
    for result in results:
        formatted_result = {
            "Numéro du lot": result.get("Numéro du lot", ""),
            "Prénom": result["Prénom"],
            "Nom": result["Nom"][0].upper() + ".",  # Initiale du nom de famille
            "Numéro du billet original": result["Numéro du billet original"],
            "Lot": result["Lot"],
            "Offert par": result["Offert par"],
            "Adresse e-mail": result["Adresse e-mail"],
        }
        export_data.append(formatted_result)

    pd.DataFrame(export_data).to_excel(export_file_path, index=False)

def reset_results():
    """Réinitialise l'historique des tirages."""
    if os.path.exists(output_file_path):
        os.remove(output_file_path)
    if os.path.exists(export_file_path):
        os.remove(export_file_path)
    st.session_state.current_lot_index = 0
    st.session_state.results = []
    st.session_state.tickets_df = load_data()[0]
    st.session_state.restricted_winners_per_lot = {}
    st.success("Historique réinitialisé avec succès.")

def format_name(name):
    """Formate les prénoms composés avec des majuscules appropriées."""
    return "-".join([part.capitalize() for part in str(name).split("-")])

def format_last_name(last_name):
    """Formate les noms de famille pour gérer les majuscules après espaces ou tirets."""
    last_name = str(last_name)
    formatted_name = " ".join(
        "-".join(part.capitalize() for part in segment.split("-"))
        for segment in last_name.split(" ")
    )
    return formatted_name

# === Normalisation & identifiants ===

def norm_text(x) -> str:
    """Normalise une chaîne pour comparaison (espaces + casse)."""
    return str(x).strip().casefold()

def person_id(row) -> str:
    """
    Identifiant unique de personne pour la restriction.
    Priorité : email (normalisé).
    Fallback : prénom+nom normalisés.
    """
    email = row.get("Adresse e-mail", "")
    email_norm = norm_text(email)
    if email_norm and email_norm != "nan":
        return f"email:{email_norm}"
    # fallback si pas d'email fiable
    prenom = norm_text(row.get("Prénom", ""))
    nom = norm_text(row.get("Nom", ""))
    return f"name:{prenom}|{nom}"

# === Gestion des lots restreints ===

restricted_lots = [
    "Pot de miel + abonnement Kazidomi", "Patchs anti-cernes", "Patchs anti-cernes + gel douche + beurre de karité", "Crème pour les mains",
    "Pot beurre de karité de poche + petite pochette", "Lot de 10 pinces et barrettes cheveux", "Savon",
    "Totebag + gel douche + beurre de karité + patch aloe vera + pince cheveux", "Jeu de rôles",
    "Gazette/enquête pour enfant espion", "Escape game à domicile", "Boucles d'oreilles", "Sweat", "Pochoirs + livret", "Cahier d'activité forêt",
    "Lot éponges lavables 4 couleurs", "Sac à dos + travel kit", "Barrette ronde + bracelet + créoles", "Créoles + bracelet", "Lunettes de soleil",
    "Boucles d'oreilles cœur", "Créoles", "Bracelet océan", "Lot affiches", "Etagère enfant", "Décoration murale", "Jeu de piste", "Box 2 repas pour 2",
    "Peluche fruits et légumes", "2 entrées enfant", "2 Kits éducatif + pochette", "Gel douche"
]

# Ensemble normalisé pour comparer sans bugs (casse/espaces)
restricted_lots_norm = {norm_text(x) for x in restricted_lots}

if "restricted_winners_per_lot" not in st.session_state:
    # dict : lot_norm -> set(person_id)
    st.session_state.restricted_winners_per_lot = {}

def draw_lots_group(tickets_df, lots_df, current_lot_index):
    """Effectue un tirage au sort pour un groupe de lots similaires."""
    if current_lot_index >= len(lots_df):
        st.warning("Tous les lots ont déjà été tirés !")
        return None, tickets_df, current_lot_index

    lot = lots_df.iloc[current_lot_index]
    group_count = 1

    while (
        current_lot_index + group_count < len(lots_df)
        and lots_df.iloc[current_lot_index + group_count]["lot"] == lot["lot"]
        and lots_df.iloc[current_lot_index + group_count]["offert par"] == lot["offert par"]
    ):
        group_count += 1

    if len(tickets_df) < 1:
        st.warning("Plus aucun ticket disponible !")
        return None, tickets_df, current_lot_index

    lot_name = lot["lot"]
    lot_name_norm = norm_text(lot_name)

    # === Cas LOT RESTREINT ===
    if lot_name_norm in restricted_lots_norm:
        if lot_name_norm not in st.session_state.restricted_winners_per_lot:
            st.session_state.restricted_winners_per_lot[lot_name_norm] = set()

        excluded_people = st.session_state.restricted_winners_per_lot[lot_name_norm]

        # Tickets éligibles = tickets dont la personne n'a PAS déjà gagné ce lot
        # On calcule aussi l'id de personne sur chaque ligne
        tmp = tickets_df.copy()
        tmp["_pid"] = tmp.apply(person_id, axis=1)

        eligible_tickets = tmp[~tmp["_pid"].isin(excluded_people)].copy()

        # Nombre de personnes éligibles (pas tickets)
        eligible_people_count = eligible_tickets["_pid"].nunique()

        if eligible_people_count == 0:
            st.warning(f"Aucun participant éligible pour le lot restreint : {lot_name}.")
            return None, tickets_df, current_lot_index + group_count  # on passe ce groupe

        if eligible_people_count < group_count:
            st.warning(
                f"Seulement {eligible_people_count} participants éligibles pour {group_count} exemplaires du lot '{lot_name}'. "
                "Certains exemplaires resteront non attribués."
            )
            group_count = eligible_people_count

        results = []

        # Pour garantir 1 personne max dans le groupe :
        # on tire une personne (pid) au hasard, puis un ticket appartenant à cette personne.
        for _ in range(group_count):
            # Recalculer les éligibles à chaque itération
            tmp = tickets_df.copy()
            tmp["_pid"] = tmp.apply(person_id, axis=1)
            eligible_tickets = tmp[~tmp["_pid"].isin(excluded_people)].copy()

            if eligible_tickets.empty:
                st.warning(f"Aucun ticket éligible pour les exemplaires restants du lot '{lot_name}'.")
                break

            # Tirer une personne parmi les pids éligibles
            pid_choices = eligible_tickets["_pid"].drop_duplicates()
            chosen_pid = pid_choices.sample(1).iloc[0]

            # Tirer un ticket de cette personne
            winner = eligible_tickets[eligible_tickets["_pid"] == chosen_pid].sample(1).iloc[0]

            results.append({
                "Prénom": format_name(winner["Prénom"]),
                "Nom": format_last_name(winner["Nom"]),
                "Lot": lot_name,
                "Offert par": lot["offert par"],
                "Adresse e-mail": winner["Adresse e-mail"],
                "Numéro du billet original": winner["Numéro du billet original"],
            })

            # Marquer la personne comme gagnante de CE lot restreint
            excluded_people.add(chosen_pid)

            # Retirer le ticket tiré du pool global (comme avant)
            tickets_df = tickets_df.drop(winner.name)

        return results, tickets_df, current_lot_index + group_count

    # === Cas LOT NON RESTREINT (tirage normal) ===
    if len(tickets_df) < group_count:
        st.warning("Pas assez de tickets pour tirer tous les gagnants !")
        return None, tickets_df, current_lot_index

    winners = tickets_df.sample(group_count)
    tickets_df = tickets_df.drop(winners.index)

    results = []
    for _, winner in winners.iterrows():
        results.append({
            "Numéro du lot": lot.get("numéro du lot", lot.get("Numero lot", lot.get("N° lot", current_lot_index + 1))),
            "Prénom": format_name(winner["Prénom"]),
            "Nom": format_last_name(winner["Nom"]),
            "Lot": lot_name,
            "Offert par": lot["offert par"],
            "Adresse e-mail": winner["Adresse e-mail"],
            "Numéro du billet original": winner["Numéro du billet original"],
        })

    return results, tickets_df, current_lot_index + group_count

# === Configuration de la barre latérale ===
st.sidebar.image(logo_afm_path, use_column_width=True)
st.sidebar.image(logo_institut_path, use_column_width=True)

# === Personnalisation des styles Streamlit ===
st.markdown(
    """
    <style>
    .tirer-button-container {
        display: flex;
        justify-content: center;
        margin: 20px 0;
    }
    div.stButton > button {
        background-color: #00B2B2; /* PANTONE 7466C */
        color: white !important; /* Couleur blanche pour le texte */
        font-size: 16px;
        font-weight: bold;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        transition: background-color 0.3s ease, color 0.3s ease;
    }
    div.stButton > button:hover {
        background-color: #008080; /* Couleur légèrement plus foncée pour l'effet hover */
        color: white !important; /* Maintenir le texte blanc au survol */
    }
    div.stButton > button:focus {
        outline: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# === Affichage principal ===
col1, col2, col3 = st.columns([1.5, 3, 1])
with col2:
    st.title("Tirage au Sort - Tombola")

tickets_df, lots_df = load_data()

if "current_lot_index" not in st.session_state:
    st.session_state.current_lot_index = 0
if "results" not in st.session_state:
    st.session_state.results = load_existing_results()
if "tickets_df" not in st.session_state:
    st.session_state.tickets_df = tickets_df

results = []

col1, col2, col3 = st.columns([5.25, 3, 5])
st.markdown("---")
with col2:
    if st.button("Tirer les prochains lots"):
        draw_results, st.session_state.tickets_df, new_index = draw_lots_group(
            st.session_state.tickets_df, lots_df, st.session_state.current_lot_index
        )
        if draw_results:
            results = draw_results
            st.session_state.results.extend(draw_results)
            save_results(st.session_state.results)
            export_results(st.session_state.results)
            st.session_state.current_lot_index = new_index

col1, col2, col3, col4, col5 = st.columns([1, 3, 1, 3, 1])

with col4:
    with st.container():
        st.markdown("### 🎉 Gagnants")
        if results:
            first_result = results[0]
            st.write(f"**Lot :** {first_result['Lot']}")
            st.write(f"**Offert par :** {first_result['Offert par']}")
            st.write("**Gagnants :**")
            for result in results:
                st.write(f"- {result['Prénom']} {result['Nom']}")
        else:
            st.info("Aucun gagnant pour le moment.")

with col2:
    with st.container():
        st.markdown("### 🎁 Prochain(s) lot(s)")
        if st.session_state.current_lot_index < len(lots_df):
            next_lot = lots_df.iloc[st.session_state.current_lot_index]
            next_lot_count = 1
            while (
                st.session_state.current_lot_index + next_lot_count < len(lots_df)
                and lots_df.iloc[st.session_state.current_lot_index + next_lot_count]["lot"] == next_lot["lot"]
                and lots_df.iloc[st.session_state.current_lot_index + next_lot_count]["offert par"] == next_lot["offert par"]
            ):
                next_lot_count += 1
            st.write(f"**Nombre de lots :** {next_lot_count}")
            st.write(f"**Lot :** {next_lot['lot']}")
            st.write(f"**Offert par :** {next_lot['offert par']}")
        else:
            st.warning("Tous les lots ont été tirés !")

st.markdown("---")
st.subheader("Historique des tirages")
if len(st.session_state.results) > 1:
    results_no_email = [
        {k: v for k, v in result.items() if k != "Adresse e-mail"}
        for result in st.session_state.results
    ]
    historical_results_df = pd.DataFrame(results_no_email).reset_index(drop=True)
    st.dataframe(historical_results_df, use_container_width=True)
else:
    st.write("Aucun tirage effectué pour le moment.")

st.markdown("---")
if st.button("Réinitialiser l'historique"):
    reset_results()
