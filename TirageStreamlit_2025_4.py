# -*- coding: utf-8 -*-
"""
Application de tirage au sort - Tombola
Gestion des lots et des tickets avec interface graphique Streamlit.

Version demandée :
- Identification des personnes POUR LES LOTS RESTREINTS basée sur "Prénom + Nom" (pas l'email)
- Lots restreints : tirage ALÉATOIRE PAR TICKET (proportionnel aux tickets)
  + une même personne (Prénom+Nom) ne peut gagner qu'une fois ce lot restreint
- Numéro de lot : 1 numéro par exemplaire (groupe de lots) + extraction robuste
- Optimisations perf : colonnes d'identifiants pré-calculées une fois (vectorisées)
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


def norm_text(x) -> str:
    """Normalise une chaîne pour comparaison (espaces + casse)."""
    return str(x).strip().casefold()


def add_person_key_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajoute une colonne _person_key basée STRICTEMENT sur prénom+nom (normalisés),
    utilisée pour les exclusions sur lots restreints.
    """
    out = df.copy()
    for col in ["Prénom", "Nom"]:
        if col not in out.columns:
            out[col] = ""

    prenom = out["Prénom"].astype(str).str.strip().str.casefold()
    nom = out["Nom"].astype(str).str.strip().str.casefold()

    out["_person_key"] = "name:" + prenom + "|" + nom
    return out


def get_lot_number(lot_row, fallback_index=None):
    """
    Récupère le numéro de lot de façon robuste, quel que soit le nom de colonne dans Lots25.xlsx.
    """
    candidates = [
        "numéro", "numero",
        "Numéro", "Numero",
        "numéro du lot", "numero du lot",
        "Numéro du lot", "Numero du lot",
        "N° lot", "N°", "N°Lot", "N° Lot",
        "Numero lot", "numéro lot", "numero lot",
    ]
    for c in candidates:
        if c in lot_row and pd.notna(lot_row[c]):
            return lot_row[c]
    return (fallback_index + 1) if fallback_index is not None else None


def load_existing_results():
    """Charge les résultats enregistrés s'ils existent. Migre l'ancienne clé 'numéro' si besoin."""
    try:
        df = pd.read_excel(output_file_path)

        if "numéro" in df.columns and "Numéro du lot" not in df.columns:
            df["Numéro du lot"] = df["numéro"]

        return df.to_dict("records")
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
            "Nom": result["Nom"][0].upper() + ".",
            "Numéro du billet original": result["Numéro du billet original"],
            "Lot": result["Lot"],
            "Offert par": result["Offert par"],
            "Adresse e-mail": result.get("Adresse e-mail", ""),
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
    st.session_state.tickets_df = add_person_key_column(load_data()[0])
    st.session_state.restricted_winners_per_lot = {}

    st.success("Historique réinitialisé avec succès.")
    st.rerun()


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


# === Lots restreints ===

restricted_lots = [
    "Pot de miel + abonnement Kazidomi", "Patchs anti-cernes", "Patchs anti-cernes + gel douche + beurre de karité", "Crème pour les mains",
    "Pot beurre de karité de poche + petite pochette", "Lot de 10 pinces et barrettes cheveux", "Savon",
    "Totebag + gel douche + beurre de karité + patch aloe vera + pince cheveux", "Jeu de rôles",
    "Gazette/enquête pour enfant espion", "Escape game à domicile", "Boucles d'oreilles", "Sweat", "Pochoirs + livret", "Cahier d'activité forêt",
    "Lot éponges lavables 4 couleurs", "Sac à dos + travel kit", "Barrette ronde + bracelet + créoles", "Créoles + bracelet", "Lunettes de soleil",
    "Boucles d'oreilles cœur", "Créoles", "Bracelet océan", "Lot affiches", "Etagère enfant", "Décoration murale", "Jeu de piste", "Box 2 repas pour 2",
    "Peluche fruits et légumes", "2 entrées enfant", "2 Kits éducatif + pochette", "Gel douche"
]

restricted_lots_norm = {norm_text(x) for x in restricted_lots}

if "restricted_winners_per_lot" not in st.session_state:
    st.session_state.restricted_winners_per_lot = {}


def draw_lots_group(tickets_df, lots_df, current_lot_index):
    """Effectue un tirage au sort pour un groupe de lots similaires."""
    if current_lot_index >= len(lots_df):
        st.warning("Tous les lots ont déjà été tirés !")
        return None, tickets_df, current_lot_index

    lot0 = lots_df.iloc[current_lot_index]
    group_count = 1

    while (
        current_lot_index + group_count < len(lots_df)
        and lots_df.iloc[current_lot_index + group_count]["lot"] == lot0["lot"]
        and lots_df.iloc[current_lot_index + group_count]["offert par"] == lot0["offert par"]
    ):
        group_count += 1

    if len(tickets_df) < 1:
        st.warning("Plus aucun ticket disponible !")
        return None, tickets_df, current_lot_index

    lot_name = lot0["lot"]
    lot_name_norm = norm_text(lot_name)

    # Groupe des lots (un numéro par ligne / exemplaire)
    lot_group = lots_df.iloc[current_lot_index: current_lot_index + group_count].reset_index(drop=True)
    lot_numbers = [
        get_lot_number(lot_group.iloc[i], fallback_index=current_lot_index + i)
        for i in range(group_count)
    ]

    # === Cas LOT RESTREINT ===
    # Tirage PAR TICKET + exclusion par (prénom+nom)
    if lot_name_norm in restricted_lots_norm:
        if lot_name_norm not in st.session_state.restricted_winners_per_lot:
            st.session_state.restricted_winners_per_lot[lot_name_norm] = set()

        # ⚠️ Ici on garde le set global, mais on va le "vider" si tout le monde est exclu
        excluded_people = st.session_state.restricted_winners_per_lot[lot_name_norm]
        results = []

        for i in range(group_count):

            # Tickets éligibles = ceux dont la personne n'a pas déjà gagné DANS CE TOUR
            eligible = tickets_df[~tickets_df["_person_key"].isin(excluded_people)]

            # ✅ Si plus personne d'éligible, on relance un tour (on ré-autorise tout le monde)
            if eligible.empty:
                excluded_people.clear()
                eligible = tickets_df  # tout le monde redevient éligible

                # si même là c'est vide -> plus aucun ticket global, donc stop
                if eligible.empty:
                    st.warning(f"Aucun ticket disponible pour attribuer les exemplaires restants du lot '{lot_name}'.")
                    break

            # ✅ Tirage aléatoire PAR TICKET
            winner = eligible.sample(1).iloc[0]
            pkey = winner["_person_key"]

            results.append({
                "Numéro du lot": lot_numbers[i],
                "Prénom": format_name(winner["Prénom"]),
                "Nom": format_last_name(winner["Nom"]),
                "Lot": lot_name,
                "Offert par": lot0["offert par"],
                "Adresse e-mail": winner.get("Adresse e-mail", ""),
                "Numéro du billet original": winner["Numéro du billet original"],
            })

            # Bloquer la personne pour le reste du tour
            excluded_people.add(pkey)

            # Retirer le ticket tiré du pool global (comme avant)
            tickets_df = tickets_df.drop(winner.name)

        return results, tickets_df, current_lot_index + group_count


    # === Cas LOT NON RESTREINT (tirage normal par ticket) ===
    if len(tickets_df) < group_count:
        st.warning("Pas assez de tickets pour tirer tous les gagnants !")
        return None, tickets_df, current_lot_index

    winners = tickets_df.sample(group_count)
    tickets_df = tickets_df.drop(winners.index)

    results = []
    for i, (_, winner) in enumerate(winners.iterrows()):
        results.append({
            "Numéro du lot": lot_numbers[i],
            "Prénom": format_name(winner["Prénom"]),
            "Nom": format_last_name(winner["Nom"]),
            "Lot": lot_name,
            "Offert par": lot0["offert par"],
            "Adresse e-mail": winner.get("Adresse e-mail", ""),
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
        background-color: #00B2B2;
        color: white !important;
        font-size: 16px;
        font-weight: bold;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        transition: background-color 0.3s ease, color 0.3s ease;
    }
    div.stButton > button:hover {
        background-color: #008080;
        color: white !important;
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
tickets_df = add_person_key_column(tickets_df)

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

            # Écritures (peuvent être lentes avec OneDrive + Excel)
            save_results(st.session_state.results)
            export_results(st.session_state.results)

            st.session_state.current_lot_index = new_index

col1, col2, col3, col4, col5 = st.columns([1, 3, 1, 3, 1])

with col2:
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

with col4:
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
if len(st.session_state.results) > 0:
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
