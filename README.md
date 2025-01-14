# Gestion de Tombola et Application de Tirage au Sort

Ce projet contient deux scripts principaux pour la gestion et le tirage au sort d'une tombola, en utilisant des données Excel et une interface utilisateur avec Streamlit.

## Fonctionnalités

1. **ProcessExcel.py** :
   - Prépare les données pour la tombola à partir d'un fichier Excel source.
   - Étend les données en dupliquant les entrées selon le nombre de tickets achetés.
   - Génère un fichier Excel contenant toutes les informations nécessaires pour le tirage.

2. **TirageStreamlit.py** :
   - Fournit une interface utilisateur interactive pour le tirage au sort des gagnants de la tombola.
   - Gère les lots et les tickets, incluant les lots restreints avec des règles spécifiques.
   - Sauvegarde les résultats dans un fichier Excel et exporte les données pour partage.

---

## Prérequis

- **Python 3.8+**
- Bibliothèques Python :
  - `pandas`
  - `streamlit`
  - `openpyxl`

Installez les dépendances avec :
```bash
pip install -r requirements.txt
```

---

## Usage

### Préparation des données

1. Placez le fichier Excel source contenant les données des participants dans le chemin défini dans `ProcessExcel.py`.
2. Exécutez le script pour générer le fichier étendu :
   ```bash
   python ProcessExcel.py
   ```

### Interface de tirage

1. Placez les fichiers générés (par exemple, les tickets et les lots) dans les chemins spécifiés dans `TirageStreamlit.py`.
2. Lancez l'application Streamlit :
   ```bash
   streamlit run TirageStreamlit.py
   ```
3. Accédez à l'interface depuis votre navigateur à l'adresse indiquée par Streamlit.

---

## Structure des fichiers

- `ProcessExcel.py` : Script de préparation des données.
- `TirageStreamlit.py` : Application de tirage au sort.
- `expanded_tombola_data.xlsx` : Fichier généré contenant les tickets étendus.
- `Lots.xlsx` : Fichier Excel des lots.
- `tirage_gagnants.xlsx` : Résultats enregistrés.
- `tirage_gagnants_export.xlsx` : Résultats formatés pour partage.

---

## Contribution

Les contributions sont les bienvenues. Veuillez créer une issue ou une pull request pour proposer des améliorations ou signaler des problèmes.