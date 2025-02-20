import streamlit as st
import pandas as pd
import io
import os
import re
import string
import base64
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile


def convert_df_to_excel(df):
    """Convertit un DataFrame en fichier Excel binaire téléchargeable."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Fichier import')
    processed_data = output.getvalue()
    return processed_data


def replace_special_chars(value):
    if pd.isna(value):
        return "Aucun"
    if isinstance(value, str) and len(value) == 1 and value in string.punctuation:
        return "Aucun"
    return value


def replace_politique_voyage(value):
    if pd.isna(value):
        return "GENERAL.DEFAULT"
    if isinstance(value, str) and len(value) == 1 and value in string.punctuation:
        return "GENERAL.DEFAULT"
    return value

def replace_genre(value):
    if isinstance(value, str) and re.search(r'ma', value, re.IGNORECASE):
        return "Mrs"
    elif isinstance(value, str) and re.search(r'mo', value, re.IGNORECASE):
        return "Mr"
    return value

def first_name(value):
    if isinstance(value, str):
        return value.capitalize()
    return value

def last_name(value):
    if isinstance(value, str):
        return value.upper()
    return value

def role(value):
    if isinstance(value, str) and re.search(r'admin|executive', value, re.IGNORECASE):
        return "executive"
    elif isinstance(value, str) and re.search(r'mana|booker', value, re.IGNORECASE):
        return "booker"
    elif isinstance(value, str) and re.search(r'compta|accountant', value, re.IGNORECASE):
        return "accountant"
    elif isinstance(value, str) and re.search(r'voyage|traveler', value, re.IGNORECASE):
        return "traveler"
    if pd.isna(value):
        return "Vide"
    return value

def langue(value):
    if isinstance(value, str) and re.search(r'fr', value, re.IGNORECASE):
        return "fr"
    elif isinstance(value, str) and re.search(r'an', value, re.IGNORECASE):
        return "en"
    elif isinstance(value, str) and re.search(r'en', value, re.IGNORECASE):
        return "en"
    elif isinstance(value, str) and re.search(r'es', value, re.IGNORECASE):
        return "es"
    elif isinstance(value, str) and re.search(r'sp', value, re.IGNORECASE):
        return "es"
    return value

def date_de_naissance(value):
    try:
        return pd.to_datetime(value).strftime('%Y-%m-%d')
    except Exception:
        return value

def clean_email(value):
    if pd.isna(value):
        return "Vide"
    if isinstance(value, str) and re.search(r'[ ,;/\']', value):
        return value
    return value

def clean_tel(value):
    if value == "nan":
        value = ""
        return value
    elif isinstance(value, str) and value.strip():
        cleaned_value = re.sub(r'[ .+]', "", value)
        return cleaned_value

def acces(value):
    if pd.isna(value):
        return "false"
    if isinstance(value, str):
        value_lower = value.lower()
        if any(kw in value_lower for kw in ["avec", "non", "f", "faux", "false"]):
            return "false"
        elif any(kw in value_lower for kw in ["sans", "oui", "v", "vrai", "true"]):
            return "true"
    return value

def clean_emails(value):
    if isinstance(value, str):
        return re.sub(r'[^a-zA-Z0-9@._]+', ' ', value).strip()
    return value


def clear_column_if_not_empty(value):
    if isinstance(value, str) and value.strip():
        return ""
    return value

def contains_special_chars_or_spaces(value):
    if isinstance(value, str) and re.search(r'[ ,;/\'\s]', value):
        return True
    return False


def process_file(file):
    """Traite le fichier Excel chargé."""
    # Charger le fichier Excel
    xls = pd.ExcelFile(file)

    # Liste des onglets à fusionner si présents
    sheets_to_merge = ["Voyageurs", "Administrateurs", "Comptables", "Managers"]

    merged_df = pd.DataFrame()

    # Fusionner les onglets spécifiés
    available_sheets = [sheet for sheet in sheets_to_merge if sheet in xls.sheet_names]
    if available_sheets:
        for sheet in available_sheets:
            df = pd.read_excel(xls, sheet_name=sheet)
            merged_df = pd.concat([merged_df, df], ignore_index=True)
    else:
        merged_df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    # Suppression de la ligne contenant "Entité de facturation"
    if "Entité de facturation" in merged_df.iloc[0].to_list():
        merged_df = merged_df.iloc[1:].reset_index(drop=True)

    # Exemple de nettoyage des colonnes
    columns_to_check = ["Centre de coût secondaire / service", "service"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(replace_special_chars)

    columns_to_check = ["Politique de voyage", "policy"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(replace_politique_voyage)

    # Transformation pour "Genre"
    columns_to_check = ["Genre", "title"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(replace_genre)

    # Transformation pour "Prénom"
    columns_to_check = ["Prénom", "firstname"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(first_name)

    # Transformation pour "Nom de famille"
    columns_to_check = ["Nom de famille", "lastname"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(last_name)

    # Transformation pour "Rôle"
    columns_to_check = ["Rôle", "roles"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(role)

    # Transformation pour "Langue"
    columns_to_check = ["Langue", "language"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(langue)

    # Transformation pour "Date de naissance"
    columns_to_check = ["Date de naissance", "birthdate"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(date_de_naissance)

    # Transformation pour "Email"
    columns_to_check = ["Email", "email"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(clean_email)

    # Transformation pour "TEL"
    columns_to_check = ["TEL", "phone"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].astype(str).apply(clean_tel)

    # Transformation pour les colonnes d'accès
    columns_to_process = [
        "Sans accès", "Peut réserver pour lui sans validation dans la politique",
        "Peut réserver pour les autres sans validation",
        "Peut réserver pour lui sans validation hors politique",
        "Peut valider dans la politique", "Peut valider hors politique",
        "Peut voir les offres hors politique", "Validation RSE",
        "Recevoir les demandes de réservations des membres de l'équipe",
        "Recevoir les confirmations de réservations des membres de l'équipe",
        "Recevoir les reçus", "Recevoir les factures périodiques"
            ]
    for col in columns_to_process:
        if col in merged_df.columns:
            merged_df[col] = merged_df[col].apply(acces)

    columns_to_check = ["Assigner valideur ", "managers"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(clean_emails)

     # Transformation pour "Recevoir tout (admin)"
    columns_to_check = ["Recevoir tout (admin)"]
    for column_name in columns_to_check:
        if column_name in merged_df.columns:
            merged_df[column_name] = merged_df[column_name].apply(clear_column_if_not_empty)

    # Formatage en texte pour toutes les colonnes d'accès
    for col in columns_to_process:
        columns_to_check = [col, "test"]
        for column_name in columns_to_check:
            if column_name in merged_df.columns:
                merged_df[column_name] = merged_df[column_name].astype(str).str.lower()

    # Filtrer les lignes où "ID" contient "Ne Pas Remplir Cette Case"
    if "ID" in merged_df.columns:
        merged_df = merged_df[merged_df["ID"] != "Ne Pas Remplir Cette Case"]

    # Réorganiser les colonnes selon l'ordre souhaité
    desired_order = [
        "ID", "Centre de coût principal", "Centre de coût secondaire / service",
        "Politique de voyage", "Genre", "Prénom", "Nom de famille", "Rôle", "Langue",
        "Date de naissance", "Email", "TEL", "Désactivé", "Sans accès",
        "Peut réserver pour lui sans validation dans la politique",
        "Peut réserver pour les autres sans validation",
        "Peut réserver pour lui sans validation hors politique",
        "Peut valider dans la politique", "Peut valider hors politique",
        "Peut voir les offres hors politique", "Validation RSE",
        "Assigner valideur ",
        "Recevoir les demandes de réservations des membres de l'équipe",
        "Recevoir les confirmations de réservations des membres de l'équipe",
        "Recevoir les reçus", "Recevoir les factures périodiques",
        "Recevoir tout (admin)", "Nom du champ perso 1 (lié au profil du voyageur)",
        "Nom du champ perso 2 (lié au profil du voyageur)"
    ]
    merged_df = merged_df.reindex(columns=[col for col in desired_order if col in merged_df.columns])

    return merged_df

# Configuration de la page
st.set_page_config(page_title="Nettoyage du fichier Implem", layout="wide")

# Titre principal
st.markdown(
    """
    <h1 style="color:#604fd7; font-size: 36px; font-weight: bold; text-align: center;">
        Nettoyage du fichier Implem
    </h1>
    <hr style="border: 1px solid #e8e6e6;">
    """,
    unsafe_allow_html=True
)

# Upload du fichier
st.markdown(
    """
    <h3 style="color:#604fd7; font-size: 24px; font-weight: bold;">
        📤 Dépose ton fichier Excel
    </h3>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("", type=["xlsx"])

if uploaded_file:
    # # Aperçu du fichier
    st.markdown(
        """
        <h3 style="color:#604fd7; font-size: 24px; font-weight: bold;">
            📂 Aperçu du fichier chargé
        </h3>
        """,
        unsafe_allow_html=True
    )

    preview_df = pd.read_excel(uploaded_file)
    st.dataframe(preview_df.head(5), use_container_width=True)

    # Traitement du fichier
    st.markdown(
        """
        <h3 style="color:#604fd7; font-size: 24px; font-weight: bold;">
            ⚙️ Traitement du fichier
        </h3>
        """,
        unsafe_allow_html=True
    )

    processed_df = process_file(uploaded_file)  # Appel à la fonction de traitement

    # Message de succès
    st.success("✅ Le fichier a été traité avec succès.")

    # Fichier final
    st.markdown(
        """
        <h3 style="color:#604fd7; font-size: 24px; font-weight: bold;">
            📥 Fichier final à télécharger
        </h3>
        """,
        unsafe_allow_html=True
    )

    # Convertir le DataFrame en CSV et l'encoder en base64
    csv = processed_df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    
    st.markdown(
        """
        <style>
        .download-btn {
            background-color: #604fd7;
            color: white !important;
            padding: 14px 28px;
            font-size: 20px;
            font-weight: bold;
            border-radius: 12px;
            text-decoration: none;
            display: inline-block;
            text-align: center;
            max-width: 500px;  /* Limiter la largeur du bouton */
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            transition: all 0.3s ease;
            cursor: pointer;
            margin: 0 auto; /* Assure que le bouton est centré */
            display: block;
        }
        .download-btn:hover {
            background-color: #503bb5;
            transform: translateY(-3px);  /* léger effet de soulèvement */
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.3);
        }
        .download-btn:active {
            transform: translateY(0);  /* effet de pression lors du clic */
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    # Récupérer le nom du fichier original sans l'extension
    original_filename = os.path.splitext(uploaded_file.name)[0]
    cleaned_filename = f"Clean {original_filename}.xlsx"


    # Bouton de téléchargement personnalisé en HTML avec le style
    st.markdown(
        f'<a href="data:file/csv;base64,{b64}" download="{cleaned_filename}" class="download-btn">Clique ici pour télécharger le fichier nettoyé</a>',
        unsafe_allow_html=True
    )
