import re
import math
from io import BytesIO

import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader

# ========= Fonctions utilitaires =========

def clean_money(s: str):
    if not s:
        return None
    s = s.replace("\u00A0", " ").replace(" ", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

def extract_first(pattern, text, group=1, default=""):
    m = re.search(pattern, text, flags=re.MULTILINE)
    return m.group(group).strip() if m else default

def parse_page(text: str, op_index: int,
               raison_sociale_demandeur: str,
               siren_demandeur: str,
               raison_sociale_pro: str,
               siren_pro: str) -> dict:
    data = {}

    # Constantes non extraites (mandataire / bonif / type trame vides)
    RAISON_SOCIALE_MANDATAIRE = ""
    SIREN_MANDATAIRE = ""
    NATURE_BONIFICATION = ""
    TYPE_TRAME = ""

    # ======= Infos générales =======
    data["Opération n°"] = op_index + 1
    data["Type de trame (TOP/TPM)"] = TYPE_TRAME
    data["Code Fiche"] = "RES-EC-104"

    data["RAISON SOCIALE \ndu demandeur"] = raison_sociale_demandeur
    data["SIREN \ndu demandeur"] = siren_demandeur

    # Référence interne = Numéro Client
    data["REFERENCE interne de l'opération"] = extract_first(r"Numéro Client\s*:\s*(.+)", text)

    # Date devis
    date_devis = extract_first(r"Date\s*:\s*([0-9/]{8,10})", text)
    data["DATE d'envoi du RAI"] = date_devis
    data["DATE D'ENGAGEMENT\nde l'opération"] = date_devis

    # Prime CEE
    prime_cee_raw = extract_first(r"Prime CEE\s*:\s*([0-9\s\u00A0,]+)\s*€", text)
    data["MONTANT de l'incitation financière CEE"] = clean_money(prime_cee_raw)

    # Mandataire (laisser vide)
    data[" Raison sociale du mandataire assurant le rôle actif et incitatif"] = RAISON_SOCIALE_MANDATAIRE
    data["Numéro SIREN du mandataire assurant le rôle actif et incitatif"] = SIREN_MANDATAIRE
    data["Nature de la bonification"] = NATURE_BONIFICATION

    # ======= Bénéficiaire / adresse =======
    after_adr = text.split("ADRESSE DES TRAVAUX :", 1)[1] if "ADRESSE DES TRAVAUX" in text else ""
    lignes = [l.strip() for l in after_adr.splitlines() if l.strip()]

    # Ligne 0 = raison sociale / nom du site
    ligne_nom_site = lignes[0] if len(lignes) >= 1 else ""

    # Représenté par
    repres = extract_first(r"Représenté par\s*:\s*(.+)", text)
    nom_benef = ""
    prenom_benef = ""
    if repres:
        parts = repres.split()
        if len(parts) >= 2:
            prenom_benef = parts[0]
            nom_benef = " ".join(parts[1:])
        else:
            nom_benef = repres

    data["NOM \ndu bénéficiaire \nde l'opération "] = nom_benef
    data["PRENOM \ndu bénéficiaire \nde l'opération"] = prenom_benef

    # CP + Ville
    cp_ville = extract_first(r"(\d{5}\s+[A-ZÉÈÎÂÔÙÇ\- ]+)", after_adr, default="")
    if cp_ville:
        cp = extract_first(r"(\d{5})", cp_ville)
        ville = cp_ville[len(cp):].strip()
    else:
        cp, ville = "", ""

    # Adresse = ligne juste avant le CP
    adresse = ""
    for i, l in enumerate(lignes):
        if re.search(r"\d{5}\s", l):
            if i > 0:
                adresse = lignes[i - 1]
            break

    data["ADRESSE \nde l'opération"] = adresse
    data["CODE POSTAL\n(sans cedex)"] = cp
    data["VILLE\n"] = ville

    # Téléphone / mail
    tel = extract_first(r"Tél\s*:\s*([0-9\-\. ]+)", text)
    mail = extract_first(r"Mail\s*:\s*(.+)", text)
    mail_clean = "" if mail.lower().startswith("néant") else mail

    data["Numéro de téléphone du bénéficiaire"] = tel
    data["Adresse de courriel du bénéficiaire"] = mail_clean

    # Partie morale (duplication)
    data["ADRESSE \nde l'opération.1"] = adresse
    data["CODE POSTAL\n(sans cedex).1"] = cp
    data["VILLE"] = ville
    data["RAISON SOCIALE \ndu bénéficiaire \nde l'opération"] = ligne_nom_site

    # SIREN via Siret
    siret_benef = extract_first(r"Siret\s*:\s*([0-9]{9,14})", text)
    siren_benef = siret_benef[:9] if len(siret_benef) >= 9 else ""
    data["SIREN"] = siren_benef
    data["ADRESSE \ndu siège social du bénéficiaire de l'opération"] = adresse
    data["CODE POSTAL\n(sans cedex).2"] = cp
    data["VILLE.1"] = ville
    data["Numéro de téléphone du bénéficiaire.1"] = tel
    data["Adresse de courriel du bénéficiaire.1"] = mail_clean

    # Professionnel (GLE)
    data["SIREN du professionnel mettant en œuvre l’opération d’économies d’énergie"] = siren_pro
    data["RAISON SOCIALE du professionnel mettant en œuvre l’opération d’économies d’énergie"] = raison_sociale_pro

    # Numéro de devis
    num_devis = extract_first(r"DEVIS\s+([^\s]+)", text)
    data["NUMERO de  devis"] = num_devis

    # Montant du devis (on met la prime CEE)
    data["MONTANT du devis (€ TTC)"] = clean_money(prime_cee_raw)

    # Pro qui figure sur le devis
    data["RAISON SOCIALE du professionnel qui figure sur le devis"] = raison_sociale_pro
    data["SIREN du professionnel qui figure sur le devis"] = siren_pro

    # Nombre de luminaires
    nb_depose = extract_first(r"Nombre de dépose\s*:\s*([0-9]+)", text)
    data["Nombre de luminaires installés ou à installer"] = int(nb_depose) if nb_depose.isdigit() else math.nan

    return data


# ========= UI STREAMLIT =========

st.set_page_config(page_title="Devis LED → Tableau RES-EC-104", layout="wide")

st.title("📊 Devis LED → Tableau RES-EC-104")
st.write(
    "Upload le **PDF de devis (1 devis par page)** et ton **modèle Excel RES-EC-104**."
)

col1, col2 = st.columns(2)

with col1:
    pdf_file = st.file_uploader("📄 PDF des devis", type=["pdf"])

with col2:
    template_file = st.file_uploader("📑 Modèle Excel RES-EC-104 (onglet 'Recensement')", type=["xlsx"])

st.markdown("### Paramètres GLE (par défaut)")

c1, c2, c3, c4 = st.columns(4)
with c1:
    raison_sociale_demandeur = st.text_input("Raison sociale demandeur", value="GLE")
with c2:
    siren_demandeur = st.text_input("SIREN demandeur", value="829067826")
with c3:
    raison_sociale_pro = st.text_input("Raison sociale professionnel", value="GLE")
with c4:
    siren_pro = st.text_input("SIREN professionnel", value="829067826")

if pdf_file and template_file:
    if st.button("🚀 Lancer l'extraction"):
        try:
            reader = PdfReader(pdf_file)
            n_pages = len(reader.pages)
            st.success(f"PDF chargé : **{n_pages} devis détectés**")

            df_template = pd.read_excel(template_file, sheet_name="Recensement")
            cols_modele = list(df_template.columns)

            rows = []
            for i in range(n_pages):
                text = reader.pages[i].extract_text()
                row = parse_page(
                    text,
                    i,
                    raison_sociale_demandeur=raison_sociale_demandeur,
                    siren_demandeur=siren_demandeur,
                    raison_sociale_pro=raison_sociale_pro,
                    siren_pro=siren_pro,
                )
                rows.append(row)

            df_new = pd.DataFrame(rows)
            df_out = pd.DataFrame(columns=cols_modele)

            for col in cols_modele:
                if col in df_new.columns:
                    df_out[col] = df_new[col]
                else:
                    df_out[col] = None

            st.subheader("Aperçu")
            st.dataframe(df_out.head())

            # Excel en mémoire
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_out.to_excel(writer, sheet_name="Recensement", index=False)
            output.seek(0)

            st.download_button(
                label="💾 Télécharger le tableau Excel",
                data=output,
                file_name="Tableau_de_recensement_rempli.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Erreur pendant l'extraction : {e}")
else:
    st.info("Uploade le PDF et le modèle Excel pour activer le bouton d'extraction.")
