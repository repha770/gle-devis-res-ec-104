import re
import math
from io import BytesIO

import streamlit as st
import pandas as pd
from PyPDF2 import PdfReader


# ========= Utilitaires =========

def clean_money(s: str):
    if not s:
        return None
    s = s.replace("\u00A0", " ").replace(" ", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def extract_first(pattern: str, text: str, group: int = 1, default: str = "") -> str:
    m = re.search(pattern, text, flags=re.MULTILINE)
    return m.group(group).strip() if m else default


def parse_page(
    text: str,
    op_index: int,
    raison_sociale_demandeur: str,
    siren_demandeur: str,
    raison_sociale_pro: str,
    siren_pro: str,
) -> dict:
    """
    Extraction page -> une ligne du tableau Recensement
    Logique validée sur le devis page 20 :

    - Bloc travaux (ADRESSE DES TRAVAUX :) :
        O,P,Q  = adresse / CP / ville (1er CP du bloc)
        T,U,V,W = NOM DU SITE + même adresse / CP / ville

    - Bloc haut droite :
        X = RAISON SOCIALE bénéficiaire (ligne après DEVIS <num>)
        Y = SIREN (9 premiers chiffres après Siret :)
        Z,AA,AB = adresse / CP / ville haut droite
        AC,AD   = Tél / Mail
    """

    data: dict = {}

    # ========= Infos générales =========
    data["Opération n°"] = op_index + 1
    data["Code Fiche"] = "RES-EC-104"
    data["Type de trame (TOP/TPM)"] = ""

    # Demandeur / mandataire
    data["RAISON SOCIALE \ndu demandeur"] = raison_sociale_demandeur
    data["SIREN \ndu demandeur"] = siren_demandeur
    data[" Raison sociale du mandataire assurant le rôle actif et incitatif"] = ""
    data["Numéro SIREN du mandataire assurant le rôle actif et incitatif"] = ""
    data["Nature de la bonification"] = "ZNI"

    # Référence / dates / montants
    ref_interne = extract_first(r"Numéro Client\s*:\s*(.+)", text)
    data["REFERENCE interne de l'opération"] = ref_interne

    date_devis = extract_first(r"Date\s*:\s*([0-9/]{8,10})", text)
    data["DATE d'envoi du RAI"] = date_devis
    data["DATE D'ENGAGEMENT\nde l'opération"] = date_devis

    prime_cee_raw = extract_first(r"Prime CEE\s*:\s*([0-9\s\u00A0,]+)\s*€", text)
    data["MONTANT de l'incitation financière CEE"] = clean_money(prime_cee_raw)

    # ========= BLOC TRAVAUX : O,P,Q et T,U,V,W =========

    bloc_travaux = ""
    if "ADRESSE DES TRAVAUX" in text:
        bloc_travaux = text.split("ADRESSE DES TRAVAUX", 1)[1]
    # couper avant Tél / Représenté par / etc.
    for stop_kw in ["Tél", "Tel", "Représenté par", "Détail Quantité", "Nombre de dépose"]:
        if stop_kw in bloc_travaux:
            bloc_travaux = bloc_travaux.split(stop_kw, 1)[0]
    # enlever les Siret éventuels
    bloc_travaux = re.sub(r"Siret\s*:\s*\d+", "", bloc_travaux)

    lignes_t = [l.strip() for l in bloc_travaux.splitlines() if l.strip()]

    # NOM DU SITE = 1re ligne du bloc travaux
    nom_site = lignes_t[0] if lignes_t else ""
    data["NOM DU SITE bénéficiaire \nde l'opération"] = nom_site

    # Trouver le PREMIER CP + ville du bloc travaux (pour O,P,Q et U,V,W)
    adr_travaux = ""
    cp_travaux = ""
    ville_travaux = ""
    for idx, li in enumerate(lignes_t):
        m = re.search(r"(\d{5})\s+(.+)", li)
        if m:
            cp_travaux = m.group(1)
            ville_travaux = m.group(2).strip()
            if idx > 0:
                adr_travaux = lignes_t[idx - 1]
            break

    # Colonnes O,P,Q => bloc travaux (1er CP)
    data["ADRESSE \nde l'opération"] = adr_travaux
    data["CODE POSTAL\n(sans cedex)"] = cp_travaux
    data["VILLE\n"] = ville_travaux

    # Colonnes T,U,V,W => NOM DU SITE + même adresse / CP / ville
    data["ADRESSE \nde l'opération.1"] = adr_travaux
    data["CODE POSTAL\n(sans cedex).1"] = cp_travaux
    data["VILLE"] = ville_travaux

    # ========= BLOC HAUT DROITE : X,Y,Z,AA,AB,AC,AD =========

    # Raison sociale bénéficiaire = ligne après "DEVIS <num>"
    rs_benef = extract_first(r"DEVIS\s+[^\s]+\s+(.+)", text)
    data["RAISON SOCIALE \ndu bénéficiaire \nde l'opération"] = rs_benef

    # SIREN = 9 premiers chiffres après Siret
    siret_benef = extract_first(r"Siret\s*:\s*([0-9]{9,14})", text)
    siren_benef = siret_benef[:9] if len(siret_benef) >= 9 else ""
    data["SIREN"] = siren_benef

    # Bloc après ce Siret, jusqu'à "Tél"
    adr_siege = ""
    cp_siege = ""
    ville_siege = ""
    if siret_benef:
        try:
            # on découpe le texte après le SIRET et avant Tél
            after_siret = text.split(siret_benef, 1)[1]
            if "Tél" in after_siret:
                after_siret = after_siret.split("Tél", 1)[0]
            lignes_haut = [l.strip() for l in after_siret.splitlines() if l.strip()]
            for idx, li in enumerate(lignes_haut):
                m = re.search(r"(\d{5})\s+(.+)", li)
                if m:
                    cp_siege = m.group(1)
                    ville_siege = m.group(2).strip()
                    if idx > 0:
                        adr_siege = lignes_haut[idx - 1]
                    break
        except Exception:
            pass

    data["ADRESSE \ndu siège social du bénéficiaire de l'opération"] = adr_siege
    data["CODE POSTAL\n(sans cedex).2"] = cp_siege
    data["VILLE.1"] = ville_siege

    # Téléphone / mail (bloc haut droite)
    tel = extract_first(r"Tél\s*:\s*(.+)", text)
    mail = extract_first(r"Mail\s*:\s*(.+)", text)
    mail_clean = "" if mail.lower().startswith("néant") else mail

    data["Numéro de téléphone du bénéficiaire"] = tel
    data["Adresse de courriel du bénéficiaire"] = mail_clean
    data["Numéro de téléphone du bénéficiaire.1"] = tel
    data["Adresse de courriel du bénéficiaire.1"] = mail_clean

    # ========= Pro (GLE ou autre) =========
    data["SIREN du professionnel mettant en œuvre l’opération d’économies d’énergie"] = siren_pro
    data["RAISON SOCIALE du professionnel mettant en œuvre l’opération d’économies d’énergie"] = raison_sociale_pro
    data["RAISON SOCIALE du professionnel qui figure sur le devis"] = raison_sociale_pro
    data["SIREN du professionnel qui figure sur le devis"] = siren_pro

    # ========= Devis / montants =========
    num_devis = extract_first(r"DEVIS\s+([^\s]+)", text)
    data["NUMERO de  devis"] = num_devis

    data["MONTANT du devis (€ TTC)"] = clean_money(prime_cee_raw)

    # Nombre de luminaires
    nb_depose = extract_first(r"Nombre de dépose\s*:\s*([0-9]+)", text)
    data["Nombre de luminaires installés ou à installer"] = int(nb_depose) if nb_depose.isdigit() else math.nan

    return data


# ========= UI STREAMLIT =========

st.set_page_config(page_title="Devis LED → Tableau RES-EC-104", layout="wide")

st.title("📊 Devis LED → Tableau RES-EC-104")

st.write(
    "Upload le **PDF de devis (1 devis par page)** et ton **modèle Excel RES-EC-104** (onglet *Recensement*)."
)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 PDF des devis", type=["pdf"])
with col2:
    template_file = st.file_uploader("📑 Modèle Excel RES-EC-104 (onglet 'Recensement')", type=["xlsx"])

st.markdown("### Paramètres (GLE par défaut)")
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
            st.success(f"PDF chargé : {n_pages} devis détectés")

            # On lit le modèle pour récupérer l'ordre et les noms de colonnes exacts
            df_template = pd.read_excel(template_file, sheet_name="Recensement")
            cols_modele = list(df_template.columns)

            rows = []
            for i in range(n_pages):
                page = reader.pages[i]
                txt = page.extract_text()
                row = parse_page(
                    txt,
                    i,
                    raison_sociale_demandeur=raison_sociale_demandeur,
                    siren_demandeur=siren_demandeur,
                    raison_sociale_pro=raison_sociale_pro,
                    siren_pro=siren_pro,
                )
                rows.append(row)

            df_new = pd.DataFrame(rows)

            # On aligne strictement sur les colonnes du modèle
            df_out = pd.DataFrame(columns=cols_modele)
            for col in cols_modele:
                if col in df_new.columns:
                    df_out[col] = df_new[col]
                else:
                    df_out[col] = None

            st.subheader("Aperçu du tableau généré")
            st.dataframe(df_out.head())

            # Export Excel
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_out.to_excel(writer, sheet_name="Recensement", index=False)
            buffer.seek(0)

            st.download_button(
                "💾 Télécharger le tableau Excel rempli",
                data=buffer,
                file_name="Tableau_de_recensement_rempli.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Erreur pendant l'extraction : {e}")
else:
    st.info("Uploade le PDF et le modèle Excel pour activer le bouton.")
