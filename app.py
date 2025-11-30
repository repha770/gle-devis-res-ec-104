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
    Extraction d'une page -> ligne du tableau Recensement
    Logique validée sur les devis page 20 (2025-1424) et 26 (2025-1418).
    """

    data: dict = {}

    # ========= Infos générales =========
    data["Opération n°"] = op_index + 1
    data["Code Fiche"] = "RES-EC-104"
    data["Type de trame (TOP/TPM)"] = ""

    data["RAISON SOCIALE \ndu demandeur"] = raison_sociale_demandeur
    data["SIREN \ndu demandeur"] = siren_demandeur
    data[" Raison sociale du mandataire assurant le rôle actif et incitatif"] = ""
    data["Numéro SIREN du mandataire assurant le rôle actif et incitatif"] = ""
    data["Nature de la bonification"] = "ZNI"

    # Référence / dates / prime
    ref_interne = extract_first(r"Numéro Client\s*:\s*(.+)", text)
    data["REFERENCE interne de l'opération"] = ref_interne

    date_devis = extract_first(r"Date\s*:\s*([0-9/]{8,10})", text)
    data["DATE d'envoi du RAI"] = date_devis
    data["DATE D'ENGAGEMENT\nde l'opération"] = date_devis

    prime_cee_raw = extract_first(r"Prime CEE\s*:\s*([0-9\s\u00A0,]+)\s*€", text)
    data["MONTANT de l'incitation financière CEE"] = clean_money(prime_cee_raw)

    # ========= BLOC ADRESSE DES TRAVAUX =========

    bloc = ""
    if "ADRESSE DES TRAVAUX :" in text:
        bloc = text.split("ADRESSE DES TRAVAUX :", 1)[1]
    elif "ADRESSE DES TRAVAUX" in text:
        bloc = text.split("ADRESSE DES TRAVAUX", 1)[1]

    # On coupe le bloc avant le tableau de détail
    for stop in ["Détail Quantité", "Detail Quantité"]:
        if stop in bloc:
            bloc = bloc.split(stop, 1)[0]
            break

    # On enlève les SIRET dans ce bloc
    bloc = re.sub(r"Siret\s*:\s*\d+", "", bloc)

    lignes_all = [l.strip() for l in bloc.splitlines() if l.strip()]

    # NOM DU SITE = 1ère ligne après "ADRESSE DES TRAVAUX :"
    nom_site = lignes_all[0] if lignes_all else ""
    data["NOM DU SITE bénéficiaire \nde l'opération"] = nom_site

    # On sépare bloc TRAVAUX et bloc HAUT-DROITE en coupant au 1er CP
    idx_cp1 = None
    cp1 = ville1 = ""
    adr1 = ""
    for idx, li in enumerate(lignes_all):
        m = re.search(r"(\d{5})\s+(.+)", li)
        if m:
            idx_cp1 = idx
            cp1 = m.group(1)
            ville1 = m.group(2).strip()
            if idx > 0:
                adr1 = lignes_all[idx - 1]
            break

    if idx_cp1 is not None:
        travaux_lines = lignes_all[: idx_cp1 + 1]
        haut_lines = lignes_all[idx_cp1 + 1 :]
    else:
        travaux_lines = lignes_all
        haut_lines = []

    # ==== Colonnes O, P, Q : bloc TRAVAUX (1er CP) ====
    data["ADRESSE \nde l'opération"] = adr1
    data["CODE POSTAL\n(sans cedex)"] = cp1
    data["VILLE\n"] = ville1

    # ==== Colonnes T, U, V, W : NOM DU SITE + mêmes valeurs que O,P,Q ====
    data["ADRESSE \nde l'opération.1"] = adr1
    data["CODE POSTAL\n(sans cedex).1"] = cp1
    data["VILLE"] = ville1

    # ========= BLOC HAUT-DROITE (adresse siège) =========

    adr2 = cp2 = ville2 = ""
    for idx, li in enumerate(haut_lines):
        m = re.search(r"(\d{5})\s+(.+)", li)
        if m:
            cp2 = m.group(1)
            ville2 = m.group(2).strip()
            # si la ligne précédente ne contient pas de CP, on la prend comme adresse,
            # sinon on prend la ligne elle-même (cas "VILLELE ANTENNE 4, 97460 SAINT-PAUL")
            if idx > 0 and not re.search(r"\d{5}\s", haut_lines[idx - 1]):
                adr2 = haut_lines[idx - 1]
            else:
                adr2 = li
            break

    # RAISON SOCIALE bénéficiaire (X)
    rs_benef = extract_first(r"DEVIS\s+[^\s]+\s+(.+)", text)
    data["RAISON SOCIALE \ndu bénéficiaire \nde l'opération"] = rs_benef

    # SIREN bénéficiaire (Y)
    siret_benef = extract_first(r"Siret\s*:\s*([0-9]{9,14})", text)
    siren_benef = siret_benef[:9] if len(siret_benef) >= 9 else ""
    data["SIREN"] = siren_benef

    # Z / AA / AB = adresse / CP / ville du bloc haut-droite (ou fallback travaux)
    data["ADRESSE \ndu siège social du bénéficiaire de l'opération"] = adr2 or adr1
    data["CODE POSTAL\n(sans cedex).2"] = cp2 or cp1
    data["VILLE.1"] = ville2 or ville1

    # ========= Téléphone / Mail =========
    tel = extract_first(r"Tél\s*:\s*(.+)", text)
    mail = extract_first(r"Mail\s*:\s*(.+)", text)
    mail_clean = "" if mail.lower().startswith("néant") else mail

    data["Numéro de téléphone du bénéficiaire"] = tel
    data["Adresse de courriel du bénéficiaire"] = mail_clean
    data["Numéro de téléphone du bénéficiaire.1"] = tel
    data["Adresse de courriel du bénéficiaire.1"] = mail_clean

    # ========= Colonnes M / N : NOM / PRENOM à partir de "Représenté par" =========
    repres = extract_first(r"Représenté par\s*:\s*(.+)", text)
    nom_benef = ""
    prenom_benef = ""
    if repres:
        repres_clean = re.sub(r",.*", "", repres).strip()
        parts = repres_clean.split()
        if len(parts) >= 2:
            nom_benef = parts[0]
            prenom_benef = " ".join(parts[1:])
        elif len(parts) == 1:
            nom_benef = parts[0]
    data["NOM \ndu bénéficiaire \nde l'opération "] = nom_benef
    data["PRENOM \ndu bénéficiaire \nde l'opération"] = prenom_benef

    # ========= Professionnel (GLE) =========
    data["SIREN du professionnel mettant en œuvre l’opération d’économies d’énergie"] = siren_pro
    data["RAISON SOCIALE du professionnel mettant en œuvre l’opération d’économies d’énergie"] = raison_sociale_pro
    data["RAISON SOCIALE du professionnel qui figure sur le devis"] = raison_sociale_pro
    data["SIREN du professionnel qui figure sur le devis"] = siren_pro

    # ========= Devis & montants =========
    num_devis = extract_first(r"DEVIS\s+([^\s]+)", text)
    data["NUMERO de  devis"] = num_devis
    data["MONTANT du devis (€ TTC)"] = clean_money(prime_cee_raw)

    nb_depose = extract_first(r"Nombre de dépose\s*:\s*([0-9]+)", text)
    data["Nombre de luminaires installés ou à installer"] = int(nb_depose) if nb_depose.isdigit() else math.nan

    return data


# ========= UI STREAMLIT =========

st.set_page_config(page_title="Devis LED → Tableau RES-EC-104", layout="wide")

st.title("📊 Devis LED → Tableau RES-EC-104")

st.write("1️⃣ Choisis le PDF des devis. 2️⃣ Je remplis automatiquement le modèle RES-EC-104 stocké dans l'app.")

pdf_file = st.file_uploader("📄 PDF des devis (1 devis par page)", type=["pdf"])

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

if pdf_file:
    if st.button("🚀 Lancer l'extraction"):
        try:
            # 1) On charge le modèle depuis le fichier stocké dans le repo
            modele_path = "modele_res_ec_104.xlsx"  # <--- mets ton fichier modèle avec ce nom dans le même dossier que app.py
            df_template = pd.read_excel(modele_path, sheet_name="Recensement")
            cols_modele = list(df_template.columns)

            # 2) Lecture du PDF
            reader = PdfReader(pdf_file)
            n_pages = len(reader.pages)
            st.success(f"PDF chargé : {n_pages} devis détectés")

            rows = []
            for i in range(n_pages):
                txt = reader.pages[i].extract_text()
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

            # 3) Aligner sur les colonnes du modèle
            df_out = pd.DataFrame(columns=cols_modele)
            for col in cols_modele:
                if col in df_new.columns:
                    df_out[col] = df_new[col]
                else:
                    df_out[col] = None

            st.subheader("Aperçu")
            st.dataframe(df_out.head())

            # 4) Export Excel
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
    st.info("Uploade le PDF des devis pour lancer l'extraction.")
