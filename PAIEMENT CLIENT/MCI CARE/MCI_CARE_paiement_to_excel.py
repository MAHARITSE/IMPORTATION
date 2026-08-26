# -*- coding: utf-8 -*-
"""
MCI_CARE_paiement_to_excel.py — société MCI CARE
================================================
Conversion des décomptes de règlement MCI CARE (PDF) en fichiers Excel
selon le modèle "Modele_Import_Reglements_Decompte_Assurance.xlsx"
(feuille Modele_Reglements, 13 colonnes).

Format traité (propre à MCI CARE) : DECOMPTE DE REGLEMENT FACTURES
  - en-tête : Facture #, Garant, Date comptable, Total prestataire
  - tableau par bénéficiaire :
      Matricule | Bénéficiaire | Date de soins | Actes | Montant réclamé |
      (#)Mtt non remboursé | Base décomptée | Ticket Modérateur |
      Montant réglé

Utilisation (double-clic sur MCI_CARE_paiement.bat, ou invite de commandes) :
    python MCI_CARE_paiement_to_excel.py                    # tous les PDF du dossier
    python MCI_CARE_paiement_to_excel.py --force            # régénérer (écrase l'Excel existant)
    python MCI_CARE_paiement_to_excel.py "mon_decompte.pdf" # un seul PDF (nom ou chemin)

Sortie : MCI CARE <MOIS> <ANNEE> <PERIODE> <MONTANT>.xlsx dans ce même dossier
    exemple : MCI CARE MAI 2026 02-03-26 à 31-03-26 471 140.xlsx
    - SOCIETE : MCI CARE (fixée dans ce script)
    - MOIS    : mois de la date comptable du décompte, en MAJUSCULES
    - ANNEE   : année de la date comptable
    - PERIODE : 1re et dernière date de soins du décompte, au format JJ-MM-AA
                (une seule date si toutes les lignes sont du même jour)
    - MONTANT : total payé (Total prestataire)

Les fichiers Excel déjà existants ne sont PAS écrasés (protection des
modifications manuelles), sauf avec l'option --force.
"""
import re
import sys
import glob
import os
import pdfplumber
import openpyxl.styles
from openpyxl import Workbook, load_workbook

# Le script se trouve dans le dossier de la société MCI CARE.
# Les PDF à convertir sont déposés dans ce même dossier.
# Le modèle Excel est dans le dossier parent "PAIEMENT CLIENT".
PDF_DIR = os.path.dirname(os.path.abspath(__file__))
MODEL = os.path.join(os.path.dirname(PDF_DIR),
                     "Modele_Import_Reglements_Decompte_Assurance.xlsx")
SHEET = "Modele_Reglements"
SOCIETE = "MCI CARE"

HEADERS = ["Ref_Decompte", "Date_Reglement", "Date_Soins", "Nom_Agent", "Matricule",
           "Numero_Facture_Prescription", "Code_Acte", "Libelle_Acte",
           "Montant_Reclame_Brut", "Ticket_Moderateur", "Montant_Paye_Regle",
           "Montant_Exclu_Rejet", "Motif_Observation"]

MONTHS = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet",
          "Aout", "Septembre", "Octobre", "Novembre", "Decembre"]


def amount_to_float(s):
    """Montant d'un PDF -> nombre. '' / None / texte non numérique -> 0.0.

    Gère les deux notations rencontrées dans les PDF :
        française : '15 000,00' / '928 750'  -> 15000.0 / 928750.0
        anglaise  : '75,000' / '1,234.56'    -> 75000.0 / 1234.56
    """
    s = re.sub(r"[^\d.,-]", "", str(s or "").replace("\u00a0", " "))
    if not s:
        return 0.0
    virgule, point = s.rfind(","), s.rfind(".")
    dernier = max(virgule, point)
    if dernier >= 0 and s.count(",") + s.count(".") == 1 \
            and len(s) - dernier - 1 == 3:
        # un seul séparateur suivi de 3 chiffres = séparateur de milliers
        s = s.replace(",", "").replace(".", "")
    elif virgule > point:
        s = s.replace(".", "").replace(",", ".")   # virgule = décimales
    else:
        s = s.replace(",", "")                     # point = décimales
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_amount(n):
    """471140.0 -> '471 140' ; 1234.5 -> '1 234,5'"""
    if abs(n - round(n)) < 0.005:
        return f"{int(round(n)):,}".replace(",", " ")
    return f"{n:,.1f}".replace(",", " ").replace(".", ",")


# Montant dans un décompte MCI : notation française ("1 153 900", "15 000,00")
# ou anglaise ("75,000", "1,234.56") selon l'édition du PDF.
MONTANT_RE = re.compile(
    r"\d{1,3}(?!\d)(?:[ \u00a0]\d{3})*(?:[.,]\d{3})?(?:[.,]\d{2})?"
    r"|\d+(?:[.,]\d+)?")


def dernier_montant(ligne):
    """Dernier montant d'une ligne de total = colonne "Montant réglé"."""
    trouves = MONTANT_RE.findall(ligne)
    return amount_to_float(trouves[-1]) if trouves else 0.0


def num(s):
    """'15 000,00' -> 15000 (int si entier)"""
    v = amount_to_float(s)
    return int(v) if v == int(v) else v


def parse_date(d):
    """'02/03/2026' -> '2026-03-02'"""
    dd, mm, yy = d.strip().split("/")
    year = int(yy)
    year += 2000 if year < 100 else 0
    return f"{year:04d}-{mm}-{dd}"


# --------------------------------------------------------------------------
# Nom du fichier Excel de sortie
#   <SOCIETE> <MOIS> <ANNEE> <PERIODE> <MONTANT>.xlsx
#   exemple : MCI CARE MAI 2026 02-03-26 à 31-03-26 471 140.xlsx
# --------------------------------------------------------------------------
# Caractères interdits dans un nom de fichier Windows
INVALIDES = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

DATE_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def date_courte(iso):
    """'2026-03-02' -> '02-03-26' (JJ-MM-AA)."""
    annee, mois, jour = iso.split("-")
    return f"{jour}-{mois}-{annee[2:]}"


def periode_soins(lignes, defaut=None):
    """Période couverte par le paiement : '02-03-26 à 31-03-26'.

    Prend la 1re et la dernière date de soins (colonne Date_Soins) des lignes
    du fichier. Une seule date si toutes les lignes tombent le même jour.
    À défaut (aucune date de soins lisible), la date de règlement.
    """
    dates = sorted(l.get("Date_Soins") or "" for l in lignes)
    dates = [d for d in dates if DATE_ISO.match(d)]
    if not dates:
        if not (defaut and DATE_ISO.match(defaut)):
            return "SANS DATE"
        dates = [defaut]
    debut, fin = dates[0], dates[-1]
    return date_courte(debut) if debut == fin \
        else f"{date_courte(debut)} à {date_courte(fin)}"


def nom_sortie(societe, date_reglement, lignes, montant):
    """Construit le nom du fichier Excel :

        <SOCIETE> <MOIS> <ANNEE> <PERIODE> <MONTANT>.xlsx
        MCI CARE MAI 2026 02-03-26 à 31-03-26 471 140.xlsx

    date_reglement : 'AAAA-MM-JJ' (date comptable du décompte).
    """
    annee, mm, _ = date_reglement.split("-")
    mois = MONTHS[int(mm) - 1].upper()
    nom = (f"{societe} {mois} {annee} "
           f"{periode_soins(lignes, date_reglement)} {fmt_amount(montant)}")
    nom = re.sub(r"\s+", " ", INVALIDES.sub(" ", nom)).strip()
    return nom + ".xlsx"


def full_pdf_text(pdf):
    return "\n".join((page.extract_text() or "") for page in pdf.pages)


# --------------------------------------------------------------------------
# Format MCI : DECOMPTE DE REGLEMENT FACTURES
# --------------------------------------------------------------------------
def parse_mci(pdf, nom_pdf):
    """Extrait le méta du décompte et les lignes du tableau Matricule...Montant réglé."""
    text = full_pdf_text(pdf)

    meta = {"facture": None, "date_reglement": None, "garant": None,
            "total_prestataire": None}
    m = re.search(r"Facture #\s*:\s*(\S+)", text)
    if m:
        meta["facture"] = m.group(1)
    m = re.search(r"Date comptable:\s*(\d{2}/\d{2}/\d{4})", text)
    if not m:
        m = re.search(r"Edition\s*:\s*(\d{2}/\d{2}/\d{4})", text)
    if m:
        meta["date_reglement"] = m.group(1)
    m = re.search(r"Garant:\s*(\S+)\s+([^\n]+)", text)
    if m:
        meta["garant"] = m.group(1) + " " + m.group(2).strip()
    # "Total prestataire:" apparait une fois par page (une page = un
    # établissement : PHIE, LABO, IMAGERIE, DISPENSAIRE...). On additionne
    # tous ces totaux pour avoir le total payé du décompte complet.
    totaux = [dernier_montant(l)
              for l in re.findall(r"Total prestataire:([^\n]*)", text)]
    if totaux:
        meta["total_prestataire"] = sum(totaux)

    lignes = []
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table or not str(table[0][0] or "").strip().startswith("Matricule"):
                continue
            for row in table[1:]:
                c = [str(x or "").strip() for x in row[:9]]
                # Ligne de données = matricule numérique + date de soins ;
                # sinon : ligne de total, "999 EXCLUSION...", description...
                if len(c) < 9 or not c[0].isdigit() \
                        or not re.fullmatch(r"\d{2}/\d{2}/\d{4}", c[2]):
                    continue
                (matricule, benef, date, acte, reclame,
                 nonremb, base, tm, regle) = c
                tm_montant = round(amount_to_float(reclame) - amount_to_float(regle))
                obs = f"Ticket modérateur {tm}"
                if amount_to_float(nonremb) > 0:
                    obs += f" ; Montant non remboursé : {fmt_amount(amount_to_float(nonremb))} Ar"
                lignes.append({
                    "Ref_Decompte": meta["facture"] or "",
                    "Date_Reglement": parse_date(meta["date_reglement"]) if meta["date_reglement"] else "",
                    "Date_Soins": parse_date(date),
                    "Nom_Agent": benef,
                    "Matricule": matricule,
                    "Numero_Facture_Prescription": meta["facture"] or "",
                    "Code_Acte": acte,
                    "Libelle_Acte": "",
                    "Montant_Reclame_Brut": num(reclame),
                    "Ticket_Moderateur": tm_montant,
                    "Montant_Paye_Regle": num(regle),
                    "Montant_Exclu_Rejet": num(nonremb),
                    "Motif_Observation": obs,
                })
    return meta, lignes


# --------------------------------------------------------------------------
# Écriture Excel (mise en forme du modèle)
# --------------------------------------------------------------------------
def style_sheet(ws):
    model_ws = load_workbook(MODEL)[SHEET]
    widths = {k: v.width for k, v in model_ws.column_dimensions.items()}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for r in range(1, ws.max_row + 1):
        for c in range(1, 14):
            ws.cell(row=r, column=c).font = openpyxl.styles.Font(name="Calibri", size=12)
    for r in range(2, ws.max_row + 1):
        for c in (9, 10, 11, 12):  # montants
            ws.cell(row=r, column=c).number_format = "#,##0"
    ws.freeze_panes = "A2"


def write_workbook(path, lignes):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.append(HEADERS)
    for l in lignes:
        ws.append([l[h] for h in HEADERS])
    style_sheet(ws)
    wb.save(path)


# --------------------------------------------------------------------------
def main():
    args = [a for a in sys.argv[1:] if a != "--force"]
    force = "--force" in sys.argv[1:]

    # PDF à convertir : ceux passés en argument, sinon tous les PDF du dossier.
    pdfs = []
    if args:
        for a in args:
            p = os.path.join(PDF_DIR, a) if not os.path.isabs(a) else a
            if os.path.isfile(p) and p.lower().endswith(".pdf"):
                pdfs.append(os.path.abspath(p))
            else:
                print(f"!! PDF introuvable : {a}")
    else:
        pdfs = sorted(glob.glob(os.path.join(PDF_DIR, "*.pdf")))
    if not pdfs:
        if args:
            print("!! Aucun des PDF demandés n'a été trouvé")
        else:
            print(f"!! Aucun PDF trouvé dans {PDF_DIR}")
        return

    for pdf_path in pdfs:
        nom_pdf = os.path.basename(pdf_path)
        try:
            pdf = pdfplumber.open(pdf_path)
        except Exception as e:
            print(f"!! {nom_pdf} : PDF illisible ({e}) -> ignoré")
            continue
        with pdf:
            text = full_pdf_text(pdf)
            if "DECOMPTE DE REGLEMENT FACTURES" not in text:
                print(f"!! {nom_pdf} : ce n'est pas un décompte MCI CARE -> ignoré "
                      f"(ce dossier est réservé aux PDF MCI CARE)")
                continue
            meta, lignes = parse_mci(pdf, nom_pdf)

        if not lignes:
            print(f"!! {nom_pdf} : aucune ligne trouvée -> ignoré")
            continue

        # --- Nom du fichier : MCI CARE MOIS ANNEE PERIODE MONTANT ---
        dr = (meta.get("date_reglement") or "").strip()
        if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", dr):
            print(f"!! {nom_pdf} : date comptable introuvable -> ignoré")
            continue
        date_reglement = parse_date(dr)          # 'AAAA-MM-JJ'

        # Montant payé = somme des "Montant réglé" des lignes (le total
        # prestataire du PDF sert seulement de contrôle, voir ci-dessous).
        total_paye = sum(l["Montant_Paye_Regle"] for l in lignes)
        ref = f"facture {meta['facture']}" if meta["facture"] else "décompte"
        if meta["total_prestataire"] is not None:
            somme = sum(l["Montant_Paye_Regle"] for l in lignes)
            if abs(somme - meta["total_prestataire"]) >= 1:
                print(f"   !! ATTENTION : {fmt_amount(somme)} Ar en lignes "
                      f"≠ total prestataire ({fmt_amount(meta['total_prestataire'])} Ar)")

        out = os.path.join(PDF_DIR,
                           nom_sortie(SOCIETE, date_reglement, lignes, total_paye))
        if os.path.exists(out) and not force:
            print(f"-- {os.path.basename(out)} : existe déjà, non écrasé "
                  f"(--force pour régénérer)  [{nom_pdf}]")
            continue
        write_workbook(out, lignes)
        print(f"OK {os.path.basename(out)} : {ref} | {len(lignes)} lignes | "
              f"Payé {fmt_amount(total_paye)} Ar  <- {nom_pdf}")


if __name__ == "__main__":
    main()
