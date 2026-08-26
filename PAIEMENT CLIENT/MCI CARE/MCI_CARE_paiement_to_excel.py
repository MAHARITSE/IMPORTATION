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

Sortie : MCI CARE <MOIS> <MONTANT>.xlsx dans ce même dossier
    exemple : MCI CARE Mai 471 140.xlsx
    - MOIS    : mois de la date comptable du décompte
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
    """'1 234,56' -> 1234.56"""
    return float(s.replace("\u00a0", " ").replace(" ", "").replace(",", "."))


def fmt_amount(n):
    """471140.0 -> '471 140' ; 1234.5 -> '1 234,5'"""
    if abs(n - round(n)) < 0.005:
        return f"{int(round(n)):,}".replace(",", " ")
    return f"{n:,.1f}".replace(",", " ").replace(".", ",")


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
    m = re.search(r"Total prestataire:\s*[\d\u00a0 ]+\s+[\d\u00a0 ]+\s+[\d\u00a0 ]+\s+-?\s*([\d\u00a0 ]+)", text)
    if m:
        meta["total_prestataire"] = amount_to_float(m.group(1))

    lignes = []
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table or not str(table[0][0] or "").strip().startswith("Matricule"):
                continue
            for row in table[1:]:
                c = [str(x or "").strip() for x in row[:9]]
                if not c[0] or not c[0][0].isdigit():
                    continue  # lignes totaux / descriptions / vides
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

        # --- Nom du fichier : MCI CARE MOIS MONTANT ---
        dr = (meta.get("date_reglement") or "")
        if len(dr) != 10:
            print(f"!! {nom_pdf} : date comptable introuvable -> ignoré")
            continue
        mois = MONTHS[int(dr[3:5]) - 1]

        total_paye = meta["total_prestataire"] or \
            sum(l["Montant_Paye_Regle"] for l in lignes)
        ref = f"facture {meta['facture']}" if meta["facture"] else "décompte"
        if meta["total_prestataire"] is not None:
            somme = sum(l["Montant_Paye_Regle"] for l in lignes)
            if abs(somme - meta["total_prestataire"]) >= 1:
                print(f"   !! ATTENTION : {fmt_amount(somme)} Ar en lignes "
                      f"≠ total prestataire ({fmt_amount(meta['total_prestataire'])} Ar)")

        out = os.path.join(PDF_DIR, f"{SOCIETE} {mois} {fmt_amount(total_paye)}.xlsx")
        if os.path.exists(out) and not force:
            print(f"-- {os.path.basename(out)} : existe déjà, non écrasé "
                  f"(--force pour régénérer)  [{nom_pdf}]")
            continue
        write_workbook(out, lignes)
        print(f"OK {os.path.basename(out)} : {ref} | {len(lignes)} lignes | "
              f"Payé {fmt_amount(total_paye)} Ar  <- {nom_pdf}")


if __name__ == "__main__":
    main()
