# -*- coding: utf-8 -*-
"""
Conversion des factures PDF SALFA (BSA) en fichiers Excel selon le modèle
"Modele_Import.xlsx" (feuille Modele_Prestations, 11 colonnes).

Usage : python3 pdf_to_excel.py
Entrée : FACTURE CLIENT/BSA *.pdf
Sortie : FACTURE CLIENT/EXCEL/BSA <MOIS>.xlsx  +  fichier consolidé
"""
import re
import glob
import os
import unicodedata
import pdfplumber
import openpyxl.styles
from openpyxl import Workbook, load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(BASE, "FACTURE CLIENT")
OUT_DIR = os.path.join(PDF_DIR, "EXCEL")
MODEL = os.path.join(PDF_DIR, "Modele_Import.xlsx")

HEADERS = ["Numero_Facture", "Date_Soins", "Nom_Agent", "Matricule", "Societe",
           "Sous_Societe", "Acte_Medicale_Prix", "Montant_Total_Brut",
           "Ticket_Moderateur", "Prise_En_Charge_Net", "Observations"]

MONTH_ORDER = ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN",
               "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"]


def amount_to_float(s):
    """'1 234,56' -> 1234.56"""
    return float(s.replace("\u00a0", " ").replace(" ", "").replace(",", "."))


def fmt_amount(n):
    """1234.0 -> '1 234' ; 1234.5 -> '1 234,5'"""
    if abs(n - round(n)) < 0.005:
        return f"{int(round(n)):,}".replace(",", " ")
    return f"{n:,.1f}".replace(",", " ").replace(".", ",")


def parse_date(d):
    """'02/01/26' -> '2026-01-02'"""
    dd, mm, yy = d.strip().split("/")
    year = int(yy)
    year += 2000 if year < 100 else 0
    return f"{year:04d}-{mm}-{dd}"


def parse_pdf(path):
    """Extrait (facture, mois, lignes) d'un PDF. Chaque ligne = dict modèle."""
    with pdfplumber.open(path) as pdf:
        facture = mois = None
        raw_rows = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for line in txt.split("\n"):
                m = re.search(r"Facture N°\s*:\s*(\S+)", line)
                if m:
                    facture = m.group(1)
                m = re.search(r"Mois de prise en charge\s*:\s*(.+)", line)
                if m:
                    mois = unicodedata.normalize("NFC", m.group(1).strip())
            for table in page.extract_tables():
                for row in table:
                    if not row or not str(row[0] or "").strip():
                        continue  # artefact de saut de page / ligne vide
                    if str(row[0]).strip() in ("N°", "Total"):
                        continue
                    raw_rows.append(row)

    lignes = []
    for row in raw_rows:
        num, date, matricule, nom_cell, actes_cell, montant, part, net = row[:8]

        # --- Nom / sous-société ---
        parts = [p.strip() for p in (nom_cell or "").split("\n") if p.strip()]
        sous = ""
        name_parts = []
        for p in parts:
            m = re.fullmatch(r"\((.+)\)", p)
            if m:
                sous = m.group(1).strip()
            else:
                name_parts.append(p)
        nom = " ".join(name_parts)
        nom_agent = f"{nom} ({sous})" if sous else nom

        # --- Actes médicaux ---
        actes = []
        total_actes = 0.0
        for a in (actes_cell or "").split("\n"):
            a = a.strip()
            if not a:
                continue
            m = re.match(r"^(.*?)\s*:\s*([\d\s\u00a0,\.]+)$", a)
            if m:
                lib, val = m.group(1).strip(), amount_to_float(m.group(2))
                total_actes += val
                actes.append(f"{lib} : {fmt_amount(val)}")
            else:
                actes.append(a)
        actes_str = "; ".join(actes)

        brut = amount_to_float(montant)
        tm = amount_to_float(part)
        net_pay = amount_to_float(net)

        # --- Observations ---
        obs = f"Facture mensuelle soins ambulatoires - {mois}"
        if abs(total_actes - brut) > 0.01:
            obs += f" ; ATTENTION : actes détaillés ({fmt_amount(total_actes)} Ar) < montant facturé (écart {fmt_amount(brut - total_actes)} Ar)"

        lignes.append({
            "Numero_Facture": facture,
            "Date_Soins": parse_date(date),
            "Nom_Agent": nom_agent,
            "Matricule": str(matricule or "").strip(),
            "Societe": "BSA",
            "Sous_Societe": sous,
            "Acte_Medicale_Prix": actes_str,
            "Montant_Total_Brut": int(brut) if brut == int(brut) else brut,
            "Ticket_Moderateur": int(tm) if tm == int(tm) else tm,
            "Prise_En_Charge_Net": int(net_pay) if net_pay == int(net_pay) else net_pay,
            "Observations": obs,
        })
    return facture, mois, lignes


def style_sheet(ws):
    """Applique la mise en forme du modèle (largeurs, police, formats)."""
    model_ws = load_workbook(MODEL)["Modele_Prestations"]
    widths = {k: v.width for k, v in model_ws.column_dimensions.items()}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(name="Calibri", size=12)
    for r in range(2, ws.max_row + 1):
        for c in range(1, 12):
            ws.cell(row=r, column=c).font = openpyxl.styles.Font(name="Calibri", size=12)
        for c in (8, 9, 10):  # montants
            ws.cell(row=r, column=c).number_format = "#,##0"
    ws.freeze_panes = "A2"


def write_workbook(path, lignes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Modele_Prestations"
    ws.append(HEADERS)
    for l in lignes:
        ws.append([l[h] for h in HEADERS])
    style_sheet(ws)
    wb.save(path)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(PDF_DIR, "BSA *.pdf")),
                  key=lambda p: MONTH_ORDER.index(
                      os.path.basename(p).replace("BSA ", "").replace(".pdf", "").upper()))
    toutes = []
    for pdf_path in pdfs:
        facture, mois, lignes = parse_pdf(pdf_path)
        if not lignes:
            print(f"!! {os.path.basename(pdf_path)} : aucune ligne trouvée")
            continue
        stem = os.path.splitext(os.path.basename(pdf_path))[0]
        out = os.path.join(OUT_DIR, f"{stem}.xlsx")
        write_workbook(out, lignes)
        s_brut = sum(l["Montant_Total_Brut"] for l in lignes)
        s_tm = sum(l["Ticket_Moderateur"] for l in lignes)
        s_net = sum(l["Prise_En_Charge_Net"] for l in lignes)
        print(f"OK {stem}.xlsx : {facture} | {mois} | {len(lignes)} lignes | "
              f"Brut {s_brut:,} | TM {s_tm:,} | Net {s_net:,}".replace(",", " "))
        for l in lignes:
            l["Observations"] = l["Observations"]  # déjà renseigné
        toutes.extend(lignes)

    consolide = os.path.join(OUT_DIR, "BSA 2026 CONSOLIDE (Janvier-Juin).xlsx")
    write_workbook(consolide, toutes)
    print(f"OK {os.path.basename(consolide)} : {len(toutes)} lignes au total")


if __name__ == "__main__":
    main()
