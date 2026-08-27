#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de conversion ASCOMA -> Excel
Ce script suit la logique des autres scripts (BSA, MCI CARE) du dossier PAIEMENT CLIENT.
Il utilise les fonctions de pdf_paiement_to_excel.py (dossier parent).

Le PDF ASCOMA est un « Décompte de Règlement Tiers Payant » :
  - les PDF à convertir sont déposés dans le sous-dossier  ASCOMA/PDF/
  - l'Excel produit est rangé dans  PAIEMENT CLIENT/<SOCIETE>/<ANNEE>/
    nom : <SOCIETE> <MOIS> <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx

Utilisation :
  python ASCOMA_to_excel.py                  tous les PDF du sous-dossier PDF/
  python ASCOMA_to_excel.py --force          régénérer (écrase l'Excel existant)
  python ASCOMA_to_excel.py "mon.pdf"        un seul PDF
"""

import os
import sys
import glob
import re
from openpyxl import Workbook
import pdfplumber

# Ajouter le répertoire parent pour importer pdf_paiement_to_excel
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from pdf_paiement_to_excel import (
    full_pdf_text,
    write_workbook,
    societe_from_content,
    num,
    parse_date,
    fmt_amount,
    amount_to_float,
    HEADERS,
    MONTANT_RE,
)

# Chemins spécifiques à ASCOMA
ICI = os.path.dirname(os.path.abspath(__file__))                    # PAIEMENT CLIENT/ASCOMA
PDF_SUBDIR = os.path.join(ICI, "PDF")                               # ASCOMA/PDF  (PDF à convertir)
PDF_DIR = ICI
MODEL = os.path.join(os.path.dirname(ICI), "Modele_Import_Reglements_Decompte_Assurance.xlsx")
SOCIETE = "ASCOMA"

DATE_JJ_MM_AAAA = re.compile(r"\d{2}/\d{2}/\d{4}")

# Montants ASCOMA : la notation « 4 761 800,0 » (une seule décimale) existe
# dans les sous-totaux — MONTANT_RE du parent s'arrête avant le « ,0 ».
MONTANT_ACOMA_RE = re.compile(
    r"\d{1,3}(?!\d)(?:[ \u00a0]\d{3})*(?:[.,]\d{3})?(?:[.,]\d{1,2})?"
    r"|\d+(?:[.,]\d+)?")


def dernier_montant_ascoma(ligne):
    """Dernier montant d'une ligne de total ASCOMA (gère « 4 761 800,0 »)."""
    trouves = MONTANT_ACOMA_RE.findall(ligne)
    return amount_to_float(trouves[-1]) if trouves else 0.0


# --------------------------------------------------------------------------
# Format ASCOMA : DECOMPTE DE REGLEMENT TIERS PAYANT
#   Colonnes du tableau :
#   Date des Soins | Matricule Bénéficiaire | Bénéficiaire | Acte Médical |
#   Code Rem. | Qté | Montant Réclamé | Montant Exclu | Base de Règlement |
#   Ticket Modérateur | Montant Réglé
# --------------------------------------------------------------------------
def parse_ascoma(pdf, nom_pdf):
    """Extrait les méta et les lignes de soin d'un décompte Tiers Payant ASCOMA."""
    text = full_pdf_text(pdf)

    meta = {"date_reglement": None, "centre": None,
            "total_net": None, "total_stotaux": None}

    # Date de règlement : « Règlement du 09/01/2025 au 09/01/2025 »
    # (peut être répartie sur plusieurs lignes), sinon « Edité le 24/01/2025 ».
    m = re.search(r"R[èe]glement\s*du\s*:?\s*(\d{2}/\d{2}/\d{4})", text)
    if not m:
        m = re.search(r"[ÉE]dit[ée]\s*le\s*(\d{2}/\d{2}/\d{4})", text)
    if m:
        meta["date_reglement"] = m.group(1)

    m = re.search(r"Centre\s*:\s*(.+?)\s+Code\s*:", text)
    if m:
        meta["centre"] = m.group(1).strip()

    # Montant NET payé (après remise) : ligne de récapitulatif qui commence
    # par la quantité et contient tous les totaux — le dernier montant de la
    # ligne est la colonne « Montant Net ».
    for ligne in text.splitlines():
        if re.match(r"\s*\d{1,4}\s+\d", ligne):
            if len(MONTANT_RE.findall(ligne)) >= 10:
                meta["total_net"] = dernier_montant_ascoma(ligne)

    # Somme des « S/Total Prestataire » (un par établissement) = total réglé
    stotaux = [dernier_montant_ascoma(l)
               for l in re.findall(r"S/Total Prestataire[^\n]*", text)]
    if stotaux:
        meta["total_stotaux"] = sum(stotaux)

    lignes = []
    for page in pdf.pages:
        for table in page.extract_tables():
            if not table or len(table[0]) < 11:
                continue
            # Tableau ASCOMA : la 1re colonne est « Date des Soins »
            entete0 = str(table[0][0] or "").strip()
            if not entete0.startswith("Date des"):
                continue
            for row in table[1:]:
                c = [str(x or "").strip() for x in row[:11]]
                if len(c) < 11 or not DATE_JJ_MM_AAAA.fullmatch(c[0]):
                    continue            # entête répétée, ligne de total, etc.
                (date_soins, matricule, benef, acte, code_rem,
                 qte, reclame, exclu, base, tm, regle) = c
                matricule = matricule.replace(" ", "").replace("\u00a0", "")
                if not matricule.isdigit():
                    continue            # pas une ligne de soin
                lignes.append({
                    "Ref_Decompte": "",
                    "Date_Reglement": parse_date(meta["date_reglement"]) if meta["date_reglement"] else "",
                    "Date_Soins": parse_date(date_soins),
                    "Nom_Agent": benef,
                    "Matricule": matricule,
                    "Numero_Facture_Prescription": "",
                    "Code_Acte": code_rem,
                    "Libelle_Acte": acte,
                    "Montant_Reclame_Brut": num(reclame),
                    "Ticket_Moderateur": num(tm),
                    "Montant_Paye_Regle": num(regle),
                    "Montant_Exclu_Rejet": num(exclu),
                    "Motif_Observation": "",
                })
    return meta, lignes


def societe_du_pdf(pdf_path, text):
    """Société = nom du sous-dossier contenant le PDF (ex : ASCOMA/BNI/...).
    Le sous-dossier PDF/ (réservé aux dépôts) et le dossier ASCOMA lui-même
    ne comptent pas : on retombe alors sur ASCOMA."""
    parent = os.path.basename(os.path.dirname(os.path.abspath(pdf_path)))
    if parent not in (os.path.basename(PDF_DIR), "PDF"):
        return parent
    return societe_from_content(text) or SOCIETE


def main():
    args = [a for a in sys.argv[1:] if a != "--force"]
    force = "--force" in sys.argv[1:]

    # PDF à convertir : ceux passés en argument, sinon tous les PDF du
    # sous-dossier PDF/ (puis du dossier ASCOMA, puis récursivement).
    if args:
        pdfs = []
        for a in args:
            if os.path.isabs(a):
                p = a
            else:
                # nom relatif : d'abord dans le sous-dossier PDF/, sinon
                # directement dans le dossier ASCOMA
                p = os.path.join(PDF_SUBDIR, a)
                if not os.path.isfile(p):
                    p = os.path.join(ICI, a)
            if os.path.isfile(p) and p.lower().endswith(".pdf"):
                pdfs.append(os.path.abspath(p))
            else:
                print(f"!! PDF introuvable : {a}")
    else:
        pdfs = sorted(glob.glob(os.path.join(PDF_SUBDIR, "*.pdf")))
        if not pdfs:
            pdfs = sorted(glob.glob(os.path.join(PDF_DIR, "*.pdf")))
        if not pdfs:
            pdfs = sorted(glob.glob(os.path.join(PDF_DIR, "**", "*.pdf"), recursive=True))
    if not pdfs:
        if args:
            print("!! Aucun des PDF demandés n'a été trouvé")
        else:
            print(f"!! Aucun PDF trouvé dans {PDF_SUBDIR}")
        return

    for pdf_path in pdfs:
        nom_pdf = os.path.basename(pdf_path)
        try:
            pdf = pdfplumber.open(pdf_path)
        except Exception as e:
            print(f"!! {nom_pdf} : PDF illisible ({e}) -> ignoré")
            continue
        with pdf:
            # C'est l'utilisateur qui classe les PDF dans le dossier ASCOMA :
            # on parse toujours, sans filtrer sur le titre du document.
            text = full_pdf_text(pdf)
            meta, lignes = parse_ascoma(pdf, nom_pdf)

        societe = societe_du_pdf(pdf_path, text)
        base_dir = os.path.dirname(ICI)                     # PAIEMENT CLIENT

        if not lignes:
            print(f"~ {nom_pdf} : aucune ligne extraite (structure differente) "
                  f"-> creation d'un fichier Excel vide avec en-têtes")
            out_dir = os.path.join(base_dir, societe)
            os.makedirs(out_dir, exist_ok=True)
            out_name = re.sub(r"\s+", " ", f"{societe} MONTANT 0Ar.xlsx").strip()
            out_path = os.path.join(out_dir, out_name)
            wb = Workbook()
            ws = wb.active
            ws.title = "Modele_Reglements"
            ws.append(HEADERS)
            ws.append([""] * len(HEADERS))
            wb.save(out_path)
            print(f"OK {out_path} : créé avec en-têtes uniquement  <- {nom_pdf}")
            continue

        # Date de règlement (JJ/MM/AAAA) -> AAAA-MM-JJ
        dr = (meta.get("date_reglement") or "").strip()
        if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", dr):
            print(f"!! {nom_pdf} : date de reglement introuvable -> ignoré")
            continue
        date_reglement = parse_date(dr)
        annee = date_reglement.split("-")[0]

        # Contrôles automatiques
        total_lignes = sum(l["Montant_Paye_Regle"] for l in lignes)
        total_paye = total_lignes
        if meta["total_stotaux"] is not None and abs(meta["total_stotaux"] - total_lignes) > 1:
            print(f"   !! ATTENTION : {fmt_amount(total_lignes)} Ar lus ≠ "
                  f"S/Total Prestataire {fmt_amount(meta['total_stotaux'])} Ar")
        if meta["total_net"] is not None:
            if meta["total_net"] <= total_lignes + 1:
                total_paye = meta["total_net"]
                remise = total_lignes - total_paye
                if remise > 1:
                    print(f"   contrôle : Montant Réglé {fmt_amount(total_lignes)} Ar "
                          f"− remise {fmt_amount(remise)} Ar = "
                          f"net {fmt_amount(total_paye)} Ar (récapitulatif)  OK")
            else:
                print(f"   !! ATTENTION : net récapitulatif {fmt_amount(meta['total_net'])} Ar "
                      f"> somme des lignes {fmt_amount(total_lignes)} Ar")

        # Mois en français (sans accent) pour le nom du fichier
        mois_map = {
            "01": "Janvier", "02": "Fevrier", "03": "Mars", "04": "Avril",
            "05": "Mai", "06": "Juin", "07": "Juillet", "08": "Aout",
            "09": "Septembre", "10": "Octobre", "11": "Novembre", "12": "Decembre",
        }
        mois_fr = mois_map.get(date_reglement.split("-")[1], "Inconnu")

        # Période de soins : 1re et dernière date (format JJ-MM-AA)
        def date_courte(iso):
            annee2, mois, jour = iso.split("-")
            return f"{jour}-{mois}-{annee2[2:]}"

        dates_iso = [l["Date_Soins"] for l in lignes
                     if re.fullmatch(r"\d{4}-\d{2}-\d{2}", l["Date_Soins"])]
        if dates_iso:
            debut, fin = dates_iso[0], dates_iso[-1]
            periode = date_courte(debut) if debut == fin \
                else f"{date_courte(debut)} à {date_courte(fin)}"
        else:
            periode = ""

        # Chemin de sortie : PAIEMENT CLIENT/<SOCIETE>/<ANNEE>/
        out_dir = os.path.join(base_dir, societe, annee)
        os.makedirs(out_dir, exist_ok=True)
        out_name = re.sub(r"\s+", " ",
                          f"{societe} {mois_fr} {annee} {periode} "
                          f"MONTANT {fmt_amount(total_paye)}Ar.xlsx").strip()
        out_path = os.path.join(out_dir, out_name)

        if os.path.exists(out_path) and not force:
            print(f"-- {out_path} : existe déjà (--force pour écraser)  [{nom_pdf}]")
            continue

        write_workbook(out_path, lignes)
        print(f"OK {out_path} : {len(lignes)} lignes | "
              f"Payé {fmt_amount(total_paye)} Ar  <- {nom_pdf}")


if __name__ == "__main__":
    main()
