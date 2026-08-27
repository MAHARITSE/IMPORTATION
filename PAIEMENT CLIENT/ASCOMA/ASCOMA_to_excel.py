#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de conversion ASCOMA -> Excel
Ce script suit la logique des autres scripts (BSA, MCI CARE) du dossier PAIEMENT CLIENT.
Il utilise les fonctions detect_kind, parse_mci, write_workbook, style_sheet definies
 dans pdf_paiement_to_excel.py (parent directory).

Le PDF ASCOMA est au format MCI (Décompte de Règlement Tiers Payant).
"""

import os
import sys
import glob
import re
import unicodedata
from openpyxl import Workbook
import pdfplumber

# Ajouter le répertoire parent pour importer pdf_paiement_to_excel
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from pdf_paiement_to_excel import (
    full_pdf_text,
    parse_mci,
    write_workbook,
    style_sheet,
    HEADERS,
    amount_to_float,
    fmt_amount,
    parse_date,
    PDF_DIR as MAIN_PDF_DIR,
    MODEL as MAIN_MODEL,
    societe_from_content,
)

# Détection de type enrichie pour prendre en compte les PDF ASCOMA dont le texte
# contient "Décompte de Règlement Tiers Payant" (avec accent) plutôt que
# exactement "DECOMPTE DE REGLEMENT FACTURES".
def detect_kind_ascoma(text):
    """Retourne 'mci' si le texte ressemble à un décompte tiers payant."""
    # Vérifications spécifiques ASCOMA
    if "Tiers Payant" in text or "Décompte de Règlement" in text or "Règlement Tiers Payant" in text:
        return "mci"
    # Ensuite on utilise la fonction du parent (qui cherche RELEVE DE REMBOURSEMENTS / DECOMPTE DE REGLEMENT FACTURES)
    from pdf_paiement_to_excel import detect_kind as _detect_kind_parent
    return _detect_kind_parent(text)

# Chemins spécifiques à ASCOMA
PDF_SUBDIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PDF")   # PAIEMENT CLIENT/ASCOMA/PDF
PDF_DIR = os.path.dirname(os.path.abspath(__file__))                         # PAIEMENT CLIENT/ASCOMA
MODEL = os.path.join(os.path.dirname(PDF_DIR), "Modele_Import_Reglements_Decompte_Assurance.xlsx")  # racine


def main():
    # Trouver tous les PDF dans le sous-dossier PDF (un niveau)
    pdfs = sorted(glob.glob(os.path.join(PDF_SUBDIR, "*.pdf")))
    if not pdfs:
        # Essayer récursif
        pdfs = sorted(glob.glob(os.path.join(PDF_DIR, "**", "*.pdf"), recursive=True))
    if not pdfs:
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
            text = full_pdf_text(pdf)
            kind = detect_kind_ascoma(text)
            if kind not in ("bsa", "mci"):
                print(f"!! {nom_pdf} : type inconnu ({kind}) -> ignoré")
                continue

            # Extraction via la fonction parse_mci existante
            meta, lignes = parse_mci(pdf, nom_pdf)
            if not lignes:
                print(f"~ {nom_pdf} : aucune ligne extraite (structure differente) -> creation d'un fichier Excel vide avec en-têtes")
                # Tout de même créer un fichier Excel avec les en-têtes pour avoir la structure
                societe = "ASCOMA"
                parent = os.path.basename(os.path.dirname(os.path.abspath(pdf_path)))
                if parent != os.path.basename(PDF_DIR):
                    societe = parent
                out_dir = os.path.join(os.path.dirname(PDF_DIR), societe)
                os.makedirs(out_dir, exist_ok=True)
                out_name = f"{societe} MONTANT 0Ar.xlsx"
                out_name = re.sub(r"\s+", " ", out_name).strip()
                out_path = os.path.join(out_dir, out_name)
                wb = Workbook()
                ws = wb.active
                ws.title = "Modele_Reglements"
                ws.append(HEADERS)
                # Ajouter une ligne indiquant que les données n'ont pas été extraites
                ws.append([""] * len(HEADERS))
                wb.save(out_path)
                print(f"OK {out_path} : créé avec en-têtes uniquement  <- {nom_pdf}")
                continue

            # Société : on peut déduire du dossier ou du contenu
            parent = os.path.basename(os.path.dirname(os.path.abspath(pdf_path)))
            if parent != os.path.basename(PDF_DIR):
                societe = parent
            else:
                societe = societe_from_content(text) or "ASCOMA"

            # Date de réglement depuis métadonnées
            dr = meta.get("date_reglement") or ""
            # Le format dans le PDF est JJ/MM/AAAA ; on convertit
            if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", dr):
                print(f"!! {nom_pdf} : date de reglement introuvable -> ignoré")
                continue
            date_reglement = parse_date(dr)  # 'AAAA-MM-JJ'

            # Année pour le sous-dossier
            annee = date_reglement.split("-")[0] if date_reglement else "0"

            # Chemin de sortie : PAIEMENT CLIENT/<SOCIETE>/<ANNEE>/
            out_dir = os.path.join(os.path.dirname(PDF_DIR), societe, annee)
            os.makedirs(out_dir, exist_ok=True)

            # Montant total payé
            total_paye = sum(l.get("Montant_Paye_Regle", 0) for l in lignes)

            # Nom de fichier selon le modèle existant : <SOCIETE> <MOIS> <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
            # On récupère le mois en français depuis la date_reglement
            mois_map = {
                "01": "Janvier", "02": "Fevrier", "03": "Mars", "04": "Avril",
                "05": "Mai", "06": "Juin", "07": "Juillet", "08": "Aout",
                "09": "Septembre", "10": "Octobre", "11": "Novembre", "12": "Decembre",
            }
            mois_num = date_reglement.split("-")[1] if date_reglement else "01"
            mois_fr = mois_map.get(mois_num, "Inconnu")

            # Période soins (1ère et dernière date de soins)
            dates_soins = [l.get("Date_Soins", "") for l in lignes if l.get("Date_Soins")]
            dates_iso = [d for d in dates_soins if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)]
            if dates_iso:
                debut, fin = dates_iso[0], dates_iso[-1]
                # Format JJ-MM-AA
                periode = f"{debut[8:10]}-{debut[5:7]} à {fin[8:10]}-{fin[5:7]}"
            else:
                # fallback: utiliser date_reglement en JJ-MM-AA
                import datetime as _dt
                try:
                    dt = _dt.datetime.strptime(date_reglement, "%Y-%m-%d")
                    periode = f"{dt.day:02d}-{dt.month:02d} à {dt.day:02d}-{dt.month:02d}"
                except Exception:
                    periode = ""

            # Construction du nom de fichier (similaire à nom_sortie)
            def date_courte(iso):
                annee2, mois, jour = iso.split("-")
                return f"{jour}-{mois}-{annee2[2:]}"

            if dates_iso:
                debut, fin = dates_iso[0], dates_iso[-1]
                if debut == fin:
                    periode = date_courte(debut)
                else:
                    periode = f"{date_courte(debut)} à {date_courte(fin)}"
            else:
                periode = ""

            # Montant formaté
            montant_fmt = fmt_amount(total_paye)

            # Nom final : SOCIETE MOIS ANNEE PERIODE MONTANT <montant>Ar.xlsx
            out_name = f"{societe} {mois_fr} {annee} {periode} MONTANT {montant_fmt}Ar.xlsx"
            # Nettoyer multiples espaces
            out_name = re.sub(r"\s+", " ", out_name).strip()
            out_path = os.path.join(out_dir, out_name)

            # Vérifier si existe déjà
            if os.path.exists(out_path):
                print(f"-- {out_path} : existe deja (--force pour ecraser)  [{nom_pdf}]")
                continue

            # Création du classeur Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Modele_Reglements"
            # En-têtes
            ws.append(HEADERS)
            # Lignes de données
            for l in lignes:
                row = [
                    l.get("Ref_Decompte", ""),
                    l.get("Date_Reglement", ""),
                    l.get("Date_Soins", ""),
                    l.get("Nom_Agent", ""),
                    l.get("Matricule", ""),
                    l.get("Numero_Facture_Prescription", ""),
                    l.get("Code_Acte", ""),
                    l.get("Libelle_Acte", ""),
                    l.get("Montant_Reclame_Brut", 0.0),
                    l.get("Ticket_Moderateur", 0.0),
                    l.get("Montant_Paye_Regle", 0.0),
                    l.get("Montant_Exclu_Rejet", 0.0),
                    l.get("Motif_Observation", ""),
                ]
                ws.append(row)
            # Mise en forme style modèle
            style_sheet(ws)
            # Sauvegarder
            wb.save(out_path)
            print(f"OK {out_path} : {len(lignes)} lignes | Paye {fmt_amount(total_paye)} Ar  <- {nom_pdf}")


if __name__ == "__main__":
    main()
