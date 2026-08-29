#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de conversion ASCOMA -> Excel
Ce script suit la logique des autres scripts (BSA, MCI CARE) du dossier PAIEMENT CLIENT.
Il utilise les fonctions de pdf_paiement_to_excel.py (dossier parent).

Le PDF ASCOMA est un « Décompte de Règlement Tiers Payant » :
  - les PDF à convertir sont déposés dans le sous-dossier  ASCOMA/PDF/
  - l'Excel produit est rangé dans
    PAIEMENT CLIENT/<SOCIETE>/<ANNEE_REGLEMENT>/<ANNEE_SOINS>/
    nom : <DATE_PAIEMENT> <SOCIETE> <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
    <ANNEE_REGLEMENT> = année du règlement, <ANNEE_SOINS> = année de la
    période de soins payée (1re date de soins) : un paiement de cette année
    peut régler des soins de l'année dernière.

Utilisation :
  python ASCOMA_to_excel.py                  tous les PDF du sous-dossier PDF/
  python ASCOMA_to_excel.py --force          régénérer (écrase l'Excel existant)
  python ASCOMA_to_excel.py "mon.pdf"        un seul PDF

PDF en erreur : si un PDF ne peut pas être converti (PDF illisible, format
non reconnu, aucune ligne trouvée, date introuvable, erreur pendant la
création de l'Excel), il est DÉPLACÉ dans le sous-dossier ASCOMA/ERREUR/
(créé automatiquement). Les PDF du dossier ERREUR ne sont pas retraités :
remettez le PDF dans PDF/ après correction pour réessayer.
"""

import os
import sys
import glob
import re
import shutil
import pdfplumber

# Ajouter le répertoire parent pour importer pdf_paiement_to_excel
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from pdf_paiement_to_excel import (
    full_pdf_text,
    write_workbook,
    num,
    parse_date,
    fmt_amount,
    amount_to_float,
    HEADERS,
    MONTANT_RE,
    nom_sortie,
    annee_soins,
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


# --------------------------------------------------------------------------
# Dossier des PDF en échec : ASCOMA/ERREUR (créé automatiquement si besoin).
# Tout PDF impossible à convertir y est DÉPLACÉ, pour ne pas le confondre
# avec les PDF en attente de conversion. Les PDF du dossier ERREUR ne sont
# jamais retraités par les conversions suivantes : après correction, on
# remet le PDF dans PDF/ pour réessayer.
# --------------------------------------------------------------------------
ERREUR_DIRNAME = "ERREUR"
ERREURS = []        # [(nom du PDF, raison)] — récapitulatif en fin de traitement


def dossier_erreur():
    """Chemin du sous-dossier ASCOMA/ERREUR (créé si absent)."""
    d = os.path.join(ICI, ERREUR_DIRNAME)
    os.makedirs(d, exist_ok=True)
    return d


def est_dans_erreur(pdf_path):
    """Vrai si le PDF se trouve déjà dans le sous-dossier ERREUR."""
    parties = os.path.normpath(os.path.abspath(pdf_path)).split(os.sep)
    return ERREUR_DIRNAME in parties


def deplacer_pdf_en_erreur(pdf_path, raison):
    """Déplace un PDF impossible à convertir dans ASCOMA/ERREUR.

    Ne bouge pas le PDF si le déplacement échoue (fichier verrouillé...) :
    le problème est affiché et le PDF sera retraité au prochain lancement.
    """
    dest = os.path.join(dossier_erreur(), os.path.basename(pdf_path))
    base, ext = os.path.splitext(dest)
    n = 1
    while os.path.exists(dest):     # ne pas écraser un PDF déjà signalé
        dest = f"{base} ({n}){ext}"
        n += 1
    try:
        shutil.move(pdf_path, dest)
        print(f"   -> PDF déplacé dans {ERREUR_DIRNAME} : "
              f"{os.path.relpath(dest, ICI)}")
        ERREURS.append((os.path.basename(pdf_path), raison))
    except Exception as e:
        print(f"   !! déplacement impossible vers {ERREUR_DIRNAME} : {e}")
        ERREURS.append((os.path.basename(pdf_path), raison + f" (non déplacé : {e})"))


def recap_erreurs():
    """Affiche le récapitulatif des PDF en erreur (fin de traitement)."""
    if ERREURS:
        print(f"\n== {len(ERREURS)} PDF en erreur, déplacés dans "
              f"le sous-dossier {ERREUR_DIRNAME} de {SOCIETE} ==")
        for nom, raison in ERREURS:
            print(f"   - {nom} : {raison}")


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
    pdfs = [p for p in pdfs if not est_dans_erreur(p)]   # ERREUR : on ne retente pas
    if not pdfs:
        if args:
            print("!! Aucun des PDF demandés n'a été trouvé")
        else:
            print(f"!! Aucun PDF trouvé dans {PDF_SUBDIR}")
        recap_erreurs()
        return

    for pdf_path in pdfs:
        nom_pdf = os.path.basename(pdf_path)
        try:
            pdf = pdfplumber.open(pdf_path)
        except Exception as e:
            print(f"!! {nom_pdf} : PDF illisible ({e})")
            deplacer_pdf_en_erreur(pdf_path, "PDF illisible ou corrompu")
            continue
        try:
            with pdf:
                # Dossier ASCOMA → parseur ASCOMA. Le contenu du PDF ne change rien.
                meta, lignes = parse_ascoma(pdf, nom_pdf)
        except Exception as e:
            print(f"!! {nom_pdf} : erreur pendant la lecture du PDF "
                  f"({type(e).__name__}: {e})")
            deplacer_pdf_en_erreur(pdf_path, "erreur de lecture (format non reconnu ?)")
            continue

        societe = SOCIETE  # toujours ASCOMA : c'est le dossier qui décide
        base_dir = os.path.dirname(ICI)                     # PAIEMENT CLIENT

        # Date de règlement (JJ/MM/AAAA) -> AAAA-MM-JJ
        dr = (meta.get("date_reglement") or "").strip()
        date_reglement = parse_date(dr) if re.fullmatch(r"\d{2}/\d{2}/\d{4}", dr) else ""
        annee = date_reglement.split("-")[0] if date_reglement else ""

        if not lignes:
            print(f"!! {nom_pdf} : aucune ligne extraite (structure differente ?)")
            deplacer_pdf_en_erreur(pdf_path,
                                   "aucune ligne de règlement trouvée dans le PDF")
            continue

        if not date_reglement:
            print(f"!! {nom_pdf} : date de reglement introuvable dans le PDF")
            deplacer_pdf_en_erreur(pdf_path, "date de règlement introuvable")
            continue

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

        # Chemin de sortie : PAIEMENT CLIENT/<SOCIETE>/<ANNEE_REGLEMENT>/<ANNEE_SOINS>/
        # (un paiement de cette année peut régler des soins de l'année dernière)
        soins = annee_soins(lignes, date_reglement)
        out_dir = os.path.join(base_dir, societe, annee, soins)
        os.makedirs(out_dir, exist_ok=True)
        out_name = nom_sortie(societe, date_reglement, lignes, total_paye)
        out_path = os.path.join(out_dir, out_name)

        if os.path.exists(out_path) and not force:
            print(f"-- {out_path} : existe déjà (--force pour écraser)  [{nom_pdf}]")
            continue

        try:
            write_workbook(out_path, lignes)
        except Exception as e:
            print(f"!! {nom_pdf} : erreur pendant la création de l'Excel "
                  f"({type(e).__name__}: {e})")
            deplacer_pdf_en_erreur(pdf_path, "erreur de création de l'Excel")
            continue
        print(f"OK {out_path} : {len(lignes)} lignes | "
              f"Payé {fmt_amount(total_paye)} Ar  <- {nom_pdf}")

    recap_erreurs()


if __name__ == "__main__":
    main()
