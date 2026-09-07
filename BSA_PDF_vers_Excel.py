# -*- coding: utf-8 -*-
"""
BSA_PDF_vers_Excel.py — société BSA
===================================
Convertit les relevés de remboursements BSA (PDF) en fichiers Excel conformes
au modèle « Modele_Import_Reglements_Decompte_Assurance.xlsx »
(feuille Modele_Reglements, 13 colonnes).

COMMENT LANCER
--------------
  • Le plus simple : double-clic sur  Lancer_BSA.bat
    (il vérifie que Python et les bibliothèques sont installés, lance le
     script, et la fenêtre reste ouverte à la fin).
  • Double-clic direct sur ce fichier .py : possible aussi, la fenêtre reste
     ouverte (message final « Appuyez sur Entrée ») et un journal est écrit
     dans journal_BSA.txt.
  • En ligne de commande :
        python BSA_PDF_vers_Excel.py                 # tous les PDF de PDF\
        python BSA_PDF_vers_Excel.py --force         # régénérer (écrase)
        python BSA_PDF_vers_Excel.py mon_releve.pdf  # un seul PDF
        python BSA_PDF_vers_Excel.py --no-pause      # ne pas attendre à la fin

Si la fenêtre se fermait toute seule auparavant, c'est que Python s'arrêtait
sur une erreur avant d'afficher quoi que ce soit (le plus souvent : les
bibliothèques pdfplumber / openpyxl absentes). Ici, toute erreur est affichée,
écrite dans journal_BSA.txt, et la fenêtre attend une touche avant de fermer.

BIBLIOTHÈQUES NÉCESSAIRES (une seule fois)
------------------------------------------
        python -m pip install pdfplumber openpyxl

FORMAT TRAITÉ : « RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE »
---------------------------------------------------------------
  - 1re page : ordre de virement — « N°: 1090270 », « Lot : 780378 »,
    « A,le 08/07/2025 », « un virement de 282 600,00 MGA » ;
  - puis un bloc par remboursement :
        ligne d'en-tête :  997712-1  23961  MAHARITSY JEAN FLORENT  CGPR  ABT ASSOCIATES (…)
                           n° décompte-ligne | matricule (facultatif) | assuré |
                           code acte | client / contrat
        ligne de données : 04/04/2025  RAZAFINAMBININA  ASSOCIATION DISPENSAIRE LOTERANA
                           20 000,00 | 0,00 | 20,00 | 18 000,00 | 2 000,00 | 0,00
                           date de soins | patient | exécutant |
                           FR.REELS | 1ERE MUT | Tx(%) | REMB | NON REMB | TPG*
        lignes suivantes : suite du nom (colonne de gauche) puis
                           libellé de l'acte / médicament (à droite).
  - fin de décompte : « Total décompte : 997712 » puis
    « Date facture: … N°004-25/BSA/ABT-A » (n° de facture SALFA du décompte) ;
  - dernière page : « Total général » (nombre de lignes + totaux).

  La lecture se fait par COORDONNÉES (colonnes du tableau), pas par texte :
  c'est ce qui permet de lire aussi les relevés passés à l'OCR, dont le texte
  est abîmé (« 20 000,00 » coupé en « 20 » + « 000,00 », dates sans « / »,
  « ADHESION: » absent, « Client: » déformé en « [C lent:| »…).

SORTIE
------
    <ANNEE_REGLEMENT>/<ANNEE_SOINS>/<DATE_PAIEMENT> BSA <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
    exemple : 2025/2025/08-07-25 BSA 2025 04-04-25 à 24-04-25 MONTANT 282 600Ar.xlsx
    - ANNEE_REGLEMENT : année du virement ; ANNEE_SOINS : année de la 1re date
      de soins (un paiement peut régler des soins de l'année précédente) ;
    - DATE_PAIEMENT : date du virement (« A,le 08/07/2025 »), format JJ-MM-AA ;
    - PERIODE : 1re et dernière date de soins du relevé ;
    - MONTANT : somme des montants payés (= montant du virement).
    Les sous-dossiers sont créés automatiquement à côté de ce script.

PDF EN ERREUR
-------------
    Un PDF illisible (scan non OCRisé, format inconnu, aucune ligne lue,
    date de virement absente…) est DÉPLACÉ dans le sous-dossier ERREUR\
    pour ne pas bloquer les conversions suivantes. Remettez-le dans PDF\
    après correction pour réessayer.

Les Excel déjà existants ne sont PAS écrasés (protection des saisies
manuelles), sauf avec --force.
"""

import glob
import os
import re
import shutil
import sys
import traceback
import unicodedata
from datetime import datetime

# ---------------------------------------------------------------------------
# Bibliothèques tierces : import protégé.
# Sans cela, un double-clic ferme la fenêtre instantanément sans explication.
# ---------------------------------------------------------------------------
MANQUANTES = []

try:
    import pdfplumber
except Exception:                                     # pragma: no cover
    pdfplumber = None
    MANQUANTES.append("pdfplumber")

try:
    from openpyxl import Workbook, load_workbook
except Exception:                                     # pragma: no cover
    Workbook = load_workbook = None
    MANQUANTES.append("openpyxl")

# ---------------------------------------------------------------------------
# Emplacements
# ---------------------------------------------------------------------------
DOSSIER = os.path.dirname(os.path.abspath(__file__))   # dossier de la société
PDF_SOUS_DOSSIER = os.path.join(DOSSIER, "PDF")        # PDF à convertir
ERREUR_SOUS_DOSSIER = os.path.join(DOSSIER, "ERREUR")  # PDF non convertis
JOURNAL = os.path.join(DOSSIER, "journal_BSA.txt")     # trace de chaque run
NOM_MODELE = "Modele_Import_Reglements_Decompte_Assurance.xlsx"
SOCIETE = "BSA"
FEUILLE = "Modele_Reglements"

HEADERS = [
    "Ref_Decompte", "Date_Reglement", "Date_Soins", "Nom_Agent", "Matricule",
    "Numero_Facture_Prescription", "Code_Acte", "Libelle_Acte",
    "Montant_Reclame_Brut", "Ticket_Moderateur", "Montant_Paye_Regle",
    "Montant_Exclu_Rejet", "Motif_Observation",
]

# Largeurs de colonnes du modèle (reprises telles quelles si le fichier
# modèle est introuvable : il ne sert qu'à la mise en forme).
LARGEURS_DEFAUT = {"A": 22.8, "B": 14.8, "C": 14.8, "D": 26.8, "E": 16.8,
                   "F": 28.8, "G": 14.8, "H": 32.8, "I": 20.8, "J": 18.8,
                   "K": 20.8, "L": 18.8, "M": 32.8}

# ---------------------------------------------------------------------------
# Géométrie du tableau du relevé.
# Les bordures ne sont pas des traits vectoriels dans ces PDF : on repère donc
# chaque colonne par sa position horizontale, en fraction de la largeur de
# page (mesurée sur les relevés BSA, portrait A4, largeur 587 à 595 points).
# Chaque mot est rangé dans la colonne où commence son bord gauche (x0).
# ---------------------------------------------------------------------------
COLONNES = [
    ("DATE",      0.000),
    ("AYANT",     0.084),   # ayant-droit / patient
    ("EXECUTANT", 0.224),
    ("ACTE",      0.410),   # code de l'acte (CGPR, PH, EB, ECH…)
    ("FR",        0.497),   # FR.REELS
    ("MUT",       0.598),   # 1ERE MUT
    ("TX",        0.694),   # Tx (%)
    ("REMB",      0.765),
    ("NONREMB",   0.822),
    ("TPG",       0.912),
    ("FIN",       1.010),
]
MONTANT_COLONNES = ("FR", "MUT", "TX", "REMB", "NONREMB", "TPG")

DATE_FR = re.compile(r"\d{2}/\d{2}/\d{4}")
BLOC_NUM = re.compile(r"^(\d{5,7})(?:[-–—](\d{1,3}))?$")
BLOC_NUM_FUSIONNE = re.compile(r"^\d{8,10}$")        # OCR : « 997712-13 » -> « 997712413 »
MATRICULE = re.compile(r"^\d{4,7}[A-Za-z]{0,2}$")    # 23961, 1084121, 16480MA…

# N° de facture SALFA :
#   « N°004-25/BSA/ABT-A », « N° 004-25/BSA/ORANGE-M », « FA-02/BFV/26-022 »,
#   « 006-25/ABT/SALFA » (le « N° » manque dans certains relevés OCR).
FACTURE_RE = re.compile(
    r"(FA-\d{2}[-/][\w/\-]*\d|N°\s*\d{2,4}[-/][\w/\-]*\w|\b\d{3}-\d{2}/[\w/\-]*\w)")

# Lignes de page d'en-tête / pied de page / totaux : jamais du contenu de bloc.
EN_TETE_RE = re.compile(
    r"(RELEVE\s+DE\s+REMBOURSEMENTS|^Lot\s*:|^N°\s*:|^Banque\s*:|^Ville\s*:"
    r"|^Page\s*:|^MADAGASCAR$|^Andraharo$|^BSA\s*/|^A\s*,?\s*le\s+\d{2}/"
    r"|AYANT|EXECUTANT|FR\s*\.?\s*REELS|1ERE\s*MUT|TPG\*|^Total\b|^Nbre\b"
    r"|Date\s+facture|Dt[\s-]*[Ff]acture|Facture\s*N°|^Total\s+g[ée]n[ée]ral)",
    re.IGNORECASE)

CARACTERES_INTERDITS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
DATE_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# ---------------------------------------------------------------------------
# Journal : tout ce qui est affiché est aussi écrit dans journal_BSA.txt,
# pour rester consultable même si la fenêtre est fermée trop vite.
# ---------------------------------------------------------------------------
_TAMPON_JOURNAL = []


def log(message=""):
    """Affiche une ligne et la mémorise pour le journal."""
    print(message)
    _TAMPON_JOURNAL.append(str(message))


def ecrire_journal():
    """Recopie la session dans journal_BSA.txt (ajout, jamais d'écrasement)."""
    try:
        with open(JOURNAL, "a", encoding="utf-8") as fh:
            fh.write("\n===== %s =====\n" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            fh.write("\n".join(_TAMPON_JOURNAL) + "\n")
    except Exception as exc:                          # un journal illisible ne
        print("(journal non écrit : %s)" % exc)      # doit jamais tout bloquer


def pause(message="Appuyez sur Entrée pour fermer cette fenêtre…"):
    """Attend une touche (double-clic). Silencieux si l'entrée n'est pas un clavier."""
    try:
        input("\n" + message)
    except (EOFError, KeyboardInterrupt, OSError):
        pass


# ---------------------------------------------------------------------------
# Nombres
# ---------------------------------------------------------------------------
def _nettoie_nombre(brut):
    """'20 000,00' -> '20000,00' ; '4.000,00' -> '4.000,00' ; '—' -> ''."""
    texte = str(brut or "").replace("\u00a0", " ").replace(" ", "")
    return re.sub(r"[^0-9.,]", "", texte)


def montant(brut, pourcentage=False):
    """Transforme un montant du PDF en nombre.

    Notations gérées : '20 000,00' -> 20000.0 ; '4.000,00' -> 4000.0 ;
    '4160.00' -> 4160.0. Un nombre sans séparateur issu de l'OCR
    ('456000', '2600000') est lu avec 2 décimales implicites -> 4560.0,
    26000.0 ; en colonne Tx(%) ('8000') il est lu comme 80,00.
    """
    s = _nettoie_nombre(brut)
    if not s:
        return 0.0
    virgules, points = s.count(","), s.count(".")
    if virgules and points:                            # '4.000,00' / '4,000.00'
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif virgules + points == 1:                       # un seul séparateur
        sep = "," if virgules else "."
        entier, decimales = s.split(sep)
        if len(decimales) == 3 and len(entier) <= 3:   # '75,000' = milliers
            s = entier + decimales
        else:
            s = entier + "." + decimales
    elif virgules > 1:                                 # '1,000,000'
        s = s.replace(",", "")
    elif points > 1:                                   # '1.000.000'
        s = s.replace(".", "")
    else:                                              # aucun séparateur (OCR)
        if pourcentage:
            valeur = float(s)
            return valeur / 100.0 if valeur > 100 else valeur
        if len(s) >= 3:
            s = s[:-2] + "." + s[-2:]
    try:
        return float(s)
    except ValueError:
        return 0.0


def entier_si_possible(valeur):
    return int(valeur) if abs(valeur - round(valeur)) < 0.005 else round(valeur, 2)


def fmt_montant(nombre):
    """928750.0 -> '928 750' ; 1234.5 -> '1 234,5' (affichage console/nom)."""
    if abs(nombre - round(nombre)) < 0.005:
        return "{:,}".format(int(round(nombre))).replace(",", " ")
    return "{:,.1f}".format(nombre).replace(",", " ").replace(".", ",")


# ---------------------------------------------------------------------------
# Dates
# ---------------------------------------------------------------------------
def date_vers_iso(brut):
    """'04/04/2025' -> '2025-04-04'.

    L'OCR perd parfois les « / » : '280042025' (= 28/04/2025) est reconnu tant
    qu'il n'existe qu'UNE seule lecture JJ MM AAAA possible ; s'il y en a
    plusieurs ('0410412025' = 04/10 ou 04/04), la date est laissée vide et
    signalée plutôt qu'inventée. Renvoie ('' , vrai) si illisible,
    (date, faux) si la date a dû être reconstruite.
    """
    texte = str(brut or "").strip()
    m = DATE_FR.search(texte)
    if m:
        jour, mois, annee = m.group(0).split("/")
        return _iso(annee, mois, jour), True

    chiffres = re.sub(r"\D", "", texte)
    if len(chiffres) < 8:
        return "", True
    # 1) lecture directe JJMMAAAA sur une fenêtre de 8 chiffres
    lectures = []
    for i in range(len(chiffres) - 7):
        cand = chiffres[i:i + 8]
        if _date_valide(cand[:2], cand[2:4], cand[4:]):
            lectures.append((cand[:2], cand[2:4], cand[4:], True))
    if not lectures:
        # 2) « / » remplacés par du bruit : JJ … MM … AAAA (2 chiffres de bruit max)
        for i in range(0, 3):
            for j in range(i + 2, min(i + 6, len(chiffres) - 5)):
                jour, mois = chiffres[i:i + 2], chiffres[j:j + 2]
                annee = chiffres[-4:]
                if _date_valide(jour, mois, annee):
                    lectures.append((jour, mois, annee, False))
    if not lectures:
        return "", True
    if len({l[:3] for l in lectures}) > 1:            # lecture ambiguë -> on n'invente pas
        return "", True
    jour, mois, annee, exacte = lectures[0]
    return _iso(annee, mois, jour), exacte


def _date_valide(jour, mois, annee):
    try:
        return (1 <= int(jour) <= 31 and 1 <= int(mois) <= 12
                and 1990 <= int(annee) <= 2100)
    except ValueError:
        return False


def _iso(annee, mois, jour):
    annee = int(annee)
    annee += 2000 if annee < 100 else 0
    return "%04d-%s-%s" % (annee, mois, jour)


def date_courte(iso):
    """'2025-04-04' -> '04-04-25' (JJ-MM-AA)."""
    annee, mois, jour = iso.split("-")
    return "%s-%s-%s" % (jour, mois, annee[2:])


def sans_accent(texte):
    texte = unicodedata.normalize("NFD", str(texte or ""))
    return "".join(c for c in texte if unicodedata.category(c) != "Mn").upper()


# ---------------------------------------------------------------------------
# Lecture des pages : mots -> lignes logiques -> colonnes
# ---------------------------------------------------------------------------
def mot_utile(mot):
    """Faux pour les résidus de tableau ('|', '][', '}', '—'…)."""
    return bool(re.search(r"[0-9A-Za-zÀ-ÿ]", mot))


def est_marqueur_client(mot):
    """Vrai pour le séparateur « Client: » du relevé, même déformé par l'OCR.

    Formes rencontrées : 'Client:', '[client', '| lient: |', '[C lent:|',
    '€ [C lent: |', '| elient: |'. Un mot n'est écarté que s'il porte un
    caractère de tableau (':', '[', ']' ou '|') ET que ses lettres forment un
    fragment de « client » — 'N°:' ou '(ce)' sont donc conservés.
    """
    if not re.search(r"[:\[\]|]", mot):
        return False
    lettres = re.sub(r"[^A-Za-z]", "", mot).lower()
    return lettres in {"c", "cl", "cli", "clie", "clien", "client", "ent",
                       "ient", "lient", "elient", "lent", "lentc", "lientc"}


def lignes_logiques(page, tolerance=4.0):
    """Regroupe les mots d'une page en lignes.

    L'OCR découpe souvent une même ligne en deux (mêmes coordonnées à 1 ou
    2 points près) : ces fragments sont recollés. Les lignes de suite d'un
    bloc sont 9 à 12 points plus bas et restent donc distinctes.
    Renvoie [{'top', 'texte', 'mots': [(x0, x1, mot), …]}] trié de haut en bas.
    """
    mots = page.extract_words() or []
    if not mots:
        return []
    mots = sorted(mots, key=lambda m: (round(m["top"], 1), m["x0"]))
    groupes, courant, top_courant = [], [], None
    for mot in mots:
        if courant and mot["top"] - top_courant > tolerance:
            groupes.append(courant)
            courant, top_courant = [], None
        if not courant:
            top_courant = mot["top"]
        courant.append(mot)
    if courant:
        groupes.append(courant)

    lignes = []
    for groupe in groupes:
        groupe = sorted(groupe, key=lambda m: m["x0"])
        lignes.append({
            "top": groupe[0]["top"],
            "texte": " ".join(m["text"] for m in groupe),
            "mots": [(m["x0"], m["x1"], m["text"]) for m in groupe],
        })
    return lignes


def colonne_de(x0, largeur):
    """Nom de la colonne dans laquelle commence le mot."""
    for index in range(len(COLONNES) - 2, -1, -1):
        if x0 >= COLONNES[index][1] * largeur:
            return COLONNES[index][0]
    return COLONNES[0][0]


def mots_par_colonne(ligne, largeur):
    """{nom de colonne: 'mots juxtaposés'} pour une ligne (résidus écartés)."""
    par_colonne = {}
    for x0, _x1, mot in ligne["mots"]:
        if not mot_utile(mot):
            continue
        par_colonne.setdefault(colonne_de(x0, largeur), []).append(mot)
    return {cle: " ".join(mots) for cle, mots in par_colonne.items()}


# ---------------------------------------------------------------------------
# Règles métier BSA (inchangées)
# ---------------------------------------------------------------------------
def calcul_montants_bsa(tx, fr_reels, remb, non_remb, tpg):
    """Détermine les 4 montants exportés à partir des 6 colonnes du relevé.

    Règle 4 (REMB = 0 et TPG = 0) prioritaire ; puis :
      1) Tx = 0, FR.REELS = REMB et NON REMB = TPG  -> tout payé ;
      2) Tx > 0 et FR.REELS = REMB                  -> tout payé ;
      3) Tx > 0, FR.REELS > REMB et TPG = 0         -> NON REMB = ticket ;
      sinon les montants du relevé sont repris tels quels.
    """
    egal = lambda a, b: abs(a - b) < 0.01

    if egal(remb, 0) and egal(tpg, 0):
        ticket, paye, exclu = 0, 0, fr_reels
    elif egal(tx, 0) and egal(fr_reels, remb) and egal(non_remb, tpg):
        ticket, paye, exclu = 0, fr_reels, 0
    elif tx > 0 and egal(fr_reels, remb):
        ticket, paye, exclu = 0, fr_reels, 0
    elif tx > 0 and fr_reels > remb and egal(tpg, 0):
        ticket, paye, exclu = non_remb, remb, 0
    else:
        ticket, paye, exclu = tpg, remb, non_remb

    return {
        "Montant_Reclame_Brut": entier_si_possible(fr_reels),
        "Ticket_Moderateur": entier_si_possible(ticket),
        "Montant_Paye_Regle": entier_si_possible(paye),
        "Montant_Exclu_Rejet": entier_si_possible(exclu),
    }


# ---------------------------------------------------------------------------
# Nom du fichier et classement
# ---------------------------------------------------------------------------
def dates_soins(lignes):
    dates = sorted(l.get("Date_Soins") or "" for l in lignes)
    return [d for d in dates if DATE_ISO.match(d)]


def periode_soins(lignes, defaut=None):
    """'04-04-25 à 24-04-25' (une seule date si tous les soins sont du même jour)."""
    dates = dates_soins(lignes) or ([defaut] if defaut and DATE_ISO.match(defaut) else [])
    if not dates:
        return "SANS DATE"
    debut, fin = dates[0], dates[-1]
    return date_courte(debut) if debut == fin else "%s à %s" % (date_courte(debut), date_courte(fin))


def annee_soins(lignes, defaut=None):
    """Année de la 1re date de soins (un paiement peut solder l'année d'avant)."""
    dates = dates_soins(lignes) or ([defaut] if defaut and DATE_ISO.match(defaut) else [])
    return dates[0].split("-")[0] if dates else "SANS DATE"


def nom_sortie(date_reglement, lignes, total_paye):
    """08-07-25 BSA 2025 04-04-25 à 24-04-25 MONTANT 282 600Ar.xlsx"""
    annee = date_reglement.split("-")[0]
    nom = "%s %s %s %s MONTANT %sAr" % (
        date_courte(date_reglement), SOCIETE, annee,
        periode_soins(lignes, date_reglement), fmt_montant(total_paye))
    nom = re.sub(r"\s+", " ", CARACTERES_INTERDITS.sub(" ", nom)).strip()
    return nom + ".xlsx"


def dossier_sortie(date_reglement, lignes):
    """<ANNEE_REGLEMENT>/<ANNEE_SOINS> à côté du script (créés si absents)."""
    chemin = os.path.join(DOSSIER, date_reglement.split("-")[0],
                          annee_soins(lignes, date_reglement))
    os.makedirs(chemin, exist_ok=True)
    return chemin


# ---------------------------------------------------------------------------
# Modèle Excel (mise en forme uniquement)
# ---------------------------------------------------------------------------
def chercher_modele():
    """Cherche le modèle : dossier du script, puis dossiers parents."""
    dossier = DOSSIER
    for _ in range(4):
        candidat = os.path.join(dossier, NOM_MODELE)
        if os.path.isfile(candidat):
            return candidat
        parent = os.path.dirname(dossier)
        if parent == dossier:
            break
        dossier = parent
    trouves = glob.glob(os.path.join(DOSSIER, "**", NOM_MODELE), recursive=True)
    return trouves[0] if trouves else None


def mise_en_forme(ws, modele):
    """Largeurs du modèle (ou valeurs par défaut), police, montants, volet figé."""
    largeurs = dict(LARGEURS_DEFAUT)
    if modele:
        try:
            feuille = load_workbook(modele)[FEUILLE]
            largeurs.update({cle: dim.width for cle, dim in feuille.column_dimensions.items()
                             if dim.width})
        except Exception as exc:
            log("   (mise en forme : modèle illisible, largeurs par défaut — %s)" % exc)
    for cle, largeur in largeurs.items():
        ws.column_dimensions[cle].width = largeur
    from openpyxl.styles import Font
    for ligne in range(1, ws.max_row + 1):
        for colonne in range(1, len(HEADERS) + 1):
            ws.cell(row=ligne, column=colonne).font = Font(name="Calibri", size=12)
    for ligne in range(2, ws.max_row + 1):
        for colonne in (9, 10, 11, 12):                # les 4 montants
            ws.cell(row=ligne, column=colonne).number_format = "#,##0"
    ws.freeze_panes = "A2"


def ecrire_excel(chemin, lignes, modele):
    wb = Workbook()
    ws = wb.active
    ws.title = FEUILLE
    ws.append(HEADERS)
    for ligne in lignes:
        ws.append([ligne.get(entete, "") for entete in HEADERS])
    mise_en_forme(ws, modele)
    wb.save(chemin)


# ---------------------------------------------------------------------------
# Analyse d'un relevé BSA
# ---------------------------------------------------------------------------
def lire_meta(lignes, texte_complet):
    """N° de relevé, lot, date du virement, montant du virement, nb de lignes."""
    meta = {"ref": None, "lot": None, "date_reglement": None, "virement": None,
            "nb_declare": None, "factures": {}, "facture_unique": None,
            "totaux_decompte": {}}

    m = re.search(r"N°\s*:?\s*(\d{4,})", texte_complet)
    if m:
        meta["ref"] = m.group(1)
    m = re.search(r"Lot\s*:?\s*(\d+)", texte_complet)
    if m:
        meta["lot"] = m.group(1)

    # Date du virement : « A,le 08/07/2025 » (et non « Le : 17/07/2025 »,
    # date d'édition du document).
    m = re.search(r"A\s*,?\s*le\s*:?\s*(\d{2}/\d{2}/\d{4})", texte_complet)
    if m:
        meta["date_reglement"] = m.group(1)
    else:
        m = re.search(r"\ble\s*:?\s*(\d{2}/\d{2}/\d{4})", texte_complet)
        if m:
            meta["date_reglement"] = m.group(1)
            log("   !! date « A,le » absente : date d'édition du document reprise")

    m = re.search(r"virement de\s*([\d\u00a0 .]+,\d{2})\s*MGA", texte_complet)
    if m:
        meta["virement"] = montant(m.group(1))
    else:
        m = re.search(r"Montant\s*:?\s*\*{0,2}\s*([\d .]+,\d{2}|\d+\.\d{2})\s*\*{0,2}",
                      texte_complet)
        if m:
            meta["virement"] = montant(m.group(1))

    m = re.search(r"Total\s+g[ée]n[ée]ral\s*:?\s*\d+\s+(\d+)", texte_complet)
    if m:
        meta["nb_declare"] = int(m.group(1))
    return meta


def lire_factures(lignes, meta):
    """N° de facture SALFA de chaque décompte.

    « Total décompte : 997712 » (parfois « . 1003884 Facture N°: », en OCR)
    suivi de « Date facture: … N°004-25/BSA/ABT-A ». La ligne « Facture N°: »
    du relevé est souvent tronquée : la ligne « Date facture » est donc
    prioritaire et écrase un n° incomplet.
    """
    decompte_courant = None
    for ligne in lignes:
        texte = ligne["texte"]
        if "facture" not in texte.lower() and "dcompte" not in texte.lower() \
                and "décompte" not in texte.lower():
            continue

        m = re.search(r"[Dd][ée]?c?ompte\s*:?\s*(\d{4,})", texte)
        if m:
            decompte_courant = m.group(1)
            meta["factures"].setdefault(decompte_courant, None)
        elif decompte_courant is None:
            # « . 1003884 Facture N°:o 005-25/ABT- » : n° de décompte isolé
            m = re.search(r"(?<![\d/])(\d{6,7})(?![\d/])", texte)
            if m:
                decompte_courant = m.group(1)
                meta["factures"].setdefault(decompte_courant, None)

        trouve = FACTURE_RE.search(texte)
        if not trouve:
            continue
        numero = trouve.group(1).strip()
        complet = bool(re.search(r"date\s*facture|dt[\s\-]*facture", texte, re.IGNORECASE))
        cle = decompte_courant
        if cle is None:
            if meta["facture_unique"] is None or complet:
                meta["facture_unique"] = numero
        elif complet or not meta["factures"].get(cle):
            meta["factures"][cle] = numero


def corriger_montants(valeurs):
    """Reconstruit FR.REELS / NON REMB quand l'OCR les a abîmés.

    Une correction n'est appliquée que si elle se recoupe avec le Tx(%) du
    relevé (REMB = FR.REELS x Tx / 100, à 1 Ar près) ; elle est alors signalée
    dans Motif_Observation et dans le journal. Renvoie (valeurs, notes).
    """
    fr, mut, tx = valeurs["FR"], valeurs["MUT"], valeurs["TX"]
    remb, nonremb, tpg = valeurs["REMB"], valeurs["NONREMB"], valeurs["TPG"]
    notes = []
    egal = lambda a, b: abs(a - b) <= 1.0
    identite = remb + max(nonremb - tpg, 0.0) + mut      # FR.REELS attendu

    if fr <= 0 and remb > 0:
        # FR.REELS illisible : on le retrouve par l'identité du relevé,
        # puis on vérifie que le Tx(%) confirme le résultat.
        candidat = round(identite)
        if candidat > 0 and (egal(tx, 0) or egal(remb, candidat * tx / 100.0)):
            valeurs["FR"] = float(candidat)
            notes.append("FR.REELS recalculé : %s Ar" % fmt_montant(candidat))
    elif fr > 0 and tx > 0 and egal(remb, fr * tx / 100.0) and not egal(fr, identite):
        # FR.REELS et REMB se recoupent avec le Tx : c'est NON REMB qui est faux.
        candidat = round(fr - remb + tpg - mut)
        if candidat >= 0:
            valeurs["NONREMB"] = float(candidat)
            notes.append("NON REMB recalculé : %s Ar" % fmt_montant(candidat))
    return valeurs, notes


def decompte_du_bloc(numero_brut, decomptes_connus):
    """'997712-3' -> '997712' ; OCR '997712413' -> '997712' (préfixe connu)."""
    if "-" in numero_brut:
        return numero_brut.split("-")[0]
    for connu in sorted(decomptes_connus, key=len, reverse=True):
        if numero_brut.startswith(connu):
            return connu
    return numero_brut[:6]


def parse_releve(pdf, nom_pdf):
    """Extrait les métadonnées du relevé et une ligne Excel par remboursement."""
    largeur_page = None
    lignes_page = []
    for page in pdf.pages:
        largeur_page = page.width
        lignes_page.extend(lignes_logiques(page))
    if not lignes_page:
        return None, [], ["aucun texte lisible dans le PDF (scan non OCRisé ?)"]

    texte_complet = "\n".join(l["texte"] for l in lignes_page)
    meta = lire_meta(lignes_page, texte_complet)
    lire_factures(lignes_page, meta)

    avertissements = []
    blocs, courant = [], None

    def fermer_courant():
        nonlocal courant
        if courant is not None:
            blocs.append(courant)
            courant = None

    for ligne in lignes_page:
        texte = ligne["texte"]

        premiere_colonne = [m for m in ligne["mots"]
                            if m[0] < 0.11 * largeur_page and mot_utile(m[2])]
        premier = premiere_colonne[0][2] if premiere_colonne else ""
        # Une ligne de données porte au moins 3 nombres dans les colonnes de
        # montants ; un en-tête de bloc n'y porte que le nom du client.
        nb_montants = sum(
            1 for x0, _x1, mot in ligne["mots"]
            if colonne_de(x0, largeur_page) in MONTANT_COLONNES
            and re.search(r"\d", mot))

        # --- total d'un décompte : « Total décompte : 997712 » + ses montants
        # (contrôle : la somme des lignes doit retomber sur ce total)
        numero_total = re.search(r"[Dd][ée]c?ompte\s*:?\s*(\d{4,})", texte)
        if not numero_total and "facture" in texte.lower():
            numero_total = re.search(r"(?<![\d/])(\d{6,7})(?![\d/])", texte)
        if numero_total and nb_montants >= 3 and \
                not re.search(r"g[ée]n[ée]ral", texte, re.IGNORECASE):
            colonnes = mots_par_colonne(ligne, largeur_page)
            meta["totaux_decompte"][numero_total.group(1)] = {
                cle: montant(colonnes.get(cle, ""), pourcentage=(cle == "TX"))
                for cle in MONTANT_COLONNES}
            fermer_courant()
            continue

        if EN_TETE_RE.search(sans_accent(texte).replace("  ", " ")) or \
                EN_TETE_RE.search(texte):
            # en-tête/pied de page, ligne de totaux : le bloc en cours s'arrête
            if not re.match(r"^\s*\d{5,7}[-–]\d{1,3}\b", texte):
                fermer_courant()
            continue

        avec_hyphen = bool(re.match(r"^\d{5,7}[-–—]\d{1,3}$", premier))
        date_lisible = bool(DATE_FR.search(premier)
                            or re.fullmatch(r"\D{0,2}\d{8,10}\D{0,2}", premier))
        est_donnees = date_lisible and nb_montants >= 3 and not avec_hyphen

        # --- ligne de données : « 04/04/2025 RAZAFINAMBININA … 20 000,00 … »
        if est_donnees:
            if courant is None:                      # en-tête de bloc illisible
                courant = {"numero": "", "matricule": "", "assure": "", "code": "",
                           "client": "", "donnees": None, "nom_suite": [], "libelle": []}
                avertissements.append("ligne de soins sans en-tête de bloc (%s)" % premier)
            if courant["donnees"] is None:
                courant["donnees"] = {"date": premier,
                                      "colonnes": mots_par_colonne(ligne, largeur_page)}
            else:
                avertissements.append("2e ligne de données pour le bloc %s (%s) — ignorée"
                                      % (courant["numero"] or "?", premier))
            continue

        # --- en-tête de bloc : « 997712-1 23961 MAHARITSY JEAN FLORENT CGPR … »
        bloc = BLOC_NUM.match(premier) if premier else None
        fusionne = BLOC_NUM_FUSIONNE.match(premier) if premier and not bloc else None
        if bloc or fusionne:
            fermer_courant()
            numero = premier
            reste = [m for m in ligne["mots"][1:] if mot_utile(m[2])]

            matricule = ""
            if reste and MATRICULE.match(reste[0][2]) and reste[0][0] < 0.24 * largeur_page:
                matricule = reste[0][2]
                reste = reste[1:]

            assure, code, client = [], "", []
            for x0, _x1, mot in reste:
                if est_marqueur_client(mot):
                    continue
                colonne = colonne_de(x0, largeur_page)
                if colonne in ("AYANT", "EXECUTANT"):
                    assure.append(mot)
                elif colonne == "ACTE":
                    if not code:
                        code = mot
                else:
                    client.append(mot)

            courant = {
                "numero": numero,
                "matricule": matricule,
                "assure": " ".join(assure),
                "code": code,
                "client": " ".join(client),
                "donnees": None, "nom_suite": [], "libelle": [],
            }
            continue

        # --- ligne de suite : nom du patient (gauche) + libellé d'acte (droite)
        if courant is None or courant["donnees"] is None:
            continue
        for x0, _x1, mot in ligne["mots"]:
            if not mot_utile(mot) or est_marqueur_client(mot):
                continue
            if x0 < 0.22 * largeur_page:
                courant["nom_suite"].append(mot)
            elif x0 >= 0.40 * largeur_page:
                courant["libelle"].append(mot)

    fermer_courant()

    # --- construction des lignes Excel ---
    decomptes_connus = {cle for cle in meta["factures"] if cle}
    decomptes_connus |= {b["numero"].split("-")[0] for b in blocs if "-" in b["numero"]}

    lignes = []
    for bloc in blocs:
        if bloc["donnees"] is None:
            avertissements.append("bloc %s sans ligne de montants — ignoré"
                                  % (bloc["numero"] or "?"))
            continue
        colonnes = bloc["donnees"]["colonnes"]
        date_soins, date_exacte = date_vers_iso(bloc["donnees"]["date"])
        if not date_soins:
            avertissements.append("bloc %s : date de soins illisible (%s) — à compléter"
                                  % (bloc["numero"] or "?", bloc["donnees"]["date"]))
        elif not date_exacte:
            avertissements.append("bloc %s : date de soins reconstruite (%s)"
                                  % (bloc["numero"] or "?", date_soins))

        valeurs = {cle: montant(colonnes.get(cle, ""), pourcentage=(cle == "TX"))
                   for cle in MONTANT_COLONNES}
        valeurs, corrections = corriger_montants(valeurs)
        montants = calcul_montants_bsa(valeurs["TX"], valeurs["FR"], valeurs["REMB"],
                                       valeurs["NONREMB"], valeurs["TPG"])

        # Contrôle de cohérence : FR.REELS = REMB + (NON REMB - TPG) + 1ERE MUT
        attendu = valeurs["REMB"] + (valeurs["NONREMB"] - valeurs["TPG"]) + valeurs["MUT"]
        ecart = valeurs["FR"] - attendu
        motif = "Prise en charge : %g%%" % valeurs["TX"]
        if valeurs["MUT"] > 0:
            motif += " ; 1ère mutuelle : %s Ar" % fmt_montant(valeurs["MUT"])
        if bloc["client"]:
            motif += " ; client : %s" % bloc["client"]
        if corrections:
            motif += " ; " + " ; ".join(corrections)
        if abs(ecart) > max(1.0, 0.01 * valeurs["FR"]):
            avertissements.append(
                "bloc %s (%s) : FR.REELS %s ≠ REMB + NON REMB - TPG + 1ERE MUT = %s "
                "(écart %s Ar) — montants à vérifier"
                % (bloc["numero"] or "?", date_soins or "sans date",
                   fmt_montant(valeurs["FR"]), fmt_montant(attendu), fmt_montant(ecart)))
            motif += " ; !! montants à vérifier"

        nom = " ".join(colonnes.get("AYANT", "").split())
        if bloc["nom_suite"]:
            nom = (nom + " " + " ".join(bloc["nom_suite"])).strip()
        if not nom:
            nom = bloc["assure"]
        nom = re.sub(r"\s+", " ", nom).strip()

        decompte = decompte_du_bloc(bloc["numero"], decomptes_connus) if bloc["numero"] else ""
        facture = (meta["factures"].get(decompte) or meta["facture_unique"] or "")

        lignes.append({
            "Ref_Decompte": meta["ref"] or "",
            "Date_Reglement": date_vers_iso(meta["date_reglement"] or "")[0],
            "Date_Soins": date_soins,
            "Nom_Agent": nom or bloc["assure"],
            "Matricule": bloc["matricule"],
            "Numero_Facture_Prescription": facture,
            "Code_Acte": bloc["code"],
            "Libelle_Acte": re.sub(r"\s+", " ", " ".join(bloc["libelle"])).strip()
                            or bloc["client"],
            **montants,
            "Motif_Observation": motif,
            "_decompte": decompte,          # interne : contrôle par décompte
        })
    return meta, lignes, avertissements


# ---------------------------------------------------------------------------
# PDF en erreur
# ---------------------------------------------------------------------------
ERREURS = []


def deplacer_en_erreur(chemin_pdf, raison):
    """Déplace un PDF non convertible dans ERREUR\ (sans écraser un homonyme)."""
    try:
        os.makedirs(ERREUR_SOUS_DOSSIER, exist_ok=True)
        dest = os.path.join(ERREUR_SOUS_DOSSIER, os.path.basename(chemin_pdf))
        base, ext = os.path.splitext(dest)
        numero = 1
        while os.path.exists(dest):
            dest = "%s (%d)%s" % (base, numero, ext)
            numero += 1
        shutil.move(chemin_pdf, dest)
        log("   -> déplacé dans ERREUR : %s" % os.path.basename(dest))
    except Exception as exc:
        raison += " (déplacement impossible : %s)" % exc
        log("   !! déplacement vers ERREUR impossible : %s" % exc)
    ERREURS.append((os.path.basename(chemin_pdf), raison))


def recapitulatif_erreurs():
    if ERREURS:
        log("")
        log("== %d PDF en erreur, déplacés dans ERREUR ==" % len(ERREURS))
        for nom, raison in ERREURS:
            log("   - %s : %s" % (nom, raison))


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------
def verifier_installation():
    """Bloque proprement (message + journal + pause) si une bibliothèque manque."""
    if not MANQUANTES:
        return True
    log("!! Bibliothèques Python manquantes : %s" % ", ".join(MANQUANTES))
    log("")
    log("   Le script ne peut pas démarrer sans elles. Installez-les une fois")
    log("   pour toutes avec la commande suivante, puis relancez le script :")
    log("")
    log("       python -m pip install %s" % " ".join(MANQUANTES))
    log("")
    log("   (si « python » n'est pas reconnu, essayez : py -3 -m pip install %s)"
        % " ".join(MANQUANTES))
    return False


def choisir_pdfs(arguments):
    """PDF demandés en argument, sinon tous ceux du sous-dossier PDF\\."""
    if arguments:
        pdfs, introuvables = [], []
        for arg in arguments:
            chemin = arg if os.path.isabs(arg) else os.path.join(PDF_SOUS_DOSSIER, arg)
            if not os.path.isfile(chemin):
                chemin = arg if os.path.isabs(arg) else os.path.join(DOSSIER, arg)
            if os.path.isfile(chemin) and chemin.lower().endswith(".pdf"):
                pdfs.append(os.path.abspath(chemin))
            else:
                introuvables.append(arg)
        for nom in introuvables:
            log("!! PDF introuvable : %s" % nom)
        return pdfs

    os.makedirs(PDF_SOUS_DOSSIER, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(PDF_SOUS_DOSSIER, "*.pdf")))
    if not pdfs:
        log("Aucun PDF à convertir dans : %s" % PDF_SOUS_DOSSIER)
        log("Déposez vos relevés BSA dans ce sous-dossier « PDF », puis relancez.")
    return pdfs


def convertir(chemin_pdf, modele, force):
    """Convertit un PDF ; renvoie Vrai si un Excel a été créé."""
    nom_pdf = os.path.basename(chemin_pdf)
    log("")
    log("--- %s" % nom_pdf)

    try:
        pdf = pdfplumber.open(chemin_pdf)
    except Exception as exc:
        log("!! PDF illisible (%s: %s)" % (type(exc).__name__, exc))
        deplacer_en_erreur(chemin_pdf, "PDF illisible ou corrompu")
        return False

    try:
        with pdf:
            meta, lignes, avertissements = parse_releve(pdf, nom_pdf)
    except Exception as exc:
        log("!! erreur de lecture (%s: %s)" % (type(exc).__name__, exc))
        log(traceback.format_exc().rstrip())
        deplacer_en_erreur(chemin_pdf, "erreur de lecture (format non reconnu ?)")
        return False

    for avert in avertissements or []:
        log("   ~ %s" % avert)
    if meta is None:
        log("!! aucune ligne de règlement trouvée (scan non OCRisé ? format inconnu ?)")
        deplacer_en_erreur(chemin_pdf, "aucun texte lisible dans le PDF")
        return False
    if not lignes:
        log("!! aucune ligne de règlement trouvée (format non reconnu ?)")
        deplacer_en_erreur(chemin_pdf, "aucune ligne de règlement trouvée")
        return False

    if not meta.get("date_reglement"):
        log("!! date de virement (« A,le JJ/MM/AAAA ») introuvable dans le PDF")
        deplacer_en_erreur(chemin_pdf, "date de virement introuvable")
        return False
    date_reglement = date_vers_iso(meta["date_reglement"])[0]

    total_paye = sum(l["Montant_Paye_Regle"] for l in lignes)
    if meta.get("virement") is not None:
        if abs(total_paye - meta["virement"]) < 1:
            log("   contrôle : %s Ar payés = montant du virement  OK" % fmt_montant(total_paye))
        else:
            log("   !! ATTENTION : %s Ar payés ≠ montant du virement (%s Ar)"
                % (fmt_montant(total_paye), fmt_montant(meta["virement"])))
    if meta.get("nb_declare") and meta["nb_declare"] != len(lignes):
        log("   !! ATTENTION : %d lignes lues, %d déclarées dans le « Total général »"
            % (len(lignes), meta["nb_declare"]))

    # Contrôle décompte par décompte : le relevé imprime lui-même le total de
    # chaque décompte (« Total décompte : 997742 … 286 560,00 »). Un écart
    # désigne le décompte dont une ligne a été abîmée par l'OCR.
    for decompte, totaux in sorted((meta.get("totaux_decompte") or {}).items()):
        annonce = totaux.get("REMB", 0.0)
        lu = sum(l["Montant_Paye_Regle"] for l in lignes if l.get("_decompte") == decompte)
        if annonce <= 0:
            continue
        if abs(lu - annonce) < 1:
            log("   décompte %s : %s Ar  OK" % (decompte, fmt_montant(lu)))
        else:
            log("   !! décompte %s : %s Ar lus ≠ %s Ar annoncés (écart %s Ar) "
                "— ligne(s) à vérifier dans ce décompte"
                % (decompte, fmt_montant(lu), fmt_montant(annonce),
                   fmt_montant(lu - annonce)))

    sortie = os.path.join(dossier_sortie(date_reglement, lignes),
                          nom_sortie(date_reglement, lignes, total_paye))
    relatif = os.path.relpath(sortie, DOSSIER)
    if os.path.exists(sortie) and not force:
        log("-- %s existe déjà, non écrasé (--force pour régénérer)" % relatif)
        return False

    try:
        ecrire_excel(sortie, lignes, modele)
    except Exception as exc:
        log("!! création de l'Excel impossible (%s: %s)" % (type(exc).__name__, exc))
        log(traceback.format_exc().rstrip())
        deplacer_en_erreur(chemin_pdf, "erreur de création de l'Excel")
        return False

    log("OK %s" % relatif)
    log("   relevé N° %s | lot %s | %d lignes | payé %s Ar"
        % (meta.get("ref") or "?", meta.get("lot") or "?", len(lignes), fmt_montant(total_paye)))
    return True


def main(argv):
    arguments = [a for a in argv if not a.startswith("--")]
    options = {a for a in argv if a.startswith("--")}
    force = "--force" in options
    sans_pause = "--no-pause" in options

    log("=" * 78)
    log("BSA — relevés PDF vers Excel — %s" % datetime.now().strftime("%d/%m/%Y %H:%M"))
    log("Dossier : %s" % DOSSIER)

    if not verifier_installation():
        return 1

    modele = chercher_modele()
    log("Modèle  : %s" % (os.path.relpath(modele, DOSSIER) if modele
                          else "introuvable (largeurs par défaut)"))

    pdfs = choisir_pdfs(arguments)
    if not pdfs:
        recapitulatif_erreurs()
        return 0

    log("PDF à convertir : %d" % len(pdfs))
    reussis = 0
    for chemin in pdfs:
        try:
            if convertir(chemin, modele, force):
                reussis += 1
        except Exception as exc:                       # un PDF ne bloque pas les autres
            log("!! erreur inattendue sur %s (%s: %s)"
                % (os.path.basename(chemin), type(exc).__name__, exc))
            log(traceback.format_exc().rstrip())
            deplacer_en_erreur(chemin, "erreur inattendue (%s)" % type(exc).__name__)

    recapitulatif_erreurs()
    log("")
    log("Terminé : %d Excel créé(s) sur %d PDF, %d en erreur."
        % (reussis, len(pdfs), len(ERREURS)))
    log("Journal : %s" % JOURNAL)
    return 0


if __name__ == "__main__":
    code = 1
    try:
        code = main(sys.argv[1:])
    except Exception:                                  # filet de sécurité global
        log("")
        log("!! ERREUR INATTENDUE — le script s'arrête :")
        log(traceback.format_exc().rstrip())
    finally:
        ecrire_journal()
        if "--no-pause" not in sys.argv[1:]:
            pause()
    sys.exit(code)
