# -*- coding: utf-8 -*-
"""
BSA_excel_to_modele.py — conversion des relevés BSA au format d'importation
============================================================================

Les fichiers Excel déposés par BSA dans le dossier ``Excel`` sont des copies
mises en page du relevé papier « RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE ».
Ils ne sont donc pas encore au format d'importation de SALFA.

Ce script lit ces relevés et crée, pour chacun, un nouveau fichier dans
``Excel/Import`` avec les 13 colonnes du modèle
``Modele_Import_Reglements_Decompte_Assurance.xlsx``.

Utilisation :
    python BSA_excel_to_modele.py                 # tous les Excel de Excel/
    python BSA_excel_to_modele.py --force         # régénérer les sorties
    python BSA_excel_to_modele.py "mon_releve.xlsx"

Les fichiers sources ne sont jamais modifiés. Les sorties sont séparées dans
``Excel/Import`` afin qu'un fichier créé par le script ne soit pas relu comme
un nouveau relevé lors de l'exécution suivante.

Règles BSA : les six montants sont lus dans l'ordre du relevé
FR.REELS, 1ERE MUT, Tx (%), REMB, NON REMB, TPG*. Les quatre colonnes de
montants du modèle sont ensuite calculées par la même fonction métier que le
convertisseur BSA des PDF (``calcul_montants_bsa``).
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import OrderedDict
from pathlib import Path
from typing import Iterable

import openpyxl

# Réutiliser les règles et l'écriture du convertisseur BSA PDF. Le script est
# placé dans le même dossier, ce qui permet aussi de le lancer par double-clic.
ICI = Path(__file__).resolve().parent
if str(ICI) not in sys.path:
    sys.path.insert(0, str(ICI))

from BSA_paiement_to_excel import (  # noqa: E402  (import après sys.path)
    amount_to_float,
    calcul_montants_bsa,
    fmt_amount,
    nom_sortie,
    parse_date,
    write_workbook,
)


INPUT_DIR = ICI / "Excel"
OUTPUT_DIR = INPUT_DIR / "Import"
SOCIETE = "BSA"

DATE_RE = re.compile(r"(?<!\d)(\d{2}/\d{2}/\d{4})(?!\d)")
BLOCK_RE = re.compile(r"(?<!\d)(\d{4,}-\d+)(?!\d)")
# Dans les relevés BSA, les montants sont imprimés avec deux décimales.
# Cette expression évite de prendre les nombres des noms de médicaments ou
# les numéros de facture du texte de l'acte.
AMOUNT_RE = re.compile(
    r"(?<![\d,])(?:\d{1,3}(?:[ \u00a0]\d{3})+|\d+),\d{2}(?!\d)"
)

# Codes de prestations fréquemment rencontrés dans les relevés BSA. Le
# parseur accepte aussi automatiquement un code court inconnu placé juste
# avant « Client: ».
KNOWN_ACTE_CODES = {
    "CG", "CGPR", "PH", "SI", "EB", "ECH", "DSO", "HMCH", "HAHO",
    "HANA", "HPHIE",
}

# N° apparaissant dans la zone « Facture N° » des décomptes. Les formes
# observées sont, par exemple, FA-02/BSA/26-050 et N°26FA0509019.
FORMAL_INVOICE_RE = re.compile(
    r"(?<![A-Z0-9])FA-\d{2}[-/][A-Z0-9]+(?:[-_/][A-Z0-9]+)*[-/]\d+"
    r"(?![A-Z0-9_-])",
    re.IGNORECASE,
)
SHORT_INVOICE_RE = re.compile(
    r"(?<![A-Z0-9])(?:N[°ºo]?\s*)?(\d{2}FA\d{5,})(?![A-Z0-9])",
    re.IGNORECASE,
)


class ReleveBSAError(ValueError):
    """Erreur de structure dans un relevé Excel BSA."""


# ---------------------------------------------------------------------------
# Lecture des valeurs Excel
# ---------------------------------------------------------------------------
def clean_text(value) -> str:
    """Convertit une cellule en texte comparable, sans perdre les retours.

    Les exports contiennent parfois un retour au milieu d'un mot à cause de
    la largeur de la cellule (par ex. ``ANDR\\nO``). Ce cas est recollé ; les
    retours entre deux vrais mots restent des espaces.
    """
    if value is None:
        return ""
    value = str(value).replace("\r\n", "\n").replace("\r", "\n")
    value = re.sub(r"(?<=[A-Za-zÀ-ÿ])\n(?=[A-Za-zÀ-ÿ](?:\s|$))", "", value)
    return value.strip()


def flat_text(value: str) -> str:
    """Remplace les retours de ligne par des espaces pour les regex."""
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def row_values(ws) -> "OrderedDict[int, list[tuple[int, str]]]":
    """Retourne les cellules non vides, regroupées par numéro de ligne.

    ``openpyxl`` ne renvoie une valeur que pour la cellule en haut à gauche
    d'une plage fusionnée : c'est exactement la représentation utile ici.
    L'ordre des colonnes est conservé pour retrouver les champs du relevé.
    """
    rows: "OrderedDict[int, list[tuple[int, str]]]" = OrderedDict()
    for row in ws.iter_rows():
        non_vides = []
        for cell in row:
            value = clean_text(cell.value)
            if value:
                non_vides.append((cell.column, value))
        if non_vides:
            rows[row[0].row] = non_vides
    return rows


def row_text(cells: Iterable[tuple[int, str]]) -> str:
    """Texte d'une ligne, dans l'ordre gauche-droite."""
    return " ".join(flat_text(value) for _, value in cells if value).strip()


def all_text(rows: "OrderedDict[int, list[tuple[int, str]]] ") -> str:
    return "\n".join(row_text(cells) for cells in rows.values())


def find_block_starts(rows):
    """Trouve les lignes d'en-tête de remboursement.

    Les exports BSA ne fusionnent pas toujours la même largeur de cellules :
    parfois l'identifiant est seul sur une ligne, parfois toute la mention
    « ADHESION » est sur cette même ligne. La recherche dans le texte de la
    ligne couvre les deux variantes.
    """
    starts = []
    for number, cells in rows.items():
        match = BLOCK_RE.search(row_text(cells))
        if match:
            starts.append((number, match.group(1)))
    return starts


def is_page_marker(text: str) -> bool:
    return bool(re.fullmatch(r"\d+\s*/\s*\d+", flat_text(text)))


def is_total_decompte(text: str) -> bool:
    return bool(re.search(r"Total\s+d[ée]compte", text, re.IGNORECASE))


def is_report_header(text: str) -> bool:
    normalized = flat_text(text).upper()
    return "RELEVE DE REMBOURSEMENTS" in normalized


def is_decompte_footer(text: str) -> bool:
    """Vrai pour la ligne de facture placée juste après un bloc.

    Selon la largeur du relevé, la ligne « Facture N° » peut être imprimée
    séparément du « Total décompte ». Elle doit donc aussi terminer la collecte
    des six montants du bloc.
    """
    return bool(re.search(r"Facture\s+N[°ºo]?\s*:", text, re.IGNORECASE))


def block_limits(rows, starts, index):
    """Retourne (fin, ligne_date, ligne_total) pour un bloc.

    La fin est exclusive. Un total ou un pied de page peut apparaître entre
    deux blocs ; il doit être exclu de la collecte des montants du bloc.
    """
    start, _ = starts[index]
    next_start = starts[index + 1][0] if index + 1 < len(starts) else max(rows, default=start) + 1
    candidates = [r for r in rows if start < r < next_start]
    date_row = None
    for r in candidates:
        text = row_text(rows[r])
        if DATE_RE.search(text):
            date_row = r
            break
    if date_row is None:
        raise ReleveBSAError(f"bloc {starts[index][1]} : date de soins introuvable")

    stop = next_start
    total_row = None
    for r in candidates:
        text = row_text(rows[r])
        if r >= date_row and is_total_decompte(text):
            stop = min(stop, r)
            total_row = r
            break
        if r >= date_row and (
            is_page_marker(text)
            or is_report_header(text)
            or is_decompte_footer(text)
        ):
            stop = min(stop, r)
            break
    # Dans un relevé valide, la date de soins est toujours conservée.
    if stop <= date_row:
        raise ReleveBSAError(f"bloc {starts[index][1]} : limites de données invalides")
    return stop, date_row, total_row


# ---------------------------------------------------------------------------
# Métadonnées et champs d'un bloc
# ---------------------------------------------------------------------------
def extract_meta(rows):
    text = all_text(rows)
    meta = {"ref": "", "date_reglement": "", "virement": None}

    match = re.search(r"N[°ºo]?\s*:\s*(\d+)", text, re.IGNORECASE)
    if match:
        meta["ref"] = match.group(1)

    # « Le : 28/08/2026 » est la date d'édition du fichier. La date de
    # règlement BSA est celle de « A , le 14/08/2026 ».
    for line in rows.values():
        line_text = row_text(line)
        match = re.search(r"\ble\s+(\d{2}/\d{2}/\d{4})", line_text, re.IGNORECASE)
        if match:
            meta["date_reglement"] = match.group(1)
            break
    if not meta["date_reglement"]:
        match = re.search(r"\ble\s+(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
        if match:
            meta["date_reglement"] = match.group(1)

    match = re.search(
        r"virement\s+de\s+([\d\s\u00a0]+),([\d]{2})\s*MGA",
        text,
        re.IGNORECASE,
    )
    if match:
        meta["virement"] = amount_to_float(match.group(1) + "," + match.group(2))
    return meta


def header_text(rows, start, date_row) -> str:
    return " ".join(
        row_text(rows[number])
        for number in rows
        if start <= number < date_row
    ).strip()


def extract_matricule_and_code(header: str):
    """Extrait le matricule après ADHESION et le code avant Client."""
    before_client = re.split(r"\bClient\s*:", header, maxsplit=1, flags=re.IGNORECASE)[0]
    adhesion = re.search(
        r"ADHESION\s*:\s*([A-Z0-9][A-Z0-9-]*)",
        before_client,
        re.IGNORECASE,
    )
    matricule = adhesion.group(1) if adhesion else ""

    code = ""
    if before_client:
        tokens = re.findall(r"[A-Z][A-Z0-9-]*", before_client.upper())
        # Le dernier jeton précédant « Client: » est le type de prestation
        # (CG, PH, EB, ...), après le nom de l'adhérent.
        for token in reversed(tokens):
            if token in KNOWN_ACTE_CODES:
                code = token
                break
        if not code and tokens:
            candidate = tokens[-1]
            if (
                len(candidate) <= 6
                and candidate not in {"ADHESION", "CLIENT"}
                and not candidate.isdigit()
            ):
                code = candidate
    return matricule, code


def amount_tokens(value: str) -> list[str]:
    return AMOUNT_RE.findall(clean_text(value))


def text_without_amounts(value: str) -> str:
    """Garde le libellé d'acte d'une cellule qui contient montant + texte."""
    value = clean_text(value)
    value = AMOUNT_RE.sub(" ", value)
    return flat_text(value)


def unique_append(items: list[str], value: str):
    value = flat_text(value)
    if not value:
        return
    if value not in items:
        items.append(value)


def parse_block(rows, start, stop, date_row, block_id, meta, total_row):
    """Convertit un bloc BSA en une ligne du modèle d'importation."""
    header = header_text(rows, start, date_row)
    matricule, code_acte = extract_matricule_and_code(header)

    data_rows = [r for r in rows if date_row <= r < stop]
    date = ""
    patient_parts: list[str] = []
    acte_parts: list[str] = []
    exec_col = None
    first_data_cells = rows[date_row]

    # La ligne de soin commence par la date, puis le nom du patient et
    # « ASSOCIATION DISPENSAIRE LOTERANA » (l'exécutant). Les colonnes exactes
    # changent entre les exports, d'où une détection par le texte.
    for col, value in first_data_cells:
        match = DATE_RE.search(value)
        if match and not date:
            date = match.group(1)
            continue
        if "ASSOCIATION DISPENSAIRE" in flat_text(value).upper():
            exec_col = col
            continue
        if amount_tokens(value):
            # Une cellule de la zone ACTE peut contenir à la fois le montant
            # FR.REELS et le libellé (par ex. « 1 000,00\\nPREDNISOLONE »).
            # Le montant est déjà collecté plus bas ; ici on conserve donc le
            # texte qui l'accompagne.
            if exec_col is not None:
                cleaned = text_without_amounts(value)
                if cleaned and "CLIENT:" not in cleaned.upper():
                    unique_append(acte_parts, cleaned)
            continue
        # Les cellules avant l'exécutant composent le nom de l'ayant-droit.
        if exec_col is None and col > 1:
            unique_append(patient_parts, value)
        # Si l'exécutant n'est pas dans cette cellule mais a déjà été vu,
        # les cellules suivantes correspondent à l'acte/libellé.
        elif exec_col is not None:
            cleaned = text_without_amounts(value)
            if cleaned and "CLIENT:" not in cleaned.upper():
                unique_append(acte_parts, cleaned)

    # Dans quelques exports, l'exécutant est une cellule fusionnée qui n'est
    # pas exactement sur la même ligne que la date. Le retrouver dans les
    # cellules du bloc évite alors de classer un libellé comme un nom.
    if exec_col is None:
        for r in data_rows:
            for col, value in rows[r]:
                if "ASSOCIATION DISPENSAIRE" in flat_text(value).upper():
                    exec_col = col
                    break
            if exec_col is not None:
                break
    if exec_col is None:
        # Le relevé reste exploitable sans ce texte : la zone de nom est à
        # gauche et l'acte à droite. La valeur 6 est celle des exports ORANGE.
        exec_col = 6

    # Collecte des suites de nom et du libellé sur les lignes fusionnées
    # suivantes. Les lignes de suite de patient sont à gauche de l'exécutant ;
    # les suites d'acte sont à droite.
    for r in data_rows:
        for col, value in rows[r]:
            if r == date_row:
                # Le nom et l'acte de la première ligne ont été traités ci-dessus.
                continue
            text = flat_text(value)
            if not text or DATE_RE.search(text):
                continue
            upper = text.upper()
            if "ASSOCIATION DISPENSAIRE" in upper or "CLIENT:" in upper:
                continue
            if amount_tokens(value):
                cleaned = text_without_amounts(value)
            else:
                cleaned = text
            if not cleaned:
                continue
            if col <= exec_col:
                unique_append(patient_parts, cleaned)
            else:
                unique_append(acte_parts, cleaned)

    if not date:
        raise ReleveBSAError(f"bloc {block_id} : date de soins introuvable")

    # Les six valeurs sont toujours dans l'ordre des colonnes du relevé.
    amounts: list[str] = []
    for r in data_rows:
        for _, value in rows[r]:
            amounts.extend(amount_tokens(value))
    if len(amounts) != 6:
        raise ReleveBSAError(
            f"bloc {block_id} : {len(amounts)} montants trouvés au lieu de 6"
        )
    fr_reels, premiere_mutuelle, tx, remb, non_remb, tpg = amounts

    tx_value = amount_to_float(tx)
    motif = f"Prise en charge : {tx_value:g}%"
    mut_value = amount_to_float(premiere_mutuelle)
    if mut_value > 0:
        motif += f" ; 1ère mutuelle : {fmt_amount(mut_value)} Ar"

    # Facture affectée après regroupement par numéro de décompte.
    facture = meta.get("factures", {}).get(block_id.split("-", 1)[0], "")
    ligne = {
        "Ref_Decompte": meta.get("ref", ""),
        "Date_Reglement": parse_date(meta["date_reglement"]),
        "Date_Soins": parse_date(date),
        "Nom_Agent": " ".join(patient_parts).strip(),
        "Matricule": matricule,
        "Numero_Facture_Prescription": facture,
        "Code_Acte": code_acte,
        "Libelle_Acte": " ".join(acte_parts).strip(),
        **calcul_montants_bsa(tx, fr_reels, remb, non_remb, tpg),
        "Motif_Observation": motif,
    }
    return ligne, total_row


# ---------------------------------------------------------------------------
# Factures de décompte et conversion complète
# ---------------------------------------------------------------------------
def invoice_candidates(text: str) -> list[str]:
    """Retourne les n° de facture explicites trouvés dans un récapitulatif."""
    found: list[tuple[int, str]] = []
    for match in FORMAL_INVOICE_RE.finditer(text):
        found.append((match.start(), match.group(0).upper()))
    for match in SHORT_INVOICE_RE.finditer(text):
        raw = match.group(0).strip()
        # Conserver le « N° » présent dans le relevé : c'est le format des
        # fichiers d'importation historiques du dossier BSA.
        prefix = "N°" if re.match(r"N[°ºo]?", raw, re.IGNORECASE) else ""
        found.append((match.start(), prefix + match.group(1).upper()))
    found.sort(key=lambda item: item[0])
    # Dédupliquer en conservant l'ordre d'apparition.
    out = []
    for _, value in found:
        if value not in out:
            out.append(value)
    return out


def extract_factures(rows):
    """Associe le n° de facture aux décomptes portant le même préfixe."""
    factures = {}
    for number, cells in rows.items():
        text = row_text(cells)
        match = re.search(r"Total\s+d[ée]compte\s*:\s*(\d{4,})", text, re.IGNORECASE)
        if not match:
            continue
        decompte = match.group(1)
        # Le n° complet peut être dans une autre cellule de la même ligne.
        candidates = invoice_candidates(text)
        if not candidates:
            # Certaines versions mettent « Facture N° » sur la ligne
            # précédente ou la ligne suivante du « Total décompte ».
            nearby_rows = [r for r in rows if number - 3 <= r <= number + 3 and r != number]
            for r in sorted(nearby_rows):
                candidates.extend(invoice_candidates(row_text(rows[r])))
        if candidates:
            factures[decompte] = candidates[0]
    return factures


def parse_source(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    if not wb.worksheets:
        raise ReleveBSAError("aucune feuille Excel")
    ws = wb.worksheets[0]
    rows = row_values(ws)
    starts = find_block_starts(rows)
    if not starts:
        raise ReleveBSAError("aucun bloc de remboursement trouvé")

    meta = extract_meta(rows)
    if not meta["ref"]:
        raise ReleveBSAError("référence du relevé introuvable")
    if not meta["date_reglement"]:
        raise ReleveBSAError("date de règlement introuvable")
    try:
        parse_date(meta["date_reglement"])
    except Exception as exc:
        raise ReleveBSAError("date de règlement invalide") from exc

    meta["factures"] = extract_factures(rows)
    lignes = []
    for index, (start, block_id) in enumerate(starts):
        stop, date_row, total_row = block_limits(rows, starts, index)
        ligne, _ = parse_block(
            rows, start, stop, date_row, block_id, meta, total_row
        )
        lignes.append(ligne)

    if not lignes:
        raise ReleveBSAError("aucune ligne de remboursement exploitable")
    return meta, lignes


def source_files(arguments: list[str]) -> list[Path]:
    if arguments:
        files = []
        for argument in arguments:
            candidate = Path(argument)
            if not candidate.is_absolute():
                # Accepter aussi bien le nom court du fichier que le chemin
                # fourni depuis la racine du dépôt ou depuis BSA/.
                candidates = [
                    INPUT_DIR / candidate,
                    ICI / candidate,
                    Path.cwd() / candidate,
                ]
                candidate = next((p for p in candidates if p.is_file()), candidates[0])
            if not candidate.is_file() or candidate.suffix.lower() != ".xlsx":
                print(f"!! Fichier Excel introuvable ou invalide : {argument}")
                continue
            files.append(candidate.resolve())
        return files
    return sorted(
        p.resolve()
        for p in INPUT_DIR.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    )


def output_path(date_reglement: str, lignes: list[dict]) -> Path:
    total = sum(ligne["Montant_Paye_Regle"] for ligne in lignes)
    # Le dossier Import est volontairement unique : les fichiers sources ont
    # déjà leur propre nom et restent visibles au même niveau dans Excel/.
    return OUTPUT_DIR / nom_sortie(SOCIETE, date_reglement, lignes, total)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Convertit les relevés Excel BSA en fichiers d'importation."
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="écraser les fichiers d'importation déjà présents",
    )
    parser.add_argument(
        "files",
        nargs="*",
        help="un ou plusieurs fichiers .xlsx de PAIEMENT CLIENT/BSA/Excel",
    )
    args = parser.parse_args(argv)

    paths = source_files(args.files)
    if not paths:
        print(f"!! Aucun fichier Excel source trouvé dans {INPUT_DIR}")
        return 1 if args.files else 0

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    errors = 0
    for path in paths:
        try:
            meta, lignes = parse_source(path)
            date_reglement = parse_date(meta["date_reglement"])
            out = output_path(date_reglement, lignes)
            total_paye = sum(ligne["Montant_Paye_Regle"] for ligne in lignes)
            if meta.get("virement") is not None:
                if abs(total_paye - meta["virement"]) < 1:
                    print(
                        f"   contrôle : {fmt_amount(total_paye)} Ar payés = "
                        f"montant du virement ({fmt_amount(meta['virement'])} Ar)  OK"
                    )
                else:
                    print(
                        f"   !! ATTENTION : {fmt_amount(total_paye)} Ar payés "
                        f"≠ montant du virement ({fmt_amount(meta['virement'])} Ar)"
                    )
            if out.exists() and not args.force:
                print(
                    f"-- {out.relative_to(ICI)} : existe déjà, non écrasé "
                    f"(--force pour régénérer)  [{path.name}]"
                )
                continue
            write_workbook(str(out), lignes)
            print(
                f"OK {out.relative_to(ICI)} : relevé N°{meta['ref']} | "
                f"{len(lignes)} lignes | Payé {fmt_amount(total_paye)} Ar "
                f"<- {path.name}"
            )
        except Exception as exc:
            errors += 1
            print(f"!! {path.name} : {type(exc).__name__}: {exc}")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
