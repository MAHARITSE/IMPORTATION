# 📚 TUTO D'UTILISATION — `pdf_to_excel.py`

Script de conversion des **factures PDF SALFA** en **fichiers Excel** selon le modèle
`Modele_Import.xlsx` (feuille `Modele_Prestations`, 11 colonnes).

---

## 1. Prérequis (une seule fois)

### a) Installer Python 3
- **Windows** : télécharger sur https://www.python.org/downloads/ → cocher **« Add Python to PATH »** à l'installation
- **Mac** : `brew install python` (ou installateur python.org)
- **Linux** : déjà installé en général

Vérifier :
```bash
python3 --version
```

### b) Installer les 2 bibliothèques nécessaires
```bash
pip install pdfplumber openpyxl
```
> ⚠️ Sur certains systèmes Linux, si erreur « externally-managed-environment » :
> ```bash
> pip install --break-system-packages pdfplumber openpyxl
> ```

---

## 2. Organisation des fichiers

```
IMPORTATION/
├── pdf_to_excel.py                      ← le script
└── FACTURE CLIENT/
    ├── Modele_Import.xlsx               ← le modèle (11 colonnes) — NE PAS RENOMMER
    ├── BSA JANVIER.pdf                  ← les PDF à convertir
    ├── BSA FEVRIER.pdf                    (nom : CLIENT MOIS.pdf)
    ├── MCI JANVIER.pdf
    ├── ...
    └── EXCEL/                           ← dossier de SORTIE (créé automatiquement)
        ├── BSA JANVIER.xlsx
        ├── MCI JANVIER.xlsx
        ├── BSA 2026 CONSOLIDE (Janvier-Juin).xlsx
        └── MCI 2026 CONSOLIDE (Janvier-Juin).xlsx
```

### ⚠️ Règle de nommage des PDF
```
<NOM_CLIENT> <MOIS>.pdf     →     MCI JANVIER.pdf   ✅
```
- **Nom du client** = préfixe avant l'espace (BSA, MCI, …) → c'est ce qui sera inscrit
  dans la colonne `Societe` de l'Excel
- **Mois** en MAJUSCULES : JANVIER, FEVRIER, MARS, AVRIL, MAI, JUIN, JUILLET, AOUT,
  SEPTEMBRE, OCTOBRE, NOVEMBRE, DECEMBRE

---

## 3. Les 3 commandes de base

Ouvrir un terminal **dans le dossier IMPORTATION** puis :

### 🔹 Commande 1 — Convertir un seul client (le plus courant)
```bash
python3 pdf_to_excel.py MCI
```
→ Convertit `MCI JANVIER.pdf` … `MCI JUIN.pdf` en 6 fichiers Excel
+ le consolidé `MCI 2026 CONSOLIDE (Janvier-Juin).xlsx`

```bash
python3 pdf_to_excel.py BSA
```
→ Idem pour BSA.

### 🔹 Commande 2 — Tout convertir d'un coup
```bash
python3 pdf_to_excel.py
```
→ Détecte **tous les clients** automatiquement (tous les préfixes de PDF trouvés)
et les convertit tous.

### 🔹 Commande 3 — Forcer la régénération
```bash
python3 pdf_to_excel.py MCI --force
```
→ **Écrase** les fichiers Excel MCI déjà existants.

> 🛡️ **Sans `--force`, les Excel existants ne sont JAMAIS écrasés** : vos modifications
> manuelles dans les Excel sont protégées. Le message affiché est alors :
> `-- MCI JANVIER.xlsx : déjà existant, non écrasé`

---

## 4. Ce que produit le script

Pour chaque PDF, une ligne Excel = un soin (agent) avec les 11 colonnes du modèle :

| Colonne | Contenu | Exemple |
|---|---|---|
| `Numero_Facture` | n° de la facture du PDF | `FA-01/MCI/26-034` |
| `Date_Soins` | date du soin | `2026-01-02` |
| `Nom_Agent` | nom + sous-société entre parenthèses | `RAFANOMEZANTSOA STEVEN (PAMF)` |
| `Matricule` | n° Mlle (si présent) | `151663` |
| `Societe` | préfixe du fichier PDF | `MCI` |
| `Sous_Societe` | employeur entre parenthèses | `UWS MADAGASCAR` |
| `Acte_Medicale_Prix` | actes détaillés `; `-séparés | `CONS : 20 000; MEDIC : 83 000` |
| `Montant_Total_Brut` | montant total | `103000` |
| `Ticket_Moderateur` | participation | `20600` |
| `Prise_En_Charge_Net` | net à payer | `82400` |
| `Observations` | facture + mois (+ alerte éventuelle) | `Facture mensuelle soins ambulatoires - Janvier 2026` |

Mise en forme identique au modèle : largeurs de colonnes, police Calibri 12,
formats de nombres `#,##0`, ligne d'en-tête figée.

---

## 5. Exemple de session complète

```bash
cd IMPORTATION

# 1) Mettre les nouveaux PDF dans "FACTURE CLIENT/"
#    ex : MCI JUILLET.pdf

# 2) Lancer la conversion
python3 pdf_to_excel.py MCI

# Résultat affiché :
# OK MCI JANVIER.xlsx : FA-01/MCI/26-034 | Janvier 2026 | 24 lignes | Brut 1 847 900 | ...
# ...
# OK MCI 2026 CONSOLIDE (Janvier-Juin).xlsx : 88 lignes au total

# 3) Les Excel sont dans "FACTURE CLIENT/EXCEL/"
```

> 💡 Le **consolidé** s'étend automatiquement : si vous ajoutez `MCI JUILLET.pdf`,
> il s'appellera `MCI 2026 CONSOLIDE (Janvier-Juillet).xlsx`.
> Pensez à supprimer l'ancien consolidé (ou utilisez `--force`).

---

## 6. Vérifications automatiques intégrées

- ✅ Les lignes coupées entre deux pages d'un PDF sont **rattachées à la bonne ligne**
- ✅ Si la somme des actes détaillés ≠ montant facturé, une mention
  `ATTENTION : actes détaillés (...) < montant facturé (écart ... Ar)` est ajoutée
  dans `Observations` → à contrôler manuellement
- ❌ Si un PDF est illisible : `!! <fichier> : aucune ligne trouvée`

**Contrôle recommandé après chaque conversion** : comparer le TOTAL `Montant_Total_Brut`
de l'Excel avec la ligne « Total » imprimée en bas de la facture PDF. Ils doivent être égaux.

---

## 7. Problèmes fréquents

| Problème | Solution |
|---|---|
| `ModuleNotFoundError: pdfplumber` | Relancer `pip install pdfplumber openpyxl` |
| `python3` n'existe pas (Windows) | Essayer `python pdf_to_excel.py MCI` |
| Le script ne trouve pas les PDF | Vérifier le nommage `CLIENT MOIS.pdf` (espace + mois en MAJUSCULES) |
| Le mois n'est pas converti | Le mois doit être écrit en toutes lettres (`MCI JANVIER.pdf`, pas `MCI 01.pdf`) |
| Excel pas mis à jour | Fichier existant protégé → utiliser `--force` |
| Accents dans le nom de fichier | Éviter ; préférer `FEVRIER` sans accent (comme les fichiers actuels) |

---

## 8. Récapitulatif express

```bash
pip install pdfplumber openpyxl        # une seule fois
python3 pdf_to_excel.py MCI            # convertir MCI
python3 pdf_to_excel.py BSA            # convertir BSA
python3 pdf_to_excel.py                # tout convertir
python3 pdf_to_excel.py MCI --force    # forcer la régénération
```
