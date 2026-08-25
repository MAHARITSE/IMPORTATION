# 📚 TUTO D'UTILISATION — `pdf_to_excel.py` (Windows)

Conversion des **factures PDF SALFA** en **fichiers Excel** selon le modèle
`Modele_Import.xlsx` (feuille `Modele_Prestations`, 11 colonnes).

---

## 1. Installation (une seule fois)

### a) Installer Python 3 sur Windows
1. Aller sur https://www.python.org/downloads/
2. Télécharger et lancer l'installateur
3. ⚠️ **COCHER « Add Python to PATH »** en bas de la fenêtre d'installation
4. Cliquer sur *Install Now*

Vérification : ouvrir l'**invite de commandes** (touche Windows → taper `cmd`) :
```
python --version
```

### b) Installer les 2 bibliothèques nécessaires
Dans l'invite de commandes :
```
pip install pdfplumber openpyxl
```

---

## 2. Organisation des fichiers

```
IMPORTATION/
├── CONVERTIR.bat                        ← double-clic = conversion (Windows)
└── FACTURE CLIENT/                      ← dossier mère
    ├── pdf_to_excel.py                  ← le script
    ├── Modele_Import.xlsx               ← le modèle (11 colonnes) — NE PAS RENOMMER
    ├── ... les PDF à convertir ...      ← ils restent tous dans ce dossier mère
    ├── BSA/                             ← fichiers Excel de la société BSA
    │   ├── BSA JANVIER 2026.xlsx
    │   └── ...
    └── MCI/                             ← fichiers Excel de la société MCI
        ├── MCI JUILLET 2026.xlsx
        └── ...
```

Le script crée automatiquement un sous-dossier pour chaque nouvelle société détectée
et y range ses fichiers Excel. Il ne déplace pas les PDF : ceux-ci restent directement
dans **`FACTURE CLIENT`**.

### 💡 Le nom du PDF n'a aucune importance
Vous pouvez déposer les PDF depuis SALFA avec n'importe quel nom
(`facture.pdf`, `téléchargement.pdf`, …). Le script **lit le contenu du PDF** :

| Information lue dans le PDF | Ligne du PDF | Exemple |
|---|---|---|
| **Société** | `Doit : MCI CARE` | 1er mot → `MCI` |
| **Mois + année** | `Mois de prise en charge : Juillet 2026` | `JUILLET` `2026` |

Et il nomme le fichier Excel tout seul :

```
PDF déposé (n'importe quel nom)  →  Excel créé : SOCIETE MOIS ANNEE.xlsx

facture.pdf   →  BSA JUILLET 2026.xlsx     (si c'est une facture BSA de juillet)
facture.pdf   →  MCI JANVIER 2026.xlsx     (si c'est une facture MCI de janvier)
```

> ⚠️ **Pas de fichier consolidé** : un PDF = un fichier Excel, rien d'autre.

---

## 3. Utilisation quotidienne — la plus simple

1. 📥 Déposer le(s) PDF dans le dossier mère **`FACTURE CLIENT`**
2. 🖱️ **Double-cliquer sur `CONVERTIR.bat`**
3. ✅ Récupérer chaque Excel dans **`FACTURE CLIENT/SOCIETE`** → `SOCIETE MOIS ANNEE.xlsx`
   (par exemple **`FACTURE CLIENT/BSA/BSA JANVIER 2026.xlsx`**)

C'est tout ! La fenêtre reste ouverte pour voir le résultat (appuyer sur une
touche pour fermer).

---

## 4. Les commandes (si vous préférez le terminal)

Ouvrir l'invite de commandes **dans le dossier IMPORTATION**
(Windows + R → `cmd` → `cd` jusqu'au dossier, ou *Maj + clic droit → Ouvrir la fenêtre PowerShell ici*) :

```
python "FACTURE CLIENT\pdf_to_excel.py"              tout convertir
python "FACTURE CLIENT\pdf_to_excel.py" MCI          uniquement les factures MCI
python "FACTURE CLIENT\pdf_to_excel.py" BSA          uniquement les factures BSA
python "FACTURE CLIENT\pdf_to_excel.py" --force      régénérer même si l'Excel existe déjà
```

> 🛡️ **Sans `--force`, les Excel existants ne sont JAMAIS écrasés** : vos
> modifications manuelles sont protégées. Message affiché dans ce cas :
> `-- MCI JUIN 2026.xlsx : existe déjà, non écrasé`

---

## 5. Ce que produit le script

Pour chaque PDF, une ligne Excel = un soin (agent) avec les 11 colonnes du modèle :

| Colonne | Contenu | Exemple |
|---|---|---|
| `Numero_Facture` | n° de la facture du PDF | `FA-01/MCI/26-034` |
| `Date_Soins` | date du soin | `2026-01-02` |
| `Nom_Agent` | nom + sous-société entre parenthèses | `RAFANOMEZANTSOA STEVEN (PAMF)` |
| `Matricule` | n° Mlle (si présent) | `151663` |
| `Societe` | société lue dans le PDF (`Doit :`) | `MCI` |
| `Sous_Societe` | employeur entre parenthèses | `UWS MADAGASCAR` |
| `Acte_Medicale_Prix` | actes détaillés séparés par `;` | `CONS : 20 000; MEDIC : 83 000` |
| `Montant_Total_Brut` | montant total | `103000` |
| `Ticket_Moderateur` | participation | `20600` |
| `Prise_En_Charge_Net` | net à payer | `82400` |
| `Observations` | facture + mois (+ alerte éventuelle) | `Facture mensuelle soins ambulatoires - Janvier 2026` |

Mise en forme identique au modèle : largeurs de colonnes, police Calibri 12,
nombres au format `#,##0`, en-tête figé.

---

## 6. Vérifications automatiques intégrées

- ✅ Lignes coupées entre deux pages d'un PDF **rattachées à la bonne ligne**
- ✅ Si la somme des actes ≠ montant facturé → mention
  `ATTENTION : ...` dans `Observations` (à contrôler manuellement)
- ❌ PDF illisible → `!! <fichier> : aucune ligne trouvée`

**Contrôle recommandé** : comparer le TOTAL `Montant_Total_Brut` de l'Excel avec
la ligne « Total » imprimée en bas de la facture PDF. Ils doivent être égaux.

---

## 7. Problèmes fréquents (Windows)

| Problème | Solution |
|---|---|
| `python` n'est pas reconnu | Réinstaller Python en cochant **« Add Python to PATH »**, ou essayer `py "FACTURE CLIENT\pdf_to_excel.py"` |
| `ModuleNotFoundError: pdfplumber` | Relancer `pip install pdfplumber openpyxl` |
| `Le fichier est utilisé par un autre processus` | Fermer l'Excel ouvert puis relancer |
| Le PDF n'est pas converti | Vérifier que le PDF est bien **directement dans `FACTURE CLIENT`** (pas dans un sous-dossier) |
| Excel pas mis à jour | Fichier existant protégé → utiliser `--force` |
| La société détectée est étrange | Elle vient du 1er mot de la ligne `Doit : ...` du PDF |

---

## 8. Récapitulatif express

```
1. pip install pdfplumber openpyxl          (une seule fois)
2. Déposer les PDF dans FACTURE CLIENT      (n'importe quel nom)
3. Double-cliquer CONVERTIR.bat
4. Excel dans FACTURE CLIENT/SOCIETE        -> SOCIETE MOIS ANNEE.xlsx
5. Les PDF restent dans FACTURE CLIENT
```
