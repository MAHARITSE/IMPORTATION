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
├── CONVERTIR.bat                        ← double-clic = menu : 1 factures / 2 paiements / 0 annuler
├── FACTURE CLIENT/                      ← dossier mère des FACTURES
│   ├── pdf_to_excel.py                  ← le script factures
│   ├── Modele_Import.xlsx               ← le modèle (11 colonnes) — NE PAS RENOMMER
│   ├── ... les PDF à convertir ...      ← ils restent tous dans ce dossier mère
│   ├── BSA/                             ← fichiers Excel de la société BSA
│   │   ├── BSA JANVIER 2026.xlsx
│   │   └── ...
│   └── MCI/                             ← fichiers Excel de la société MCI
│       ├── MCI JUILLET 2026.xlsx
│       └── ...
└── PAIEMENT CLIENT/                     ← paiements des assurances (voir §9)
    ├── Modele_Import_Reglements_Decompte_Assurance.xlsx  — NE PAS RENOMMER
    ├── BSA/                             ← PDF + Excel + script de la société BSA
    │   ├── BSA_paiement_to_excel.py     ← le script paiements BSA
    │   ├── BSA_paiement.bat             ← double-clic = conversion BSA
    │   ├── PDF/                         ← les PDF BSA se déposent ICI
    │   └── ERREUR/                      ← (créé si besoin) les PDF en erreur sont déplacés ICI
    ├── MCI CARE/                        ← PDF + Excel + script de MCI CARE
    │   ├── MCI_CARE_paiement_to_excel.py ← le script paiements MCI CARE
    │   ├── MCI_CARE_paiement.bat         ← double-clic = conversion MCI CARE
    │   ├── PDF/                         ← les PDF MCI CARE se déposent ICI
    │   └── ERREUR/                      ← (créé si besoin) les PDF en erreur sont déplacés ICI
    └── ASCOMA/                          ← PDF + Excel + script d'ASCOMA
        ├── ASCOMA_to_excel.py           ← le script paiements ASCOMA
        ├── ASCOMA_paiement.bat          ← double-clic = conversion ASCOMA
        ├── PDF/                         ← les PDF ASCOMA se déposent ICI
        └── ERREUR/                      ← (créé si besoin) les PDF en erreur sont déplacés ICI
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
   (pour les paiements : dans le sous-dossier **`PDF/`** de la société —
   `PAIEMENT CLIENT/BSA/PDF`, `MCI CARE/PDF` ou `ASCOMA/PDF`)
2. 🖱️ **Double-cliquer sur `CONVERTIR.bat`**
3. ⌨️ Au lancement, **taper le numéro voulu puis Entrée** :
   - **`1`** pour convertir les **FACTURES** (dossier FACTURE CLIENT)
   - **`2`** pour convertir les **PAIEMENTS** (dossier PAIEMENT CLIENT : BSA, MCI CARE, ASCOMA)
   - **`0`** pour **annuler** (aucune conversion)
4. ⌨️ Répondre à la question **« Ecraser les Excel existants pour les
   régénérer ? (O/N) »** :
   - **`O`** (oui) → les Excel existants sont **écrasés** et régénérés (= `--force`)
   - **`N`** (non) → les Excel existants sont **conservés**, jamais écrasés
5. ✅ Récupérer chaque Excel dans **`FACTURE CLIENT/SOCIETE`** → `SOCIETE MOIS ANNEE.xlsx`
   (par exemple **`FACTURE CLIENT/BSA/BSA JANVIER 2026.xlsx`**)

C'est tout ! La fenêtre reste ouverte pour voir le résultat (appuyer sur une
touche pour fermer). En cas de mauvaise touche, le menu (ou la question) se
réaffiche.

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
>
> 💡 C'est exactement ce que fait la question **O/N** de `CONVERTIR.bat` :
> répondre `O` revient à relancer les scripts avec `--force`.

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
3. Double-cliquer CONVERTIR.bat             puis taper 1 (factures),
                                            2 (paiements) ou 0 (annuler)
4. Répondre O (écraser les Excel existants) ou N (les conserver)
5. Excel dans FACTURE CLIENT/SOCIETE        -> SOCIETE MOIS ANNEE.xlsx
6. Les PDF restent dans FACTURE CLIENT
```

---

## 9. PAIEMENT CLIENT — décomptes de règlements des assurances

Le dossier **`PAIEMENT CLIENT`** contient **un script Python + un .bat par
société**, chacun adapté au format de paiement de sa société, selon le modèle
**`Modele_Import_Reglements_Decompte_Assurance.xlsx`**
(feuille `Modele_Reglements`, 13 colonnes).

### Organisation

```
PAIEMENT CLIENT/
├── Modele_Import_Reglements_Decompte_Assurance.xlsx   ← NE PAS RENOMMER
├── BSA/                                   ← société BSA (relevés de remboursements)
│   ├── BSA_paiement_to_excel.py           ← script du format BSA
│   ├── BSA_paiement.bat                   ← double-clic = conversion BSA
│   ├── PDF/                               ← les PDF BSA à convertir (dépôt ici)
│   │   └── 17-04-2026 BFV-SG 1129370 ....pdf
│   ├── ERREUR/                            ← (créé si besoin) PDF en erreur déplacés ici
│   └── 2026/                              ← dossier de l'ANNÉE du règlement (créé auto)
│       └── 17-04-26 BSA 2026 27-01-26 à 23-02-26 MONTANT 928 750Ar.xlsx ← l'Excel produit
├── MCI CARE/                              ← société MCI CARE (décomptes de règlement)
│   ├── MCI_CARE_paiement_to_excel.py      ← script du format MCI CARE
│   ├── MCI_CARE_paiement.bat              ← double-clic = conversion MCI CARE
│   ├── PDF/                               ← les PDF MCI CARE à convertir (dépôt ici)
│   │   └── DISPENSAIRE LOTERANA ....pdf
│   ├── ERREUR/                            ← (créé si besoin) PDF en erreur déplacés ici
│   └── 2026/                              ← dossier de l'ANNÉE du règlement (créé auto)
│       └── 02-05-26 MCI CARE 2026 02-03-26 à 31-03-26 MONTANT 471 140Ar.xlsx ← l'Excel produit
└── ASCOMA/                                ← société ASCOMA (décomptes tiers payant)
    ├── ASCOMA_to_excel.py                 ← script du format ASCOMA
    ├── ASCOMA_paiement.bat                ← double-clic = conversion ASCOMA
    ├── PDF/                               ← les PDF ASCOMA à convertir (dépôt ici)
    │   └── 7 035 543 23-01-25.pdf
    ├── ERREUR/                            ← (créé si besoin) PDF en erreur déplacés ici
    └── 2025/                              ← dossier de l'ANNÉE du règlement (créé auto)
        └── 09-01-25 ASCOMA 2025 13-05-24 à 31-08-24 MONTANT 7 035 543Ar.xlsx ← l'Excel produit
```

**Règle : c'est le DOSSIER qui décide, pas le contenu du PDF.**

| PDF déposé dans | Traité comme |
|---|---|
| `PAIEMENT CLIENT/BSA/` (ou `BSA/PDF/`) | **BSA** |
| `PAIEMENT CLIENT/MCI CARE/` (ou `MCI CARE/PDF/`) | **MCI CARE** |
| `PAIEMENT CLIENT/ASCOMA/` (ou `ASCOMA/PDF/`) | **ASCOMA** |

Chaque société a son script et son .bat. On dépose le PDF dans le sous-dossier
**`PDF/` de sa société**, puis on double-clique sur le .bat (ou `CONVERTIR.bat` choix 2).
Chaque Excel créé est rangé dans un sous-dossier **au nom de l'année du
règlement** (`BSA/2026/…`, `MCI CARE/2025/…`, `ASCOMA/2025/…`), créé
automatiquement si besoin.
Les PDF déposés directement dans le dossier de la société (ancienne habitude)
sont quand même trouvés et convertis.
Le nom du PDF lui-même n'a aucune importance.

### Utilisation

**`CONVERTIR.bat`** demande au lancement quoi convertir — **taper le numéro
voulu puis Entrée** :

```
1  -  les FACTURES   (dossier FACTURE CLIENT)
2  -  les PAIEMENTS  (dossier PAIEMENT CLIENT)
0  -  annuler
```

Le choix `2` lance **chaque société de paiement une à une** (BSA, MCI CARE
puis ASCOMA). Après le choix `1` ou `2`, `CONVERTIR.bat` demande :

```
Ecraser les Excel existants pour les regenerer ? (O/N)
```

- **`O`** : les Excel déjà présents sont **écrasés** et régénérés
  (les scripts sont relancés avec `--force`) ;
- **`N`** : ils sont **conservés** — seul le nouveau est créé
  (comportement par défaut, vos modifications manuelles sont protégées).

Ou, société par société, double-cliquer sur le .bat du dossier :

```
PAIEMENT CLIENT\BSA\BSA_paiement.bat              tout convertir pour BSA
PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement.bat    tout convertir pour MCI CARE
PAIEMENT CLIENT\ASCOMA\ASCOMA_paiement.bat        tout convertir pour ASCOMA
```

Ou en invite de commandes (dans le dossier de la société) :

```
python BSA_paiement_to_excel.py                    tous les PDF du sous-dossier BSA\PDF
python BSA_paiement_to_excel.py --force            régénérer (écrase l'Excel existant)
python BSA_paiement_to_excel.py "mon_releve.pdf"   un seul PDF

python MCI_CARE_paiement_to_excel.py               tous les PDF du sous-dossier MCI CARE\PDF
python MCI_CARE_paiement_to_excel.py --force       régénérer
python MCI_CARE_paiement_to_excel.py "mon_decompte.pdf"   un seul PDF

python ASCOMA_to_excel.py                          tous les PDF du sous-dossier ASCOMA\PDF
python ASCOMA_to_excel.py --force                  régénérer
python ASCOMA_to_excel.py "mon_decompte.pdf"       un seul PDF
```

### Nom du fichier Excel créé : DATE DE PAIEMENT + SOCIÉTÉ + ANNÉE + PÉRIODE + MONTANT

```
exemple : BSA/2026/17-04-26 BSA 2026 27-01-26 à 23-02-26 MONTANT 928 750Ar.xlsx
                 └── date paiement ──┘ └┬┘ └┬┘ └────────┬─────────┘ └──────┬──────┘
   dossier de l'année                    société année    intervalle des         montant
                                                           dates de soins         payé

exemple : MCI CARE/2026/02-05-26 MCI CARE 2026 02-03-26 à 31-03-26 MONTANT 471 140Ar.xlsx

exemple : ASCOMA/2025/09-01-25 ASCOMA 2025 13-05-24 à 31-08-24 MONTANT 7 035 543Ar.xlsx
```

| Pièce | Où le script la lit |
|---|---|
| **Dossier année** | année du règlement : les Excel sont rangés dans `SOCIETE/<ANNÉE>/` (`BSA/2026/…`) |
| **Société** | fixée dans le script du dossier (`BSA` / `MCI CARE` / `ASCOMA`) |
| **Date de paiement** | BSA : date du virement (`A , le 17/04/2026`) · MCI : `Date comptable` · ASCOMA : `Règlement du JJ/MM/AAAA` (sinon `Edité le`) ; écrite au début du nom au format `JJ-MM-AA` |
| **Intervalle des dates** | 1re et dernière date de soins des lignes (colonne `Date_Soins`), au format `JJ-MM-AA`. Une seule date si toutes les lignes tombent le même jour (`17-04-26 BSA 2026 17-04-26 MONTANT 20 000Ar.xlsx`) |
| **Montant** | BSA : somme des `REMB` (= montant du virement) · MCI : somme des « Montant réglé » de toutes les pages du décompte · ASCOMA : « Montant Net » du récapitulatif final (= Montant réglé − remise, le montant réellement payé). Il est écrit entre le mot `MONTANT` et le suffixe `Ar` : `MONTANT 928 750Ar` |

Le mois n'est plus écrit dans le nom du fichier : il est remplacé par la date
de paiement au début du nom. Les dates de l'intervalle gardent l'année sur
2 chiffres (`25` = 2025), même quand elles ne sont pas de la même année que
le paiement :
`MCI CARE/2026/03-01-26 MCI CARE 2026 03-11-25 à 16-11-25 MONTANT 206 000Ar.xlsx`
(le dossier suit l'année du **règlement**, ici 2026, pas celle des soins).

### Colonnes produites (13)

| Colonne | BSA (relevé de remboursements) | MCI (décompte de règlement) | ASCOMA (tiers payant) |
|---|---|---|---|
| `Ref_Decompte` | N° du relevé (`1129370`) | N° de la facture réglée, brut, comme les autres sociétés (`FA-02/MCI/26-047` — pas de suffixe « /10L ») | (vide : pas de n° dans le PDF) |
| `Date_Reglement` | date du virement | date comptable | « Règlement du » (sinon « Edité le ») |
| `Date_Soins` | date du soin (colonne DATE) | date de soins | date de soins |
| `Nom_Agent` | nom du patient (aligné à la date du soin) | bénéficiaire | bénéficiaire |
| `Matricule` | n° ADHESION | matricule | n° bénéficiaire (espaces retirés) |
| `Numero_Facture_Prescription` | facture SALFA de chaque décompte (ex `FA-02/BFV/26-022`, ou `N°006-25/BFV/BSA/SA` en format 2025) | même facture | (vide) |
| `Code_Acte` | CG, PH, ECH, EB, DC, SI, SUP... | PH, LABO, ... | Code Rem. (`1`, `2`, `3`...) |
| `Libelle_Acte` | médicament / acte détaillé | (vide si non détaillé) | acte médical (`CONSULT. GENERALISTE`, `PHARMACIE`...) |
| `Montant_Reclame_Brut` | FR.REELS | Montant réclamé | Montant réclamé |
| `Ticket_Moderateur` | Tx (%) — taux de prise en charge | Montant réclamé − Montant réglé | Ticket modérateur |
| `Montant_Paye_Regle` | REMB | Montant réglé | Montant réglé |
| `Montant_Exclu_Rejet` | **le reste = FR.REELS − REMB** (= 0 si FR.REELS = REMB) | Mtt non remboursé | Montant exclu |
| `Motif_Observation` | `Prise en charge : 95%` | `Ticket modérateur 10%` | (vide) |

### Vérifications automatiques

- ✅ BSA : **règle du reste** — le reste (`Montant_Exclu_Rejet`) = `FR.REELS − REMB`,
  donc **si FR.REELS = REMB alors le reste = 0** ; le reste est contrôlé sur
  chaque ligne contre `NON_REMB − TPG` (le NON_REMB du PDF inclut le TPG) :
  tout écart est signalé `!! INCOHÉRENCE ...` (console + `Motif_Observation`)
  et un récapitulatif s'affiche :
  `cohérence : 51 lignes vérifiées — si FR.REELS = REMB alors le reste = 0  OK`
- ✅ BSA : somme des REMB = montant du virement annoncé
- ✅ BSA : nombre de lignes lues = nombre déclaré dans le « Total général »
- ✅ MCI : somme des « Montant réglé » = « Total prestataire »
  (les totaux de toutes les pages du décompte sont additionnés : un PDF couvre
  parfois plusieurs établissements — PHIE, LABO, IMAGERIE, DISPENSAIRE)
- ✅ ASCOMA : somme des « Montant réglé » = somme des « S/Total Prestataire »
  (un par établissement)
- ✅ ASCOMA : le « Montant Net » du récapitulatif final = Montant réglé − remise
  (c'est ce montant réellement payé qui est écrit dans le nom du fichier ;
  un message affiche le détail, ex : `Montant Réglé 7 565 100 Ar − remise 529 557 Ar = net 7 035 543 Ar`)
- ❌ Si écart → ligne `!! ATTENTION : ...` affichée (à contrôler manuellement)
- 🛡️ Sans `--force`, les Excel existants ne sont JAMAIS écrasés

### 🚨 PDF en erreur — le sous-dossier `ERREUR` de chaque société

Si un PDF ne peut **pas** être converti en Excel, le script ne s'arrête pas :
il **crée automatiquement le sous-dossier `ERREUR` dans le dossier de la
société** (`PAIEMENT CLIENT\BSA\ERREUR`, `MCI CARE\ERREUR`, `ASCOMA\ERREUR`)
et y **déplace le PDF fautif**. Cas concernés :

| Cas | Message affiché |
|---|---|
| PDF illisible ou corrompu | `!! ... : PDF illisible (...)` |
| Format non reconnu / aucune ligne de règlement trouvée | `!! ... : aucune ligne de règlement trouvée (format non reconnu ?)` |
| Date de règlement introuvable dans le PDF | `!! ... : date de règlement introuvable dans le PDF` |
| Erreur pendant la création de l'Excel (ex : fichier ouvert dans Excel, verrouillé) | `!! ... : erreur pendant la création de l'Excel (...)` |

En fin de traitement, un récapitulatif liste tous les PDF en erreur :

```
== 1 PDF en erreur, déplacés dans le sous-dossier ERREUR de BSA ==
   - mon_fichier.pdf : PDF illisible ou corrompu
```

**Que faire ensuite ?**

1. Ouvrir le PDF dans `SOCIETE\ERREUR` pour voir ce qui cloche
   (mauvais dossier ? PDF tronqué ? pas un décompte de paiement ? Excel cible ouvert ?)
2. Corriger le problème, puis **remettre le PDF dans le sous-dossier `PDF/`**
   de la société et relancer la conversion.
3. Les PDF du dossier `ERREUR` ne sont **jamais retraités** par les
   conversions suivantes (pas de messages d'erreur répétés). Si un PDF y
   arrive par erreur, il suffit de le resortir. S'il y a déjà un fichier du
   même nom dans `ERREUR`, le nouveau est renommé `nom (1).pdf`, `nom (2).pdf`…
4. Un PDF déplacé dans `ERREUR` n'a **pas** produit d'Excel : aucun fichier
   incomplet n'est créé dans le dossier de l'année.
