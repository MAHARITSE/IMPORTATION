# Vérification de cohérence - Importation paiement BSA

## Date de vérification
28 août 2026

## Règle de cohérence BSA (la règle du reste)

> **Pour BSA : si FR.REELS = REMB alors le reste = 0.**

Le « reste » (montant non remboursé, colonne **Montant_Exclu_Rejet**) est
défini par :

```
le reste = FR.REELS - REMB
```

Conséquences :
- si **FR.REELS = REMB** (tout est remboursé) → **le reste = 0** ;
- sinon le reste = la part non remboursée, et l'identité comptable
  **Montant_Reclame_Brut = Montant_Paye_Regle + Montant_Exclu_Rejet**
  est respectée sur chaque ligne.

## Règles de mapping BSA

Pour l'importation des paiements BSA, les colonnes du modèle Excel sont mappées comme suit :

| Colonne Excel | Source PDF BSA | Description |
|---------------|----------------|-------------|
| **Montant_Reclame_Brut** | FR.REELS | Frais réels réclamés |
| **Ticket_Moderateur** | Tx (%) | Taux de prise en charge (pourcentage) |
| **Montant_Paye_Regle** | REMB | Montant remboursé par BSA |
| **Montant_Exclu_Rejet** | FR.REELS − REMB | **Le reste** (= 0 si FR.REELS = REMB) |

### Note importante sur le calcul
Dans le PDF BSA, la colonne **NON_REMB** inclut le TPG (ticket modérateur).
Le PDF décompose donc le reste en : `NON_REMB − TPG`.
Cette valeur sert de **contrôle croisé** : elle doit être **égale au reste**
(FR.REELS − REMB). Tout écart d'au moins 1 Ar est signalé comme
**INCOHÉRENCE** (console + colonne Motif_Observation).

## Vérification de cohérence

### Formules vérifiées sur chaque ligne
```
1) le reste = FR.REELS - REMB              (règle du reste)
2) si FR.REELS = REMB  =>  reste = 0       (cas tout remboursé)
3) NON_REMB - TPG = le reste               (contrôle croisé avec le PDF)
4) Montant_Reclame_Brut = Montant_Paye_Regle + Montant_Exclu_Rejet
```

### Résultats de vérification

**Fichier testé :** `[2025-10-01] 16-09-2025 BFV-SG 1127394 SALFATU LOT69_092025.pdf`
**Excel produit :** `BSA/2025/16-09-25 BSA 2025 05-06-25 à 29-07-25 MONTANT 777 650Ar.xlsx`

- **Total lignes traitées :** 51
- **Lignes avec FR.REELS = REMB :** 43 → **toutes avec reste = 0** ✅ (43/43)
- **Violations de la règle « FR.REELS = REMB → reste = 0 » :** 0 ✅
- **Incohérences détectées :** 0 ✅
- **Identité Brut = Payé + Reste :** 0 violation ✅
- **Totaux :** FR.REELS 888 500 Ar = REMB 777 650 Ar + reste 110 850 Ar ✅
- **Contrôle virement :** 777 650 Ar payés = montant du virement (777 650 Ar) ✅

Toutes les lignes respectent la règle de cohérence.

## Modifications apportées au script

### Fichier modifié
`PAIEMENT CLIENT/BSA/BSA_paiement_to_excel.py`

### Changements
1. **Ticket_Moderateur** : changé de `TPG` vers `Tx (%)`
   - Avant : `"Ticket_Moderateur": num(tpg)`
   - Après : `"Ticket_Moderateur": num(tx)`

2. **Montant_Exclu_Rejet = le reste** :
   - Avant : `NON_REMB - TPG` (valeur du PDF, sans garantie d'identité comptable)
   - Après : `FR.REELS - REMB` — **si FR.REELS = REMB alors le reste = 0**,
     garanti par construction sur chaque ligne

3. **Vérification de cohérence** : ajout de la fonction `controle_reste()`
   - Calcule le reste = FR.REELS − REMB et vérifie la règle
     « si FR.REELS = REMB alors le reste = 0 »
   - Contrôle croisé : le reste doit être égal à NON_REMB − TPG
   - Signale aussi un REMB supérieur à FR.REELS
   - Message `!! INCOHÉRENCE ...` en console + note dans Motif_Observation
   - Récapitulatif par fichier :
     `cohérence : 51 lignes vérifiées — si FR.REELS = REMB alors le reste = 0  OK`

4. **Documentation** : mise à jour des commentaires et docstrings

## Exemples de données

### Cas 1 : Tx = 100% (prise en charge totale)
```
Date       : 2025-06-05
Nom        : RAZANAJATOVO ALPHONSE
FR.REELS   : 80 000
Tx (%)     : 100
REMB       : 80 000
NON_REMB   : 20 000
TPG        : 20 000
─────────────────────────────────────
FR.REELS = REMB  →  le reste = 0 ✅
Ticket_Mod : 100
Exclu_Rejet: 0
Cohérence  : 80 000 = 80 000 + 0 ✅
```

### Cas 2 : Tx = 95% (prise en charge partielle)
```
Date       : 2025-07-02
Nom        : RAZAFINIHATRAINA ROGER
FR.REELS   : 15 000
Tx (%)     : 95
REMB       : 14 250
NON_REMB   : 750
TPG        : 0
─────────────────────────────────────
FR.REELS ≠ REMB  →  le reste = 15 000 - 14 250 = 750
Ticket_Mod : 95
Exclu_Rejet: 750
Cohérence  : 15 000 = 14 250 + 750 ✅
```

### Cas 3 : Acte limité avec Tx = 95%
```
Date       : 2025-06-28
Nom        : RATSIMBA NANTENAINA ALLAN
FR.REELS   : 20 000
Tx (%)     : 95
REMB       : 15 000 (limité par acte)
NON_REMB   : 5 000
TPG        : 0
─────────────────────────────────────
le reste   : 20 000 - 15 000 = 5 000
Ticket_Mod : 95
Exclu_Rejet: 5 000
Cohérence  : 20 000 = 15 000 + 5 000 ✅
```

### Cas 4 : Tx = 95% mais tout remboursé (NON_REMB = TPG)
```
Date       : 2025-06-19
FR.REELS   : 20 000
Tx (%)     : 95
REMB       : 20 000
NON_REMB   : 5 000
TPG        : 5 000
─────────────────────────────────────
FR.REELS = REMB  →  le reste = 0 ✅
Exclu_Rejet: 0  (NON_REMB - TPG = 5 000 - 5 000 = 0 : contrôle croisé OK)
Cohérence  : 20 000 = 20 000 + 0 ✅
```

## Utilisation

Le script peut maintenant être utilisé en toute confiance :
```bash
python "PAIEMENT CLIENT/BSA/BSA_paiement_to_excel.py"
```

La règle « si FR.REELS = REMB alors le reste = 0 » est **garantie par
construction** (Montant_Exclu_Rejet = FR.REELS − REMB) et toute incohérence
des données du PDF est automatiquement signalée lors de l'exécution.
