# Vérification de cohérence - Importation paiement BSA

## Date de vérification
28 août 2026

## Règles de mapping BSA

Pour l'importation des paiements BSA, les colonnes du modèle Excel sont mappées comme suit :

| Colonne Excel | Source PDF BSA | Description |
|---------------|----------------|-------------|
| **Montant_Reclame_Brut** | FR.REELS | Frais réels réclamés |
| **Ticket_Moderateur** | Tx (%) | Taux de prise en charge (pourcentage) |
| **Montant_Paye_Regle** | REMB | Montant remboursé par BSA |
| **Montant_Exclu_Rejet** | NON_REMB - TPG | Montant exclu réel (calculé) |

### Note importante sur le calcul
Dans le PDF BSA, la colonne **NON_REMB** inclut le TPG (ticket modérateur).  
Le **vrai montant exclu** = NON_REMB - TPG

## Vérification de cohérence

### Formule de cohérence
Pour chaque ligne de données :
```
FR.REELS = REMB + (NON_REMB - TPG)
```

Soit en termes Excel :
```
Montant_Reclame_Brut = Montant_Paye_Regle + Montant_Exclu_Rejet
```

### Résultats de vérification

**Fichier testé :** `[2025-10-01] 16-09-2025 BFV-SG 1127394 SALFATU LOT69_092025.pdf`

- **Total lignes traitées :** 51
- **Incohérences détectées :** 0 ✅
- **Contrôle virement :** 777 650 Ar payés = montant du virement (777 650 Ar) ✅

Toutes les lignes respectent la formule de cohérence.

## Modifications apportées au script

### Fichier modifié
`PAIEMENT CLIENT/BSA/BSA_paiement_to_excel.py`

### Changements
1. **Ticket_Moderateur** : changé de `TPG` vers `Tx (%)`
   - Avant : `"Ticket_Moderateur": num(tpg)`
   - Après : `"Ticket_Moderateur": num(tx)`

2. **Vérification de cohérence** : ajout d'un contrôle automatique par ligne
   - Vérifie que `FR.REELS = REMB + (NON_REMB - TPG)` pour chaque ligne
   - Affiche un message d'alerte si incohérence détectée

3. **Documentation** : mise à jour des commentaires et docstrings

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
Ticket_Mod : 100
Exclu_Rejet: 0 (20 000 - 20 000)
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
Ticket_Mod : 95
Exclu_Rejet: 750 (750 - 0)
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
Ticket_Mod : 95
Exclu_Rejet: 5 000 (5 000 - 0)
Cohérence  : 20 000 = 15 000 + 5 000 ✅
```

## Utilisation

Le script peut maintenant être utilisé en toute confiance :
```bash
python "PAIEMENT CLIENT/BSA/BSA_paiement_to_excel.py"
```

Toute incohérence dans les données sera automatiquement signalée lors de l'exécution.
