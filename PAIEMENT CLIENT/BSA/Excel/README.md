# Relevés Excel BSA

Les fichiers `.xlsx` déposés directement dans ce dossier sont les relevés BSA
mis en page (et non encore le format d'importation).

Pour les convertir selon `Modele_Import_Reglements_Decompte_Assurance.xlsx` :

```bat
cd ..
python BSA_excel_to_modele.py
```

Ou double-cliquez sur `BSA_paiement.bat` dans le dossier parent.

Les fichiers convertis sont créés dans le sous-dossier `Import` :

- `Import/14-08-26 BSA 2026 09-05-26 MONTANT 38 240Ar.xlsx`
- `Import/20-08-26 BSA 2026 14-02-26 à 12-05-26 MONTANT 1 226 080Ar.xlsx`

Chaque fichier contient la feuille `Modele_Reglements` et ses 13 colonnes.
Les relevés sources ne sont jamais modifiés. Utilisez `--force` pour régénérer
un fichier d'importation existant.
