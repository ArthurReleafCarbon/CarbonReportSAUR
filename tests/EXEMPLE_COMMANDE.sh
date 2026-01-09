#!/bin/bash
# Exemple de commande pour tester la génération de rapport

# Remplacer par le chemin vers votre fichier Excel
EXCEL_FILE="votre_fichier.xlsx"

# Année du bilan
ANNEE=2024

# Lancer le test
python tests/test_generation_rapport.py "$EXCEL_FILE" $ANNEE

# Le rapport généré sera dans : tests/output/rapport_test.docx

echo ""
echo "✅ Rapport généré dans : tests/output/rapport_test.docx"
echo "📂 Ouvrir le dossier : open tests/output/"
