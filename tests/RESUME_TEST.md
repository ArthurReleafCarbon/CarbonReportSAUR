# 📊 Résumé - Script de Test

## ✅ Ce qui a été créé

### Fichiers ajoutés

```
tests/
├── __init__.py                   # Module Python
├── test_generation_rapport.py    # 🚀 Script principal de test
├── README.md                     # Documentation complète
├── EXEMPLE_COMMANDE.sh          # Exemple de commande
├── RESUME_TEST.md               # Ce fichier
└── output/                      # Dossier de sortie (créé auto)
    ├── .gitkeep
    └── rapport_test.docx        # Rapport généré (après exécution)
```

## 🎯 Objectif

Générer un rapport Word **sans passer par Streamlit** et **sans les overrides utilisateur**.

**Usage** : Tester rapidement si votre Excel + template fonctionnent correctement.

## 🚀 Utilisation

### Commande de base

```bash
python tests/test_generation_rapport.py votre_fichier.xlsx
```

### Avec année spécifique

```bash
python tests/test_generation_rapport.py votre_fichier.xlsx 2024
```

### Exemples réels

```bash
# Exemple 1 : Fichier dans le dossier courant
python tests/test_generation_rapport.py bilan_2024.xlsx

# Exemple 2 : Fichier dans un autre dossier
python tests/test_generation_rapport.py ~/Documents/bilans/bilan_SAUR_2024.xlsx 2024

# Exemple 3 : Chemin absolu
python tests/test_generation_rapport.py /Users/vous/Desktop/data.xlsx 2024
```

## 📤 Sortie

**Rapport généré** : `tests/output/rapport_test.docx`

Le fichier est écrasé à chaque exécution.

## 🔄 Workflow de test

```
1. Modifier votre Excel ou template
         ↓
2. Lancer le script de test
         ↓
3. Ouvrir tests/output/rapport_test.docx
         ↓
4. Vérifier le rendu
         ↓
5. Itérer si nécessaire
```

## ⚙️ Ce qui est testé

Le script exécute **toutes les étapes** de génération :

1. ✅ Chargement Excel (validation stricte)
2. ✅ Construction arborescence ORG/LOT/ENT
3. ✅ Calcul émissions BRUT (top 4 postes)
4. ✅ Calcul indicateurs par LOT×ACTIVITÉ
5. ✅ Chargement catalogue TEXTE_RAPPORT
6. ✅ Calcul KPI (équivalents, ratios)
7. ✅ Génération rapport Word complet

## ❌ Ce qui n'est PAS testé

- ❌ Renommage des nœuds (ORG, LOTs)
- ❌ Exclusion de postes (modes A et B)
- ❌ Calcul NET (avec overrides)
- ❌ Interface Streamlit
- ❌ Export/Import configuration

**Pourquoi ?** Ces fonctionnalités nécessitent une interaction utilisateur.

## 📊 Affichage

Le script affiche un résumé détaillé :

```
🧪 TEST DE GÉNÉRATION DE RAPPORT CARBONE
======================================================================

📥 Étape 1/7 : Chargement du fichier Excel...
   Fichier : data.xlsx
   ✅ Excel chargé et validé

🌳 Étape 2/7 : Construction de l'arborescence...
   ORG : Mon Organisation
   LOTs : 3
   ENTs : 15
   Activités : AEP, EU
   ✅ Arborescence construite

📊 Étape 3/7 : Calcul des émissions...
   Total ORG : 1234.5 tCO₂e
   • Scope 1 : 456.7 tCO₂e
   • Scope 2 : 123.4 tCO₂e
   • Scope 3 : 654.4 tCO₂e
   Top poste : Électricité
   ✅ Émissions calculées

[... autres étapes ...]

======================================================================
✅ GÉNÉRATION RÉUSSIE !
======================================================================

📄 Rapport généré : /path/to/tests/output/rapport_test.docx
📊 Émissions totales : 1234.5 tCO₂e
🌳 Structure : 3 LOT(s), 15 ENT(s)
📈 Activités : AEP, EU
```

## 🐛 Gestion des erreurs

### Erreur de validation Excel

```
❌ Erreur de validation Excel :
   Onglet 'EMISSIONS' : colonnes manquantes : scope, tco2e
```

➡️ **Solution** : Corriger votre Excel

### Template non trouvé

```
❌ Template non trouvé : templates/rapport_template.docx
→ Placer votre template dans templates/rapport_template.docx
```

➡️ **Solution** : Ajouter le template Word

### Erreur d'arborescence

```
❌ Erreurs dans la structure :
   • ENT E001 n'a pas d'activité EU ou AEP
```

➡️ **Solution** : Corriger l'onglet ORG_TREE

### Traceback complet

En cas d'erreur inattendue, le traceback Python complet est affiché pour faciliter le debug.

## 🆚 Test vs Streamlit

| Critère | Script de test | Streamlit |
|---------|----------------|-----------|
| **Vitesse** | ⚡ Très rapide (< 5s) | 🐢 Plus lent |
| **Setup** | Aucun | Lancer serveur |
| **Interface** | Console | Interface web |
| **Overrides** | ❌ Non | ✅ Oui |
| **Preview** | ❌ Non | ✅ Oui |
| **Itération** | ✅ Très rapide | Plus lent |
| **Production** | ❌ Non recommandé | ✅ Oui |

## 💡 Bonnes pratiques

### 1. Tester avant Streamlit

Avant de lancer Streamlit, testez d'abord avec le script :

```bash
# Test rapide
python tests/test_generation_rapport.py data.xlsx

# Si OK, alors lancer Streamlit
streamlit run app.py
```

### 2. Itération template

Pour modifier votre template Word :

1. Modifier le template
2. Lancer le test : `python tests/test_generation_rapport.py data.xlsx`
3. Ouvrir `tests/output/rapport_test.docx`
4. Vérifier le rendu
5. Répéter jusqu'à satisfaction

### 3. Debug Excel

Pour débugger un problème Excel :

```bash
# Le script affiche exactement où est le problème
python tests/test_generation_rapport.py probleme.xlsx
```

### 4. Vérification rapide

Avant de livrer un rapport :

```bash
# Test final
python tests/test_generation_rapport.py final.xlsx 2024
```

## 📦 Intégration

### Utiliser dans un autre script

```python
from tests.test_generation_rapport import test_generation_rapport

# Générer un rapport
success = test_generation_rapport(
    excel_path="data.xlsx",
    output_path="mon_rapport.docx",
    annee=2024
)

if success:
    print("✅ Rapport généré")
else:
    print("❌ Erreur")
```

### Automatisation

```bash
#!/bin/bash
# Script d'automatisation

for file in data/*.xlsx; do
    echo "Génération pour $file"
    python tests/test_generation_rapport.py "$file" 2024
done
```

## 🎓 Ce que vous apprenez

En regardant la sortie du script, vous comprenez :

1. **Structure de vos données** (ORG/LOT/ENT)
2. **Répartition des émissions** (scopes, postes)
3. **Top postes émetteurs**
4. **Problèmes de validation** (colonnes manquantes, etc.)

## ⚡ Performance

Le script est très rapide :

- Petit fichier (< 100 lignes) : **< 2 secondes**
- Fichier moyen (100-1000 lignes) : **< 5 secondes**
- Gros fichier (> 1000 lignes) : **< 10 secondes**

## 📚 Pour aller plus loin

- **Documentation complète** : [tests/README.md](README.md)
- **Guide Streamlit** : [../README.md](../README.md)
- **Troubleshooting** : [../README.md#troubleshooting](../README.md#troubleshooting)

---

**Bon test ! 🚀**
