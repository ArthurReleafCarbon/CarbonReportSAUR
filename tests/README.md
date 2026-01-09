# 🧪 Tests - Génération de Rapports

Ce dossier contient les scripts de test pour générer des rapports sans passer par Streamlit.

## 📋 Scripts disponibles

### `test_generation_rapport.py`

Script de test complet qui simule la génération d'un rapport Word **sans les overrides** (pas de renommage, pas d'exclusion de postes).

**Utilité** : Tester rapidement si votre Excel et votre template génèrent correctement un rapport.

## 🚀 Utilisation

### Commande basique

```bash
python tests/test_generation_rapport.py votre_fichier.xlsx
```

### Avec une année spécifique

```bash
python tests/test_generation_rapport.py votre_fichier.xlsx 2024
```

### Exemple complet

```bash
# Avec votre fichier Excel
python tests/test_generation_rapport.py ~/Downloads/bilan_carbone_2024.xlsx 2024
```

## 📤 Sortie

Le rapport généré sera dans : **`tests/output/rapport_test.docx`**

## 📊 Ce que fait le script

Le script exécute toutes les étapes de génération :

1. ✅ **Chargement Excel** - Validation du format
2. ✅ **Arborescence** - Construction ORG/LOT/ENT
3. ✅ **Calculs émissions** - BRUT uniquement (top 4 postes)
4. ✅ **Calculs indicateurs** - Par LOT×ACTIVITÉ
5. ✅ **Catalogue contenus** - Chargement TEXTE_RAPPORT
6. ✅ **Calcul KPI** - Équivalents et ratios
7. ✅ **Génération Word** - Rapport final

## ⚙️ Configuration

### Modifier le nombre de top postes

Dans le script, ligne avec `calculate_brut(top_n=4)` :

```python
results_brut = emission_calc.calculate_brut(top_n=5)  # Top 5 au lieu de 4
```

### Changer le chemin de sortie

```python
# Dans la fonction test_generation_rapport()
output_path = "mon_dossier/mon_rapport.docx"
test_generation_rapport(excel_path, output_path=output_path)
```

## 📝 Affichage détaillé

Le script affiche des informations détaillées pendant l'exécution :

```
🧪 TEST DE GÉNÉRATION DE RAPPORT CARBONE
======================================================================

📥 Étape 1/7 : Chargement du fichier Excel...
   Fichier : data.xlsx
   ✅ Excel chargé et validé

🌳 Étape 2/7 : Construction de l'arborescence...
   ORG : SAUR IDF
   LOTs : 3
   ENTs : 12
   Activités : AEP, EU
   ✅ Arborescence construite

📊 Étape 3/7 : Calcul des émissions...
   Total ORG : 1234.5 tCO₂e
   • Scope 1 : 456.7 tCO₂e
   • Scope 2 : 123.4 tCO₂e
   • Scope 3 : 654.4 tCO₂e
   Top poste : Électricité
   ✅ Émissions calculées

[...]

✅ GÉNÉRATION RÉUSSIE !
📄 Rapport généré : /path/to/tests/output/rapport_test.docx
📊 Émissions totales : 1234.5 tCO₂e
🌳 Structure : 3 LOT(s), 12 ENT(s)
📈 Activités : AEP, EU
```

## 🐛 En cas d'erreur

### Erreur "Template non trouvé"

```
❌ Template non trouvé : templates/rapport_template.docx
→ Placer votre template dans templates/rapport_template.docx
```

**Solution** : Ajouter votre template Word dans `templates/rapport_template.docx`

### Erreur de validation Excel

```
❌ Erreur de validation Excel :
   Onglet 'EMISSIONS' : colonnes manquantes : scope
```

**Solution** : Vérifier que votre Excel contient tous les onglets et colonnes requis (voir README.md principal)

### Erreur de structure arborescence

```
❌ Erreurs dans la structure :
   • ENT E001 n'a pas d'activité EU ou AEP
```

**Solution** : Corriger l'onglet ORG_TREE dans votre Excel

## 🔄 Workflow de test

1. **Préparer** votre fichier Excel
2. **Lancer** le script de test
3. **Ouvrir** le rapport généré dans `tests/output/`
4. **Vérifier** que tout s'affiche correctement
5. **Itérer** si nécessaire en modifiant Excel ou template
6. **Relancer** le test

## 📦 Structure du dossier

```
tests/
├── README.md                    # Cette documentation
├── test_generation_rapport.py   # Script de test principal
└── output/                      # Rapports générés (créé automatiquement)
    └── rapport_test.docx        # Dernier rapport généré
```

## 💡 Conseils

- **Test rapide** : Ce script est beaucoup plus rapide que de passer par Streamlit
- **Itération** : Parfait pour tester des modifications de template ou d'Excel
- **Debugging** : Les erreurs sont affichées avec le traceback complet
- **Sans overrides** : Le rapport est généré en mode BRUT (pas de modifications utilisateur)

## 🆚 Différence avec Streamlit

| Fonctionnalité | Test script | Streamlit |
|----------------|-------------|-----------|
| Chargement Excel | ✅ | ✅ |
| Calculs émissions | ✅ BRUT uniquement | ✅ BRUT + NET |
| Renommage nœuds | ❌ | ✅ |
| Exclusion postes | ❌ | ✅ |
| Génération Word | ✅ | ✅ |
| Preview interactif | ❌ | ✅ |
| Export/Import config | ❌ | ✅ |
| **Vitesse** | ⚡ Très rapide | 🐢 Plus lent |
| **Usage** | 🧪 Tests & debug | 👥 Production |

## 🎯 Cas d'usage

**Utiliser le script de test quand :**
- ✅ Vous développez/modifiez le template Word
- ✅ Vous testez un nouveau fichier Excel
- ✅ Vous déboguez un problème de génération
- ✅ Vous voulez générer rapidement sans UI

**Utiliser Streamlit quand :**
- ✅ Vous voulez renommer des nœuds
- ✅ Vous voulez exclure des postes
- ✅ Vous voulez prévisualiser avant génération
- ✅ Vous voulez sauvegarder votre configuration

---

**Bon test ! 🚀**
