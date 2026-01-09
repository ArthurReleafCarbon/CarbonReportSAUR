# Instructions pour ajouter vos fichiers

## 📁 Fichiers à ajouter

### 1. Template Word

**Emplacement** : `templates/rapport_template.docx`

Placez votre template Word avec les placeholders définis dans le brief :

**Placeholders globaux :**
- `{{annee}}`, `{{ORG_NAME}}`, `{{TOTAL_EMISSIONS}}`, etc.

**Blocs répétables :**
```
[[START_LOT]]
  [[START_ACTIVITY]]
    [[START_POST]]
    [[END_POST]]
  [[END_ACTIVITY]]
[[END_LOT]]
```

### 2. Assets (images)

**Emplacement** : `assets/`

**Fichiers requis :**
- `assets/logo_org.png` : Logo de votre organisation
- `assets/digesteur_schema.png` : Schéma du digesteur

**Fichiers optionnels :**
- `assets/icones/` : Icônes pour les postes (si utilisé)

### 3. Fichier Excel de données

**Format attendu** : Voir README.md section "Format Excel attendu"

**Comment l'utiliser :**
1. Lancer l'application : `streamlit run app.py`
2. Uploader votre fichier via l'interface web

## 🚀 Démarrage

Une fois tous les fichiers en place :

```bash
# 1. Installer les dépendances
pip install -r requirements.txt

# 2. Placer votre template Word dans templates/
cp votre_template.docx templates/rapport_template.docx

# 3. Placer vos assets dans assets/
cp votre_logo.png assets/logo_org.png
cp votre_schema.png assets/digesteur_schema.png

# 4. Lancer l'application
streamlit run app.py
```

## ✅ Checklist avant de commencer

- [ ] Template Word dans `templates/rapport_template.docx`
- [ ] Logo dans `assets/logo_org.png`
- [ ] Schéma digesteur dans `assets/digesteur_schema.png`
- [ ] Fichier Excel de données prêt
- [ ] Dépendances Python installées

## 📝 Structure du template Word

Votre template doit contenir :

1. **Section globale ORG** avec placeholders :
   - `{{ORG_NAME}}`, `{{TOTAL_EMISSIONS}}`, etc.
   - Graphiques : `{{chart_emissions_scope_org}}`, etc.

2. **Blocs LOT** (si applicable) :
   ```
   [[START_LOT]]
   Nom du LOT : {{LOT_NAME}}
   [[END_LOT]]
   ```

3. **Blocs ACTIVITÉ** par LOT :
   ```
   [[START_ACTIVITY]]
   Activité : {{ENT_ACTIVITY}}
   [[END_ACTIVITY]]
   ```

4. **Blocs POSTE détaillés** :
   ```
   [[START_POST]]
   {{POST_TITLE}}
   {{POST_TEXT}}
   {{POST_CHART_1}}
   {{POST_TABLE_1}}
   [[END_POST]]
   ```

## 🔧 Personnalisation

### Modifier le nombre de top postes

Par défaut : 4 postes

Pour changer : Modifier `top_n=4` dans [app.py](app.py) ligne ~380

### Ajouter des graphiques

1. Ajouter la CHART_KEY dans `TEXTE_RAPPORT` Excel
2. Implémenter le générateur dans [src/chart_generators.py](src/chart_generators.py)
3. Ajouter la key dans `SUPPORTED_CHART_KEYS`

### Ajouter des tableaux

1. Ajouter la TABLE_KEY dans `TEXTE_RAPPORT` Excel
2. Implémenter le générateur dans [src/table_generators.py](src/table_generators.py)
3. Ajouter la key dans `SUPPORTED_TABLE_KEYS`

## 📞 Support

En cas de problème :
1. Vérifier les logs dans la console Streamlit
2. Consulter le README.md section "Troubleshooting"
3. Vérifier que tous les fichiers sont au bon emplacement
