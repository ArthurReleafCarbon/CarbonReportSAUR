# 🚀 Démarrage rapide - Carbon Report Generator

## 🧪 Test rapide (recommandé pour débuter)

**Testez la génération sans passer par Streamlit :**

```bash
# Tester avec votre fichier Excel
python tests/test_generation_rapport.py votre_fichier.xlsx 2024
```

✅ Rapport généré dans : `tests/output/rapport_test.docx`
✅ Très rapide pour itérer sur votre template
✅ Pas besoin de l'interface Streamlit

➡️ **Voir [tests/README.md](tests/README.md) pour plus de détails**

---

## 📱 Application complète (Streamlit)

### En 3 étapes

### 1️⃣ Vérifier l'installation

```bash
python check_setup.py
```

Ce script vérifie que tout est bien configuré.

### 2️⃣ Ajouter vos fichiers

**Template Word :**
```bash
cp votre_template.docx templates/rapport_template.docx
```

**Images :**
```bash
cp votre_logo.png assets/logo_org.png
cp votre_schema.png assets/digesteur_schema.png
```

### 3️⃣ Lancer l'application

```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur !

## 📋 Checklist rapide

Avant de commencer, assurez-vous d'avoir :

- [ ] Python 3.8+ installé
- [ ] Dépendances installées (`pip install -r requirements.txt`)
- [ ] Template Word dans `templates/`
- [ ] Images dans `assets/`
- [ ] Fichier Excel de données prêt

## 🎯 Workflow dans l'app

1. **Uploader** votre fichier Excel
2. **Vérifier** l'arborescence et les calculs dans "Aperçu"
3. **Configurer** les renommages et exclusions dans "Configuration"
4. **Générer** le rapport dans "Génération"
5. **Télécharger** le rapport .docx généré

## 📚 Documentation complète

- [README.md](README.md) - Documentation utilisateur complète
- [INSTRUCTIONS.md](INSTRUCTIONS.md) - Guide détaillé pour ajouter vos fichiers
- [PROJET_TERMINE.md](PROJET_TERMINE.md) - Récapitulatif du projet

## 🆘 Problèmes courants

**L'app ne démarre pas ?**
```bash
pip install -r requirements.txt --force-reinstall
```

**Template non trouvé ?**
```bash
ls -la templates/rapport_template.docx
# Doit afficher le fichier
```

**Erreur de validation Excel ?**
→ Vérifiez que tous les onglets requis sont présents (voir README.md)

## 💡 Astuce

Lancez `python check_setup.py` après chaque modification pour vérifier que tout est OK !

---

**Prêt à générer vos rapports carbone ! 🌍**
