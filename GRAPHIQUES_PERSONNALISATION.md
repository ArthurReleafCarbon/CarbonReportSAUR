# 🎨 Personnalisation des Graphiques

## Vue d'ensemble

Une nouvelle page **"🎨 Graphiques"** a été ajoutée à l'application Streamlit pour permettre la prévisualisation et la personnalisation de tous les graphiques avant la génération du rapport.

## Fonctionnalités

### 1. Prévisualisation en temps réel
- Visualisation de tous les graphiques dans une grille organisée
- Aperçu instantané des modifications
- Organisation par onglets (Globaux, LOT, Autres)

### 2. Personnalisation disponible

Pour chaque graphique, vous pouvez modifier :

- **📝 Titre** : Personnaliser le titre du graphique
- **🎨 Couleurs** : Choisir parmi plusieurs palettes prédéfinies
  - Verte (défaut)
  - Bleue
  - Rouge/Orange
  - Personnalisée (à venir)
- **📊 Légende** : Afficher ou masquer la légende

### 3. Application automatique

Les personnalisations sont automatiquement appliquées lors de la génération du rapport final.

## Utilisation

### Étape 1 : Charger les données
1. Dans la section principale, uploadez votre fichier Excel
2. Attendez que les données soient chargées

### Étape 2 : Accéder à la page Graphiques
1. Cliquez sur "🎨 Graphiques" dans le menu de navigation
2. Les graphiques se chargeront automatiquement

### Étape 3 : Personnaliser
1. Cliquez sur un graphique pour le développer
2. Modifiez les paramètres dans le panneau de droite :
   - Changez le titre
   - Sélectionnez une palette de couleurs
   - Activez/désactivez la légende
3. Cliquez sur "💾 Appliquer" pour sauvegarder

### Étape 4 : Générer le rapport
1. Retournez à la section "📄 Génération"
2. Générez le rapport comme d'habitude
3. Les graphiques personnalisés seront automatiquement intégrés

## Architecture technique

### Fichiers modifiés/créés

1. **`src/streamlit_charts_page.py`** (NOUVEAU)
   - Page Streamlit dédiée aux graphiques
   - Gestion de la personnalisation
   - Prévisualisation en temps réel

2. **`app.py`** (MODIFIÉ)
   - Ajout de l'import de la nouvelle page
   - Ajout de "🎨 Graphiques" dans le menu
   - Initialisation des personnalisations dans `init_session_state()`
   - Stockage des `poste_labels` dans session_state

### Structure du session_state

```python
st.session_state.chart_customization = {
    'chart_key': {
        'title': str,
        'colors': List[str],  # Liste de codes couleurs hex
        'show_legend': bool
    },
    # ... pour chaque graphique
}
```

### Graphiques supportés

- `FILE_EAU_BREAKDOWN` : Répartition file eau STEP
- `EM_INDIRECTES_SPLIT` : Émissions indirectes
- `chart_emissions_scope_org` : Scopes ORG
- `chart_contrib_lot` : Contribution des LOTs
- `chart_emissions_total_org` : Contribution des postes
- `chart_emissions_elec_org` : Électricité par activité

## Rétrocompatibilité

✅ **Aucun impact sur le code existant**

- La génération de rapport sans personnalisation fonctionne exactement comme avant
- Les graphiques utilisent les valeurs par défaut si aucune personnalisation n'est définie
- Le comportement actuel de l'application est préservé à 100%

## Évolutions futures possibles

- [ ] Palettes de couleurs personnalisées (sélecteur de couleur)
- [ ] Export/Import des configurations de graphiques
- [ ] Prévisualisation côte-à-côte (avant/après)
- [ ] Modification de la taille des graphiques
- [ ] Personnalisation des polices
- [ ] Ajout de notes/annotations sur les graphiques

## Troubleshooting

### Les graphiques ne s'affichent pas
- Assurez-vous que les données sont chargées (section Aperçu)
- Vérifiez que les émissions sont calculées

### Les modifications ne sont pas appliquées
- Cliquez bien sur "💾 Appliquer" après chaque modification
- Vérifiez que vous êtes dans la bonne section lors de la génération

### Erreur lors de la prévisualisation
- Vérifiez les logs dans la console
- Assurez-vous que toutes les dépendances sont installées
- Rechargez la page si nécessaire
