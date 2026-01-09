# Carbon Report Generator - SAUR

Application Streamlit pour générer automatiquement des rapports de bilan carbone.

## Installation

```bash
# Cloner le repository
git clone <repo-url>
cd carbonReportSAUR

# Créer un environnement virtuel (recommandé)
python -m venv venv
source venv/bin/activate  # Sur Windows: venv\Scripts\activate

# Installer les dépendances
pip install -r requirements.txt
```

## Structure du projet

```
.
├── src/                      # Code source
│   ├── excel_loader.py       # Chargement et validation Excel
│   ├── tree.py               # Arborescence ORG/LOT/ENT
│   ├── calc_emissions.py     # Calculs émissions BRUT/NET
│   ├── calc_indicators.py    # Calculs indicateurs
│   ├── content_catalog.py    # Catalogue contenus rapport
│   ├── chart_generators.py   # Générateurs graphiques (matplotlib)
│   ├── table_generators.py   # Générateurs tableaux Word
│   ├── kpi_calculators.py    # Calculs KPI et équivalents
│   ├── word_renderer.py      # Moteur de rendu Word
│   └── word_blocks.py        # Gestion blocs répétables
├── templates/                # Templates Word
│   └── rapport_template.docx # ⚠️ PLACER VOTRE TEMPLATE ICI
├── assets/                   # Assets (logos, images)
│   ├── logo_org.png          # ⚠️ PLACER VOS ASSETS ICI
│   ├── digesteur_schema.png
│   └── icones/
├── output/                   # Rapports générés
├── app.py                    # Application Streamlit principale
├── requirements.txt          # Dépendances Python
└── README.md                 # Cette documentation
```

## Démarrage rapide

### 1. Préparer vos fichiers

**Template Word** : Placer votre template `rapport_template.docx` dans le dossier `templates/`

**Assets** : Placer vos images (logo, schémas) dans le dossier `assets/`
- `logo_org.png` : Logo de l'organisation
- `digesteur_schema.png` : Schéma du digesteur
- Autres icônes dans `assets/icones/`

### 2. Lancer l'application

```bash
streamlit run app.py
```

L'application s'ouvrira automatiquement dans votre navigateur à l'adresse `http://localhost:8501`

### 3. Workflow complet

1. **Upload Excel** : Charger votre fichier Excel au format standard
2. **Aperçu** : Vérifier l'arborescence et les résultats calculés
3. **Configuration** :
   - Renommer les nœuds (ORG, LOTs)
   - Configurer l'affichage des postes (masquer/exclure)
4. **Export/Import** : Sauvegarder/charger vos configurations
5. **Génération** : Générer le rapport Word final

## Format Excel attendu

### Onglets requis

#### ORG_TREE
Arborescence de l'organisation

Colonnes : `node_id`, `parent_id`, `node_type`, `node_name`, `activity`

**Règles** :
- `node_type` ∈ {ORG, LOT, ENT}
- `activity` ∈ {EU, AEP, NA}
- ORG = racine (parent_id vide)
- LOT = optionnel (entre ORG et ENT)
- ENT = feuilles avec activity EU ou AEP

#### EMISSIONS
Émissions niveau L1 (postes principaux)

Colonnes : `node_id`, `scope`, `poste_l1_code`, `tco2e`, `comment`

#### EMISSIONS_L2
Détails niveau 2 pour certains postes

Colonnes : `node_id`, `poste_l1_code`, `poste_l2`, `tco2e`

#### POSTES_REF
Référentiel des postes L1

Colonnes : `poste_l1_code`, `poste_l1_label`, `commentaire`

#### POSTES_L2_REF
Référentiel des postes L2

Colonnes : `poste_l1_code`, `poste_l2`, `poste_l2_order`

#### INDICATORS
Indicateurs par LOT/ORG et activité

Colonnes : `node_id`, `activity`, `indicator_code`, `value`, `unit`, `comment`

#### INDICATORS_REF
Référentiel des indicateurs

Colonnes : `indicator_code`, `indicator_label`, `default_unit`, `activity_scope`, `display_order`

#### EMISSIONS_EVITEES
Émissions évitées (affichées séparément)

Colonnes : `node_id`, `typologie`, `tco2e`

#### TEXTE_RAPPORT
Catalogue des textes et éléments par poste

Colonnes : `poste_l1_code`, `value`, `icone`, `CHART_KEY`, `IMAGE_KEY`, `TABLE_KEY`, `activity`, `DETAIL_SOUCE`

#### ICONE (optionnel)
Icônes associées aux postes

Colonnes : `poste_l1`

## Fonctionnalités

### Calculs automatiques

- **Agrégations** : ORG, LOT×ACTIVITÉ (EU/AEP)
- **Scopes** : Scope 1, 2, 3
- **Top postes** : Top 4 postes émetteurs (configurable)
- **KPI** : kgCO₂e/m³, équivalents (vols, personnes)

### Deux modes de calcul

- **BRUT** : Calcul direct depuis Excel sans modifications
- **NET** : Calcul après application des overrides utilisateur

### Gestion des postes

**Mode A** : Masquer mais inclure dans les totaux
- `show_in_report=False`, `include_in_totals=True`
- Le poste n'apparaît pas dans le rapport mais ses émissions comptent

**Mode B** : Masquer et exclure des totaux
- `show_in_report=False`, `include_in_totals=False`
- Le poste est complètement retiré (avec note de traçabilité)

### Graphiques supportés

- `TRAVAUX_BREAKDOWN` : Répartition des travaux
- `FILE_EAU_BREAKDOWN` : File eau STEP (N-N20, N2O, CH4)
- `EM_INDIRECTES_SPLIT` : Émissions indirectes
- `chart_emissions_scope_org` : Camembert scopes ORG
- `chart_contrib_lot` : Contribution des LOTs
- `chart_emissions_total_org` : Bar chart scopes ORG
- `chart_batonnet_inter_lot_top3` : Top 3 postes inter-LOT
- `chart_pie_scope_entity_activity` : Scopes LOT×ACT
- `chart_pie_postes_entity_activity` : Postes LOT×ACT

### Tableaux supportés

- `EM_INDIRECTES_TABLE` : Tableau émissions indirectes détaillé

### Images supportées

- `DIGESTEUR_SCHEMA` : Schéma du digesteur
- `ORG_LOGO` : Logo organisation

## Template Word

### Placeholders simples

Format : `{{PLACEHOLDER}}`

Exemples :
- `{{annee}}` : Année du bilan
- `{{ORG_NAME}}` : Nom organisation
- `{{TOTAL_EMISSIONS}}` : Total tCO₂e
- `{{TOTAL_EMISSIONS_S1}}`, `{{TOTAL_EMISSIONS_S2}}`, `{{TOTAL_EMISSIONS_S3}}`
- `{{kpi_EU_1}}`, `{{kpi_EU_2}}` : KPI activités
- `{{TOP_POSTE_1}}`, `{{TOP_POSTE_2}}`, `{{TOP_POSTE_3}}` : Top 3 postes

### Blocs répétables

Format : `[[START_XXX]]` ... `[[END_XXX]]`

**Blocs LOT** :
```
[[START_LOT]]
### {{LOT_NAME}}
...
[[END_LOT]]
```

**Blocs ACTIVITÉ** :
```
[[START_ACTIVITY]]
#### {{ENT_ACTIVITY}}
...
[[END_ACTIVITY]]
```

**Blocs POSTE détaillé** :
```
[[START_POST]]
### {{POST_TITLE}}
{{POST_TEXT}}
{{POST_TABLE_1}}
{{POST_CHART_1}}
{{POST_IMAGE_1}}
[[END_POST]]
```

### Placeholders images/graphiques

Les placeholders d'images doivent être **seuls sur leur ligne** :

```
{{chart_emissions_scope_org}}

{{ORG_LOGO}}

{{POST_CHART_1}}
```

## Cas d'usage supportés

- ✅ Organisation sans LOT (ORG → ENT direct)
- ✅ Organisation avec LOTs (ORG → LOT → ENT)
- ✅ Activité EU uniquement
- ✅ Activité AEP uniquement
- ✅ Mix EU + AEP
- ✅ Renommage des nœuds
- ✅ Exclusion de postes (Modes A et B)
- ✅ Export/Import configuration

## Troubleshooting

### Erreur "Template non trouvé"
→ Vérifier que `rapport_template.docx` est bien dans le dossier `templates/`

### Erreur de validation Excel
→ Vérifier que tous les onglets requis sont présents avec les bonnes colonnes

### Graphique ne s'affiche pas
→ Vérifier que le placeholder est seul sur sa ligne dans le template Word

### Émissions à 0
→ Vérifier les données dans l'onglet EMISSIONS et que les node_id correspondent

## Support

Pour toute question ou problème :
- Consulter la documentation technique dans chaque module Python
- Vérifier les logs d'erreur dans Streamlit
- Contacter le support technique

## Licence

© SAUR - Tous droits réservés
