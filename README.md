# Carbon Report Generator - SAUR

Application Streamlit pour la generation automatique de rapports de bilan carbone au format Word, a partir de donnees Excel.

## Installation

```bash
git clone <repo-url>
cd carbonReportSAUR

python -m venv venv
source venv/bin/activate  # Windows : venv\Scripts\activate

pip install -r requirements.txt
```

Dependances : `streamlit`, `pandas`, `openpyxl`, `python-docx`, `matplotlib`, `Pillow`

## Demarrage rapide

### Lancer l'application Streamlit

```bash
streamlit run app_v1.py
```

L'interface s'ouvre dans le navigateur (`http://localhost:8501`). Workflow :

1. Choisir l'annee du bilan
2. Uploader le template Word (`.docx`)
3. Uploader le fichier Excel (`.xlsx`)
4. Cliquer sur "Generer le rapport"
5. Telecharger le rapport genere

### Test rapide sans Streamlit

```bash
python tests/test_rapport.py <fichier_excel.xlsx> [annee]
```

Le rapport est genere dans `tests/output/`. Ce mode est utile pour iterer rapidement sur le template ou deboguer sans interface graphique.

### Verifier l'installation

```bash
python check_setup.py
```

---

## Structure du projet

```
carbonReportSAUR/
|-- app_v1.py                  # Application Streamlit (point d'entree)
|-- src/
|   |-- flat_loader.py         # Chargement/validation Excel (onglet DATA)
|   |-- tree.py                # Arborescence ORG/LOT/ENT
|   |-- calc_emissions.py      # Calculs emissions (BRUT/NET, scopes, top postes)
|   |-- calc_indicators.py     # Calculs indicateurs par LOT x ACTIVITE
|   |-- content_catalog.py     # Catalogue contenus du rapport (TEXTE_RAPPORT)
|   |-- chart_generators.py    # Graphiques matplotlib (camemberts, barres, tableaux)
|   |-- table_generators.py    # Tableaux Word
|   |-- kpi_calculators.py     # KPI, equivalents, formatage nombres
|   |-- word_renderer.py       # Moteur de rendu Word (~1600 lignes)
|   |-- word_blocks.py         # Gestion blocs repetables (duplication/remplacement)
|   +-- streamlit_charts_page.py
|-- templates/                 # Template Word
|-- assets/                    # Images (logos, schemas, icones, polices)
|-- tests/
|   |-- test_rapport.py        # Test E2E de generation
|   |-- unit/                  # Tests unitaires
|   +-- output/                # Rapports de test generes
|-- output/                    # Rapports generes via Streamlit
+-- requirements.txt
```

---

## Format Excel

L'application attend un fichier Excel avec un onglet `DATA` principal, plus des onglets optionnels (`TEXTE_RAPPORT`, `BEGES`).

### Onglet DATA (requis)

Tableau plat contenant toutes les donnees du bilan.

Colonnes requises : `Organisation`, `Lot`, `Entite`, `Annee`, `Categorie`, `Poste`, `Quantite`, `Unite`, `Emissions_kgCO2`

Le FlatLoader transforme ce tableau plat en DataFrames internes (arborescence, emissions, indicateurs, etc.).

**Regles sur l'arborescence** :
- **ORG** : racine unique, deduite de la colonne `Organisation`
- **LOT** : niveau intermediaire optionnel, deduit de la colonne `Lot`
- **ENT** : feuilles, deduites de la colonne `Entite`, avec activite EU (eaux usees) ou AEP (eau potable)

**Categories speciales** :
- `indicateur` : les lignes de cette categorie sont traitees comme des indicateurs (volumes, etc.)
- `emissions evitees` : les lignes de cette categorie alimentent la section emissions evitees du rapport

### Onglet TEXTE_RAPPORT (requis)

Catalogue des textes et elements par poste.

Colonnes : `poste_l1_code`, `value`, `icone`, `CHART_KEY`, `IMAGE_KEY`, `TABLE_KEY`, `activity`, `DETAIL_SOUCE`

### Onglet BEGES (optionnel)

Tableau BEGES reglementaire, rendu en annexe du rapport. Charge automatiquement si present.

---

## Template Word

Le template utilise des placeholders et des blocs repetables que le moteur de rendu remplace automatiquement.

### Placeholders simples

Format : `{{PLACEHOLDER}}`

| Placeholder | Description |
|-------------|-------------|
| `{{annee}}` | Annee du bilan |
| `{{ORG_NAME}}` | Nom de l'organisation |
| `{{TOTAL_EMISSIONS}}` | Total tCO2e |
| `{{TOTAL_EMISSIONS_S1}}` | Total Scope 1 |
| `{{TOTAL_EMISSIONS_S2}}` | Total Scope 2 |
| `{{TOTAL_EMISSIONS_S3}}` | Total Scope 3 |
| `{{kpi_EU_1}}`, `{{kpi_EU_2}}` | KPI activites |
| `{{TOP_POSTE_1}}` a `{{TOP_POSTE_3}}` | Top 3 postes emetteurs |
| `{{CHAUFFAGE_TOTAL}}` | Total chauffage tCO2e |
| `{{CHAUFFAGE_PERCENTAGE}}` | % chauffage par rapport au total AEP |
| `{{EVITEES_TOTAL}}` | Total emissions evitees |

### Blocs repetables

Format : `[[START_XXX]]` ... `[[END_XXX]]`

Les blocs sont dupliques automatiquement selon les donnees. Ils s'imbriquent de l'exterieur vers l'interieur :

```
[[START_LOT]]
  Nom du LOT : {{LOT_NAME}}

  [[START_ACTIVITY]]
    Activite : {{ENT_ACTIVITY}}

    [[START_POST]]
      {{POST_TITLE}}
      {{POST_TEXT}}
      {{POST_CHART_1}}
      {{POST_TABLE_1}}
      {{POST_IMAGE_1}}
    [[END_POST]]

    [[START_OTHER_POST]]
      {{OTHER_POST_TITLE}} - {{OTHER_POST_TCO2E}} tCO2e
      {{OTHER_POST_TEXT}}
    [[END_OTHER_POST]]

    [[START_EVITEES]]
      {{EVITEES_TOTAL}}
      {{EVITEES_TABLE}}
    [[END_EVITEES]]

  [[END_ACTIVITY]]
[[END_LOT]]

[[START_CHAUFFAGE_INCLUS]]
  {{CHAUFFAGE_TOTAL}}
  {{CHAUFFAGE_PERCENTAGE}}
  {{PIE_CHART_CHAUFFAGE_INCLU}}
[[END_CHAUFFAGE_INCLUS]]
```

**Blocs conditionnels** :
- `[[START_CHAUFFAGE_INCLUS]]` : affiche uniquement si des entites AEP existent
- `[[START_EVITEES]]` : affiche uniquement si la feuille EMISSIONS_EVITEES contient des donnees

**Cas sans LOT** : les blocs `[[START_ACTIVITY]]` sont traites directement au niveau ORG.

### Placeholders images/graphiques

Les placeholders d'images doivent etre **seuls sur leur ligne** dans le template :

```
{{chart_emissions_scope_org}}

{{ORG_LOGO}}

{{PIE_CHART_CHAUFFAGE_INCLU}}
```

Un placeholder sur une ligne partagee avec du texte sera traite comme un remplacement texte, pas une image.

---

## Graphiques et tableaux

### Graphiques supportes (CHART_KEY)

| Cle | Description |
|-----|-------------|
| `TRAVAUX_BREAKDOWN` | Repartition des travaux (barres horizontales) |
| `FILE_EAU_BREAKDOWN` | File eau STEP - N-N2O, N2O, CH4 (camembert) |
| `EM_INDIRECTES_SPLIT` | Emissions indirectes (camembert) |
| `chart_emissions_scope_org` | Repartition par scope - ORG |
| `chart_contrib_lot` | Contribution des LOTs |
| `chart_emissions_total_org` | Contribution des postes - ORG |
| `chart_emissions_elec_org` | Electricite par activite |
| `chart_batonnet_inter_lot_top3` | Top 3 postes inter-LOT (barres groupees) |
| `chart_pie_scope_entity_activity` | Scopes par LOT x activite |
| `chart_pie_postes_entity_activity` | Postes par LOT x activite |
| `BEGES_TABLE` | Tableau BEGES reglementaire (image) |

### Tableaux supportes (TABLE_KEY)

| Cle | Description |
|-----|-------------|
| `EM_INDIRECTES_TABLE` | Tableau emissions indirectes detaille |

### Images supportees (IMAGE_KEY)

| Cle | Description |
|-----|-------------|
| `DIGESTEUR_SCHEMA` | Schema du digesteur |
| `ORG_LOGO` | Logo de l'organisation |

Les graphiques sont generes via matplotlib avec la police Poppins (chargee depuis `assets/police/`) et une palette verte a 6 couleurs.

---

## Architecture technique

### Flux de donnees

```
Excel (.xlsx)
    |
    v
FlatLoader                  -- chargement et validation
    |
    v
OrganizationTree            -- hierarchie ORG -> LOT -> ENT
    |
    v
EmissionCalculator          -- aggregation scopes, top postes
IndicatorCalculator         -- indicateurs par LOT x activite
ContentCatalog              -- textes, charts, images par poste
KPICalculator               -- kgCO2e/m3, equivalents
    |
    v
WordRenderer                -- rendu Word avec blocs repetables
    |
    v
Rapport .docx
```

### Modes de calcul

- **BRUT** : calcul direct depuis l'Excel sans modification. Utilise par le script de test et par V1.
- **NET** : calcul avec overrides utilisateur (exclusion/masquage de postes). Structure presente dans le code mais non exposee dans l'interface V1.

### Rendu Word : ordre des operations

1. Remplacement des placeholders simples (texte)
2. Insertion du logo ORG
3. Duplication des blocs LOT
4. Pour chaque LOT : duplication des blocs ACTIVITY, POST, OTHER_POST
5. Traitement de la section chauffage inclus
6. Traitement de la section emissions evitees
7. Nettoyage de tous les marqueurs `[[START/END_XXX]]`
8. Insertion des graphiques ORG
9. Traitement de l'annexe BEGES
10. Nettoyage des placeholders restants

### Degradation gracieuse

Le moteur de rendu ne leve jamais d'exception pour du contenu manquant :
- Image absente : placeholder ignore puis nettoye
- Donnees L2 absentes : graphique/tableau non genere
- Bloc sans donnees : supprime automatiquement
- Donnees partielles : rapport toujours valide

### Formatage des nombres

Convention francaise avec espace insecable comme separateur de milliers (ex: "19 445"). Voir `KPICalculator.format_number()`.

---

## Tests

### Test E2E

```bash
python tests/test_rapport.py <fichier_excel.xlsx> [annee]
```

Execute toutes les etapes : chargement Excel, arborescence, calculs emissions/indicateurs/KPI, generation Word. Le rapport est enregistre dans `tests/output/`.

Verifications automatiques :
- Aucun marqueur `[[START_` ou `[[END_` restant
- Tous les noms de LOT presents dans le rapport
- Nombre correct d'occurrences par activite

### Tests unitaires

```bash
python tests/unit/test_word_renderer_blocks.py
```

Tests rapides qui verifient l'existence et l'accessibilite des methodes de WordRenderer sans generer de rapport.

---

## Troubleshooting

**Template non trouve** : verifier que le fichier `.docx` est bien dans `templates/` ou uploader le directement via l'interface.

**Erreur de validation Excel** : verifier que tous les onglets requis sont presents avec les bonnes colonnes.

**Graphique qui ne s'affiche pas** : verifier que le placeholder est seul sur sa ligne dans le template Word.

**Emissions a 0** : verifier les donnees dans l'onglet EMISSIONS et la correspondance des `node_id` avec ORG_TREE.

---

## Personnalisation

### Modifier le nombre de top postes

Le nombre de top postes (defaut : 4) est configure via le parametre `top_n` dans `app_v1.py` et `test_rapport.py`.

### Ajouter un nouveau type de graphique

1. Ajouter la methode dans `ChartGenerator` (`src/chart_generators.py`)
2. Ajouter la cle dans `SUPPORTED_CHART_KEYS` (`src/word_renderer.py`)
3. Mapper dans l'Excel via la colonne `CHART_KEY` de l'onglet `TEXTE_RAPPORT`
4. Ajouter le placeholder dans le template : `{{POST_CHART_1}}` ou `{{chart_xxx}}`

### Ajouter des assets

Placer les fichiers dans le dossier `assets/` :
- `assets/logo_org.png` ou `assets/ORG_LOGO.png` : logo de l'organisation
- `assets/digesteur_schema.png` : schema du digesteur
- `assets/icones/` : icones des postes
- `assets/police/` : fichiers de police Poppins

---

(c) SAUR - Tous droits reserves
