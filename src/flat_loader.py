"""
Module de chargement d'un fichier Excel au format simplifié (2 onglets : DATA + TEXTE_RAPPORT).
Transforme le tableau plat en 9 DataFrames compatibles avec le format standard de l'application.
"""

import re
import unicodedata
import pandas as pd
from typing import Dict, List, Optional, Tuple, Set
from pathlib import Path

from src.excel_loader import ExcelValidationError


# ---------------------------------------------------------------------------
# Colonnes attendues dans l'onglet DATA
# ---------------------------------------------------------------------------
FLAT_REQUIRED_COLUMNS = [
    'Organisation', 'Lot', 'Entité', 'Année',
    'Catégorie', 'Poste', 'Quantité', 'Unité', 'Emissions_kgCO2'
]

# ---------------------------------------------------------------------------
# Catégories spéciales (non-émission)
# ---------------------------------------------------------------------------
CATEGORIE_INDICATEUR = 'indicateur'
CATEGORIE_EMISSIONS_EVITEES = 'émissions évitées'

# ---------------------------------------------------------------------------
# Mapping fixe : catégorie → scope
# ---------------------------------------------------------------------------
SCOPE_BY_CATEGORY = {
    'Travaux': 3,
    'Réactifs': 3,
    "Imports d'eau": 3,
    'Achats de services': 3,
    'File eau': 3,
    'Compostage des boues': 3,
    'Rejets dans le milieu': 3,
    'Energie': 2,
    'Immobilisations': 3,
    'Digesteur': 3,
    'Utilisation': 3,
    'Déplacements domicile-travail': 3,
    'Déplacements pro': 3,
    'Déchets': 3,
    'Fret sortant': 1,
}

# ---------------------------------------------------------------------------
# Mapping fixe : catégorie pipeline → label rapport
# ---------------------------------------------------------------------------
L1_LABEL_MAP = {
    'Travaux': 'Intrants - Travaux',
    'Réactifs': 'Intrants - Réactifs',
    "Imports d'eau": "Intrants - Imports d'eau",
    'Achats de services': 'Intrants - Autres intrants',
    'File eau': 'File eau des STEP',
    'Compostage des boues': 'Emissions indirectes liées au compostage',
    'Rejets dans le milieu': 'Emissions indirectes liées aux rejets',
    'Energie': 'Electricité',
    'Immobilisations': 'Immobilisations',
    'Digesteur': 'Digesteur',
    'Utilisation': 'Epandage des boues',
    'Déplacements domicile-travail': 'Déplacements domicile-travail',
    'Déplacements pro': 'Parc automobile',
    'Déchets': 'Déchets',
    'Fret sortant': 'Fret sortant',
}

# ---------------------------------------------------------------------------
# Mapping fixe : nom indicateur pipeline → code interne (pour KPI)
# ---------------------------------------------------------------------------
INDICATOR_CODE_MAP = {
    "m3 d'eau distribués": 'VOL_EAU_DISTRIB',
    "m3 d'eau distribues": 'VOL_EAU_DISTRIB',
    "m3 d'eau assainis": 'VOL_EAU_EPURE',
    "m3 d'eau épurés": 'VOL_EAU_EPURE',
    "m3 d'eau epures": 'VOL_EAU_EPURE',
    "nombre de branchements": 'NB_BRANCHEMENTS',
    "nombre d'habitants": 'NB_HAB_DESSERVIS',
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _slugify(text: str) -> str:
    """Normalise un texte en slug : supprime accents, remplace espaces/tirets par _, majuscules."""
    normalized = unicodedata.normalize('NFD', text)
    ascii_text = normalized.encode('ascii', 'ignore').decode('ascii')
    slug = re.sub(r'[^a-zA-Z0-9]+', '_', ascii_text).strip('_').upper()
    return slug


def _make_poste_code(category_name: str) -> str:
    """Génère un poste_l1_code stable depuis un nom de catégorie.

    Ex: 'Travaux' -> 'P_TRAVAUX', 'File eau' -> 'P_FILE_EAU'
    """
    return f'P_{_slugify(category_name)}'


def _make_node_id_ent(lot_name: str, activity: str) -> str:
    return f'ENT_{_slugify(lot_name)}_{activity}'


def _make_node_id_lot(lot_name: str) -> str:
    return f'LOT_{_slugify(lot_name)}'


def _make_indicator_code(indicator_name: str) -> str:
    """Convertit un nom d'indicateur pipeline en code interne."""
    normalized = indicator_name.strip().lower()
    if normalized in INDICATOR_CODE_MAP:
        return INDICATOR_CODE_MAP[normalized]
    return f'IND_{_slugify(indicator_name)}'


# ---------------------------------------------------------------------------
# FlatLoader
# ---------------------------------------------------------------------------

class FlatLoader:
    """
    Charge un fichier Excel au format simplifié (2 onglets : DATA + TEXTE_RAPPORT)
    et produit le même Dict[str, pd.DataFrame] que ExcelLoader.load().
    """

    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.data: Dict[str, pd.DataFrame] = {}
        self.validation_errors: List[str] = []
        self.validation_warnings: List[str] = []

    def load(self) -> Dict[str, pd.DataFrame]:
        """Charge et transforme le fichier Excel simplifié en 9 DataFrames standard."""
        if not self.file_path.exists():
            raise ExcelValidationError(f"Fichier non trouvé : {self.file_path}")

        try:
            excel_file = pd.ExcelFile(self.file_path)

            # 1. Valider les onglets
            self._validate_sheets(excel_file.sheet_names)
            if self.validation_errors:
                raise ExcelValidationError(
                    "Erreurs de validation :\n" + "\n".join(self.validation_errors)
                )

            # 2. Lire et valider DATA
            data_df = pd.read_excel(excel_file, sheet_name='DATA')
            self._validate_flat_columns(data_df)
            if self.validation_errors:
                raise ExcelValidationError(
                    "Erreurs de validation :\n" + "\n".join(self.validation_errors)
                )

            # 3. Nettoyer les données
            data_df = self._clean_data(data_df)

            # 4. Classifier les lignes
            emission_rows, indicator_rows, evitees_rows = self._classify_rows(data_df)

            # 5. Construire les 9 DataFrames
            self.data['ORG_TREE'] = self._build_org_tree(emission_rows, indicator_rows)
            self.data['EMISSIONS'] = self._build_emissions(emission_rows)
            self.data['EMISSIONS_L2'] = self._build_emissions_l2(emission_rows)
            self.data['POSTES_REF'] = self._build_postes_ref(emission_rows)
            self.data['POSTES_L2_REF'] = self._build_postes_l2_ref(emission_rows)
            self.data['INDICATORS'] = self._build_indicators(indicator_rows)
            self.data['INDICATORS_REF'] = self._build_indicators_ref(indicator_rows)
            self.data['EMISSIONS_EVITEES'] = self._build_emissions_evitees(evitees_rows)

            # 6. TEXTE_RAPPORT : lecture directe + mapping codes
            emission_categories = set(emission_rows['Catégorie'].dropna().unique())
            self.data['TEXTE_RAPPORT'] = self._build_texte_rapport(
                excel_file, emission_categories
            )

            # 7. Validation finale des schemas
            self._validate_output_schemas()

            if self.validation_errors:
                raise ExcelValidationError(
                    "Erreurs de validation :\n" + "\n".join(self.validation_errors)
                )

            return self.data

        except Exception as e:
            if isinstance(e, ExcelValidationError):
                raise
            raise ExcelValidationError(f"Erreur lors du chargement du fichier : {str(e)}")

    def get_validation_report(self) -> Tuple[List[str], List[str]]:
        """Retourne les erreurs et warnings de validation."""
        return self.validation_errors, self.validation_warnings

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def _validate_sheets(self, sheet_names: List[str]) -> None:
        if 'DATA' not in sheet_names:
            self.validation_errors.append("Onglet 'DATA' manquant")
        if 'TEXTE_RAPPORT' not in sheet_names:
            self.validation_errors.append("Onglet 'TEXTE_RAPPORT' manquant")

    def _validate_flat_columns(self, df: pd.DataFrame) -> None:
        missing = set(FLAT_REQUIRED_COLUMNS) - set(df.columns)
        if missing:
            self.validation_errors.append(
                f"Onglet 'DATA' : colonnes manquantes : {', '.join(sorted(missing))}"
            )

    def _validate_output_schemas(self) -> None:
        """Vérifie que chaque DataFrame de sortie a les colonnes attendues."""
        from src.excel_loader import REQUIRED_SHEETS
        for sheet_name, required_cols in REQUIRED_SHEETS.items():
            if sheet_name in self.data:
                df = self.data[sheet_name]
                missing = set(required_cols) - set(df.columns)
                if missing:
                    self.validation_errors.append(
                        f"Schema '{sheet_name}' : colonnes manquantes après transformation : "
                        f"{', '.join(sorted(missing))}"
                    )

    # ------------------------------------------------------------------
    # Nettoyage
    # ------------------------------------------------------------------

    def _clean_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Nettoyage basique du DataFrame DATA."""
        df = df.copy()
        # Supprimer les lignes entièrement vides
        df = df.dropna(how='all')
        # Strip les colonnes texte
        for col in ['Organisation', 'Lot', 'Entité', 'Catégorie', 'Poste', 'Unité']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan': None, 'None': None, '': None})
        # Convertir les colonnes numériques
        for col in ['Quantité', 'Emissions_kgCO2']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
        return df

    # ------------------------------------------------------------------
    # Classification des lignes
    # ------------------------------------------------------------------

    def _classify_rows(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Sépare DATA en lignes émission, indicateur, émissions évitées."""
        cat_lower = df['Catégorie'].str.lower().str.strip()

        indicator_mask = cat_lower == CATEGORIE_INDICATEUR
        evitees_mask = cat_lower.str.contains(
            'missions évitées|missions evitees|emissions evitees',
            regex=True, na=False
        )
        emission_mask = ~indicator_mask & ~evitees_mask

        emission_rows = df[emission_mask].copy()
        indicator_rows = df[indicator_mask].copy()
        evitees_rows = df[evitees_mask].copy()

        # Warnings pour catégories d'émission inconnues
        if len(emission_rows) > 0:
            unknown_cats = set(emission_rows['Catégorie'].dropna().unique()) - set(SCOPE_BY_CATEGORY.keys())
            for cat in unknown_cats:
                self.validation_warnings.append(
                    f"Catégorie inconnue '{cat}' : scope par défaut = 3, label = nom brut"
                )

        return emission_rows, indicator_rows, evitees_rows

    # ------------------------------------------------------------------
    # Builders
    # ------------------------------------------------------------------

    def _build_org_tree(self, emission_rows: pd.DataFrame,
                        indicator_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit ORG_TREE depuis les colonnes Organisation/Lot/Entité."""
        # Combiner les deux sources pour trouver toutes les combinaisons
        all_rows = pd.concat([emission_rows, indicator_rows], ignore_index=True)
        if len(all_rows) == 0:
            return pd.DataFrame(columns=['node_id', 'parent_id', 'node_type', 'node_name', 'activity'])

        org_name = all_rows['Organisation'].dropna().iloc[0]

        rows = []
        # Nœud ORG racine
        rows.append({
            'node_id': 'ORG_1',
            'parent_id': None,
            'node_type': 'ORG',
            'node_name': org_name,
            'activity': 'NA'
        })

        # Trouver les couples (Lot, Entité) uniques
        lot_ent_pairs = all_rows[['Lot', 'Entité']].dropna().drop_duplicates()
        unique_lots = lot_ent_pairs['Lot'].unique()

        for lot_name in sorted(unique_lots):
            lot_id = _make_node_id_lot(lot_name)
            rows.append({
                'node_id': lot_id,
                'parent_id': 'ORG_1',
                'node_type': 'LOT',
                'node_name': lot_name,
                'activity': 'NA'
            })

            activities = lot_ent_pairs[lot_ent_pairs['Lot'] == lot_name]['Entité'].unique()
            for activity in sorted(activities):
                ent_id = _make_node_id_ent(lot_name, activity)
                rows.append({
                    'node_id': ent_id,
                    'parent_id': lot_id,
                    'node_type': 'ENT',
                    'node_name': f'{lot_name} - {activity}',
                    'activity': activity
                })

        return pd.DataFrame(rows)

    def _build_emissions(self, emission_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit EMISSIONS : agrégé au niveau L1 (node_id, scope, poste_l1_code)."""
        if len(emission_rows) == 0:
            return pd.DataFrame(columns=['node_id', 'scope', 'poste_l1_code', 'tco2e', 'comment'])

        df = emission_rows.copy()
        df['node_id'] = df.apply(
            lambda r: _make_node_id_ent(r['Lot'], r['Entité']), axis=1
        )
        df['scope'] = df['Catégorie'].map(SCOPE_BY_CATEGORY).fillna(3).astype(int)
        df['poste_l1_code'] = df['Catégorie'].apply(_make_poste_code)
        df['tco2e'] = df['Emissions_kgCO2'] / 1000.0

        # Agréger les sous-postes L2 au niveau L1
        aggregated = df.groupby(
            ['node_id', 'scope', 'poste_l1_code'], as_index=False
        ).agg({'tco2e': 'sum'})
        aggregated['comment'] = ''

        return aggregated[['node_id', 'scope', 'poste_l1_code', 'tco2e', 'comment']]

    def _build_emissions_l2(self, emission_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit EMISSIONS_L2 : chaque ligne est un détail L2."""
        if len(emission_rows) == 0:
            return pd.DataFrame(columns=['node_id', 'poste_l1_code', 'poste_l2', 'tco2e'])

        df = emission_rows.copy()
        df['node_id'] = df.apply(
            lambda r: _make_node_id_ent(r['Lot'], r['Entité']), axis=1
        )
        df['poste_l1_code'] = df['Catégorie'].apply(_make_poste_code)
        df['poste_l2'] = df['Poste']
        df['tco2e'] = df['Emissions_kgCO2'] / 1000.0

        # Agréger les doublons éventuels
        aggregated = df.groupby(
            ['node_id', 'poste_l1_code', 'poste_l2'], as_index=False
        ).agg({'tco2e': 'sum'})

        return aggregated[['node_id', 'poste_l1_code', 'poste_l2', 'tco2e']]

    def _build_postes_ref(self, emission_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit POSTES_REF : un poste L1 par catégorie unique."""
        if len(emission_rows) == 0:
            return pd.DataFrame(columns=['poste_l1_code', 'poste_l1_label', 'commentaire'])

        categories = emission_rows['Catégorie'].dropna().unique()
        rows = []
        for cat in sorted(categories):
            rows.append({
                'poste_l1_code': _make_poste_code(cat),
                'poste_l1_label': L1_LABEL_MAP.get(cat, cat),
                'commentaire': ''
            })

        df = pd.DataFrame(rows).drop_duplicates(subset='poste_l1_code')
        return df

    def _build_postes_l2_ref(self, emission_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit POSTES_L2_REF : ordre séquentiel par catégorie."""
        if len(emission_rows) == 0:
            return pd.DataFrame(columns=['poste_l1_code', 'poste_l2', 'poste_l2_order'])

        rows = []
        for cat, group in emission_rows.groupby('Catégorie'):
            subcats = group['Poste'].dropna().unique()
            for order, subcat in enumerate(sorted(subcats), 1):
                rows.append({
                    'poste_l1_code': _make_poste_code(cat),
                    'poste_l2': subcat,
                    'poste_l2_order': order
                })

        df = pd.DataFrame(rows).drop_duplicates(subset=['poste_l1_code', 'poste_l2'])
        return df

    def _build_indicators(self, indicator_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit INDICATORS depuis les lignes Indicateur."""
        if len(indicator_rows) == 0:
            return pd.DataFrame(columns=['node_id', 'activity', 'indicator_code',
                                         'value', 'unit', 'comment'])

        rows = []
        for _, row in indicator_rows.iterrows():
            indicator_name = row['Poste']
            if pd.isna(indicator_name) or indicator_name is None:
                continue

            code = _make_indicator_code(indicator_name)
            node_id = _make_node_id_ent(row['Lot'], row['Entité'])

            rows.append({
                'node_id': node_id,
                'activity': row['Entité'],
                'indicator_code': code,
                'value': row['Quantité'],
                'unit': row['Unité'] if pd.notna(row.get('Unité')) else '',
                'comment': ''
            })

        return pd.DataFrame(rows)

    def _build_indicators_ref(self, indicator_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit INDICATORS_REF : une entrée par indicateur unique."""
        if len(indicator_rows) == 0:
            return pd.DataFrame(columns=['indicator_code', 'indicator_label',
                                         'default_unit', 'activity_scope', 'display_order'])

        rows = []
        seen: Set[str] = set()
        order = 0

        for _, row in indicator_rows.iterrows():
            indicator_name = row['Poste']
            if pd.isna(indicator_name) or indicator_name is None:
                continue

            code = _make_indicator_code(indicator_name)
            if code in seen:
                continue
            seen.add(code)
            order += 1

            rows.append({
                'indicator_code': code,
                'indicator_label': indicator_name,
                'default_unit': row['Unité'] if pd.notna(row.get('Unité')) else '',
                'activity_scope': 'BOTH',
                'display_order': order
            })

        return pd.DataFrame(rows)

    def _build_emissions_evitees(self, evitees_rows: pd.DataFrame) -> pd.DataFrame:
        """Construit EMISSIONS_EVITEES depuis les lignes Émissions évitées."""
        if len(evitees_rows) == 0:
            return pd.DataFrame(columns=['node_id', 'typologie', 'tco2e'])

        rows = []
        for _, row in evitees_rows.iterrows():
            node_id = _make_node_id_ent(row['Lot'], row['Entité'])
            rows.append({
                'node_id': node_id,
                'typologie': row['Poste'] if pd.notna(row.get('Poste')) else '',
                'tco2e': row['Emissions_kgCO2'] / 1000.0
            })

        return pd.DataFrame(rows)

    def _build_texte_rapport(self, excel_file: pd.ExcelFile,
                             emission_categories: Set[str]) -> pd.DataFrame:
        """Lit TEXTE_RAPPORT et transforme poste_l1_code si nécessaire."""
        df = pd.read_excel(excel_file, sheet_name='TEXTE_RAPPORT')

        # Si les poste_l1_code dans TEXTE_RAPPORT correspondent aux noms de catégorie
        # bruts du pipeline, les convertir en codes générés
        category_lower_set = {cat.strip().lower() for cat in emission_categories if cat}

        def maybe_transform_code(code):
            if pd.isna(code):
                return code
            code_str = str(code).strip()
            if code_str.lower() in category_lower_set:
                # Retrouver le nom original avec la bonne casse
                for cat in emission_categories:
                    if cat.strip().lower() == code_str.lower():
                        return _make_poste_code(cat)
                return _make_poste_code(code_str)
            return code_str

        df['poste_l1_code'] = df['poste_l1_code'].apply(maybe_transform_code)
        return df
