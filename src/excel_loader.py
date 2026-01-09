"""
Module de chargement et validation des fichiers Excel.
Vérifie la présence des onglets et colonnes requis selon le format standard.
"""

import pandas as pd
from typing import Dict, List, Optional, Tuple
from pathlib import Path


class ExcelValidationError(Exception):
    """Exception levée lors d'erreurs de validation du fichier Excel."""
    pass


# Définition des onglets requis et leurs colonnes exactes
REQUIRED_SHEETS = {
    'ORG_TREE': ['node_id', 'parent_id', 'node_type', 'node_name', 'activity'],
    'EMISSIONS': ['node_id', 'scope', 'poste_l1_code', 'tco2e', 'comment'],
    'EMISSIONS_L2': ['node_id', 'poste_l1_code', 'poste_l2', 'tco2e'],
    'POSTES_L2_REF': ['poste_l1_code', 'poste_l2', 'poste_l2_order'],
    'POSTES_REF': ['poste_l1_code', 'poste_l1_label', 'commentaire'],
    'INDICATORS': ['node_id', 'activity', 'indicator_code', 'value', 'unit', 'comment'],
    'INDICATORS_REF': ['indicator_code', 'indicator_label', 'default_unit', 'activity_scope', 'display_order'],
    'EMISSIONS_EVITEES': ['node_id', 'typologie', 'tco2e'],
    'TEXTE_RAPPORT': ['poste_l1_code', 'value', 'icone', 'CHART_KEY', 'IMAGE_KEY', 'TABLE_KEY', 'activity', 'DETAIL_SOUCE'],
}

OPTIONAL_SHEETS = {
    'ICONE': ['poste_l1'],
}


class ExcelLoader:
    """
    Charge et valide un fichier Excel au format standard pour le bilan carbone.
    """

    def __init__(self, file_path: str):
        """
        Initialise le loader avec le chemin du fichier Excel.

        Args:
            file_path: Chemin vers le fichier Excel à charger
        """
        self.file_path = Path(file_path)
        self.data: Dict[str, pd.DataFrame] = {}
        self.validation_errors: List[str] = []
        self.validation_warnings: List[str] = []

    def load(self) -> Dict[str, pd.DataFrame]:
        """
        Charge le fichier Excel et valide sa structure.

        Returns:
            Dictionnaire {nom_onglet: DataFrame}

        Raises:
            ExcelValidationError: Si le fichier n'est pas valide
        """
        if not self.file_path.exists():
            raise ExcelValidationError(f"Fichier non trouvé : {self.file_path}")

        try:
            # Charger tous les onglets
            excel_file = pd.ExcelFile(self.file_path)
            available_sheets = excel_file.sheet_names

            # Vérifier les onglets requis
            self._validate_required_sheets(available_sheets)

            # Charger les onglets requis
            for sheet_name, required_columns in REQUIRED_SHEETS.items():
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                self._validate_columns(sheet_name, df, required_columns)
                self.data[sheet_name] = df

            # Charger les onglets optionnels s'ils existent
            for sheet_name, required_columns in OPTIONAL_SHEETS.items():
                if sheet_name in available_sheets:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    self._validate_columns(sheet_name, df, required_columns, optional=True)
                    self.data[sheet_name] = df
                else:
                    self.validation_warnings.append(f"Onglet optionnel '{sheet_name}' absent")

            # Validations métier
            self._validate_business_rules()

            if self.validation_errors:
                error_msg = "\n".join(self.validation_errors)
                raise ExcelValidationError(f"Erreurs de validation :\n{error_msg}")

            return self.data

        except Exception as e:
            if isinstance(e, ExcelValidationError):
                raise
            raise ExcelValidationError(f"Erreur lors du chargement du fichier : {str(e)}")

    def _validate_required_sheets(self, available_sheets: List[str]) -> None:
        """Vérifie que tous les onglets requis sont présents."""
        missing_sheets = set(REQUIRED_SHEETS.keys()) - set(available_sheets)
        if missing_sheets:
            self.validation_errors.append(
                f"Onglets manquants : {', '.join(sorted(missing_sheets))}"
            )

    def _validate_columns(self, sheet_name: str, df: pd.DataFrame,
                         required_columns: List[str], optional: bool = False) -> None:
        """
        Vérifie que toutes les colonnes requises sont présentes dans l'onglet.

        Args:
            sheet_name: Nom de l'onglet
            df: DataFrame à valider
            required_columns: Liste des colonnes requises
            optional: Si True, ne génère que des warnings
        """
        df_columns = set(df.columns)
        missing_columns = set(required_columns) - df_columns

        if missing_columns:
            msg = f"Onglet '{sheet_name}' : colonnes manquantes : {', '.join(sorted(missing_columns))}"
            if optional:
                self.validation_warnings.append(msg)
            else:
                self.validation_errors.append(msg)

    def _validate_business_rules(self) -> None:
        """Valide les règles métier sur les données chargées."""

        # Validation ORG_TREE
        if 'ORG_TREE' in self.data:
            org_tree = self.data['ORG_TREE']

            # Vérifier les types de nœuds valides
            valid_node_types = {'ORG', 'LOT', 'ENT'}
            invalid_types = set(org_tree['node_type'].dropna().unique()) - valid_node_types
            if invalid_types:
                self.validation_errors.append(
                    f"ORG_TREE : node_type invalides : {', '.join(invalid_types)}"
                )

            # Vérifier les activités valides
            valid_activities = {'EU', 'AEP', 'NA'}
            invalid_activities = set(org_tree['activity'].dropna().unique()) - valid_activities
            if invalid_activities:
                self.validation_errors.append(
                    f"ORG_TREE : activity invalides : {', '.join(invalid_activities)}"
                )

            # Vérifier qu'il y a au moins un nœud ORG
            org_nodes = org_tree[org_tree['node_type'] == 'ORG']
            if len(org_nodes) == 0:
                self.validation_errors.append("ORG_TREE : aucun nœud de type 'ORG' trouvé")

            # Vérifier que les ENT ont une activité définie (pas NA)
            ent_without_activity = org_tree[
                (org_tree['node_type'] == 'ENT') &
                ((org_tree['activity'].isna()) | (org_tree['activity'] == 'NA'))
            ]
            if len(ent_without_activity) > 0:
                self.validation_errors.append(
                    f"ORG_TREE : {len(ent_without_activity)} ENT sans activité (EU/AEP) définie"
                )

        # Validation EMISSIONS
        if 'EMISSIONS' in self.data:
            emissions = self.data['EMISSIONS']

            # Vérifier les scopes valides
            valid_scopes = {1, 2, 3, '1', '2', '3'}
            invalid_scopes = set(emissions['scope'].dropna().unique()) - valid_scopes
            if invalid_scopes:
                self.validation_errors.append(
                    f"EMISSIONS : scopes invalides : {', '.join(map(str, invalid_scopes))}"
                )

        # Validation TEXTE_RAPPORT
        if 'TEXTE_RAPPORT' in self.data:
            texte_rapport = self.data['TEXTE_RAPPORT']

            # Vérifier les activités valides
            valid_activities = {'EU', 'AEP', 'BOTH'}
            invalid_activities = set(texte_rapport['activity'].dropna().unique()) - valid_activities
            if invalid_activities:
                self.validation_warnings.append(
                    f"TEXTE_RAPPORT : activity invalides : {', '.join(invalid_activities)}"
                )

    def get_validation_report(self) -> Tuple[List[str], List[str]]:
        """
        Retourne les erreurs et warnings de validation.

        Returns:
            Tuple (erreurs, warnings)
        """
        return self.validation_errors, self.validation_warnings

    def get_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """
        Retourne un onglet spécifique.

        Args:
            sheet_name: Nom de l'onglet

        Returns:
            DataFrame ou None si l'onglet n'existe pas
        """
        return self.data.get(sheet_name)

    def get_all_data(self) -> Dict[str, pd.DataFrame]:
        """Retourne toutes les données chargées."""
        return self.data


def load_and_validate_excel(file_path: str) -> Tuple[Dict[str, pd.DataFrame], List[str], List[str]]:
    """
    Fonction utilitaire pour charger et valider un fichier Excel.

    Args:
        file_path: Chemin vers le fichier Excel

    Returns:
        Tuple (données, erreurs, warnings)

    Raises:
        ExcelValidationError: Si le fichier n'est pas valide
    """
    loader = ExcelLoader(file_path)
    data = loader.load()
    errors, warnings = loader.get_validation_report()
    return data, errors, warnings
