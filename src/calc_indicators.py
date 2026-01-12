"""
Module de calcul des indicateurs.
Gère les indicateurs par LOT×ACTIVITÉ ou ORG×ACTIVITÉ.
"""

import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass, field

from .tree import OrganizationTree


@dataclass
class IndicatorValue:
    """Valeur d'un indicateur."""
    indicator_code: str
    indicator_label: str
    value: float
    unit: str
    comment: Optional[str] = None


@dataclass
class IndicatorResult:
    """Résultat des indicateurs pour un périmètre donné."""
    node_id: str
    node_name: str
    activity: str  # EU ou AEP
    indicators: Dict[str, IndicatorValue] = field(default_factory=dict)

    def add_indicator(self, indicator: IndicatorValue):
        """Ajoute un indicateur."""
        self.indicators[indicator.indicator_code] = indicator

    def get_indicator(self, indicator_code: str) -> Optional[IndicatorValue]:
        """Récupère un indicateur par son code."""
        return self.indicators.get(indicator_code)


class IndicatorCalculator:
    """
    Calcule les indicateurs à tous les niveaux.
    Un indicateur appartient à une activité + un LOT, ou activité + ORG si pas de LOT.
    """

    def __init__(self, tree: OrganizationTree, indicators_df: pd.DataFrame,
                 indicators_ref_df: pd.DataFrame):
        """
        Initialise le calculateur.

        Args:
            tree: Arborescence de l'organisation
            indicators_df: DataFrame INDICATORS
            indicators_ref_df: DataFrame INDICATORS_REF
        """
        self.tree = tree
        self.indicators_df = indicators_df
        self.indicators_ref_df = indicators_ref_df

        # Préparer les référentiels
        self._prepare_references()

    def _prepare_references(self):
        """Prépare les dictionnaires de référence."""
        self.indicator_info = {}
        for _, row in self.indicators_ref_df.iterrows():
            self.indicator_info[row['indicator_code']] = {
                'label': row['indicator_label'],
                'default_unit': row['default_unit'],
                'activity_scope': row.get('activity_scope', None),
                'display_order': row.get('display_order', 999)
            }

    def calculate(self) -> Dict[str, IndicatorResult]:
        """
        Calcule tous les indicateurs.

        Returns:
            Dictionnaire {key: IndicatorResult}
            Keys format: 'LOT_{lot_id}_{activity}' ou 'ORG_{activity}'
        """
        results = {}

        if self.tree.has_lots():
            # Cas avec LOTs : indicateurs par LOT × ACTIVITÉ
            for lot in self.tree.get_lots():
                activities = self.tree.get_lot_activities(lot.node_id)
                for activity in activities:
                    key = f"LOT_{lot.node_id}_{activity}"
                    result = self._calculate_lot_activity(lot.node_id, lot.node_name, activity)
                    if result and len(result.indicators) > 0:
                        results[key] = result
        else:
            # Cas sans LOT : indicateurs par ORG × ACTIVITÉ
            org = self.tree.get_org()
            activities = self.tree.get_org_activities()
            for activity in activities:
                key = f"ORG_{activity}"
                result = self._calculate_org_activity(org.node_id, org.node_name, activity)
                if result and len(result.indicators) > 0:
                    results[key] = result

        return results

    def _calculate_lot_activity(self, lot_id: str, lot_name: str,
                               activity: str) -> Optional[IndicatorResult]:
        """
        Calcule les indicateurs pour LOT × ACTIVITÉ.
        Agrège les indicateurs de tous les ENT enfants du LOT.

        Args:
            lot_id: ID du LOT
            lot_name: Nom du LOT
            activity: Activité (EU ou AEP)

        Returns:
            IndicatorResult ou None si pas d'indicateurs
        """
        # Récupérer tous les ENT enfants de ce LOT avec cette activité
        ent_ids = self.tree.get_ent_ids_by_activity(lot_id, activity)

        if not ent_ids:
            # Fallback : chercher directement par LOT (si données au niveau LOT)
            mask = (self.indicators_df['node_id'] == lot_id) & \
                   (self.indicators_df['activity'] == activity)
            lot_indicators = self.indicators_df[mask]
        else:
            # Filtrer les indicateurs pour tous les ENT de ce LOT
            mask = (self.indicators_df['node_id'].isin(ent_ids)) & \
                   (self.indicators_df['activity'] == activity)
            lot_indicators = self.indicators_df[mask]

        if len(lot_indicators) == 0:
            return None

        result = IndicatorResult(
            node_id=lot_id,
            node_name=lot_name,
            activity=activity
        )

        # Grouper par indicator_code et sommer les valeurs
        # (car plusieurs ENT peuvent contribuer au même indicateur)
        grouped = lot_indicators.groupby('indicator_code').agg({
            'value': 'sum',
            'unit': 'first',  # Prendre la première unité (devrait être identique)
            'comment': 'first'  # Prendre le premier commentaire
        }).reset_index()

        # Ajouter chaque indicateur
        for _, row in grouped.iterrows():
            indicator_code = row['indicator_code']
            info = self.indicator_info.get(indicator_code, {})

            indicator = IndicatorValue(
                indicator_code=indicator_code,
                indicator_label=info.get('label', indicator_code),
                value=float(row['value']),
                unit=row['unit'] if pd.notna(row['unit']) else info.get('default_unit', ''),
                comment=row['comment'] if pd.notna(row['comment']) else None
            )
            result.add_indicator(indicator)

        return result

    def _calculate_org_activity(self, org_id: str, org_name: str,
                               activity: str) -> Optional[IndicatorResult]:
        """
        Calcule les indicateurs pour ORG × ACTIVITÉ (cas sans LOT).

        Args:
            org_id: ID de l'ORG
            org_name: Nom de l'ORG
            activity: Activité (EU ou AEP)

        Returns:
            IndicatorResult ou None si pas d'indicateurs
        """
        # Filtrer les indicateurs pour cette ORG et cette activité
        mask = (self.indicators_df['node_id'] == org_id) & \
               (self.indicators_df['activity'] == activity)

        org_indicators = self.indicators_df[mask]

        if len(org_indicators) == 0:
            return None

        result = IndicatorResult(
            node_id=org_id,
            node_name=org_name,
            activity=activity
        )

        # Ajouter chaque indicateur
        for _, row in org_indicators.iterrows():
            indicator_code = row['indicator_code']
            info = self.indicator_info.get(indicator_code, {})

            indicator = IndicatorValue(
                indicator_code=indicator_code,
                indicator_label=info.get('label', indicator_code),
                value=float(row['value']),
                unit=row['unit'] if pd.notna(row['unit']) else info.get('default_unit', ''),
                comment=row['comment'] if pd.notna(row['comment']) else None
            )
            result.add_indicator(indicator)

        return result

    def get_sorted_indicators(self, result: IndicatorResult) -> List[IndicatorValue]:
        """
        Retourne les indicateurs triés par display_order.

        Args:
            result: Résultat d'indicateurs

        Returns:
            Liste d'IndicatorValue triés
        """
        indicators = list(result.indicators.values())

        # Trier par display_order
        indicators.sort(key=lambda ind: self.indicator_info.get(
            ind.indicator_code, {}
        ).get('display_order', 999))

        return indicators
