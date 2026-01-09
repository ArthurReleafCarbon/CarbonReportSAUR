"""
Module de calcul des émissions.
Gère les agrégations ORG, LOT×ACTIVITÉ, scopes, top postes et calculs BRUT/NET.
"""

import pandas as pd
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field

from .tree import OrganizationTree, TreeNode


@dataclass
class EmissionOverrides:
    """Configuration des overrides utilisateur pour les postes."""
    # Renommage des nœuds : {node_id: nouveau_nom}
    node_renames: Dict[str, str] = field(default_factory=dict)

    # Gestion des postes L1 : {poste_l1_code: config}
    # show_in_report: afficher dans le rapport
    # include_in_totals: inclure dans les totaux
    poste_config: Dict[str, Dict[str, bool]] = field(default_factory=dict)

    def get_node_name(self, node_id: str, original_name: str) -> str:
        """Retourne le nom du nœud (renommé ou original)."""
        return self.node_renames.get(node_id, original_name)

    def is_poste_shown(self, poste_l1_code: str) -> bool:
        """Vérifie si un poste doit être affiché dans le rapport."""
        config = self.poste_config.get(poste_l1_code, {})
        return config.get('show_in_report', True)

    def is_poste_included(self, poste_l1_code: str) -> bool:
        """Vérifie si un poste doit être inclus dans les totaux."""
        config = self.poste_config.get(poste_l1_code, {})
        return config.get('include_in_totals', True)

    def set_poste_config(self, poste_l1_code: str, show_in_report: bool, include_in_totals: bool):
        """Configure l'affichage et l'inclusion d'un poste."""
        self.poste_config[poste_l1_code] = {
            'show_in_report': show_in_report,
            'include_in_totals': include_in_totals
        }

    def get_excluded_postes(self) -> List[str]:
        """Retourne la liste des postes exclus des totaux."""
        return [
            code for code, config in self.poste_config.items()
            if not config.get('include_in_totals', True)
        ]


@dataclass
class EmissionResult:
    """Résultat de calcul d'émissions pour un périmètre donné."""
    node_id: str
    node_name: str
    activity: Optional[str]  # None pour ORG global, EU/AEP pour LOT×ACT

    # Totaux
    total_tco2e: float = 0.0
    scope1_tco2e: float = 0.0
    scope2_tco2e: float = 0.0
    scope3_tco2e: float = 0.0

    # Émissions par poste L1
    emissions_by_poste: Dict[str, float] = field(default_factory=dict)

    # Top postes (liste ordonnée)
    top_postes: List[Tuple[str, float]] = field(default_factory=list)

    # Autres postes (hors top)
    other_postes: List[Tuple[str, float]] = field(default_factory=list)

    def get_scope_percentage(self, scope: int) -> float:
        """Calcule le pourcentage d'un scope par rapport au total."""
        if self.total_tco2e == 0:
            return 0.0
        scope_value = getattr(self, f'scope{scope}_tco2e', 0.0)
        return (scope_value / self.total_tco2e) * 100


class EmissionCalculator:
    """
    Calcule les émissions à tous les niveaux (ORG, LOT×ACTIVITÉ).
    Gère les calculs BRUT (sans overrides) et NET (avec overrides).
    """

    def __init__(self, tree: OrganizationTree, emissions_df: pd.DataFrame,
                 postes_ref_df: pd.DataFrame):
        """
        Initialise le calculateur.

        Args:
            tree: Arborescence de l'organisation
            emissions_df: DataFrame EMISSIONS
            postes_ref_df: DataFrame POSTES_REF
        """
        self.tree = tree
        self.emissions_df = emissions_df
        self.postes_ref_df = postes_ref_df

        # Préparer les données
        self._prepare_data()

    def _prepare_data(self):
        """Prépare les données pour les calculs."""
        # Normaliser les scopes en int
        self.emissions_df['scope'] = self.emissions_df['scope'].astype(int)

        # Créer un dictionnaire des labels de postes
        self.poste_labels = {}
        for _, row in self.postes_ref_df.iterrows():
            self.poste_labels[row['poste_l1_code']] = row['poste_l1_label']

    def calculate_brut(self, top_n: int = 4) -> Dict[str, EmissionResult]:
        """
        Calcule les émissions BRUT (sans overrides).

        Args:
            top_n: Nombre de top postes à calculer

        Returns:
            Dictionnaire {key: EmissionResult}
            Keys format: 'ORG', 'LOT_{lot_id}_{activity}'
        """
        return self._calculate(EmissionOverrides(), top_n)

    def calculate_net(self, overrides: EmissionOverrides, top_n: int = 4) -> Dict[str, EmissionResult]:
        """
        Calcule les émissions NET (avec overrides).

        Args:
            overrides: Configuration des overrides
            top_n: Nombre de top postes à calculer

        Returns:
            Dictionnaire {key: EmissionResult}
        """
        return self._calculate(overrides, top_n)

    def _calculate(self, overrides: EmissionOverrides, top_n: int) -> Dict[str, EmissionResult]:
        """
        Effectue les calculs d'émissions.

        Args:
            overrides: Configuration des overrides
            top_n: Nombre de top postes

        Returns:
            Dictionnaire des résultats
        """
        results = {}

        # 1. Calcul au niveau ORG (total global)
        org_result = self._calculate_org(overrides, top_n)
        results['ORG'] = org_result

        # 2. Calcul par LOT × ACTIVITÉ (si LOTs présents)
        if self.tree.has_lots():
            for lot in self.tree.get_lots():
                activities = self.tree.get_lot_activities(lot.node_id)
                for activity in activities:
                    key = f"LOT_{lot.node_id}_{activity}"
                    result = self._calculate_lot_activity(lot, activity, overrides, top_n)
                    results[key] = result
        else:
            # Pas de LOT : calculer directement ORG × ACTIVITÉ
            activities = self.tree.get_org_activities()
            for activity in activities:
                key = f"ORG_{activity}"
                result = self._calculate_org_activity(activity, overrides, top_n)
                results[key] = result

        return results

    def _calculate_org(self, overrides: EmissionOverrides, top_n: int) -> EmissionResult:
        """Calcule les émissions au niveau ORG (tous ENT confondus)."""
        org = self.tree.get_org()
        all_ents = self.tree.get_ents()
        ent_ids = [ent.node_id for ent in all_ents]

        return self._aggregate_emissions(
            node_id=org.node_id,
            node_name=overrides.get_node_name(org.node_id, org.node_name),
            ent_ids=ent_ids,
            activity=None,
            overrides=overrides,
            top_n=top_n
        )

    def _calculate_org_activity(self, activity: str, overrides: EmissionOverrides,
                                top_n: int) -> EmissionResult:
        """Calcule les émissions ORG × ACTIVITÉ (cas sans LOT)."""
        org = self.tree.get_org()
        ent_ids = self.tree.get_ent_ids_by_activity(org.node_id, activity)

        return self._aggregate_emissions(
            node_id=org.node_id,
            node_name=overrides.get_node_name(org.node_id, org.node_name),
            ent_ids=ent_ids,
            activity=activity,
            overrides=overrides,
            top_n=top_n
        )

    def _calculate_lot_activity(self, lot: TreeNode, activity: str,
                               overrides: EmissionOverrides, top_n: int) -> EmissionResult:
        """Calcule les émissions LOT × ACTIVITÉ."""
        ent_ids = self.tree.get_ent_ids_by_activity(lot.node_id, activity)

        return self._aggregate_emissions(
            node_id=lot.node_id,
            node_name=overrides.get_node_name(lot.node_id, lot.node_name),
            ent_ids=ent_ids,
            activity=activity,
            overrides=overrides,
            top_n=top_n
        )

    def _aggregate_emissions(self, node_id: str, node_name: str, ent_ids: List[str],
                           activity: Optional[str], overrides: EmissionOverrides,
                           top_n: int) -> EmissionResult:
        """
        Agrège les émissions pour un ensemble d'ENT.

        Args:
            node_id: ID du nœud (ORG ou LOT)
            node_name: Nom du nœud
            ent_ids: Liste des IDs d'ENT à agréger
            activity: Activité (None pour ORG global)
            overrides: Overrides à appliquer
            top_n: Nombre de top postes

        Returns:
            Résultat d'émissions agrégé
        """
        result = EmissionResult(
            node_id=node_id,
            node_name=node_name,
            activity=activity
        )

        if len(ent_ids) == 0:
            return result

        # Filtrer les émissions pour ces ENT
        mask = self.emissions_df['node_id'].isin(ent_ids)
        emissions = self.emissions_df[mask].copy()

        # Appliquer les overrides sur les postes
        if overrides:
            # Filtrer les postes à inclure dans les totaux
            included_mask = emissions['poste_l1_code'].apply(
                lambda code: overrides.is_poste_included(code)
            )
            emissions_for_totals = emissions[included_mask]
        else:
            emissions_for_totals = emissions

        # Calculer les totaux par scope
        result.scope1_tco2e = emissions_for_totals[emissions_for_totals['scope'] == 1]['tco2e'].sum()
        result.scope2_tco2e = emissions_for_totals[emissions_for_totals['scope'] == 2]['tco2e'].sum()
        result.scope3_tco2e = emissions_for_totals[emissions_for_totals['scope'] == 3]['tco2e'].sum()
        result.total_tco2e = result.scope1_tco2e + result.scope2_tco2e + result.scope3_tco2e

        # Agréger par poste L1 (pour les postes inclus)
        poste_groups = emissions_for_totals.groupby('poste_l1_code')['tco2e'].sum()
        result.emissions_by_poste = poste_groups.to_dict()

        # Calculer les top postes (triés par émissions décroissantes)
        sorted_postes = sorted(result.emissions_by_poste.items(), key=lambda x: x[1], reverse=True)

        # Filtrer les postes à afficher dans le rapport
        if overrides:
            sorted_postes = [(code, value) for code, value in sorted_postes
                           if overrides.is_poste_shown(code)]

        result.top_postes = sorted_postes[:top_n]
        result.other_postes = sorted_postes[top_n:]

        return result

    def get_poste_label(self, poste_l1_code: str) -> str:
        """Retourne le label d'un poste L1."""
        return self.poste_labels.get(poste_l1_code, poste_l1_code)

    def get_emissions_l2(self, ent_ids: List[str], poste_l1_code: str,
                        emissions_l2_df: pd.DataFrame) -> pd.DataFrame:
        """
        Récupère les émissions de niveau 2 pour un poste L1 donné.

        Args:
            ent_ids: Liste des IDs d'ENT
            poste_l1_code: Code du poste L1
            emissions_l2_df: DataFrame EMISSIONS_L2

        Returns:
            DataFrame des émissions L2 agrégées par poste_l2
        """
        mask = (emissions_l2_df['node_id'].isin(ent_ids)) & \
               (emissions_l2_df['poste_l1_code'] == poste_l1_code)

        l2_data = emissions_l2_df[mask].copy()

        if len(l2_data) == 0:
            return pd.DataFrame(columns=['poste_l2', 'tco2e'])

        # Agréger par poste_l2
        aggregated = l2_data.groupby('poste_l2')['tco2e'].sum().reset_index()
        return aggregated.sort_values('tco2e', ascending=False)
