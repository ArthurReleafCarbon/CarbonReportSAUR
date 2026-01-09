"""
Module de génération de graphiques avec Matplotlib.
Supporte tous les CHART_KEY définis dans le brief.
"""

import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Backend sans interface graphique
import pandas as pd
from io import BytesIO
from typing import Optional, List
import numpy as np

from .calc_emissions import EmissionResult


class ChartGenerator:
    """Générateur de graphiques pour le rapport."""

    def __init__(self):
        """Initialise le générateur avec les styles par défaut."""
        # Style général
        plt.style.use('seaborn-v0_8-darkgrid')
        self.colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A994E', '#BC4B51']
        self.dpi = 150

    def generate_chart(self, chart_key: str, data: pd.DataFrame,
                      **kwargs) -> Optional[BytesIO]:
        """
        Génère un graphique selon la key spécifiée.

        Args:
            chart_key: Clé du graphique à générer
            data: Données pour le graphique
            **kwargs: Arguments additionnels

        Returns:
            BytesIO contenant l'image PNG ou None si erreur
        """
        if chart_key == 'TRAVAUX_BREAKDOWN':
            return self.generate_travaux_breakdown(data)
        elif chart_key == 'FILE_EAU_BREAKDOWN':
            return self.generate_file_eau_breakdown(data)
        elif chart_key == 'EM_INDIRECTES_SPLIT':
            return self.generate_em_indirectes_split(data)
        elif chart_key == 'chart_emissions_scope_org':
            return self.generate_scope_pie(data, **kwargs)
        elif chart_key == 'chart_contrib_lot':
            return self.generate_lot_contribution(data, **kwargs)
        elif chart_key == 'chart_emissions_total_org':
            return self.generate_total_emissions_bar(data, **kwargs)
        elif chart_key == 'chart_emissions_elec_org':
            return self.generate_elec_emissions(data, **kwargs)
        elif chart_key == 'chart_batonnet_inter_lot_top3':
            return self.generate_inter_lot_top3(data, **kwargs)
        elif chart_key == 'chart_pie_scope_entity_activity':
            return self.generate_scope_pie_entity(data, **kwargs)
        elif chart_key == 'chart_pie_postes_entity_activity':
            return self.generate_postes_pie_entity(data, **kwargs)
        else:
            # Key non supportée - retourner None gracieusement
            return None

    def generate_travaux_breakdown(self, data: pd.DataFrame) -> Optional[BytesIO]:
        """
        Graphique TRAVAUX_BREAKDOWN : répartition des travaux par type.

        Args:
            data: DataFrame avec colonnes ['poste_l2', 'tco2e']

        Returns:
            Image PNG en BytesIO
        """
        if data.empty:
            return None

        fig, ax = plt.subplots(figsize=(10, 6), dpi=self.dpi)

        # Bar chart horizontal
        y_pos = np.arange(len(data))
        ax.barh(y_pos, data['tco2e'], color=self.colors[:len(data)])
        ax.set_yticks(y_pos)
        ax.set_yticklabels(data['poste_l2'])
        ax.set_xlabel('Émissions (tCO₂e)')
        ax.set_title('Répartition des émissions - Travaux')
        ax.invert_yaxis()  # Plus gros en haut

        plt.tight_layout()

        # Sauvegarder dans BytesIO
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_file_eau_breakdown(self, data: pd.DataFrame) -> Optional[BytesIO]:
        """
        Graphique FILE_EAU_BREAKDOWN : répartition file eau STEP (N-N20, N2O, CH4).

        Args:
            data: DataFrame avec colonnes ['poste_l2', 'tco2e']

        Returns:
            Image PNG en BytesIO
        """
        if data.empty:
            return None

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)

        # Pie chart
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        ax.pie(data['tco2e'], labels=data['poste_l2'], autopct='%1.1f%%',
               colors=colors[:len(data)], startangle=90)
        ax.set_title('Répartition des émissions - File eau STEP')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_em_indirectes_split(self, data: pd.DataFrame) -> Optional[BytesIO]:
        """
        Graphique EM_INDIRECTES_SPLIT : répartition émissions indirectes.

        Args:
            data: DataFrame avec colonnes ['poste_l2', 'tco2e']

        Returns:
            Image PNG en BytesIO
        """
        if data.empty:
            return None

        fig, ax = plt.subplots(figsize=(10, 6), dpi=self.dpi)

        # Bar chart
        x_pos = np.arange(len(data))
        ax.bar(x_pos, data['tco2e'], color=self.colors[:len(data)])
        ax.set_xticks(x_pos)
        ax.set_xticklabels(data['poste_l2'], rotation=45, ha='right')
        ax.set_ylabel('Émissions (tCO₂e)')
        ax.set_title('Répartition des émissions indirectes')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_scope_pie(self, emission_result: EmissionResult, **kwargs) -> Optional[BytesIO]:
        """
        Graphique camembert des scopes (ORG).

        Args:
            emission_result: Résultat d'émissions ORG

        Returns:
            Image PNG en BytesIO
        """
        scopes = ['Scope 1', 'Scope 2', 'Scope 3']
        values = [
            emission_result.scope1_tco2e,
            emission_result.scope2_tco2e,
            emission_result.scope3_tco2e
        ]

        # Filtrer les scopes à 0
        filtered_data = [(s, v) for s, v in zip(scopes, values) if v > 0]
        if not filtered_data:
            return None

        scopes, values = zip(*filtered_data)

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        ax.pie(values, labels=scopes, autopct='%1.1f%%', colors=colors[:len(values)],
               startangle=90)
        ax.set_title('Répartition par scope')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_lot_contribution(self, lot_data: List[tuple], **kwargs) -> Optional[BytesIO]:
        """
        Graphique contribution des LOTs.

        Args:
            lot_data: Liste de tuples (lot_name, tco2e)

        Returns:
            Image PNG en BytesIO
        """
        if not lot_data:
            return None

        names, values = zip(*lot_data)

        fig, ax = plt.subplots(figsize=(10, 6), dpi=self.dpi)
        x_pos = np.arange(len(names))
        ax.bar(x_pos, values, color=self.colors[:len(names)])
        ax.set_xticks(x_pos)
        ax.set_xticklabels(names, rotation=45, ha='right')
        ax.set_ylabel('Émissions (tCO₂e)')
        ax.set_title('Contribution des LOTs')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_total_emissions_bar(self, emission_result: EmissionResult, **kwargs) -> Optional[BytesIO]:
        """
        Graphique bar des émissions totales par scope.

        Args:
            emission_result: Résultat d'émissions

        Returns:
            Image PNG en BytesIO
        """
        scopes = ['Scope 1', 'Scope 2', 'Scope 3']
        values = [
            emission_result.scope1_tco2e,
            emission_result.scope2_tco2e,
            emission_result.scope3_tco2e
        ]

        fig, ax = plt.subplots(figsize=(10, 6), dpi=self.dpi)
        x_pos = np.arange(len(scopes))
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        ax.bar(x_pos, values, color=colors)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(scopes)
        ax.set_ylabel('Émissions (tCO₂e)')
        ax.set_title('Émissions totales par scope')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_elec_emissions(self, data: pd.DataFrame, **kwargs) -> Optional[BytesIO]:
        """
        Graphique émissions électricité (placeholder).

        Args:
            data: Données émissions électricité

        Returns:
            Image PNG en BytesIO
        """
        # TODO: À implémenter selon les données disponibles
        return None

    def generate_inter_lot_top3(self, top_postes_data: List[tuple], **kwargs) -> Optional[BytesIO]:
        """
        Graphique bâtonnet inter-LOT top 3 postes.

        Args:
            top_postes_data: Liste de tuples (poste_name, tco2e)

        Returns:
            Image PNG en BytesIO
        """
        if not top_postes_data:
            return None

        names, values = zip(*top_postes_data[:3])  # Top 3

        fig, ax = plt.subplots(figsize=(10, 6), dpi=self.dpi)
        x_pos = np.arange(len(names))
        ax.bar(x_pos, values, color=self.colors[:len(names)])
        ax.set_xticks(x_pos)
        ax.set_xticklabels(names, rotation=45, ha='right')
        ax.set_ylabel('Émissions (tCO₂e)')
        ax.set_title('Top 3 postes émetteurs')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_scope_pie_entity(self, emission_result: EmissionResult, **kwargs) -> Optional[BytesIO]:
        """
        Camembert scopes pour LOT × ACTIVITÉ.

        Args:
            emission_result: Résultat d'émissions LOT×ACT

        Returns:
            Image PNG en BytesIO
        """
        return self.generate_scope_pie(emission_result)

    def generate_postes_pie_entity(self, emission_result: EmissionResult, **kwargs) -> Optional[BytesIO]:
        """
        Camembert des postes pour LOT × ACTIVITÉ.

        Args:
            emission_result: Résultat d'émissions LOT×ACT

        Returns:
            Image PNG en BytesIO
        """
        if not emission_result.emissions_by_poste:
            return None

        # Prendre top 5 postes + regrouper le reste
        sorted_postes = sorted(emission_result.emissions_by_poste.items(),
                             key=lambda x: x[1], reverse=True)

        if len(sorted_postes) <= 5:
            labels = [p[0] for p in sorted_postes]
            values = [p[1] for p in sorted_postes]
        else:
            labels = [p[0] for p in sorted_postes[:5]] + ['Autres']
            values = [p[1] for p in sorted_postes[:5]] + [sum(p[1] for p in sorted_postes[5:])]

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        ax.pie(values, labels=labels, autopct='%1.1f%%',
               colors=self.colors[:len(values)], startangle=90)
        ax.set_title(f'Répartition par poste - {emission_result.node_name}')

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer
