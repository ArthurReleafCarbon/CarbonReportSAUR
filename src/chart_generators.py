"""
Module de génération de graphiques avec Matplotlib.
Supporte tous les CHART_KEY définis dans le brief.
"""

import matplotlib.pyplot as plt
import matplotlib
from matplotlib import font_manager as fm
matplotlib.use('Agg')  # Backend sans interface graphique
import pandas as pd
from io import BytesIO
from typing import Optional, List
import numpy as np
from pathlib import Path

from .calc_emissions import EmissionResult


class ChartGenerator:
    """Générateur de graphiques pour le rapport."""

    def __init__(self):
        """Initialise le générateur avec les styles par défaut."""
        # Style général
        plt.style.use('default')
        self.colors = ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8', '#E9F7F4']
        self.dpi = 150
        self._load_fonts()
        plt.rcParams["axes.grid"] = False
        plt.rcParams["axes.facecolor"] = "white"
        plt.rcParams["figure.facecolor"] = "white"
        self._pie_textprops = {
            "fontproperties": self.body_font,
            "fontweight": "bold",
            "fontsize": plt.rcParams.get("font.size", 10) + 1
        }

    def _style_axes(self, ax):
        """Applique un style sans cadres ni axes."""
        ax.grid(False)
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(length=0)

    def _pie_autopct(self, pct: float) -> str:
        """Formatte les % de pie chart sans décimales."""
        return f"{int(round(pct))}%"

    def _load_fonts(self):
        """Charge les polices Poppins pour les graphiques."""
        font_dir = Path(__file__).resolve().parent.parent / "assets" / "police"
        if font_dir.exists():
            for font_path in font_dir.glob("*.ttf"):
                try:
                    fm.fontManager.addfont(str(font_path))
                except Exception:
                    pass

        plt.rcParams["font.family"] = "Poppins"
        plt.rcParams["font.weight"] = "normal"
        plt.rcParams["mathtext.fontset"] = "custom"
        plt.rcParams["mathtext.rm"] = "Poppins"
        plt.rcParams["mathtext.it"] = "Poppins:italic"
        plt.rcParams["mathtext.bf"] = "Poppins:bold"
        self.title_font = fm.FontProperties(family="Poppins", weight="bold", size=12)
        self.body_font = fm.FontProperties(family="Poppins", weight="normal")

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
            return self.generate_total_emissions_pie(data, **kwargs)
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
        ax.set_xlabel('Émissions (tCO$_2$e)')
        ax.set_title('Répartition des émissions - Travaux', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)
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

        # Pie chart avec légende à droite et pourcentages à l'extérieur
        colors = ['#2E86AB', '#A23B72', '#F18F01']
        explode = [0.05] * len(data)
        wedges, texts, autotexts = ax.pie(
            data['tco2e'],
            labels=None,  # Pas de labels sur le graphique, on utilise la légende
            autopct=self._pie_autopct,
            textprops=self._pie_textprops,
            colors=colors[:len(data)],
            startangle=90,
            explode=explode,
            pctdistance=1.15  # Placer les pourcentages à l'extérieur
        )

        # Mettre les pourcentages en blanc pour meilleure lisibilité
        for autotext in autotexts:
            autotext.set_color('black')

        # Ajouter la légende à droite
        ax.legend(
            wedges,
            data['poste_l2'],
            loc='center left',
            bbox_to_anchor=(1.05, 0.5),
            frameon=False,
            prop=self.body_font
        )

        ax.set_title('Répartition des émissions - File eau STEP', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

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

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)

        # Pie chart (légende à droite, pas de labels autour du pie)
        explode = [0.05] * len(data)
        wedges, texts, autotexts = ax.pie(
            data['tco2e'],
            labels=None,
            autopct=self._pie_autopct,
            textprops=self._pie_textprops,
            colors=self.colors[:len(data)],
            startangle=90,
            explode=explode
        )
        ax.legend(
            wedges,
            data['poste_l2'],
            loc='center left',
            bbox_to_anchor=(1.05, 0.5),
            frameon=False,
            prop=self.body_font
        )
        ax.set_title('Répartition des émissions indirectes', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

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

        # Couleurs pour les scopes (bleu clair, bleu moyen, bleu foncé)
        scope_colors = ['#89CFF0', '#5B9BD5', '#4472C4']

        # Créer l'explode pour ajouter de l'espace entre les morceaux
        explode = [0.05] * len(values)

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        wedges, texts, autotexts = ax.pie(
            values,
            labels=None,
            autopct=self._pie_autopct,
            textprops=self._pie_textprops,
            colors=scope_colors[:len(values)],
            startangle=90,
            explode=explode
        )
        for text in autotexts:
            text.set_color("white")
        ax.legend(
            wedges,
            scopes,
            loc='center left',
            bbox_to_anchor=(1.05, 0.5),
            frameon=False,
            prop=self.body_font
        )

        # Utiliser le nom de l'organisation si disponible
        org_name = kwargs.get('org_name', 'ORG')
        ax.set_title(f'Répartition par scope - {org_name}', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_lot_contribution(self, lot_data: List[tuple], **kwargs) -> Optional[BytesIO]:
        """
        Graphique contribution des LOTs (pie chart).

        Args:
            lot_data: Liste de tuples (lot_name, tco2e)

        Returns:
            Image PNG en BytesIO
        """
        if not lot_data:
            return None

        names, values = zip(*lot_data)

        # Créer le pie chart avec les mêmes couleurs que les autres graphiques
        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        explode = [0.05] * len(names)

        wedges, texts, autotexts = ax.pie(
            values,
            labels=names,
            autopct=self._pie_autopct,
            textprops=self._pie_textprops,
            colors=self.colors[:len(names)],
            startangle=90,
            explode=explode
        )

        # Mettre les pourcentages en blanc et gras
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')

        # Utiliser le nom de l'organisation si disponible
        org_name = kwargs.get('org_name', 'ORG')
        ax.set_title(f'Contribution des LOTs - {org_name}', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_total_emissions_pie(self, emission_result: EmissionResult,
                                     poste_labels: dict = None, **kwargs) -> Optional[BytesIO]:
        """
        Graphique PIE de la contribution de chaque poste L1 au niveau ORG.

        Args:
            emission_result: Résultat d'émissions ORG
            poste_labels: Dictionnaire {code: label}

        Returns:
            Image PNG en BytesIO
        """
        if not emission_result.emissions_by_poste:
            return None

        # Trier par valeur décroissante
        sorted_postes = sorted(emission_result.emissions_by_poste.items(),
                              key=lambda x: x[1], reverse=True)

        # Prendre top 5 + regrouper le reste
        if len(sorted_postes) <= 5:
            labels = []
            values = []
            for code, value in sorted_postes:
                label = poste_labels.get(code, code) if poste_labels else code
                labels.append(label)
                values.append(value)
        else:
            labels = []
            values = []
            for code, value in sorted_postes[:5]:
                label = poste_labels.get(code, code) if poste_labels else code
                labels.append(label)
                values.append(value)
            # Regrouper le reste
            labels.append('Autres')
            values.append(sum(v for _, v in sorted_postes[5:]))

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        explode = [0.05] * len(values)
        ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=self.colors[:len(values)], startangle=90, explode=explode)
        ax.set_title('Contribution des postes - ORG', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_elec_emissions(self, emissions_by_activity: dict, **kwargs) -> Optional[BytesIO]:
        """
        Graphique PIE des émissions électricité par activité (EU vs AEP).

        Args:
            emissions_by_activity: Dict {activity: tco2e} pour le poste électricité

        Returns:
            Image PNG en BytesIO
        """
        if not emissions_by_activity:
            return None

        # Filtrer les valeurs > 0
        filtered = {k: v for k, v in emissions_by_activity.items() if v > 0}
        if not filtered:
            return None

        labels = list(filtered.keys())
        values = list(filtered.values())

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        colors = ['#2E86AB', '#A23B72']
        explode = [0.05] * len(values)
        ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=colors[:len(values)], startangle=90, explode=explode)
        ax.set_title('Répartition émissions Électricité par activité', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_elec_emissions_by_lot(self, emissions_by_lot: dict, **kwargs) -> Optional[BytesIO]:
        """
        Graphique PIE des émissions électricité par LOT.

        Args:
            emissions_by_lot: Dict {lot_name: tco2e} pour le poste électricité

        Returns:
            Image PNG en BytesIO
        """
        if not emissions_by_lot:
            return None

        # Filtrer les valeurs > 0
        filtered = {k: v for k, v in emissions_by_lot.items() if v > 0}
        if not filtered:
            return None

        labels = list(filtered.keys())
        values = list(filtered.values())

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        explode = [0.05] * len(values)
        ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=self.colors[:len(values)], startangle=90, explode=explode)
        ax.set_title('Répartition émissions Électricité par LOT', fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_inter_lot_top3(self, top3_by_lot: dict, **kwargs) -> Optional[BytesIO]:
        """
        Graphique bâtonnet groupé: top 3 postes comparés entre LOTs.

        Affiche une barre pour chaque LOT, regroupées par poste d'émission.

        Args:
            top3_by_lot: Dict {poste_name: {lot_name: tco2e}}

        Returns:
            Image PNG en BytesIO
        """
        if not top3_by_lot:
            return None

        # Extraire postes et LOTs
        postes = list(top3_by_lot.keys())
        all_lots = set()
        for poste_lots in top3_by_lot.values():
            all_lots.update(poste_lots.keys())
        lots = sorted(list(all_lots))

        if not lots or not postes:
            return None

        # Préparer les données pour chaque LOT
        x = np.arange(len(postes))
        width = 0.8 / len(lots)  # Largeur de chaque barre

        fig, ax = plt.subplots(figsize=(12, 7), dpi=self.dpi)

        # Tracer une barre pour chaque LOT
        lot_colors = self.colors[:max(1, len(lots))]
        for i, lot in enumerate(lots):
            values = [top3_by_lot[poste].get(lot, 0) for poste in postes]
            offset = width * i - (width * len(lots) / 2) + width / 2
            bars = ax.bar(x + offset, values, width, label=lot,
                          color=lot_colors[i % len(lot_colors)])

            for bar, value in zip(bars, values):
                if value <= 0:
                    continue
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + (max(values) * 0.02 if max(values) else 1),
                    f"{int(round(value)):,}".replace(",", " "),
                    ha='center',
                    va='bottom',
                    color=lot_colors[i % len(lot_colors)],
                    fontproperties=self.body_font
                )

        ax.set_xlabel('Poste d\'émission', fontproperties=self.body_font)
        ax.set_ylabel('Émissions (tCO$_2$e)', fontproperties=self.body_font)
        ax.set_title('Comparaison de l’impact entre lots – Postes\nmajeurs',
                     fontproperties=self.title_font, pad=24)
        ax.set_xticks(x)
        ax.set_xticklabels(postes, rotation=0, ha='center', fontproperties=self.body_font)
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15),
                  ncol=min(3, len(lots)), frameon=False)
        self._style_axes(ax)

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

    def generate_postes_pie_entity(self, emission_result: EmissionResult,
                                   poste_labels: dict = None, **kwargs) -> Optional[BytesIO]:
        """
        Camembert des postes pour LOT × ACTIVITÉ.

        Args:
            emission_result: Résultat d'émissions LOT×ACT
            poste_labels: Dictionnaire {code: label}

        Returns:
            Image PNG en BytesIO
        """
        if not emission_result.emissions_by_poste:
            return None

        # Prendre top 5 postes + regrouper le reste
        sorted_postes = sorted(emission_result.emissions_by_poste.items(),
                             key=lambda x: x[1], reverse=True)

        if len(sorted_postes) <= 5:
            labels = []
            values = []
            for code, value in sorted_postes:
                label = poste_labels.get(code, code) if poste_labels else code
                labels.append(label)
                values.append(value)
        else:
            labels = []
            values = []
            for code, value in sorted_postes[:5]:
                label = poste_labels.get(code, code) if poste_labels else code
                labels.append(label)
                values.append(value)
            # Regrouper le reste
            labels.append('Autres')
            values.append(sum(v for _, v in sorted_postes[5:]))

        fig, ax = plt.subplots(figsize=(8, 8), dpi=self.dpi)
        explode = [0.05] * len(values)
        ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=self.colors[:len(values)], startangle=90, explode=explode)
        ax.set_title(f'Répartition par poste - {emission_result.node_name}',
                     fontproperties=self.title_font, pad=20)
        self._style_axes(ax)

        plt.tight_layout()

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_reactif_breakdown(self, df: pd.DataFrame) -> BytesIO:
        """
        Génère un graphique donut de répartition des réactifs.

        Args:
            df: DataFrame avec colonnes 'poste_l2' et 'tco2e'

        Returns:
            Buffer image PNG
        """
        if df is None or df.empty:
            return None

        # Grouper et sommer par poste_l2
        grouped = df.groupby('poste_l2', as_index=False)['tco2e'].sum()
        grouped = grouped.sort_values('tco2e', ascending=False)

        # Calculer les pourcentages
        total = grouped['tco2e'].sum()
        if total == 0:
            return None

        grouped['percentage'] = (grouped['tco2e'] / total * 100).round(0).astype(int)

        # Préparer les données pour le graphique
        labels = grouped['poste_l2'].tolist()
        sizes = grouped['tco2e'].tolist()
        percentages = grouped['percentage'].tolist()

        # Créer la figure
        fig, ax = plt.subplots(figsize=(8, 6))

        # Palette de couleurs (inspirée de l'image)
        colors = ['#1b4d3e', '#2d8b6b', '#f4c542', '#e8a87c']
        # Étendre la palette si nécessaire
        while len(colors) < len(labels):
            colors.append(f'#{hash(labels[len(colors)]) % 0xFFFFFF:06x}')

        # Créer le donut
        explode = [0.05] * len(labels)
        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=None,  # Pas de labels sur le graphique
            autopct=lambda pct: f'{int(pct)}%' if pct > 5 else '',
            startangle=90,
            colors=colors[:len(labels)],
            wedgeprops={'width': 0.4, 'edgecolor': 'white', 'linewidth': 2},
            textprops={'fontsize': 14, 'weight': 'bold', 'color': 'white'},
            explode=explode
        )

        # Ajouter le titre
        ax.set_title(
            'Répartition de l\'empreinte - Zoom sur les réactifs',
            fontsize=16,
            weight='bold',
            fontproperties=self.title_font,
            pad=20
        )

        # Créer la légende à droite
        legend_labels = [f"{label}" for label in labels]
        ax.legend(
            legend_labels,
            loc='center left',
            bbox_to_anchor=(1, 0.5),
            fontsize=12,
            frameon=False
        )

        plt.tight_layout()

        # Sauvegarder dans un buffer
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer
