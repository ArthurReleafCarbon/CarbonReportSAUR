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

    # Tailles de figure centralisées (largeur, hauteur en inches)
    FIGSIZE_PIE = (8, 6)        # Tous les camemberts (scope, postes, contribution, etc.)
    FIGSIZE_BAR = (8, 6)        # Barres horizontales (travaux breakdown)
    FIGSIZE_GROUPED_BAR = (8, 6)   # Barres groupées (inter-lot top3)
    FIGSIZE_DONUT = (8, 6)      # Donuts (réactifs)
    FIGSIZE_TABLE_WIDTH = 12.0  # Largeur du tableau BEGES (hauteur dynamique)
    DPI = 150

    def __init__(self):
        """Initialise le générateur avec les styles par défaut."""
        # Style général
        plt.style.use('default')
        self.colors = ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8', '#E9F7F4']
        self.dpi = self.DPI
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
        self.title_font = fm.FontProperties(family="Poppins", weight="bold", size=14)
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
        elif chart_key == 'BEGES_TABLE':
            return self.generate_beges_table_image(data)
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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_BAR, dpi=self.dpi)

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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)

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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)

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

        # Couleurs pour les scopes (bleu clair, bleu moyen, bleu foncé)
        scope_colors = ['#89CFF0', '#5B9BD5', '#4472C4']

        # Garder tous les scopes pour la légende, mais ne tracer que ceux > 0
        all_scopes = list(scopes)
        all_values = list(values)

        pie_indices = [i for i, v in enumerate(all_values) if v > 0]
        if not pie_indices:
            return None

        pie_values = [all_values[i] for i in pie_indices]
        pie_colors = [scope_colors[i] for i in pie_indices]
        explode = [0.05] * len(pie_values)

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
        wedges, texts, autotexts = ax.pie(
            pie_values,
            labels=None,
            autopct=self._pie_autopct,
            textprops=self._pie_textprops,
            colors=pie_colors,
            startangle=90,
            explode=explode
        )
        for text in autotexts:
            text.set_color("white")

        # Légende : inclure TOUS les scopes (même ceux à 0)
        import matplotlib.patches as mpatches
        legend_handles = [
            mpatches.Patch(
                color=scope_colors[i],
                label=f'{all_scopes[i]} ({all_values[i]:,.0f} tCO₂e)'.replace(',', ' ')
            )
            for i in range(len(all_scopes))
        ]
        ax.legend(
            handles=legend_handles,
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
        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
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
        ax.set_title('Contribution des lots du contrat', fontproperties=self.title_font, pad=20)
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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
        explode = [0.05] * len(values)
        ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=self.colors[:len(values)], startangle=90, explode=explode)
        ax.set_title("Contribution des postes sur l'ensemble du contrat", fontproperties=self.title_font, pad=20)
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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_GROUPED_BAR, dpi=self.dpi)

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
                                   poste_labels: dict = None,
                                   title_override: str = None, **kwargs) -> Optional[BytesIO]:
        """
        Camembert des postes pour LOT × ACTIVITÉ.

        Args:
            emission_result: Résultat d'émissions LOT×ACT
            poste_labels: Dictionnaire {code: label}
            title_override: Titre personnalisé (remplace le titre par défaut)

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

        fig, ax = plt.subplots(figsize=self.FIGSIZE_PIE, dpi=self.dpi)
        explode = [0.05] * len(values)
        wedges, texts, autotexts = ax.pie(values, labels=labels, autopct=self._pie_autopct,
               textprops=self._pie_textprops, colors=self.colors[:len(values)], startangle=90, explode=explode)
        for autotext in autotexts:
            autotext.set_color('white')
        title = title_override or f'Répartition par poste - {emission_result.node_name}'
        ax.set_title(title, fontproperties=self.title_font, pad=20)
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
        fig, ax = plt.subplots(figsize=self.FIGSIZE_DONUT)

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
            textprops={'fontsize': 10, 'weight': 'bold', 'color': 'white'},
            explode=explode,
            pctdistance=0.75  # Positionner les % au milieu de l'anneau (0.75 pour donut width=0.4)
        )

        # Ajouter le titre
        ax.set_title(
            'Répartition de l\'empreinte - Zoom sur les réactifs',
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

    def generate_beges_table_image(self, beges_df: pd.DataFrame) -> Optional[BytesIO]:
        """
        Génère une image de tableau BEGES réglementaire.

        Rendu matplotlib d'un tableau stylisé avec catégories colorées,
        sous-totaux et total. Cohérent avec la charte graphique du rapport.

        Args:
            beges_df: DataFrame avec les colonnes de la feuille BEGES

        Returns:
            BytesIO contenant l'image PNG ou None si données vides
        """
        if beges_df is None or beges_df.empty:
            return None

        # Normaliser les noms de colonnes
        df = beges_df.copy()
        col_map = {}
        for col in df.columns:
            col_lower = col.strip().lower()
            if 'catégorie' in col_lower or 'categorie' in col_lower:
                col_map[col] = 'categorie'
            elif 'numéro' in col_lower or 'numero' in col_lower:
                col_map[col] = 'numero'
            elif 'poste' in col_lower:
                col_map[col] = 'poste'
            elif 'co2' in col_lower:
                col_map[col] = 'co2'
        df = df.rename(columns=col_map)

        # Filtrer les lignes entièrement vides
        rows = []
        for _, row in df.iterrows():
            cat = row.get('categorie', '')
            num = str(row.get('numero', '')).strip()
            poste = row.get('poste', '')
            co2 = row.get('co2', None)

            # Transformer NaN en chaîne vide
            if pd.isna(cat):
                cat = ''
            if pd.isna(num) or num == 'nan':
                num = ''
            if pd.isna(poste):
                poste = ''
            if pd.isna(co2):
                co2_str = ''
            else:
                co2_str = f"{co2:,.1f}".replace(',', ' ')

            # Sauter les lignes complètement vides
            if not cat and not num and not poste and not co2_str:
                continue

            rows.append({
                'categorie': str(cat).strip(),
                'numero': num,
                'poste': str(poste).strip(),
                'co2': co2_str,
            })

        if not rows:
            return None

        # Couleurs
        color_header = '#0B3B2E'
        color_category = '#1A5C4A'
        color_subtotal = '#D5D5D5'
        color_total = '#0B3B2E'
        color_row_even = '#F5F9F7'
        color_row_odd = '#FFFFFF'
        text_white = '#FFFFFF'
        text_dark = '#1A1A1A'
        color_border = '#CCCCCC'

        # Dimensions
        n_rows = len(rows) + 1  # +1 pour l'en-tête
        row_height = 0.45
        fig_height = n_rows * row_height + 1.0
        fig_width = self.FIGSIZE_TABLE_WIDTH

        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=self.dpi)
        ax.set_xlim(0, fig_width)
        ax.set_ylim(0, n_rows)
        ax.axis('off')

        # Largeurs de colonnes (proportions)
        col_widths = [3.8, 1.0, 5.8, 1.4]  # categorie, numero, poste, co2
        col_starts = [0]
        for w in col_widths[:-1]:
            col_starts.append(col_starts[-1] + w)

        headers = ["Catégories d'émissions", "N°", "Postes d'émissions", "CO2 (t CO2e)"]

        # Dessiner l'en-tête
        y = n_rows - 1
        ax.add_patch(plt.Rectangle((0, y), fig_width, 1,
                                    facecolor=color_header, edgecolor='none'))
        for j, header in enumerate(headers):
            ha = 'right' if j == 3 else 'left'
            x_pos = col_starts[j] + (col_widths[j] - 0.1 if j == 3 else 0.15)
            ax.text(x_pos, y + 0.5, header, color=text_white,
                    fontproperties=self.title_font, fontsize=9,
                    ha=ha, va='center')

        # Dessiner les lignes de données
        for i, row_data in enumerate(rows):
            y = n_rows - 2 - i
            is_category = bool(row_data['categorie'])
            is_subtotal = row_data['numero'].lower().startswith('sous')
            is_total = row_data['numero'].upper() == 'TOTAL'

            # Couleur de fond
            if is_total:
                bg_color = color_total
                txt_color = text_white
                font_weight = 'bold'
                font_size = 9
            elif is_category:
                bg_color = color_category
                txt_color = text_white
                font_weight = 'bold'
                font_size = 8.5
            elif is_subtotal:
                bg_color = color_subtotal
                txt_color = text_dark
                font_weight = 'bold'
                font_size = 8.5
            else:
                bg_color = color_row_even if i % 2 == 0 else color_row_odd
                txt_color = text_dark
                font_weight = 'normal'
                font_size = 8

            # Rectangle de fond
            ax.add_patch(plt.Rectangle((0, y), fig_width, 1,
                                        facecolor=bg_color, edgecolor=color_border,
                                        linewidth=0.5))

            # Texte de chaque colonne
            values = [row_data['categorie'], row_data['numero'],
                      row_data['poste'], row_data['co2']]
            for j, val in enumerate(values):
                if not val:
                    continue
                ha = 'right' if j == 3 else 'left'
                x_pos = col_starts[j] + (col_widths[j] - 0.15 if j == 3 else 0.15)
                fp = fm.FontProperties(family="Poppins", weight=font_weight, size=font_size)
                ax.text(x_pos, y + 0.5, val, color=txt_color,
                        fontproperties=fp, ha=ha, va='center')

        # Bordure extérieure
        ax.add_patch(plt.Rectangle((0, 0), fig_width, n_rows,
                                    facecolor='none', edgecolor=color_header,
                                    linewidth=1.5))

        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight',
                    pad_inches=0.1)
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer

    def generate_evitees_table_image(self, evitees_df: pd.DataFrame) -> Optional[BytesIO]:
        """
        Génère une image de tableau des émissions évitées.

        Agrège par typologie et affiche un tableau stylisé avec total.

        Args:
            evitees_df: DataFrame avec colonnes ['node_id', 'typologie', 'tco2e']

        Returns:
            BytesIO contenant l'image PNG ou None si données vides
        """
        if evitees_df is None or evitees_df.empty:
            return None

        # Agréger par typologie
        grouped = evitees_df.groupby('typologie', sort=False)['tco2e'].sum().reset_index()
        grouped = grouped.sort_values('tco2e', ascending=False)

        if grouped.empty:
            return None

        total = grouped['tco2e'].sum()

        # Préparer les lignes : données + total
        rows = []
        for _, row in grouped.iterrows():
            typ = str(row['typologie']).strip() if pd.notna(row['typologie']) else ''
            tco2e_val = row['tco2e']
            rows.append({
                'typologie': typ,
                'tco2e': f"{tco2e_val:,.1f}".replace(',', ' '),
                'is_total': False,
            })
        rows.append({
            'typologie': 'Total émissions évitées',
            'tco2e': f"{total:,.1f}".replace(',', ' '),
            'is_total': True,
        })

        # Couleurs (même charte que BEGES)
        color_header = '#0B3B2E'
        color_total = '#0B3B2E'
        color_row_even = '#F5F9F7'
        color_row_odd = '#FFFFFF'
        text_white = '#FFFFFF'
        text_dark = '#1A1A1A'
        color_border = '#CCCCCC'

        # Dimensions
        n_rows = len(rows) + 1  # +1 pour l'en-tête
        row_height = 0.45
        fig_height = n_rows * row_height + 1.0
        fig_width = 8.0  # Plus étroit que BEGES (2 colonnes seulement)

        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=self.dpi)
        ax.set_xlim(0, fig_width)
        ax.set_ylim(0, n_rows)
        ax.axis('off')

        # Largeurs de colonnes
        col_widths = [5.8, 2.2]  # typologie, tco2e
        col_starts = [0, col_widths[0]]

        headers = ["Typologie", "tCO₂e évitées"]

        # Dessiner l'en-tête
        y = n_rows - 1
        ax.add_patch(plt.Rectangle((0, y), fig_width, 1,
                                    facecolor=color_header, edgecolor='none'))
        for j, header in enumerate(headers):
            ha = 'right' if j == 1 else 'left'
            x_pos = col_starts[j] + (col_widths[j] - 0.15 if j == 1 else 0.15)
            ax.text(x_pos, y + 0.5, header, color=text_white,
                    fontproperties=self.title_font, fontsize=9,
                    ha=ha, va='center')

        # Dessiner les lignes de données
        for i, row_data in enumerate(rows):
            y = n_rows - 2 - i

            if row_data['is_total']:
                bg_color = color_total
                txt_color = text_white
                font_weight = 'bold'
                font_size = 9
            else:
                bg_color = color_row_even if i % 2 == 0 else color_row_odd
                txt_color = text_dark
                font_weight = 'normal'
                font_size = 8.5

            ax.add_patch(plt.Rectangle((0, y), fig_width, 1,
                                        facecolor=bg_color, edgecolor=color_border,
                                        linewidth=0.5))

            values = [row_data['typologie'], row_data['tco2e']]
            for j, val in enumerate(values):
                if not val:
                    continue
                ha = 'right' if j == 1 else 'left'
                x_pos = col_starts[j] + (col_widths[j] - 0.15 if j == 1 else 0.15)
                fp = fm.FontProperties(family="Poppins", weight=font_weight, size=font_size)
                ax.text(x_pos, y + 0.5, val, color=txt_color,
                        fontproperties=fp, ha=ha, va='center')

        # Bordure extérieure
        ax.add_patch(plt.Rectangle((0, 0), fig_width, n_rows,
                                    facecolor='none', edgecolor=color_header,
                                    linewidth=1.5))

        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=self.dpi, bbox_inches='tight',
                    pad_inches=0.1)
        img_buffer.seek(0)
        plt.close(fig)

        return img_buffer
