"""
Page Streamlit pour la prévisualisation et personnalisation des graphiques.
"""

import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO
from typing import Dict, Any, Optional
import pandas as pd

from .chart_generators import ChartGenerator
from .calc_emissions import EmissionResult


def init_chart_customization():
    """Initialise les paramètres de personnalisation des graphiques dans session_state."""
    if 'chart_customization' not in st.session_state:
        st.session_state.chart_customization = {
            'FILE_EAU_BREAKDOWN': {
                'title': 'Répartition des émissions - File eau STEP',
                'colors': ['#2E86AB', '#A23B72', '#F18F01'],
                'show_legend': True
            },
            'EM_INDIRECTES_SPLIT': {
                'title': 'Répartition des émissions indirectes',
                'colors': ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8'],
                'show_legend': True
            },
            'chart_emissions_scope_org': {
                'title': 'Répartition par scope - {org_name}',
                'colors': ['#89CFF0', '#5B9BD5', '#4472C4'],
                'show_legend': True
            },
            'chart_contrib_lot': {
                'title': 'Contribution des LOTs - {org_name}',
                'colors': ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8'],
                'show_legend': True
            },
            'chart_emissions_total_org': {
                'title': 'Contribution des postes - ORG',
                'colors': ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8'],
                'show_legend': False
            },
            'chart_emissions_elec_org': {
                'title': 'Répartition émissions Électricité par activité',
                'colors': ['#2E86AB', '#A23B72'],
                'show_legend': False
            }
        }


def get_chart_customization(chart_key: str) -> Dict[str, Any]:
    """
    Récupère les paramètres de personnalisation pour un graphique.

    Args:
        chart_key: Clé du graphique

    Returns:
        Dictionnaire avec les paramètres de personnalisation
    """
    if 'chart_customization' not in st.session_state:
        init_chart_customization()

    return st.session_state.chart_customization.get(chart_key, {})


def update_chart_customization(chart_key: str, params: Dict[str, Any]):
    """
    Met à jour les paramètres de personnalisation pour un graphique.

    Args:
        chart_key: Clé du graphique
        params: Nouveaux paramètres
    """
    if 'chart_customization' not in st.session_state:
        init_chart_customization()

    if chart_key not in st.session_state.chart_customization:
        st.session_state.chart_customization[chart_key] = {}

    st.session_state.chart_customization[chart_key].update(params)


def display_chart_preview(chart_key: str, chart_data: Any, chart_gen: ChartGenerator,
                          org_name: str = "ORG", **kwargs):
    """
    Affiche la prévisualisation d'un graphique avec options de personnalisation.

    Args:
        chart_key: Clé du graphique
        chart_data: Données du graphique
        chart_gen: Générateur de graphiques
        org_name: Nom de l'organisation
        **kwargs: Arguments supplémentaires pour la génération
    """
    custom = get_chart_customization(chart_key)

    with st.expander(f"📊 {chart_key}", expanded=False):
        col1, col2 = st.columns([2, 1])

        with col1:
            # Prévisualisation du graphique
            try:
                # Générer le graphique avec les paramètres actuels
                kwargs['org_name'] = org_name
                img_buffer = chart_gen.generate_chart(chart_key, chart_data, **kwargs)

                if img_buffer:
                    st.image(img_buffer, use_container_width=True)
                else:
                    st.warning("Pas de données disponibles pour ce graphique")
            except Exception as e:
                st.error(f"Erreur lors de la génération : {str(e)}")

        with col2:
            st.subheader("Personnalisation")

            # Titre
            new_title = st.text_input(
                "Titre",
                value=custom.get('title', ''),
                key=f"title_{chart_key}"
            )

            # Légende
            show_legend = st.checkbox(
                "Afficher la légende",
                value=custom.get('show_legend', True),
                key=f"legend_{chart_key}"
            )

            # Couleurs (simple pour l'instant)
            st.write("**Couleurs** (prédéfinies)")
            color_preset = st.selectbox(
                "Palette",
                ["Verte (défaut)", "Bleue", "Rouge/Orange", "Personnalisée"],
                key=f"colors_{chart_key}"
            )

            # Bouton pour appliquer les changements
            if st.button("💾 Appliquer", key=f"apply_{chart_key}"):
                colors = custom.get('colors', [])
                if color_preset == "Bleue":
                    colors = ['#89CFF0', '#5B9BD5', '#4472C4', '#2E5090', '#1C3D70']
                elif color_preset == "Rouge/Orange":
                    colors = ['#FF6B6B', '#FFA07A', '#FF8C42', '#D9534F', '#C9302C']
                elif color_preset == "Verte (défaut)":
                    colors = ['#0B3B2E', '#3F9B83', '#62CC7B', '#8AD2C5', '#CDEFE8']

                update_chart_customization(chart_key, {
                    'title': new_title,
                    'colors': colors,
                    'show_legend': show_legend
                })
                st.success("✅ Paramètres enregistrés")
                st.rerun()


def display_charts_page():
    """Affiche la page de prévisualisation et personnalisation des graphiques."""
    st.title("🎨 Prévisualisation et Personnalisation des Graphiques")

    # Vérifier que les données sont chargées
    if not st.session_state.get('data_loaded', False):
        st.warning("⚠️ Veuillez d'abord charger un fichier Excel dans la section principale")
        return

    # Initialiser les personnalisations
    init_chart_customization()

    st.markdown("""
    Cette page vous permet de prévisualiser tous les graphiques qui seront générés dans le rapport
    et de les personnaliser selon vos besoins.

    **Modifications disponibles :**
    - 📝 Titres des graphiques
    - 🎨 Palettes de couleurs
    - 📊 Affichage de la légende
    """)

    st.markdown("---")

    # Récupérer les données nécessaires
    org_result = st.session_state.get('results_net', {}).get('ORG')
    lot_results = st.session_state.get('results_net', {})
    tree = st.session_state.get('tree')
    emissions_df = st.session_state.get('excel_data', {}).get('emissions')

    if not org_result:
        st.info("ℹ️ Calculez d'abord les émissions dans la section Aperçu")
        return

    # Créer le générateur de graphiques
    chart_gen = ChartGenerator()
    org_name = org_result.node_name if org_result else "ORG"

    # Onglets pour organiser les graphiques
    tab1, tab2, tab3 = st.tabs(["📊 Graphiques Globaux", "🏢 Graphiques LOT", "⚙️ Autres"])

    with tab1:
        st.subheader("Graphiques au niveau Organisation")

        # Graphique des scopes
        if org_result:
            display_chart_preview(
                'chart_emissions_scope_org',
                org_result,
                chart_gen,
                org_name=org_name
            )

        # Graphique de contribution des postes
        if org_result:
            poste_labels = st.session_state.get('poste_labels', {})
            display_chart_preview(
                'chart_emissions_total_org',
                org_result,
                chart_gen,
                org_name=org_name,
                poste_labels=poste_labels
            )

    with tab2:
        st.subheader("Graphiques au niveau LOT")

        # Graphique de contribution des LOTs
        if tree and tree.has_lots():
            lot_totals = {}
            for lot in tree.get_lots():
                total = 0.0
                for activity in ['EU', 'AEP']:
                    key = f"LOT_{lot.node_id}_{activity}"
                    if key in lot_results:
                        total += lot_results[key].total_tco2e
                if total > 0:
                    lot_totals[lot.node_name] = total

            if lot_totals:
                lot_data = list(lot_totals.items())
                display_chart_preview(
                    'chart_contrib_lot',
                    lot_data,
                    chart_gen,
                    org_name=org_name
                )

    with tab3:
        st.subheader("Autres graphiques")

        # File eau breakdown
        if emissions_df is not None and not emissions_df.empty:
            file_eau_data = emissions_df[
                emissions_df['poste_l2'].str.contains('N2O|CH4|N-N2O', case=False, na=False)
            ].groupby('poste_l2')['tco2e'].sum().reset_index()

            if not file_eau_data.empty:
                display_chart_preview(
                    'FILE_EAU_BREAKDOWN',
                    file_eau_data,
                    chart_gen,
                    org_name=org_name
                )

        # Émissions indirectes
        if emissions_df is not None and not emissions_df.empty:
            indirect_data = emissions_df[
                emissions_df['scope'] == 3
            ].groupby('poste_l2')['tco2e'].sum().reset_index()

            if not indirect_data.empty:
                display_chart_preview(
                    'EM_INDIRECTES_SPLIT',
                    indirect_data,
                    chart_gen,
                    org_name=org_name
                )

    st.markdown("---")
    st.info("💡 **Astuce :** Les personnalisations seront automatiquement appliquées lors de la génération du rapport.")
