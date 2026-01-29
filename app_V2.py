"""
Application Streamlit pour la génération de rapports carbone.
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
from typing import Dict, Optional

# Imports des modules
from src.excel_loader import ExcelLoader, ExcelValidationError
from src.flat_loader import FlatLoader
from src.tree import OrganizationTree
from src.calc_emissions import EmissionCalculator, EmissionOverrides
from src.calc_indicators import IndicatorCalculator
from src.content_catalog import ContentCatalog
from src.kpi_calculators import KPICalculator
from src.word_renderer import WordRenderer
from src.streamlit_charts_page import display_charts_page, init_chart_customization


# Configuration de la page
st.set_page_config(
    page_title="Générateur Rapport Carbone",
    page_icon="🌍",
    layout="wide"
)


def init_session_state():
    """Initialise les variables de session."""
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'tree' not in st.session_state:
        st.session_state.tree = None
    if 'emission_calc' not in st.session_state:
        st.session_state.emission_calc = None
    if 'indicator_calc' not in st.session_state:
        st.session_state.indicator_calc = None
    if 'content_catalog' not in st.session_state:
        st.session_state.content_catalog = None
    if 'overrides' not in st.session_state:
        st.session_state.overrides = EmissionOverrides()
    if 'results_brut' not in st.session_state:
        st.session_state.results_brut = None
    if 'results_net' not in st.session_state:
        st.session_state.results_net = None
    if 'poste_labels' not in st.session_state:
        st.session_state.poste_labels = {}

    # Initialiser les personnalisations de graphiques
    init_chart_customization()


def load_excel_file(uploaded_file) -> bool:
    """
    Charge et valide un fichier Excel.

    Args:
        uploaded_file: Fichier uploadé via Streamlit

    Returns:
        True si succès, False sinon
    """
    try:
        with st.spinner("Chargement et validation du fichier Excel..."):
            # Sauvegarder temporairement le fichier
            temp_path = Path("temp_upload.xlsx")
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Auto-détection du format : simplifié (DATA) ou standard (9 onglets)
            excel_file = pd.ExcelFile(str(temp_path))
            if 'DATA' in excel_file.sheet_names:
                loader = FlatLoader(str(temp_path))
            else:
                loader = ExcelLoader(str(temp_path))
            data = loader.load()

            # Récupérer warnings
            errors, warnings = loader.get_validation_report()

            if warnings:
                st.warning("Avertissements de validation :")
                for warning in warnings:
                    st.warning(f"⚠️ {warning}")

            # Construire l'arborescence
            tree = OrganizationTree(data['ORG_TREE'])

            # Valider la structure
            tree_errors = tree.validate_structure()
            if tree_errors:
                st.error("Erreurs dans la structure de l'arborescence :")
                for error in tree_errors:
                    st.error(f"❌ {error}")
                return False

            # Créer les calculateurs
            emission_calc = EmissionCalculator(
                tree,
                data['EMISSIONS'],
                data['POSTES_REF']
            )

            indicator_calc = IndicatorCalculator(
                tree,
                data['INDICATORS'],
                data['INDICATORS_REF']
            )

            content_catalog = ContentCatalog(data['TEXTE_RAPPORT'])

            # Stocker dans session_state
            st.session_state.excel_data = data
            st.session_state.tree = tree
            st.session_state.emission_calc = emission_calc
            st.session_state.indicator_calc = indicator_calc
            st.session_state.content_catalog = content_catalog
            st.session_state.poste_labels = emission_calc.poste_labels
            st.session_state.data_loaded = True

            # Calculer les résultats BRUT
            st.session_state.results_brut = emission_calc.calculate_brut(top_n=4)
            st.session_state.indicator_results = indicator_calc.calculate()

            # Supprimer le fichier temporaire
            temp_path.unlink()

            st.success("✅ Fichier chargé et validé avec succès !")
            return True

    except ExcelValidationError as e:
        st.error(f"❌ Erreur de validation : {str(e)}")
        return False
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement : {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return False


def display_tree():
    """Affiche l'arborescence de l'organisation."""
    st.subheader("📊 Arborescence de l'organisation")

    tree = st.session_state.tree
    org = tree.get_org()

    st.text(tree.print_tree())

    # Statistiques
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Nombre de LOTs", len(tree.get_lots()))
    with col2:
        st.metric("Nombre d'ENTs", len(tree.get_ents()))
    with col3:
        activities = tree.get_org_activities()
        st.metric("Activités", ", ".join(sorted(activities)))


def display_results():
    """Affiche les résultats calculés."""
    st.subheader("📈 Résultats des émissions")

    results = st.session_state.results_brut
    if not results:
        st.warning("Aucun résultat calculé")
        return

    # Résultat ORG
    if 'ORG' in results:
        org_result = results['ORG']
        st.markdown("### Organisation globale")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total", f"{org_result.total_tco2e:.1f} tCO₂e")
        with col2:
            st.metric("Scope 1", f"{org_result.scope1_tco2e:.1f} tCO₂e")
        with col3:
            st.metric("Scope 2", f"{org_result.scope2_tco2e:.1f} tCO₂e")
        with col4:
            st.metric("Scope 3", f"{org_result.scope3_tco2e:.1f} tCO₂e")

        # Top postes
        if org_result.top_postes:
            st.markdown("**Top postes émetteurs :**")
            emission_calc = st.session_state.emission_calc
            for i, (code, value) in enumerate(org_result.top_postes, 1):
                label = emission_calc.get_poste_label(code)
                st.write(f"{i}. {label}: {value:.1f} tCO₂e")

    # Résultats LOT × ACTIVITÉ
    tree = st.session_state.tree
    if tree.has_lots():
        st.markdown("### Résultats par LOT et ACTIVITÉ")

        for lot in tree.get_lots():
            with st.expander(f"📦 {lot.node_name}"):
                activities = tree.get_lot_activities(lot.node_id)

                for activity in sorted(activities):
                    key = f"LOT_{lot.node_id}_{activity}"
                    if key in results:
                        result = results[key]
                        st.markdown(f"**Activité {activity}**")

                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Total", f"{result.total_tco2e:.1f} tCO₂e")
                        with col2:
                            st.metric("Scope 1", f"{result.scope1_tco2e:.1f} tCO₂e")
                        with col3:
                            st.metric("Scope 2", f"{result.scope2_tco2e:.1f} tCO₂e")
                        with col4:
                            st.metric("Scope 3", f"{result.scope3_tco2e:.1f} tCO₂e")


def display_overrides_ui():
    """Affiche l'interface de gestion des overrides."""
    st.subheader("⚙️ Configuration du rapport")

    tree = st.session_state.tree
    emission_calc = st.session_state.emission_calc
    overrides = st.session_state.overrides

    # Renommage des nœuds
    st.markdown("### Renommer les nœuds")

    org = tree.get_org()
    new_org_name = st.text_input(
        f"Nom de l'organisation",
        value=overrides.get_node_name(org.node_id, org.node_name),
        key=f"rename_{org.node_id}"
    )
    if new_org_name != org.node_name:
        overrides.node_renames[org.node_id] = new_org_name

    if tree.has_lots():
        st.markdown("**LOTs :**")
        for lot in tree.get_lots():
            new_lot_name = st.text_input(
                f"LOT: {lot.node_name}",
                value=overrides.get_node_name(lot.node_id, lot.node_name),
                key=f"rename_{lot.node_id}"
            )
            if new_lot_name != lot.node_name:
                overrides.node_renames[lot.node_id] = new_lot_name

    # Gestion des postes L1
    st.markdown("### Gestion des postes émetteurs")

    st.info("""
    **Mode A** : Masquer mais garder dans les totaux (show=False, include=True)
    **Mode B** : Masquer et exclure des totaux (show=False, include=False)
    """)

    # Récupérer tous les postes L1 uniques
    all_postes = set()
    for _, row in st.session_state.excel_data['EMISSIONS'].iterrows():
        all_postes.add(row['poste_l1_code'])

    for poste_code in sorted(all_postes):
        label = emission_calc.get_poste_label(poste_code)

        col1, col2, col3 = st.columns([3, 1, 1])

        with col1:
            st.write(f"**{label}** ({poste_code})")

        with col2:
            show = st.checkbox(
                "Afficher",
                value=overrides.is_poste_shown(poste_code),
                key=f"show_{poste_code}"
            )

        with col3:
            include = st.checkbox(
                "Inclure totaux",
                value=overrides.is_poste_included(poste_code),
                key=f"include_{poste_code}"
            )

        # Mettre à jour les overrides
        overrides.set_poste_config(poste_code, show, include)

    # Recalculer les résultats NET
    if st.button("🔄 Recalculer avec les modifications"):
        with st.spinner("Recalcul en cours..."):
            st.session_state.results_net = emission_calc.calculate_net(overrides, top_n=4)
            st.success("✅ Résultats recalculés !")
            st.rerun()


def display_export_import():
    """Affiche l'interface d'export/import des overrides."""
    st.subheader("💾 Export / Import overrides")

    overrides = st.session_state.overrides

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Export")
        if st.button("📥 Exporter les overrides"):
            # Créer le JSON
            export_data = {
                'node_renames': overrides.node_renames,
                'poste_config': overrides.poste_config
            }

            json_str = json.dumps(export_data, indent=2, ensure_ascii=False)

            st.download_button(
                label="💾 Télécharger overrides.json",
                data=json_str,
                file_name="overrides.json",
                mime="application/json"
            )

    with col2:
        st.markdown("#### Import")
        uploaded_json = st.file_uploader("📤 Importer overrides.json", type=['json'])

        if uploaded_json is not None:
            try:
                import_data = json.load(uploaded_json)

                # Charger les overrides
                overrides.node_renames = import_data.get('node_renames', {})
                overrides.poste_config = import_data.get('poste_config', {})

                st.success("✅ Overrides importés avec succès !")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Erreur lors de l'import : {str(e)}")


def generate_report():
    """Génère le rapport Word."""
    st.subheader("📄 Génération du rapport")

    # Vérifier que le template existe
    template_path = Path("templates/rapport_template.docx")
    if not template_path.exists():
        st.error("❌ Template non trouvé : templates/rapport_template.docx")
        st.info("Veuillez placer votre template Word dans le dossier templates/")
        return

    # Paramètres
    annee = st.number_input("Année du bilan", min_value=2000, max_value=2100, value=2024)

    if st.button("🚀 Générer le rapport"):
        with st.spinner("Génération du rapport en cours..."):
            try:
                # Préparer le contexte
                results = st.session_state.results_net or st.session_state.results_brut
                overrides = st.session_state.overrides

                # Calculer les KPI m³
                kpi_calc = KPICalculator()
                kpi_m3_eu = None
                kpi_m3_aep = None

                # Filtrer les résultats et indicateurs par activité
                tree = st.session_state.tree
                indicator_results_dict = st.session_state.indicator_results

                # Collecter tous les résultats EU
                eu_results_list = []
                eu_indicators_list = []
                for key, result in results.items():
                    if '_EU' in key:
                        eu_results_list.append(result)
                    if key in indicator_results_dict:
                        ind_result = indicator_results_dict[key]
                        if ind_result.activity == 'EU':
                            eu_indicators_list.append(ind_result)

                # Collecter tous les résultats AEP
                aep_results_list = []
                aep_indicators_list = []
                for key, result in results.items():
                    if '_AEP' in key:
                        aep_results_list.append(result)
                    if key in indicator_results_dict:
                        ind_result = indicator_results_dict[key]
                        if ind_result.activity == 'AEP':
                            aep_indicators_list.append(ind_result)

                # Calculer le KPI EU global
                if eu_results_list and eu_indicators_list:
                    # Sommer toutes les émissions EU
                    total_eu_tco2e = sum(r.total_tco2e for r in eu_results_list)
                    # Créer un EmissionResult fictif avec le total
                    from src.calc_emissions import EmissionResult
                    eu_total_result = EmissionResult(
                        node_id='ORG',
                        node_name='ORG',
                        activity='EU',
                        total_tco2e=total_eu_tco2e,
                        scope1_tco2e=sum(r.scope1_tco2e for r in eu_results_list),
                        scope2_tco2e=sum(r.scope2_tco2e for r in eu_results_list),
                        scope3_tco2e=sum(r.scope3_tco2e for r in eu_results_list),
                        emissions_by_poste={},
                        top_postes=[],
                        other_postes=[]
                    )
                    kpi_m3_eu = kpi_calc.calculate_kpi_m3_eu(eu_total_result, eu_indicators_list)

                # Calculer le KPI AEP global
                if aep_results_list and aep_indicators_list:
                    # Sommer toutes les émissions AEP
                    total_aep_tco2e = sum(r.total_tco2e for r in aep_results_list)
                    # Créer un EmissionResult fictif avec le total
                    from src.calc_emissions import EmissionResult
                    aep_total_result = EmissionResult(
                        node_id='ORG',
                        node_name='ORG',
                        activity='AEP',
                        total_tco2e=total_aep_tco2e,
                        scope1_tco2e=sum(r.scope1_tco2e for r in aep_results_list),
                        scope2_tco2e=sum(r.scope2_tco2e for r in aep_results_list),
                        scope3_tco2e=sum(r.scope3_tco2e for r in aep_results_list),
                        emissions_by_poste={},
                        top_postes=[],
                        other_postes=[]
                    )
                    kpi_m3_aep = kpi_calc.calculate_kpi_m3_aep(aep_total_result, aep_indicators_list)

                context = {
                    'annee': annee,
                    'org_result': results.get('ORG'),
                    'lot_results': {k: v for k, v in results.items() if k.startswith('LOT_')},
                    'has_lots': st.session_state.tree.has_lots(),
                    'poste_labels': st.session_state.emission_calc.poste_labels,
                    'top_n': 4,
                    'overrides': overrides,
                    'emissions_df': st.session_state.excel_data.get('EMISSIONS'),
                    'emissions_l2_df': st.session_state.excel_data.get('EMISSIONS_L2'),
                    'content_catalog': st.session_state.content_catalog,
                    'tree': st.session_state.tree,
                    'indicator_results': st.session_state.indicator_results,
                    'kpi_m3_eu': kpi_m3_eu,
                    'kpi_m3_aep': kpi_m3_aep
                }

                # Générer le rapport
                renderer = WordRenderer(
                    template_path=str(template_path),
                    assets_path="assets"
                )

                doc = renderer.render(context)

                # Récupérer le nom de l'organisation (avec renommage si applicable)
                org = st.session_state.tree.get_org()
                org_name = overrides.get_node_name(org.node_id, org.node_name)

                # Nettoyer le nom de l'organisation pour le nom de fichier (enlever caractères spéciaux)
                import re
                org_name_clean = re.sub(r'[^\w\s-]', '', org_name).strip().replace(' ', '_')

                # Sauvegarder
                output_path = Path("output/rapport_genere.docx")
                output_path.parent.mkdir(exist_ok=True)
                renderer.save(str(output_path))

                st.success("✅ Rapport généré avec succès !")

                # Proposer le téléchargement avec le nom formaté
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 Télécharger le rapport",
                        data=f,
                        file_name=f"Rapport Bilan Carbone {org_name_clean} {annee}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"❌ Erreur lors de la génération : {str(e)}")
                import traceback
                st.error(traceback.format_exc())


def main():
    """Fonction principale de l'application."""
    init_session_state()

    st.title("🌍 Générateur de Rapport Carbone")
    st.markdown("Application de génération automatique de rapports de bilan carbone")

    # Sidebar
    with st.sidebar:
        st.header("Navigation")

        # Upload fichier Excel
        st.subheader("1️⃣ Charger les données")
        uploaded_file = st.file_uploader(
            "Fichier Excel",
            type=['xlsx'],
            help="Fichier Excel au format standard"
        )

        if uploaded_file is not None and not st.session_state.data_loaded:
            if load_excel_file(uploaded_file):
                st.rerun()

        if st.session_state.data_loaded:
            st.success("✅ Données chargées")

            # Menu de navigation
            page = st.radio(
                "Sections",
                ["📊 Aperçu", "⚙️ Configuration", "🎨 Graphiques", "💾 Export/Import", "📄 Génération"]
            )
        else:
            page = None
            st.info("👆 Chargez un fichier Excel pour commencer")

    # Contenu principal
    if page == "📊 Aperçu":
        display_tree()
        st.markdown("---")
        display_results()

    elif page == "⚙️ Configuration":
        display_overrides_ui()

    elif page == "🎨 Graphiques":
        display_charts_page()

    elif page == "💾 Export/Import":
        display_export_import()

    elif page == "📄 Génération":
        generate_report()


if __name__ == "__main__":
    main()
