"""
Application Streamlit V1 - Version simplifiée pour la génération de rapports carbone.
Interface épurée : upload template + excel → génération directe.
"""

import streamlit as st
from pathlib import Path
from io import BytesIO

# Imports des modules
from src.flat_loader import FlatLoader, ExcelValidationError
from src.tree import OrganizationTree
from src.calc_emissions import EmissionCalculator, EmissionOverrides, EmissionResult
from src.calc_indicators import IndicatorCalculator
from src.content_catalog import ContentCatalog
from src.kpi_calculators import KPICalculator
from src.word_renderer import WordRenderer


# Configuration de la page
st.set_page_config(
    page_title="Générateur Rapport Carbone",
    page_icon="🌍",
    layout="centered"
)

# CSS personnalisé pour la charte graphique
st.markdown("""
<style>


    /* Bouton vert personnalisé */
    .stButton > button {
        background-color: #2D5016 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        padding: 12px 30px !important;
        border: none !important;
        font-size: 16px !important;
        width: 100% !important;
    }

    .stButton > button:hover {
        background-color: #1F3810 !important;
        color: white !important;
    }

    /* Bouton de téléchargement */
    .stDownloadButton > button {
        background-color: #2D5016 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        padding: 12px 30px !important;
        border: none !important;
        font-size: 16px !important;
        width: 100% !important;
    }

    .stDownloadButton > button:hover {
        background-color: #1F3810 !important;
        color: white !important;
    }

    /* Message de succès personnalisé */
    .success-message {
        background-color: #D4EDDA;
        color: #155724;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #28A745;
        margin: 20px 0;
        font-weight: bold;
    }

    /* Message d'erreur personnalisé */
    .error-message {
        background-color: #F8D7DA;
        color: #721C24;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #DC3545;
        margin: 20px 0;
        font-weight: bold;
    }

    /* Centrer le contenu */
    .main-container {
        max-width: 1000px;
        margin: 0 auto;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialise les variables de session."""
    if 'template_file' not in st.session_state:
        st.session_state.template_file = None
    if 'excel_file' not in st.session_state:
        st.session_state.excel_file = None
    if 'report_generated' not in st.session_state:
        st.session_state.report_generated = False
    if 'report_data' not in st.session_state:
        st.session_state.report_data = None
    if 'error_message' not in st.session_state:
        st.session_state.error_message = None
    if 'org_name' not in st.session_state:
        st.session_state.org_name = None


def generate_report_v1(template_file, excel_file, annee: int = 2024):
    """
    Génère le rapport Word (version simplifiée V1).

    Args:
        template_file: Fichier template Word uploadé
        excel_file: Fichier Excel uploadé
        annee: Année du bilan

    Returns:
        BytesIO: Document Word généré, ou None si erreur
    """
    try:
        # 1. Sauvegarder temporairement les fichiers
        temp_excel_path = Path("temp_upload.xlsx")
        temp_template_path = Path("temp_template.docx")

        with open(temp_excel_path, "wb") as f:
            f.write(excel_file.getbuffer())

        with open(temp_template_path, "wb") as f:
            f.write(template_file.getbuffer())

        # 2. Charger et valider l'Excel
        loader = FlatLoader(str(temp_excel_path))
        data = loader.load()

        # Vérifier les warnings
        errors, warnings = loader.get_validation_report()
        if errors:
            raise ExcelValidationError(f"Erreurs de validation : {', '.join(errors)}")

        # 3. Construire l'arborescence
        tree = OrganizationTree(data['ORG_TREE'])

        # Valider la structure
        tree_errors = tree.validate_structure()
        if tree_errors:
            raise ValueError(f"Erreurs dans l'arborescence : {', '.join(tree_errors)}")

        # Stocker le nom de l'organisation dans session_state pour le nom du fichier
        org = tree.get_org()
        import streamlit as st
        st.session_state.org_name = org.node_name

        # 4. Créer les calculateurs
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

        # 5. Calculer les résultats (avec overrides auto)
        auto_overrides = loader.get_auto_overrides()
        if auto_overrides.poste_config:
            results_brut = emission_calc.calculate_net(auto_overrides, top_n=4)
        else:
            results_brut = emission_calc.calculate_brut(top_n=4)
        indicator_results = indicator_calc.calculate()

        # 6. Calculer les KPI m³
        kpi_calc = KPICalculator()
        kpi_m3_eu = None
        kpi_m3_aep = None

        # Collecter tous les résultats EU
        eu_results_list = []
        eu_indicators_list = []
        for key, result in results_brut.items():
            if '_EU' in key:
                eu_results_list.append(result)
            if key in indicator_results:
                ind_result = indicator_results[key]
                if ind_result.activity == 'EU':
                    eu_indicators_list.append(ind_result)

        # Collecter tous les résultats AEP
        aep_results_list = []
        aep_indicators_list = []
        for key, result in results_brut.items():
            if '_AEP' in key:
                aep_results_list.append(result)
            if key in indicator_results:
                ind_result = indicator_results[key]
                if ind_result.activity == 'AEP':
                    aep_indicators_list.append(ind_result)

        # Calculer le KPI EU global
        if eu_results_list and eu_indicators_list:
            total_eu_tco2e = sum(r.total_tco2e for r in eu_results_list)
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
            total_aep_tco2e = sum(r.total_tco2e for r in aep_results_list)
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

        # 6b. Texte de comparaison volumes EU/AEP
        eu_first_result = eu_results_list[0] if eu_results_list else None
        aep_first_result = aep_results_list[0] if aep_results_list else None
        eu_ind_first = eu_indicators_list[0] if eu_indicators_list else None
        aep_ind_first = aep_indicators_list[0] if aep_indicators_list else None
        activity_comparison = kpi_calc.generate_activity_volume_comparison_text(
            eu_first_result, aep_first_result, eu_ind_first, aep_ind_first
        )

        # 7. Calculer les données chauffage AEP
        aep_with_chauffage = emission_calc.calculate_aep_with_chauffage()
        chauffage_total = emission_calc.get_chauffage_total()
        org_with_chauffage = emission_calc.calculate_org_with_chauffage()

        # 8. Préparer le contexte pour le rendu
        overrides = EmissionOverrides()  # Vide pour V1

        context = {
            'annee': annee,
            'org_result': results_brut.get('ORG'),
            'lot_results': {k: v for k, v in results_brut.items() if k.startswith('LOT_') or k.startswith('ORG_')},
            'has_lots': tree.has_lots(),
            'poste_labels': emission_calc.poste_labels,
            'top_n': 4,
            'overrides': overrides,
            'emissions_df': data.get('EMISSIONS'),
            'emissions_l2_df': data.get('EMISSIONS_L2'),
            'content_catalog': content_catalog,
            'tree': tree,
            'indicator_results': indicator_results,
            'kpi_m3_eu': kpi_m3_eu,
            'kpi_m3_aep': kpi_m3_aep,
            'aep_with_chauffage_result': aep_with_chauffage,
            'chauffage_total_tco2e': chauffage_total,
            'org_with_chauffage_result': org_with_chauffage,
            'beges_df': data.get('BEGES'),
            'emissions_evitees_df': data.get('EMISSIONS_EVITEES'),
            'activity_volume_comparison_text': activity_comparison,
        }

        # 9. Générer le rapport Word
        renderer = WordRenderer(
            template_path=str(temp_template_path),
            assets_path="assets"
        )

        doc = renderer.render(context)

        # 10. Sauvegarder dans un BytesIO
        output_buffer = BytesIO()
        renderer.doc.save(output_buffer)
        output_buffer.seek(0)

        # 11. Nettoyer les fichiers temporaires
        temp_excel_path.unlink(missing_ok=True)
        temp_template_path.unlink(missing_ok=True)

        return output_buffer

    except Exception as e:
        # Nettoyer les fichiers temporaires en cas d'erreur
        Path("temp_upload.xlsx").unlink(missing_ok=True)
        Path("temp_template.docx").unlink(missing_ok=True)
        raise e


def main():
    """Fonction principale de l'application V1."""
    init_session_state()

    # Titre principal
    st.title("🌍 Générateur de Rapport Carbone pour SAUR")
    st.markdown("---")

    # Conteneur principal
    st.markdown('<div class="main-container">', unsafe_allow_html=True)

    # Année du bilan
    annee = st.number_input(
        "📅 Année du bilan",
        min_value=2000,
        max_value=2100,
        value=2024,
        step=1
    )

    st.markdown("<br>", unsafe_allow_html=True)

    # Colonnes pour les uploads
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown('<div class="upload-container">', unsafe_allow_html=True)
        st.markdown('<div class="upload-title">📄 Template du rapport</div>', unsafe_allow_html=True)

        template_file = st.file_uploader(
            "",
            type=['docx'],
            key="template_uploader",
            label_visibility="collapsed"
        )

        if template_file:
            st.success(f"✅ {template_file.name}")
            st.session_state.template_file = template_file

        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="upload-container">', unsafe_allow_html=True)
        st.markdown('<div class="upload-title">📊 Fichier Excel</div>', unsafe_allow_html=True)

        excel_file = st.file_uploader(
            "",
            type=['xlsx', 'xlsm'],
            key="excel_uploader",
            label_visibility="collapsed"
        )

        if excel_file:
            st.success(f"✅ {excel_file.name}")
            st.session_state.excel_file = excel_file

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)

    # Bouton de génération (toujours visible, désactivé si fichiers manquants)
    can_generate = st.session_state.template_file and st.session_state.excel_file

    if st.button("🚀 Générer le rapport", disabled=not can_generate):
        if can_generate:
            # Réinitialiser les états
            st.session_state.report_generated = False
            st.session_state.report_data = None
            st.session_state.error_message = None

            # Afficher le loader
            with st.spinner("⏳ Génération du rapport en cours..."):
                try:
                    # Générer le rapport
                    report_buffer = generate_report_v1(
                        st.session_state.template_file,
                        st.session_state.excel_file,
                        annee
                    )

                    # Succès
                    st.session_state.report_generated = True
                    st.session_state.report_data = report_buffer
                    st.session_state.error_message = None

                except Exception as e:
                    # Erreur
                    st.session_state.report_generated = False
                    st.session_state.report_data = None
                    st.session_state.error_message = str(e)

            # Forcer le rafraîchissement pour afficher le résultat
            st.rerun()

    # Afficher le message de succès
    if st.session_state.report_generated and st.session_state.report_data:
        st.markdown(
            '<div class="success-message">✅ Rapport généré avec succès</div>',
            unsafe_allow_html=True
        )

        # Bouton de téléchargement
        # Nettoyer le nom de l'organisation pour le nom de fichier
        import re
        org_name = st.session_state.org_name or "Organisation"
        org_name_clean = re.sub(r'[^\w\s-]', '', org_name).strip().replace(' ', '_')

        st.download_button(
            label="📥 Télécharger le rapport",
            data=st.session_state.report_data,
            file_name=f"Rapport Bilan Carbone {org_name_clean} {annee}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Afficher le message d'erreur
    if st.session_state.error_message:
        st.markdown(
            f'<div class="error-message">❌ Erreur dans la génération du rapport<br><small>{st.session_state.error_message}</small></div>',
            unsafe_allow_html=True
        )

    st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #7F8C8D;'>© SAUR - Générateur de Rapport Carbone V1.0</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
