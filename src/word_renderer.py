"""
Module de rendu Word.
Gère le remplacement des placeholders, la duplication des blocs répétables,
l'insertion d'images et tableaux, et le nettoyage des placeholders vides.
"""

import re
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from typing import Dict, List, Optional, Any
from pathlib import Path
from io import BytesIO

from .calc_emissions import EmissionResult, EmissionOverrides
from .calc_indicators import IndicatorResult
from .chart_generators import ChartGenerator
from .table_generators import TableGenerator
from .kpi_calculators import KPICalculator
from .content_catalog import ContentCatalog


class WordRenderer:
    """
    Moteur de rendu Word.
    Remplace les placeholders, duplique les blocs, insère images/tableaux.
    """

    def __init__(self, template_path: str, assets_path: str):
        """
        Initialise le renderer.

        Args:
            template_path: Chemin vers le template Word
            assets_path: Chemin vers le dossier assets
        """
        self.template_path = Path(template_path)
        self.assets_path = Path(assets_path)
        self.doc = None

        # Générateurs
        self.chart_gen = ChartGenerator()
        self.table_gen = TableGenerator()
        self.kpi_calc = KPICalculator()

    def load_template(self):
        """Charge le template Word."""
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template non trouvé : {self.template_path}")
        self.doc = Document(self.template_path)

    def render(self, context: Dict[str, Any]) -> Document:
        """
        Effectue le rendu complet du document.

        Args:
            context: Dictionnaire de contexte contenant toutes les données

        Returns:
            Document Word rendu
        """
        self.load_template()

        # 1. Remplacer les placeholders simples globaux
        self._replace_simple_placeholders(context)

        # 2. Dupliquer et remplir les blocs LOT
        self._process_lot_blocks(context)

        # 3. Insérer les graphiques ORG
        self._insert_org_charts(context)

        # 4. Nettoyer les placeholders vides
        self._clean_empty_placeholders()

        return self.doc

    def _replace_simple_placeholders(self, context: Dict[str, Any]):
        """
        Remplace les placeholders simples de type {{PLACEHOLDER}}.

        Args:
            context: Dictionnaire de contexte
        """
        replacements = self._build_simple_replacements(context)

        # Parcourir tous les paragraphes
        for paragraph in self.doc.paragraphs:
            self._replace_in_paragraph(paragraph, replacements)

        # Parcourir les tableaux
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replacements)

    def _build_simple_replacements(self, context: Dict[str, Any]) -> Dict[str, str]:
        """
        Construit le dictionnaire des remplacements simples.

        Args:
            context: Contexte global

        Returns:
            Dictionnaire {placeholder: valeur}
        """
        replacements = {}

        # Année
        replacements['{{annee}}'] = str(context.get('annee', '2024'))

        # ORG - AVEC FORMAT ESPACE ET ARRONDI
        org_result = context.get('org_result')
        if org_result:
            replacements['{{ORG_NAME}}'] = org_result.node_name
            replacements['{{TOTAL_EMISSIONS}}'] = self.kpi_calc.format_number(org_result.total_tco2e)
            replacements['{{TOTAL_EMISSIONS_S1}}'] = self.kpi_calc.format_number(org_result.scope1_tco2e)
            replacements['{{TOTAL_EMISSIONS_S2}}'] = self.kpi_calc.format_number(org_result.scope2_tco2e)
            replacements['{{TOTAL_EMISSIONS_S3}}'] = self.kpi_calc.format_number(org_result.scope3_tco2e)
            replacements['{{pourc_s3_org}}'] = self.kpi_calc.format_number(org_result.get_scope_percentage(3))

        # KPI m³ - NOUVEAUX NOMS
        kpi_m3_eu = context.get('kpi_m3_eu')
        kpi_m3_aep = context.get('kpi_m3_aep')
        if kpi_m3_eu is not None:
            replacements['{{kpi_M3_EU}}'] = f"{kpi_m3_eu:.2f} kgCO₂e/m³".replace(".", ",")
        else:
            replacements['{{kpi_M3_EU}}'] = "N/A"

        if kpi_m3_aep is not None:
            replacements['{{kpi_M3_AEP}}'] = f"{kpi_m3_aep:.2f} kgCO₂e/m³".replace(".", ",")
        else:
            replacements['{{kpi_M3_AEP}}'] = "N/A"

        # Équivalents - AVEC FORMAT ESPACE
        if org_result:
            flights = self.kpi_calc.calculate_flight_equivalent(org_result.total_tco2e)
            persons = self.kpi_calc.calculate_person_equivalent(org_result.total_tco2e)
            replacements['{{kpi_1}}'] = self.kpi_calc.format_number(flights)
            replacements['{{kpi_2}}'] = self.kpi_calc.format_number(persons)

        # Top postes ORG
        if org_result and org_result.top_postes:
            top_3 = org_result.top_postes[:3]
            poste_labels = context.get('poste_labels', {})
            for i, (code, _) in enumerate(top_3, 1):
                label = poste_labels.get(code, code)
                replacements[f'{{{{TOP_POSTE_{i}}}}}'] = label

        # Texte de comparaison volumes
        comp_text = context.get('activity_volume_comparison_text', '')
        replacements['{{ACTIVITY_VOLUME_COMPARISON_TEXT}}'] = comp_text

        # Top postes longueur
        top_n = context.get('top_n', 4)
        replacements['{{TOP_POSTES_LONGUEUR}}'] = str(top_n)

        return replacements

    def _replace_in_paragraph(self, paragraph, replacements: Dict[str, str]):
        """
        Remplace les placeholders dans un paragraphe.

        Args:
            paragraph: Paragraphe Word
            replacements: Dictionnaire de remplacements
        """
        # Utiliser le texte complet du paragraphe pour gérer les placeholders fragmentés
        full_text = paragraph.text

        # Vérifier s'il y a des placeholders
        if '{{' not in full_text:
            return

        # Remplacer les placeholders
        new_text = full_text
        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, value)

        # Si le texte a changé, le mettre à jour
        if new_text != full_text:
            # Conserver le style du premier run
            if paragraph.runs:
                style = paragraph.runs[0].style
                font = paragraph.runs[0].font
                # Supprimer tous les runs
                for run in paragraph.runs:
                    run.text = ''
                # Ajouter le nouveau texte
                new_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                new_run.text = new_text
                if style:
                    new_run.style = style
            else:
                paragraph.text = new_text

    def _process_lot_blocks(self, context: Dict[str, Any]):
        """
        Traite les blocs LOT répétables.

        Args:
            context: Contexte global
        """
        # Récupérer les résultats LOT
        lot_results = context.get('lot_results', {})
        has_lots = context.get('has_lots', False)

        if not has_lots:
            # Cas sans LOT : traiter les blocs ACTIVITY au niveau ORG
            self._process_org_activity_blocks(context)
            return

        # Trouver les blocs [[START_LOT]] ... [[END_LOT]]
        # (Implémentation simplifiée - à compléter avec la logique de duplication)
        # Pour le MVP, on va traiter bloc par bloc

        # TODO: Implémenter la duplication des blocs LOT
        # Pour l'instant, on remplace juste les placeholders LOT dans le document

    def _process_org_activity_blocks(self, context: Dict[str, Any]):
        """
        Traite les blocs ACTIVITY au niveau ORG (cas sans LOT).

        Args:
            context: Contexte global
        """
        # TODO: Implémenter la duplication des blocs ACTIVITY au niveau ORG
        pass

    def _insert_org_charts(self, context: Dict[str, Any]):
        """
        Insère les graphiques au niveau ORG.

        Args:
            context: Contexte global
        """
        org_result = context.get('org_result')
        if not org_result:
            return

        poste_labels = context.get('poste_labels', {})
        tree = context.get('tree')

        # 1. Chart emissions scope (camembert scopes)
        charts = {
            '{{chart_emissions_scope_org}}': self.chart_gen.generate_scope_pie(org_result),
        }

        # 2. Chart emissions total (PIE des postes L1) - CORRIGÉ
        charts['{{chart_emissions_total_org}}'] = self.chart_gen.generate_total_emissions_pie(
            org_result, poste_labels=poste_labels
        )

        # 3. Chart contribution LOT (bar chart)
        lot_results = context.get('lot_results', {})
        if lot_results and tree and tree.has_lots():
            # Pour chaque LOT, sommer EU + AEP
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
                charts['{{chart_contrib_lot}}'] = self.chart_gen.generate_lot_contribution(lot_data)

        # 4. Chart électricité par activité (PIE EU vs AEP) - NOUVEAU
        # Chercher le poste "électricité" dans les données
        elec_by_activity = {}
        for key, result in lot_results.items():
            if result.activity in ['EU', 'AEP']:
                # Chercher le poste électricité
                for poste_code, tco2e in result.emissions_by_poste.items():
                    # Adapter selon le nom exact du poste électricité dans ton Excel
                    if 'ELEC' in poste_code.upper() or 'ELECTRICITE' in poste_labels.get(poste_code, '').upper():
                        if result.activity not in elec_by_activity:
                            elec_by_activity[result.activity] = 0.0
                        elec_by_activity[result.activity] += tco2e

        if elec_by_activity:
            charts['{{chart_emissions_elec_org}}'] = self.chart_gen.generate_elec_emissions(elec_by_activity)

        # 5. Top 3 inter-lot (bar chart)
        if org_result.top_postes:
            top_data = [(poste_labels.get(code, code), value)
                       for code, value in org_result.top_postes[:3]]
            charts['{{chart_batonnet_inter_lot_top3}}'] = self.chart_gen.generate_inter_lot_top3(top_data)

        # Insérer les graphiques dans le document
        for placeholder, img_buffer in charts.items():
            if img_buffer:
                self._insert_image(placeholder, img_buffer)

    def _insert_image(self, placeholder: str, img_buffer: BytesIO, width: float = 5.0):
        """
        Insère une image à la place d'un placeholder.

        Args:
            placeholder: Placeholder à remplacer
            img_buffer: Buffer contenant l'image
            width: Largeur en inches
        """
        for paragraph in self.doc.paragraphs:
            if placeholder in paragraph.text:
                # Supprimer le placeholder
                paragraph.clear()
                # Insérer l'image
                run = paragraph.add_run()
                run.add_picture(img_buffer, width=Inches(width))
                return

        # Chercher dans les tableaux
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if placeholder in paragraph.text:
                            paragraph.clear()
                            run = paragraph.add_run()
                            run.add_picture(img_buffer, width=Inches(width))
                            return

    def _clean_empty_placeholders(self):
        """Supprime les paragraphes contenant des placeholders non remplacés."""
        paragraphs_to_remove = []

        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            # Si le paragraphe contient uniquement un placeholder {{...}}
            if re.match(r'^\{\{[A-Z_0-9]+\}\}$', text):
                paragraphs_to_remove.append(paragraph)

        # Supprimer les paragraphes
        for paragraph in paragraphs_to_remove:
            self._delete_paragraph(paragraph)

    def _delete_paragraph(self, paragraph):
        """Supprime un paragraphe du document."""
        p = paragraph._element
        p.getparent().remove(p)

    def save(self, output_path: str):
        """
        Sauvegarde le document rendu.

        Args:
            output_path: Chemin de sortie
        """
        if self.doc is None:
            raise ValueError("Document non chargé. Appelez render() d'abord.")

        self.doc.save(output_path)
