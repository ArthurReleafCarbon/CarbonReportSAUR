"""
Module de génération de tableaux Word.
Supporte tous les TABLE_KEY définis dans le brief.
"""

import pandas as pd
from docx.table import Table
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Optional, List


class TableGenerator:
    """Générateur de tableaux pour le rapport Word."""

    def __init__(self):
        """Initialise le générateur."""
        pass

    def generate_table(self, table_key: str, data: pd.DataFrame, doc_table: Table) -> bool:
        """
        Remplit un tableau Word selon la key spécifiée.

        Args:
            table_key: Clé du tableau à générer
            data: Données pour le tableau
            doc_table: Objet Table de python-docx à remplir

        Returns:
            True si le tableau a été généré, False sinon
        """
        if table_key == 'EM_INDIRECTES_TABLE':
            return self.generate_em_indirectes_table(data, doc_table)
        else:
            # Key non supportée
            return False

    def generate_em_indirectes_table(self, data: pd.DataFrame, doc_table: Table) -> bool:
        """
        Génère le tableau EM_INDIRECTES_TABLE.

        Args:
            data: DataFrame avec colonnes ['poste_l2', 'tco2e']
            doc_table: Table Word à remplir

        Returns:
            True si succès
        """
        if data.empty:
            return False

        # S'assurer que la table a au moins 2 colonnes
        if len(doc_table.columns) < 2:
            return False

        # Header (si pas déjà dans le template)
        if len(doc_table.rows) == 0:
            header_row = doc_table.add_row()
            header_row.cells[0].text = "Sous-catégorie"
            header_row.cells[1].text = "tCO₂e"
            self._style_header_row(header_row)

        # Ajouter les données
        for _, row in data.iterrows():
            data_row = doc_table.add_row()
            data_row.cells[0].text = str(row['poste_l2'])
            data_row.cells[1].text = f"{row['tco2e']:.2f}"

        # Appliquer le style
        self._apply_table_style(doc_table)

        return True

    def _style_header_row(self, row):
        """Applique le style à une ligne d'en-tête."""
        for cell in row.cells:
            # Gras
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(11)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Fond gris
            shading_elm = self._get_or_create_shading(cell)
            shading_elm.fill = "D9D9D9"

    def _apply_table_style(self, table: Table):
        """Applique le style général à un tableau."""
        # Bordures et alignement
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

    def _get_or_create_shading(self, cell):
        """Récupère ou crée l'élément shading pour une cellule."""
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        tc = cell._element
        tcPr = tc.get_or_add_tcPr()
        shading = tcPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')

        if shading is None:
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>')
            tcPr.append(shading)

        return shading

    def create_simple_table(self, data: List[List[str]], headers: Optional[List[str]] = None) -> pd.DataFrame:
        """
        Crée un DataFrame simple pour un tableau.

        Args:
            data: Données du tableau
            headers: En-têtes optionnels

        Returns:
            DataFrame
        """
        if headers:
            return pd.DataFrame(data, columns=headers)
        else:
            return pd.DataFrame(data)
