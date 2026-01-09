"""
Module de gestion du catalogue de contenus pour le rapport.
Lit TEXTE_RAPPORT et résout les keys CHART/TABLE/IMAGE par poste.
"""

import pandas as pd
from typing import Dict, List, Optional
from dataclasses import dataclass


@dataclass
class PosteContent:
    """Contenu associé à un poste L1."""
    poste_l1_code: str
    text: str
    icone: Optional[str] = None
    chart_key: Optional[str] = None
    image_key: Optional[str] = None
    table_key: Optional[str] = None
    activity: str = 'BOTH'  # EU, AEP, ou BOTH

    def matches_activity(self, target_activity: str) -> bool:
        """
        Vérifie si ce contenu correspond à l'activité cible.

        Args:
            target_activity: Activité cible (EU ou AEP)

        Returns:
            True si le contenu s'applique à cette activité
        """
        if self.activity == 'BOTH':
            return True
        return self.activity == target_activity


class ContentCatalog:
    """
    Catalogue des contenus de rapport pour chaque poste L1.
    Gère la résolution des textes, graphiques, tableaux et images.
    """

    # Keys supportées
    SUPPORTED_CHART_KEYS = {
        'TRAVAUX_BREAKDOWN',
        'FILE_EAU_BREAKDOWN',
        'EM_INDIRECTES_SPLIT'
    }

    SUPPORTED_TABLE_KEYS = {
        'EM_INDIRECTES_TABLE'
    }

    SUPPORTED_IMAGE_KEYS = {
        'DIGESTEUR_SCHEMA'
    }

    def __init__(self, texte_rapport_df: pd.DataFrame):
        """
        Initialise le catalogue.

        Args:
            texte_rapport_df: DataFrame TEXTE_RAPPORT
        """
        self.texte_rapport_df = texte_rapport_df
        self._build_catalog()

    def _build_catalog(self):
        """Construit le catalogue à partir du DataFrame."""
        # Dictionnaire : {poste_l1_code: [PosteContent]}
        self.catalog: Dict[str, List[PosteContent]] = {}

        for _, row in self.texte_rapport_df.iterrows():
            poste_code = row['poste_l1_code']

            # Nettoyer les valeurs None/NaN
            chart_key = row['CHART_KEY'] if pd.notna(row['CHART_KEY']) and row['CHART_KEY'].strip() else None
            image_key = row['IMAGE_KEY'] if pd.notna(row['IMAGE_KEY']) and row['IMAGE_KEY'].strip() else None
            table_key = row['TABLE_KEY'] if pd.notna(row['TABLE_KEY']) and row['TABLE_KEY'].strip() else None
            icone = row['icone'] if pd.notna(row['icone']) and row['icone'].strip() else None
            activity = row['activity'] if pd.notna(row['activity']) else 'BOTH'

            content = PosteContent(
                poste_l1_code=poste_code,
                text=str(row['value']) if pd.notna(row['value']) else '',
                icone=icone,
                chart_key=chart_key,
                image_key=image_key,
                table_key=table_key,
                activity=activity
            )

            if poste_code not in self.catalog:
                self.catalog[poste_code] = []

            self.catalog[poste_code].append(content)

    def get_content(self, poste_l1_code: str, activity: str) -> Optional[PosteContent]:
        """
        Récupère le contenu pour un poste L1 et une activité donnée.

        Args:
            poste_l1_code: Code du poste L1
            activity: Activité (EU ou AEP)

        Returns:
            PosteContent correspondant ou None
        """
        if poste_l1_code not in self.catalog:
            return None

        contents = self.catalog[poste_l1_code]

        # Chercher d'abord un match exact sur l'activité
        for content in contents:
            if content.activity == activity:
                return content

        # Sinon fallback sur BOTH
        for content in contents:
            if content.activity == 'BOTH':
                return content

        # Si rien trouvé, retourner le premier
        return contents[0] if contents else None

    def has_chart(self, poste_l1_code: str, activity: str) -> bool:
        """Vérifie si un poste a un graphique associé."""
        content = self.get_content(poste_l1_code, activity)
        return content is not None and content.chart_key is not None

    def has_table(self, poste_l1_code: str, activity: str) -> bool:
        """Vérifie si un poste a un tableau associé."""
        content = self.get_content(poste_l1_code, activity)
        return content is not None and content.table_key is not None

    def has_image(self, poste_l1_code: str, activity: str) -> bool:
        """Vérifie si un poste a une image associée."""
        content = self.get_content(poste_l1_code, activity)
        return content is not None and content.image_key is not None

    def get_chart_key(self, poste_l1_code: str, activity: str) -> Optional[str]:
        """Retourne la CHART_KEY pour un poste/activité."""
        content = self.get_content(poste_l1_code, activity)
        return content.chart_key if content else None

    def get_table_key(self, poste_l1_code: str, activity: str) -> Optional[str]:
        """Retourne la TABLE_KEY pour un poste/activité."""
        content = self.get_content(poste_l1_code, activity)
        return content.table_key if content else None

    def get_image_key(self, poste_l1_code: str, activity: str) -> Optional[str]:
        """Retourne la IMAGE_KEY pour un poste/activité."""
        content = self.get_content(poste_l1_code, activity)
        return content.image_key if content else None

    def get_text(self, poste_l1_code: str, activity: str) -> str:
        """Retourne le texte pour un poste/activité."""
        content = self.get_content(poste_l1_code, activity)
        return content.text if content else ''

    def is_chart_supported(self, chart_key: str) -> bool:
        """Vérifie si une CHART_KEY est supportée."""
        return chart_key in self.SUPPORTED_CHART_KEYS

    def is_table_supported(self, table_key: str) -> bool:
        """Vérifie si une TABLE_KEY est supportée."""
        return table_key in self.SUPPORTED_TABLE_KEYS

    def is_image_supported(self, image_key: str) -> bool:
        """Vérifie si une IMAGE_KEY est supportée."""
        return image_key in self.SUPPORTED_IMAGE_KEYS

    def get_all_postes(self) -> List[str]:
        """Retourne la liste de tous les postes L1 dans le catalogue."""
        return list(self.catalog.keys())
