"""
Module de calcul des KPI et textes générés.
Calcule les équivalents (vols, empreinte FR) et génère les textes de comparaison.
"""

from typing import Dict, Optional
from .calc_emissions import EmissionResult
from .calc_indicators import IndicatorResult


class KPICalculator:
    """Calculateur de KPI et textes générés."""

    # Constantes pour les équivalences
    CO2_PER_FLIGHT_PARIS_NY = 1.0  # tCO2e par vol (à ajuster selon données réelles)
    CO2_PER_PERSON_YEAR_FR = 10.0  # tCO2e/an/personne en France (à ajuster)

    def __init__(self):
        """Initialise le calculateur."""
        pass

    def calculate_flight_equivalent(self, total_tco2e: float) -> float:
        """
        Calcule l'équivalent en nombre de vols Paris-New York.

        Args:
            total_tco2e: Émissions totales en tCO2e

        Returns:
            Nombre de vols équivalents
        """
        return total_tco2e / self.CO2_PER_FLIGHT_PARIS_NY

    def calculate_person_equivalent(self, total_tco2e: float) -> float:
        """
        Calcule l'équivalent en empreinte carbone de Français par an.

        Args:
            total_tco2e: Émissions totales en tCO2e

        Returns:
            Nombre de personnes équivalent
        """
        return total_tco2e / self.CO2_PER_PERSON_YEAR_FR

    def calculate_kpi_eu_1(self, emission_result: EmissionResult,
                          indicator_result: Optional[IndicatorResult]) -> Optional[float]:
        """
        Calcule le KPI EU 1 : kgCO2e/m³ eau épurée.

        Args:
            emission_result: Résultat d'émissions EU
            indicator_result: Résultat d'indicateurs EU (contient volume eau épurée)

        Returns:
            KPI en kgCO2e/m³ ou None si pas de données
        """
        if indicator_result is None:
            return None

        # Chercher l'indicateur de volume d'eau épurée
        # (code à adapter selon le nom exact dans INDICATORS_REF)
        volume_indicator = indicator_result.get_indicator('VOL_EAU_EPUREE')
        if volume_indicator is None:
            return None

        volume_m3 = volume_indicator.value
        if volume_m3 == 0:
            return None

        # Convertir tCO2e en kgCO2e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / volume_m3

    def calculate_kpi_eu_2(self, emission_result: EmissionResult,
                          indicator_result: Optional[IndicatorResult]) -> Optional[float]:
        """
        Calcule le KPI AEP 1 : kgCO2e/m³ eau distribuée.

        Args:
            emission_result: Résultat d'émissions AEP
            indicator_result: Résultat d'indicateurs AEP (contient volume eau distribuée)

        Returns:
            KPI en kgCO2e/m³ ou None si pas de données
        """
        if indicator_result is None:
            return None

        # Chercher l'indicateur de volume d'eau distribuée
        volume_indicator = indicator_result.get_indicator('VOL_EAU_DISTRIBUEE')
        if volume_indicator is None:
            return None

        volume_m3 = volume_indicator.value
        if volume_m3 == 0:
            return None

        # Convertir tCO2e en kgCO2e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / volume_m3

    def generate_activity_volume_comparison_text(self,
                                                 eu_result: Optional[EmissionResult],
                                                 aep_result: Optional[EmissionResult],
                                                 eu_indicators: Optional[IndicatorResult],
                                                 aep_indicators: Optional[IndicatorResult]) -> str:
        """
        Génère le texte de comparaison des volumes d'activité.

        Args:
            eu_result: Résultat émissions EU
            aep_result: Résultat émissions AEP
            eu_indicators: Indicateurs EU
            aep_indicators: Indicateurs AEP

        Returns:
            Texte généré
        """
        lines = []

        # Volumes EU
        if eu_result and eu_indicators:
            vol_epuree = eu_indicators.get_indicator('VOL_EAU_EPUREE')
            if vol_epuree:
                lines.append(f"L'activité Assainissement a traité {vol_epuree.value:,.0f} {vol_epuree.unit} "
                           f"pour un total de {eu_result.total_tco2e:,.1f} tCO₂e.")

        # Volumes AEP
        if aep_result and aep_indicators:
            vol_distribuee = aep_indicators.get_indicator('VOL_EAU_DISTRIBUEE')
            if vol_distribuee:
                lines.append(f"L'activité Eau Potable a distribué {vol_distribuee.value:,.0f} {vol_distribuee.unit} "
                           f"pour un total de {aep_result.total_tco2e:,.1f} tCO₂e.")

        # Comparaison si les deux activités sont présentes
        if eu_result and aep_result and len(lines) == 2:
            if eu_result.total_tco2e > aep_result.total_tco2e:
                ratio = eu_result.total_tco2e / aep_result.total_tco2e if aep_result.total_tco2e > 0 else 0
                lines.append(f"L'activité Assainissement représente {ratio:.1f}x les émissions de l'activité Eau Potable.")
            else:
                ratio = aep_result.total_tco2e / eu_result.total_tco2e if eu_result.total_tco2e > 0 else 0
                lines.append(f"L'activité Eau Potable représente {ratio:.1f}x les émissions de l'activité Assainissement.")

        return " ".join(lines) if lines else ""

    def generate_top_postes_list_text(self, top_postes: list, poste_labels: Dict[str, str]) -> str:
        """
        Génère le texte de liste des top postes.

        Args:
            top_postes: Liste de tuples (poste_l1_code, tco2e)
            poste_labels: Dictionnaire {code: label}

        Returns:
            Texte formaté
        """
        if not top_postes:
            return ""

        lines = []
        for i, (code, tco2e) in enumerate(top_postes, 1):
            label = poste_labels.get(code, code)
            lines.append(f"{i}. **{label}** : {tco2e:,.1f} tCO₂e")

        return "\n".join(lines)

    def format_kpi(self, value: Optional[float], unit: str, decimals: int = 2) -> str:
        """
        Formate un KPI pour affichage.

        Args:
            value: Valeur du KPI
            unit: Unité
            decimals: Nombre de décimales

        Returns:
            Texte formaté
        """
        if value is None:
            return "N/A"

        return f"{value:,.{decimals}f} {unit}"

    def generate_excluded_postes_note(self, excluded_postes: list,
                                      poste_labels: Dict[str, str]) -> str:
        """
        Génère une note sur les postes exclus des totaux.

        Args:
            excluded_postes: Liste des codes de postes exclus
            poste_labels: Dictionnaire {code: label}

        Returns:
            Texte de note
        """
        if not excluded_postes:
            return ""

        poste_names = [poste_labels.get(code, code) for code in excluded_postes]
        postes_str = ", ".join(poste_names)

        return (f"Note : Les postes suivants ont été exclus des totaux de ce rapport : {postes_str}. "
               f"Cette exclusion a été effectuée pour une meilleure représentativité des émissions.")
