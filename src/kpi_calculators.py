"""
Module de calcul des KPI et textes générés.
Calcule les équivalents (vols, empreinte FR) et génère les textes de comparaison.
"""

from typing import Dict, Optional, List
from .calc_emissions import EmissionResult
from .calc_indicators import IndicatorResult


class KPICalculator:
    """Calculateur de KPI et textes générés."""

    # Constantes pour les équivalences
    CO2_PER_FLIGHT_PARIS_NY = 1.75  # tCO2e par vol Paris-New York
    CO2_PER_PERSON_YEAR_FR = 9.0  # tCO2e/an/personne en France

    def __init__(self):
        """Initialise le calculateur."""
        pass

    def format_number(self, value: float, decimals: int = 0) -> str:
        """
        Formate un nombre avec espace comme séparateur de milliers et arrondi.

        Args:
            value: Valeur à formater
            decimals: Nombre de décimales (0 par défaut)

        Returns:
            Nombre formaté (ex: "19 445")
        """
        if decimals == 0:
            # Arrondir à l'entier
            rounded = round(value)
            # Formater avec espace comme séparateur
            return f"{rounded:,}".replace(",", " ")
        else:
            # Avec décimales
            return f"{value:,.{decimals}f}".replace(",", " ")

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

    def sum_volumes_by_activity(self, indicator_results: List[IndicatorResult],
                                  indicator_code: str) -> float:
        """
        Somme tous les volumes pour un indicator_code donné.

        Args:
            indicator_results: Liste de résultats d'indicateurs
            indicator_code: Code de l'indicateur à sommer

        Returns:
            Volume total
        """
        total_volume = 0.0
        for result in indicator_results:
            indicator = result.get_indicator(indicator_code)
            if indicator:
                total_volume += indicator.value
        return total_volume

    def calculate_kpi_m3_eu(self, emission_result: EmissionResult,
                            indicator_results: List[IndicatorResult]) -> Optional[float]:
        """
        Calcule le KPI EU : kgCO2e/m³ eau épurée.
        Somme TOUS les volumes EU de tous les périmètres.

        Args:
            emission_result: Résultat d'émissions EU (total)
            indicator_results: Liste de TOUS les résultats d'indicateurs EU

        Returns:
            KPI en kgCO2e/m³ ou None si pas de données
        """
        if not indicator_results:
            return None

        # Sommer TOUS les volumes d'eau épurée
        total_volume_m3 = self.sum_volumes_by_activity(indicator_results, 'VOL_EAU_EPURE')

        if total_volume_m3 == 0:
            return None

        # Convertir tCO2e en kgCO2e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / total_volume_m3

    def calculate_kpi_m3_aep(self, emission_result: EmissionResult,
                             indicator_results: List[IndicatorResult]) -> Optional[float]:
        """
        Calcule le KPI AEP : kgCO2e/m³ eau distribuée.
        Somme TOUS les volumes AEP de tous les périmètres.

        Args:
            emission_result: Résultat d'émissions AEP (total)
            indicator_results: Liste de TOUS les résultats d'indicateurs AEP

        Returns:
            KPI en kgCO2e/m³ ou None si pas de données
        """
        if not indicator_results:
            return None

        # Sommer TOUS les volumes d'eau distribuée
        total_volume_m3 = self.sum_volumes_by_activity(indicator_results, 'VOL_EAU_DISTRIB')

        if total_volume_m3 == 0:
            return None

        # Convertir tCO2e en kgCO2e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / total_volume_m3

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
            vol_epuree = eu_indicators.get_indicator('VOL_EAU_EPURE')
            if vol_epuree:
                lines.append(f"L'activité Assainissement a traité {vol_epuree.value:,.0f} {vol_epuree.unit} "
                           f"pour un total de {eu_result.total_tco2e:,.1f} tCO₂e.")

        # Volumes AEP
        if aep_result and aep_indicators:
            vol_distribuee = aep_indicators.get_indicator('VOL_EAU_DISTRIB')
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

    def calculate_kpi_m3_entity(self, emission_result: EmissionResult,
                                indicator_result: IndicatorResult,
                                activity: str) -> Optional[float]:
        """
        Calcule le KPI kgCO₂e/m³ pour une entité spécifique (LOT×ACTIVITÉ).

        Args:
            emission_result: Résultat d'émissions pour cette entité
            indicator_result: Résultat d'indicateurs pour cette entité
            activity: Activité (EU ou AEP)

        Returns:
            KPI en kgCO₂e/m³ ou None si pas de données
        """
        if not indicator_result:
            return None

        # Sélectionner le bon indicateur de volume selon l'activité
        volume_code = 'VOL_EAU_EPURE' if activity == 'EU' else 'VOL_EAU_DISTRIB'
        volume_indicator = indicator_result.get_indicator(volume_code)

        if not volume_indicator or volume_indicator.value == 0:
            return None

        # Convertir tCO₂e en kgCO₂e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / volume_indicator.value

    def calculate_kpi_hab_entity(self, emission_result: EmissionResult,
                                 indicator_result: IndicatorResult) -> Optional[float]:
        """
        Calcule le KPI kgCO₂e/habitant pour une entité spécifique (LOT×ACTIVITÉ).

        Args:
            emission_result: Résultat d'émissions pour cette entité
            indicator_result: Résultat d'indicateurs pour cette entité

        Returns:
            KPI en kgCO₂e/habitant ou None si pas de données
        """
        if not indicator_result:
            return None

        # Récupérer le nombre d'habitants desservis
        hab_indicator = indicator_result.get_indicator('NB_HAB_DESSERVIS')

        if not hab_indicator or hab_indicator.value == 0:
            return None

        # Convertir tCO₂e en kgCO₂e
        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / hab_indicator.value

    def calculate_kpi_branch_entity(self, emission_result: EmissionResult,
                                     indicator_result: IndicatorResult) -> Optional[float]:
        """
        Calcule le KPI kgCO₂e/branchement pour une entité (LOT×ACTIVITÉ).

        Args:
            emission_result: Résultat d'émissions pour cette entité
            indicator_result: Résultat d'indicateurs pour cette entité

        Returns:
            KPI en kgCO₂e/branchement ou None si pas de données
        """
        if not indicator_result:
            return None

        branch_indicator = indicator_result.get_indicator('NB_BRANCHEMENTS')

        if not branch_indicator or branch_indicator.value == 0:
            return None

        kg_co2e = emission_result.total_tco2e * 1000
        return kg_co2e / branch_indicator.value

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
