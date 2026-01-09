"""
Module de gestion de l'arborescence ORG/LOT/ENT.
Reconstruit l'arbre hiérarchique et fournit des helpers de navigation.
"""

import pandas as pd
from typing import Dict, List, Optional, Set
from dataclasses import dataclass, field


@dataclass
class TreeNode:
    """Représente un nœud dans l'arborescence."""
    node_id: str
    parent_id: Optional[str]
    node_type: str  # ORG, LOT, ENT
    node_name: str
    activity: Optional[str]  # EU, AEP, NA (uniquement pour ENT)
    children: List['TreeNode'] = field(default_factory=list)
    parent: Optional['TreeNode'] = None

    def is_org(self) -> bool:
        """Vérifie si le nœud est de type ORG."""
        return self.node_type == 'ORG'

    def is_lot(self) -> bool:
        """Vérifie si le nœud est de type LOT."""
        return self.node_type == 'LOT'

    def is_ent(self) -> bool:
        """Vérifie si le nœud est de type ENT."""
        return self.node_type == 'ENT'

    def has_activity(self, activity: str) -> bool:
        """Vérifie si le nœud a une activité spécifique."""
        return self.activity == activity

    def get_depth(self) -> int:
        """Retourne la profondeur du nœud dans l'arbre."""
        depth = 0
        current = self.parent
        while current is not None:
            depth += 1
            current = current.parent
        return depth

    def get_path(self) -> List['TreeNode']:
        """Retourne le chemin depuis la racine jusqu'à ce nœud."""
        path = []
        current = self
        while current is not None:
            path.insert(0, current)
            current = current.parent
        return path

    def __repr__(self) -> str:
        return f"TreeNode(id={self.node_id}, type={self.node_type}, name={self.node_name})"


class OrganizationTree:
    """
    Gère l'arborescence complète de l'organisation.
    Fournit des méthodes de navigation et de filtrage.
    """

    def __init__(self, org_tree_df: pd.DataFrame):
        """
        Initialise l'arbre à partir du DataFrame ORG_TREE.

        Args:
            org_tree_df: DataFrame de l'onglet ORG_TREE
        """
        self.nodes: Dict[str, TreeNode] = {}
        self.root: Optional[TreeNode] = None
        self._build_tree(org_tree_df)

    def _build_tree(self, df: pd.DataFrame) -> None:
        """Construit l'arborescence à partir du DataFrame."""
        # Créer tous les nœuds
        for _, row in df.iterrows():
            node = TreeNode(
                node_id=str(row['node_id']),
                parent_id=str(row['parent_id']) if pd.notna(row['parent_id']) else None,
                node_type=str(row['node_type']),
                node_name=str(row['node_name']),
                activity=str(row['activity']) if pd.notna(row['activity']) else None
            )
            self.nodes[node.node_id] = node

        # Établir les relations parent-enfant
        for node in self.nodes.values():
            if node.parent_id is None or node.parent_id == '' or node.parent_id == 'nan':
                # C'est la racine
                if self.root is not None:
                    raise ValueError(f"Plusieurs racines détectées : {self.root.node_id} et {node.node_id}")
                self.root = node
            else:
                parent = self.nodes.get(node.parent_id)
                if parent is None:
                    raise ValueError(f"Parent {node.parent_id} introuvable pour le nœud {node.node_id}")
                parent.children.append(node)
                node.parent = parent

        if self.root is None:
            raise ValueError("Aucune racine (nœud ORG sans parent) trouvée dans l'arborescence")

        # Vérifier que la racine est bien de type ORG
        if not self.root.is_org():
            raise ValueError(f"La racine doit être de type ORG, trouvé : {self.root.node_type}")

    def get_node(self, node_id: str) -> Optional[TreeNode]:
        """Retourne un nœud par son ID."""
        return self.nodes.get(node_id)

    def get_org(self) -> TreeNode:
        """Retourne le nœud ORG racine."""
        if self.root is None:
            raise ValueError("Arbre non initialisé")
        return self.root

    def get_lots(self) -> List[TreeNode]:
        """Retourne tous les nœuds de type LOT."""
        return [node for node in self.nodes.values() if node.is_lot()]

    def get_ents(self) -> List[TreeNode]:
        """Retourne tous les nœuds de type ENT."""
        return [node for node in self.nodes.values() if node.is_ent()]

    def has_lots(self) -> bool:
        """Vérifie si l'organisation a des LOTs."""
        return len(self.get_lots()) > 0

    def get_children(self, node_id: str, node_type: Optional[str] = None) -> List[TreeNode]:
        """
        Retourne les enfants directs d'un nœud.

        Args:
            node_id: ID du nœud parent
            node_type: Filtre optionnel sur le type de nœud (ORG, LOT, ENT)

        Returns:
            Liste des nœuds enfants
        """
        node = self.get_node(node_id)
        if node is None:
            return []

        children = node.children
        if node_type is not None:
            children = [c for c in children if c.node_type == node_type]

        return children

    def get_descendants(self, node_id: str, node_type: Optional[str] = None) -> List[TreeNode]:
        """
        Retourne tous les descendants d'un nœud (récursif).

        Args:
            node_id: ID du nœud parent
            node_type: Filtre optionnel sur le type de nœud

        Returns:
            Liste de tous les descendants
        """
        node = self.get_node(node_id)
        if node is None:
            return []

        descendants = []
        self._collect_descendants(node, descendants, node_type)
        return descendants

    def _collect_descendants(self, node: TreeNode, result: List[TreeNode],
                            node_type: Optional[str] = None) -> None:
        """Collecte récursivement les descendants d'un nœud."""
        for child in node.children:
            if node_type is None or child.node_type == node_type:
                result.append(child)
            self._collect_descendants(child, result, node_type)

    def get_ents_by_activity(self, parent_node_id: str, activity: str) -> List[TreeNode]:
        """
        Retourne tous les ENT descendants d'un nœud avec une activité spécifique.

        Args:
            parent_node_id: ID du nœud parent (ORG ou LOT)
            activity: Activité recherchée (EU ou AEP)

        Returns:
            Liste des ENT avec l'activité spécifiée
        """
        all_ents = self.get_descendants(parent_node_id, node_type='ENT')
        return [ent for ent in all_ents if ent.has_activity(activity)]

    def get_lot_activities(self, lot_id: str) -> Set[str]:
        """
        Retourne les activités présentes dans un LOT (via ses ENT).

        Args:
            lot_id: ID du LOT

        Returns:
            Set des activités (EU, AEP) présentes dans ce LOT
        """
        ents = self.get_descendants(lot_id, node_type='ENT')
        activities = {ent.activity for ent in ents if ent.activity in ['EU', 'AEP']}
        return activities

    def get_org_activities(self) -> Set[str]:
        """
        Retourne les activités présentes au niveau ORG.

        Returns:
            Set des activités (EU, AEP) présentes dans l'organisation
        """
        all_ents = self.get_ents()
        activities = {ent.activity for ent in all_ents if ent.activity in ['EU', 'AEP']}
        return activities

    def get_ent_ids_by_activity(self, parent_node_id: str, activity: str) -> List[str]:
        """
        Retourne les IDs des ENT descendants avec une activité spécifique.

        Args:
            parent_node_id: ID du nœud parent
            activity: Activité recherchée

        Returns:
            Liste des node_id des ENT
        """
        ents = self.get_ents_by_activity(parent_node_id, activity)
        return [ent.node_id for ent in ents]

    def print_tree(self, node: Optional[TreeNode] = None, indent: int = 0) -> str:
        """
        Génère une représentation textuelle de l'arbre.

        Args:
            node: Nœud de départ (racine si None)
            indent: Niveau d'indentation

        Returns:
            Représentation textuelle de l'arbre
        """
        if node is None:
            node = self.root

        if node is None:
            return "Arbre vide"

        lines = []
        prefix = "  " * indent

        activity_str = f" [{node.activity}]" if node.activity and node.activity != 'NA' else ""
        lines.append(f"{prefix}{node.node_type}: {node.node_name} (ID: {node.node_id}){activity_str}")

        for child in node.children:
            lines.append(self.print_tree(child, indent + 1))

        return "\n".join(lines)

    def validate_structure(self) -> List[str]:
        """
        Valide la structure de l'arbre et retourne les erreurs trouvées.

        Returns:
            Liste des erreurs de validation
        """
        errors = []

        # Vérifier que tous les ENT ont une activité
        for ent in self.get_ents():
            if ent.activity not in ['EU', 'AEP']:
                errors.append(f"ENT {ent.node_id} ({ent.node_name}) n'a pas d'activité EU ou AEP")

        # Vérifier que les LOT sont bien entre ORG et ENT
        for lot in self.get_lots():
            if lot.parent is None or not lot.parent.is_org():
                errors.append(f"LOT {lot.node_id} ({lot.node_name}) n'a pas d'ORG comme parent")

            ent_children = [c for c in lot.children if c.is_ent()]
            if len(ent_children) == 0:
                errors.append(f"LOT {lot.node_id} ({lot.node_name}) n'a pas d'ENT enfant")

        # Vérifier qu'il n'y a pas de cycles
        visited = set()
        for node in self.nodes.values():
            path = []
            current = node
            while current is not None:
                if current.node_id in path:
                    errors.append(f"Cycle détecté impliquant le nœud {current.node_id}")
                    break
                path.append(current.node_id)
                current = current.parent

        return errors
