# Tests Unitaires

Ce dossier contient les tests unitaires du projet carbonReportSAUR.

## Tests Disponibles

### test_word_renderer_blocks.py

Tests unitaires pour les nouvelles fonctionnalités de WordRenderer :
- Vérification de l'existence de toutes les méthodes (15 nouvelles méthodes)
- Vérification que les méthodes sont callable
- Tests rapides sans générer de rapport
- Couverture : Logo ORG, Blocs LOT/ACTIVITY/POST/OTHER_POST

**Exécution** :
```bash
python tests/unit/test_word_renderer_blocks.py
```

**Résultat attendu** :
```
✅ 3 tests réussis, 0 échoués
```

## Différence avec test_rapport.py

- **Tests unitaires** (ce dossier) : Tests rapides, techniques, sans Excel ni rapport Word
- **test_rapport.py** : Test d'intégration E2E complet avec Excel et génération du rapport Word

## Ajouter un Nouveau Test Unitaire

Créez un nouveau fichier `test_xxx.py` dans ce dossier avec la structure :

```python
#!/usr/bin/env python3
import sys
from pathlib import Path

# Ajouter le dossier racine au path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.your_module import YourClass

def test_your_feature():
    # Votre test ici
    pass

if __name__ == "__main__":
    test_your_feature()
```
