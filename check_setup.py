#!/usr/bin/env python3
"""
Script de vérification de l'installation et de la configuration.
Lance ce script pour vérifier que tout est prêt avant de démarrer l'app.
"""

import sys
from pathlib import Path


def check_python_version():
    """Vérifie la version de Python."""
    print("🐍 Vérification version Python...")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 8:
        print(f"   ✅ Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"   ❌ Python {version.major}.{version.minor} (requis: >= 3.8)")
        return False


def check_dependencies():
    """Vérifie que les dépendances sont installées."""
    print("\n📦 Vérification des dépendances...")

    required = [
        'streamlit',
        'pandas',
        'openpyxl',
        'docx',
        'matplotlib',
        'PIL'
    ]

    missing = []
    for module in required:
        try:
            if module == 'docx':
                __import__('docx')
            elif module == 'PIL':
                __import__('PIL')
            else:
                __import__(module)
            print(f"   ✅ {module}")
        except ImportError:
            print(f"   ❌ {module} (manquant)")
            missing.append(module)

    if missing:
        print(f"\n   ⚠️  Installer les dépendances manquantes :")
        print(f"   pip install -r requirements.txt")
        return False

    return True


def check_directories():
    """Vérifie que les dossiers nécessaires existent."""
    print("\n📁 Vérification des dossiers...")

    required_dirs = [
        'src',
        'templates',
        'assets',
        'output'
    ]

    all_ok = True
    for dir_name in required_dirs:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"   ✅ {dir_name}/")
        else:
            print(f"   ❌ {dir_name}/ (manquant)")
            all_ok = False

    return all_ok


def check_source_files():
    """Vérifie que les modules sources sont présents."""
    print("\n💻 Vérification des modules sources...")

    required_files = [
        'src/__init__.py',
        'src/excel_loader.py',
        'src/tree.py',
        'src/calc_emissions.py',
        'src/calc_indicators.py',
        'src/content_catalog.py',
        'src/chart_generators.py',
        'src/table_generators.py',
        'src/kpi_calculators.py',
        'src/word_renderer.py',
        'src/word_blocks.py',
        'app.py',
        'requirements.txt'
    ]

    all_ok = True
    for file_name in required_files:
        file_path = Path(file_name)
        if file_path.exists():
            print(f"   ✅ {file_name}")
        else:
            print(f"   ❌ {file_name} (manquant)")
            all_ok = False

    return all_ok


def check_template():
    """Vérifie que le template Word est présent."""
    print("\n📄 Vérification du template Word...")

    template_path = Path('templates/rapport_template.docx')
    if template_path.exists():
        print(f"   ✅ rapport_template.docx trouvé")
        return True
    else:
        print(f"   ⚠️  rapport_template.docx NON TROUVÉ")
        print(f"   → Placer votre template dans templates/rapport_template.docx")
        return False


def check_assets():
    """Vérifie que les assets sont présents."""
    print("\n🖼️  Vérification des assets...")

    required_assets = [
        'assets/logo_org.png',
        'assets/digesteur_schema.png'
    ]

    all_ok = True
    for asset in required_assets:
        asset_path = Path(asset)
        if asset_path.exists():
            print(f"   ✅ {asset}")
        else:
            print(f"   ⚠️  {asset} NON TROUVÉ")
            all_ok = False

    if not all_ok:
        print(f"   → Placer vos images dans le dossier assets/")

    return all_ok


def main():
    """Fonction principale."""
    print("=" * 60)
    print("🔍 VÉRIFICATION DE L'INSTALLATION")
    print("=" * 60)

    checks = {
        'Python': check_python_version(),
        'Dépendances': check_dependencies(),
        'Dossiers': check_directories(),
        'Modules sources': check_source_files(),
        'Template Word': check_template(),
        'Assets': check_assets()
    }

    print("\n" + "=" * 60)
    print("📊 RÉSUMÉ")
    print("=" * 60)

    for check_name, result in checks.items():
        status = "✅" if result else "❌"
        print(f"{status} {check_name}")

    all_ok = all(checks.values())

    print("\n" + "=" * 60)

    if all_ok:
        print("🎉 TOUT EST PRÊT !")
        print("\nVous pouvez lancer l'application avec :")
        print("   streamlit run app.py")
    else:
        print("⚠️  CONFIGURATION INCOMPLÈTE")
        print("\nActions requises :")

        if not checks['Python']:
            print("   - Installer Python >= 3.8")

        if not checks['Dépendances']:
            print("   - Installer les dépendances : pip install -r requirements.txt")

        if not checks['Template Word']:
            print("   - Ajouter le template Word dans templates/rapport_template.docx")

        if not checks['Assets']:
            print("   - Ajouter les images dans assets/")

        print("\nConsultez INSTRUCTIONS.md pour plus de détails.")

    print("=" * 60)

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
