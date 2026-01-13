# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Carbon Report Generator (SAUR) - A Streamlit application that automatically generates carbon footprint reports in Word format from Excel data. The application builds hierarchical organizational structures (ORG/LOT/ENT), calculates emissions across scopes and activities (EU/AEP), and produces customized reports with charts, tables, and images.

## Development Commands

### Running the Application

```bash
# Launch full Streamlit application
streamlit run app.py

# Quick test without Streamlit (BRUT mode only)
python tests/test_rapport.py <excel_file.xlsx> <year>
```

### Testing

```bash
# Run unit tests
python tests/unit/test_word_renderer_blocks.py

# End-to-end test with sample data
python tests/test_rapport.py <excel_file.xlsx>
```

### Setup Verification

```bash
# Check installation and configuration
python check_setup.py
```

## Architecture

### Core Data Flow

1. **Excel Loading** (`excel_loader.py`) - Validates and loads 9+ required Excel sheets
2. **Tree Construction** (`tree.py`) - Builds hierarchical ORG→LOT→ENT structure with activity types
3. **Emission Calculation** (`calc_emissions.py`) - Aggregates emissions at all levels, handles BRUT/NET modes
4. **Indicator Calculation** (`calc_indicators.py`) - Computes indicators by LOT×ACTIVITY
5. **Content Cataloging** (`content_catalog.py`) - Maps emissions data to report elements (text, charts, images)
6. **Word Rendering** (`word_renderer.py`) - Generates final Word document with block duplication

### Two Calculation Modes

**BRUT Mode**: Direct calculations from Excel without user modifications
- Used by test scripts
- No renaming, no post exclusions
- Baseline for comparison

**NET Mode**: Calculations with user overrides applied
- Used by Streamlit app
- Supports node renaming (via `EmissionOverrides.node_renames`)
- Supports post configuration (via `EmissionOverrides.poste_config`)
  - Mode A: `show_in_report=False, include_in_totals=True` (hidden but counted)
  - Mode B: `show_in_report=False, include_in_totals=False` (fully excluded)

### Hierarchical Structure

**3-Level Hierarchy**:
- **ORG** (root): Organization level, one per report
- **LOT** (optional): Intermediate grouping level, multiple per ORG
- **ENT** (leaves): Entity level with activity type (EU or AEP)

**Activity Types**:
- **EU**: Eaux usées (wastewater)
- **AEP**: Assainissement eau potable (drinking water)
- **NA**: Not applicable (non-leaf nodes)

**Key Constraint**: Emissions are calculated at ORG level AND at LOT×ACTIVITY level. Each calculation produces an `EmissionResult` with scopes, top posts, and other posts.

### Word Template System

The Word renderer uses a sophisticated block duplication system:

**Block Hierarchy** (outer to inner):
1. `[[START_LOT]]...[[END_LOT]]` - Duplicated for each LOT
2. `[[START_ACTIVITY]]...[[END_ACTIVITY]]` - Duplicated for each activity (EU/AEP)
3. `[[START_POST]]...[[END_POST]]` - Duplicated for top N emitting posts
4. `[[START_OTHER_POST]]...[[END_OTHER_POST]]` - Duplicated for non-top posts

**Processing Order**: Must process from outermost blocks inward to avoid index conflicts after duplication. After each duplication, blocks must be re-scanned to get updated indices.

**Critical Pattern**: `duplicate_block()` → `re-scan blocks` → `process each block`

### Chart Generation

Charts are generated using matplotlib via `ChartGenerator`:
- **ORG-level charts**: `chart_emissions_scope_org`, `chart_contrib_lot`, etc.
- **LOT×ACTIVITY charts**: `chart_pie_scope_entity_activity`, `chart_pie_postes_entity_activity`
- **POST-specific charts**: `TRAVAUX_BREAKDOWN`, `FILE_EAU_BREAKDOWN`, `EM_INDIRECTES_SPLIT`

Charts are customizable via `st.session_state.chart_customizations` in the Streamlit UI.

## Excel Data Format

### Required Sheets

1. **ORG_TREE**: Hierarchical structure (`node_id`, `parent_id`, `node_type`, `node_name`, `activity`)
2. **EMISSIONS**: L1 emissions data (`node_id`, `scope`, `poste_l1_code`, `tco2e`, `comment`)
3. **EMISSIONS_L2**: L2 detail for breakdown charts
4. **POSTES_REF**: L1 post reference (`poste_l1_code`, `poste_l1_label`, `commentaire`)
5. **POSTES_L2_REF**: L2 post reference for detail views
6. **INDICATORS**: Indicators by node and activity
7. **INDICATORS_REF**: Indicator definitions
8. **EMISSIONS_EVITEES**: Avoided emissions
9. **TEXTE_RAPPORT**: Content catalog linking posts to text/images/charts

### Optional Sheets

- **ICONE**: Icon mapping for posts

## Key Implementation Details

### BlockProcessor Pattern

The `word_blocks.py` module provides a `BlockProcessor` class that wraps python-docx operations:
- `find_block(start_marker, end_marker)` - Locate block boundaries
- `duplicate_block(start_idx, end_idx, n_copies)` - Clone paragraphs
- `replace_in_block(start_idx, end_idx, replacements)` - Replace placeholders within boundaries
- `remove_block_markers(start_marker, end_marker)` - Clean up `[[XXX]]` markers

### Image Insertion

Images must be on their own line in the template for proper replacement:
```
{{chart_emissions_scope_org}}

{{ORG_LOGO}}
```

Image placeholders NOT on their own line will be treated as text replacements.

### Top Posts Calculation

The emission calculator determines top N posts (default: 4) by:
1. Filtering posts where `include_in_totals=True`
2. Sorting by tCO2e descending
3. Taking top N
4. Storing remainder in `other_postes`

This happens at both ORG and LOT×ACTIVITY levels.

### Content Catalog Mapping

The `ContentCatalog` reads `TEXTE_RAPPORT` to map each `poste_l1_code` to:
- `value`: Descriptive text
- `CHART_KEY`: Which chart to generate (e.g., `TRAVAUX_BREAKDOWN`)
- `IMAGE_KEY`: Which static image to insert (e.g., `DIGESTEUR_SCHEMA`)
- `TABLE_KEY`: Which table to generate (e.g., `EM_INDIRECTES_TABLE`)
- `activity`: Filter for EU/AEP/both
- `DETAIL_SOURCE`: Where to pull L2 data

## Common Patterns

### Adding a New Chart Type

1. Add method to `ChartGenerator` (e.g., `generate_my_chart()`)
2. Add chart key to `SUPPORTED_CHART_KEYS` in `word_renderer.py`
3. Map chart in `TEXTE_RAPPORT` Excel sheet via `CHART_KEY` column
4. Placeholder in template: `{{POST_CHART_1}}` or `{{chart_my_chart}}`

### Handling Missing Data Gracefully

The renderer never raises exceptions for missing content. Instead:
- Missing images: placeholder left as-is (cleaned later)
- Missing charts: placeholder left as-is (cleaned later)
- Missing L2 data: chart/table not generated
- Empty blocks: deleted automatically

This allows partial data to still produce a valid report.

### EmissionResult Structure

Each emission calculation produces an `EmissionResult` containing:
- `node_id`, `node_name`, `activity`
- `total_tco2e`, `scope1_tco2e`, `scope2_tco2e`, `scope3_tco2e`
- `emissions_by_poste`: Dict[str, float]
- `top_postes`: List[Tuple[str, float]] - Top N posts
- `other_postes`: List[Tuple[str, float]] - Remaining posts

## Testing Strategy

**Unit Tests**: Located in `tests/unit/`, test individual components
**Integration Tests**: `tests/test_rapport.py` - Full end-to-end generation
**Quick Iteration**: Use test script instead of Streamlit for rapid template/data testing

## File Organization

- `src/`: All Python modules (8+ files)
- `templates/`: Word template (`rapport_template.docx`)
- `assets/`: Static images (logos, schemas, icons)
- `output/`: Generated reports (Streamlit)
- `tests/output/`: Test-generated reports
- `app.py`: Streamlit application entry point

## Dependencies

Core: `streamlit`, `pandas`, `openpyxl`, `python-docx`, `matplotlib`, `Pillow`

All dependencies in `requirements.txt` - install with: `pip install -r requirements.txt`
