# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Carbon Report Generator (SAUR) — a Streamlit application that generates carbon footprint reports in Word format from Excel data. It builds hierarchical organizational structures (ORG/LOT/ENT), calculates emissions across scopes and activities (EU/AEP), and produces reports with charts, tables, and images.

## Development Commands

```bash
# Launch Streamlit application (latest version)
streamlit run app_v1.py

# Quick test without Streamlit (BRUT mode only)
python tests/test_rapport.py <excel_file.xlsx> <year>

# Run unit tests
python tests/unit/test_word_renderer_blocks.py

# Check installation and configuration
python check_setup.py
```

Note: `app_v1.py` is the current and only version.

## Architecture

### Core Data Flow

1. **Excel Loading** (`flat_loader.py`) — Validates and loads Excel (DATA sheet) into internal DataFrames
2. **Tree Construction** (`tree.py`) — Builds hierarchical ORG→LOT→ENT structure with activity types
3. **Emission Calculation** (`calc_emissions.py`) — Aggregates emissions at all levels, handles BRUT/NET modes
4. **Indicator Calculation** (`calc_indicators.py`) — Computes indicators by LOT×ACTIVITY
5. **Content Cataloging** (`content_catalog.py`) — Maps emissions data to report elements (text, charts, images)
6. **Word Rendering** (`word_renderer.py`, `word_blocks.py`) — Generates final Word document with block duplication

`word_renderer.py` is the largest module (~1,576 lines, 44 methods). When modifying report generation logic, this is the primary file.

### Two Calculation Modes

**BRUT Mode**: Direct calculations from Excel without user modifications. Used by test scripts.

**NET Mode**: Calculations with user overrides applied (Streamlit app). Supports:
- Node renaming via `EmissionOverrides.node_renames`
- Post configuration via `EmissionOverrides.poste_config`:
  - Mode A: `show_in_report=False, include_in_totals=True` (hidden but counted)
  - Mode B: `show_in_report=False, include_in_totals=False` (fully excluded)

### Hierarchical Structure

**3-Level Hierarchy**: ORG (root, one per report) → LOT (optional intermediate grouping) → ENT (leaves with activity EU or AEP)

**Activity Types**: EU (eaux usées/wastewater), AEP (eau potable/drinking water), NA (non-leaf nodes)

**Key Constraint**: Emissions are calculated at ORG level AND at LOT×ACTIVITY level. Each produces an `EmissionResult` with scopes, top posts (default: 4), and other posts.

### Word Template System

The renderer uses nested block duplication processed from **outermost to innermost** to avoid index conflicts:

1. `[[START_LOT]]...[[END_LOT]]` — Duplicated for each LOT
2. `[[START_ACTIVITY]]...[[END_ACTIVITY]]` — Duplicated for each activity (EU/AEP)
3. `[[START_POST]]...[[END_POST]]` — Duplicated for top N emitting posts
4. `[[START_OTHER_POST]]...[[END_OTHER_POST]]` — Duplicated for remaining posts

**Critical Pattern**: `duplicate_block()` → re-scan blocks for updated indices → process each block

**Image/chart placeholders** (e.g., `{{chart_emissions_scope_org}}`, `{{ORG_LOGO}}`) must be **alone on their own line** in the template. Placeholders not on their own line are treated as text replacements.

### Graceful Degradation

The renderer never raises exceptions for missing content. Missing images/charts are left as-is then cleaned up. Missing L2 data skips chart/table generation. Empty blocks are deleted. Partial data always produces a valid report.

## Key Patterns

### Adding a New Chart Type

1. Add method to `ChartGenerator` in `chart_generators.py`
2. Add chart key to `SUPPORTED_CHART_KEYS` in `word_renderer.py`
3. Map chart in `TEXTE_RAPPORT` Excel sheet via `CHART_KEY` column
4. Add placeholder in template: `{{POST_CHART_1}}` or `{{chart_my_chart}}`

### Number Formatting

Numbers use French convention with space as thousands separator (e.g., "19 445"). See `KPICalculator.format_number()` in `kpi_calculators.py`.

### Chart Customization

Charts use matplotlib with custom Poppins fonts (loaded from `assets/police/`). A 6-color green palette is the default. Customizable at runtime via `st.session_state.chart_customizations` in the Streamlit UI (`streamlit_charts_page.py`).

## File Organization

- `src/` — All Python modules (11 files)
- `templates/` — Word template (`rapport_template.docx`)
- `assets/` — Static images (logos, schemas, icons, fonts)
- `output/` — Generated reports (Streamlit)
- `tests/output/` — Test-generated reports
- `app_v1.py` — Streamlit application entry point

## Dependencies

Core: `streamlit`, `pandas`, `openpyxl`, `python-docx`, `matplotlib`, `Pillow`

Install with: `pip install -r requirements.txt`
