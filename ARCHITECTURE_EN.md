# RateApp Architecture

## Overview

RateApp follows a layered architecture (**UI → Business Logic → Integration Services**) built on top of PySide6. The user interface focuses on presentation, while calculations, import/export routines, and external integrations live in dedicated modules. Data from Trados/Smartcat, Outlook, Excel, and Microsoft 365 flows through a central pipeline and results in production-ready proposals.

Guiding principles:
- **Separation of concerns** — UI components delegate business rules and I/O to `logic/` and `services/`.
- **Reusability** — import/export services power both the main window and embedded utilities such as the rates panel.
- **Packaging readiness** — `resource_utils.py`, `env_loader.py`, and bundled templates guarantee identical behaviour in development and PyInstaller builds.

---

## Application layers

### 1. Entry point
**`main.py`**
- loads environment variables and configures logging (`env_loader.load_application_env`, `logging_utils.setup_logging`);
- records application launch via `activity_logger`;
- creates `QApplication`, installs the icon, and shows `TranslationCostCalculator`;
- ensures stray Excel processes are closed on exit (`excel_process.close_excel_processes`).

### 2. User interface (`gui/`)
PySide6 presentation layer:
- **`main_window.py`** — main window composition, menus, drag & drop handling, integration with `ProjectManager`;
- **`panels/`** — left/right panes hosting the project card, language pairs, and calculation summaries;
- **`language_pair.py`**, **`additional_services.py`** — widgets for pair-specific and additional services calculations;
- **`project_setup_widget.py`** — "Project setup & management" block with configurable tasks, discounts, and markups;
- **`project_manager_dialog.py`** — project manager selector/editor dialog;
- **`rates_manager_window.py`** — embedded rates panel with Excel mapping and mismatch highlighting;
- **`drop_areas.py`** — drag & drop helpers for XML/MSG/Excel files;
- **`styles.py`**, **`utils.py`** — shared styling and formatting helpers.

Widgets rely on the `LanguagePairsMixin` and functions from `logic/` for calculations, imports, and exports.

### 3. Business logic (`logic/`)
Core domain, calculations, and integration logic.

**Project management**
- `project_manager.py` — orchestrates save/load, Excel/PDF export, and state resets;
- `project_data.py` — immutable project snapshot consumed by exporters and logs;
- `project_io.py` — JSON serialisation/deserialisation with compatibility checks;
- `pm_store.py` — project manager history storage;
- `user_config.py` — persisted user preferences (UI language, defaults).

**Calculations & reference data**
- `calculations.py` — totals, discount/markup aggregation, VAT handling;
- `online_rates.py` — exchange-rate storage and conversions;
- `language_pairs.py`, `language_codes.py` — language pair models and normalisation utilities;
- `legal_entities.py` + `legal_entities.json` — legal entity metadata and template paths.

**Import pipeline**
- `importers.py` — high-level facades for reports and rate cards;
- `sc_xml_parser.py`, `trados_xml_parser.py`, `xml_parser_common.py` — CAT report parsers;
- `outlook_import/` — Outlook `.msg` processing (card fields, tables, attachments);
- `excel_process.py`, `rates_importer.py` — Excel parsing, normalisation, and rate application;
- `ms_graph_client.py` — Microsoft Graph access (MSAL auth, file lookup by path or `fileId`).

**Export pipeline**
- `excel_exporter.py` — template-driven Excel proposal generation;
- `pdf_exporter.py` — Excel → PDF conversion via COM (Windows);
- `activity_logger.py` — Markdown-formatted activity log for diagnostics.

**Infrastructure**
- `env_loader.py`, `service_config.py` — environment and path configuration;
- `logging_utils.py`, `activity_logger.py` — logging configuration and structured records;
- `progress.py` — progress indicators for long-running tasks;
- `outlook_com_cache.py` — Outlook COM cache maintenance on Windows;
- `app_info.py` — application metadata (version, author).

### 4. Integration services (`services/`)
Lightweight adapters that keep legacy components and external APIs isolated:
- `excel_export.py` — exports Qt table models (LogTab, MemoQ, legacy tabs) to Excel with auto-formatting;
- `ms_graph.py` — high-level wrapper over `ms_graph_client` returning `pandas.DataFrame` objects and logging failures.

---

## Supporting directories & assets
- **`templates/`** — Excel templates, branding, and resources used by exporters.
- **`rates1/`** — embedded legacy rate tabs reused inside `rates_manager_window.py`.
- **`utils/`** — lightweight UI utilities (`history.py`, `theme.py`).
- **`tests/`** — pytest suite covering calculations, importers, exporters, and helpers.
- **`resource_utils.py`** — central resource locator for development and PyInstaller.
- **`requirements.txt`**, **`main.spec`**, **`rateapp.ico`**, **`RateApp.exe.sha256`** — build infrastructure files.

---

## Data flow

### 1. Import
```text
User action (Drag & Drop / file dialog / Microsoft Graph)
    ↓
logic/importers.py
    ↓
Format-specific parsers (XML, Outlook MSG, Excel)
    ↓
LanguagePair widgets / ProjectSetupWidget / AdditionalServicesWidget
    ↓
ProjectManager + ProjectData
```
**Sources:**
- Trados & Smartcat XML reports;
- Outlook `.msg` files (project card, volume tables, attachments);
- Local or Microsoft Graph Excel rate cards.

### 2. Processing
```text
ProjectManager (state aggregation)
    ↓
calculations.py (totals, discounts, VAT)
    ↓
online_rates.py (currency conversion)
    ↓
project_data.py (snapshot for export/logging)
```
Reference providers (`legal_entities`, `pm_store`, `user_config`) supply metadata and user defaults.

### 3. Export
```text
User request (Export menu)
    ↓
project_manager.py
    ↓
excel_exporter.py  ──→  templates/
    └── pdf_exporter.py (optional, Windows)
    ↓
Output files (Excel / PDF / JSON)
```
`activity_logger` records structured entries alongside optional project snapshots.

### 4. External integrations
```text
Microsoft Graph requests ← ms_graph_client.py → MSAL authentication
    ↓
services/ms_graph.py (normalisation & caching)
```

---

## Design considerations
- **Maintainability** — clear separation between UI, business logic, and services simplifies changes.
- **Extensibility** — new parsers, templates, and data providers can be added without touching existing UI components.
- **Diagnostics** — enhanced logging (`logging_utils`, `activity_logger`) captures user actions for troubleshooting.
- **Packaging** — consistent resource discovery enables seamless PyInstaller builds.
