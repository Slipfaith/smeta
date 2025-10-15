# RateApp Architecture

## Overview

RateApp is a desktop application built with PySide6 that follows a layered architecture pattern: **UI → Business Logic → Integration Services**. The application manages translation project data from multiple sources (Trados/Smartcat reports, Outlook emails, Excel files, SharePoint) and produces professional commercial proposals in Excel and PDF formats.

### Core Principles

- **Separation of Concerns**: UI components delegate all business logic and I/O operations to dedicated modules
- **Centralized Data Processing**: All project data flows through a unified processing pipeline
- **Modular Integration**: External services are wrapped in lightweight adapters
- **Resource Management**: Supports both development and PyInstaller-packaged deployment modes

---

## Architecture Layers

### 1. Entry Point
**`main.py`**
- Application initialization
- Configuration and localization setup
- Main window instantiation

### 2. User Interface Layer (`gui/`)
Handles all user interactions and visual presentation using PySide6 widgets.

**Structure:**
```
gui/
├── __init__.py          # Widget registration
├── main_window.py       # Main window composition
├── panels/              # Feature-specific panels
│   ├── Project card
│   ├── Language pairs
│   ├── Additional services
│   └── Activity log
├── dialogs/             # Modal dialogs
│   ├── File selection
│   ├── Settings
│   └── Confirmations
└── models/              # Qt data models
    └── Tables and trees for calculations
```

**Responsibilities:**
- Render project data from `logic/` models
- Capture user input and trigger business operations
- Display progress indicators and validation feedback

### 3. Business Logic Layer (`logic/`)
Contains all core business rules, calculations, and data transformations.

#### Core Modules

**Project Management:**
- `project_manager.py` — Central project state aggregator, coordinates cards, language pairs, services, and calculations
- `project_io.py` — Project serialization/deserialization with version compatibility
- `pm_store.py` — Project manager contact history and auto-population

**Calculations & Financial:**
- `calculations.py` — Volume, currency, discount, tax, and total cost computations
- `online_rates.py` — Currency exchange rates storage and updates

**Import Pipeline:**
- `importers.py` — High-level import facades
- `sc_xml_parser.py` — Smartcat report parser
- `trados_xml_parser.py` — Trados report parser
- `outlook_import/` — Email `.msg` file parser (headers, body, tables)
- `xml_parser_common.py` — Shared XML parsing utilities
- `excel_process.py` — Excel data reading and normalization

**Export Pipeline:**
- `excel_exporter.py` — Commercial proposal generation from templates
- Template-based rendering system

**Reference Data:**
- `language_pairs.py` — Language pair models and user directory management
- `language_codes.py` — Language code normalization with `langcodes`/`Babel`
- `legal_entities.py` + `legal_entities.json` — Legal entity reference data

**Microsoft Graph Integration:**
- `ms_graph_client.py` — Low-level MS Graph API client (MSAL authentication, file operations by path and `fileId`)

**Infrastructure:**
- `service_config.py` — Environment configuration (`.env` loading, paths)
- `user_config.py` — User preferences (default currency, UI language)
- `translation_config.py` — Interface and terminology localization
- `logging_utils.py` — Logging configuration and log file management
- `progress.py` — Long-running operation progress tracking

### 4. Integration Services Layer (`services/`)
Lightweight adapters for external APIs and legacy component support.

**Modules:**
- `excel_export.py` — Export Qt table models to Excel with auto-formatting (supports legacy tabs, LogTab, MemoQ)
- `ms_graph.py` — Compatibility wrapper over `logic.ms_graph_client`, normalizes scopes, error logging, returns `DataFrame` objects

---

## Supporting Components

### Resource Management
**`utils/resource_utils.py`**
- Locates templates and translations
- Handles both development and PyInstaller bundle modes

### Static Resources
**`templates/`**
- Excel templates for commercial proposals
- PDF generation assets

### Update System
**`updater/`**
- Release metadata management
- Update package verification and installation
- Version checking

### Testing
**`tests/`**
- Unit tests for calculations
- Parser validation
- Service function coverage

### Build & Deployment
- `requirements.txt` — External dependencies with version constraints
- `main.spec` — PyInstaller build configuration (binaries, templates, translations)
- `rateapp.ico` — Application icon
- `RateApp.exe.sha256` — Build artifact checksums

---

## Data Flow Architecture

### 1. Import Phase
```
User Action (Drag & Drop / MS Graph)
    ↓
importers.py (Import Facade)
    ↓
Format-Specific Parsers (XML/MSG/Excel)
    ↓
Language Pair Models
    ↓
project_manager.py (Data Aggregation)
```

**Sources:**
- Trados/Smartcat XML reports
- Outlook `.msg` files
- Excel spreadsheets
- SharePoint via Microsoft Graph

### 2. Processing Phase
```
project_manager.py (State Coordination)
    ↓
calculations.py (Cost Computation)
    ↓
online_rates.py (Currency Conversion)
    ↓
project_io.py (State Persistence)
```

**Reference Data Integration:**
- `language_pairs.py` — Language combinations
- `pm_store.py` — Project manager contacts
- `user_config.py` — User preferences
- `legal_entities.py` — Company information

### 3. Export Phase
```
Export Request
    ↓
excel_exporter.py (Template Processing)
    ├→ services/excel_export.py (Legacy Format Support)
    └→ pdf_exporter.py (PDF Generation)
    ↓
Output Files (Excel/PDF)
```

### 4. External Integration
```
Update Checks → updater/ → Release Metadata
    ↓
ms_graph_client.py ← MSAL Auth → Microsoft Graph API
    ↓
services/ms_graph.py (Data Normalization)
```

---

## Design Benefits

### Maintainability
- **Clear boundaries** between UI, logic, and services
- **Independent evolution** of business rules without UI changes
- **Testable components** with minimal mocking requirements

### Extensibility
- **Plugin-ready parser system** for new report formats
- **Service abstraction** allows swapping integrations
- **Template-based exports** enable customization without code changes

### Reusability
- **Service layer** supports both new and legacy components
- **Shared utilities** prevent code duplication
- **Modular importers** can be composed for complex workflows

---

## Technology Stack

- **UI Framework**: PySide6 (Qt for Python)
- **Language Processing**: `langcodes`, `Babel`
- **Excel Operations**: `openpyxl`, `pandas`
- **Authentication**: MSAL (Microsoft Authentication Library)
- **API Integration**: Microsoft Graph API
- **Build Tool**: PyInstaller
- **Testing**: `pytest` (standard unit testing)

---

## Development Guidelines

### Adding New Features
1. Implement business logic in `logic/` with unit tests
2. Create UI components in `gui/` that consume logic models
3. Add integration adapters in `services/` if external APIs are needed
4. Update configuration in `service_config.py` or `user_config.py`

### Adding New Import Formats
1. Create parser in `logic/` (follow existing XML parser patterns)
2. Register parser in `importers.py` facade
3. Add format detection logic
4. Include test cases with sample files

### Modifying Export Templates
1. Update template files in `templates/`
2. Adjust rendering logic in `excel_exporter.py`
3. Test with various project configurations

---

## Future Considerations

- **API Layer**: RESTful API for external integrations
- **Database Backend**: Replace JSON persistence with SQL database
- **Cloud Storage**: Direct integration with cloud storage providers
- **Collaborative Features**: Multi-user project editing
- **Analytics**: Usage tracking and reporting dashboard