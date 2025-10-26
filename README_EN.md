# RateApp

> Professional toolkit for estimating translation projects

RateApp is a Python/PySide6 desktop application that automates the preparation of commercial proposals for translation agencies. It consolidates data from Trados/Smartcat, Outlook, Excel, and SharePoint/OneDrive, calculates multilingual projects with discounts and markups, and produces client-ready deliverables.

---

## ✨ Key Features

### 📋 Project card & contacts
- Store project name, client, contact person, email, legal entity, and billing currency
- Switch the interface language (RU/EN) and report display language independently
- Maintain a project manager directory with history and auto-fill
- Configure VAT rates and "new words only" calculation mode

### 🌍 Language pairs & rates
- Create and edit pairs while keeping user-defined language dictionaries per platform
- Normalise language codes via `langcodes` and `Babel`
- Use the dedicated **Rates Panel** window to review grids, highlight mismatches, and map Excel values
- Apply bulk edits for discounts, markups, and service blocks across multiple pairs

### 📥 Data import
#### CAT reports
- Drag & drop Trados and Smartcat XML reports straight into the main window
- Automatic distribution of volume categories with warning highlights

#### Outlook emails
- Parse `.msg` files (via `extract_msg`) to pre-fill project cards and service tables
- Rebuild the Outlook COM cache on Windows when required for reliable parsing

#### Rate cards
- Load Excel sheets (`R1_*`, `R2_*`) and apply them to the selected language pairs
- Integrate with Microsoft 365 via Microsoft Graph using MSAL authentication by path or `fileId`

### 💰 Financial tooling
- Recalculate totals for every service block with discounts/markups and currency rounding
- Convert currencies with `online_rates.py` while allowing manual overrides
- Manage "Project setup & management" fees with flexible tasks, discounts, and markups
- Track additional services with custom units and grouped tables

### 📤 Export & templates
- Save/load project snapshots in JSON with backward compatibility
- Generate commercial offers and internal rate tables in Excel using branded templates
- Export PDF documents through the Excel COM interface on Windows
- Customise styles and resources stored in `templates/`

### 📊 Logging & diagnostics
- Structured Markdown logs for launches and user actions stored in `logs/` or `~/.smeta/logs`
- Quick access to the latest log via **Project → Open Log**
- Automatic cleanup of orphaned Excel processes on application exit

---

## 🏗️ Project layout

```text
.
├── main.py                      # Application entry point
├── gui/                         # PySide6 widgets and window composition
│   ├── main_window.py           # Main window & menus
│   ├── panels/                  # Left/right panels with forms and tables
│   ├── additional_services.py   # Additional services widget
│   ├── project_setup_widget.py  # Project setup & management costs
│   ├── project_manager_dialog.py# Project manager selector dialog
│   ├── rates_manager_window.py  # Embedded rates management window
│   └── ...
├── logic/                       # Business logic & services
│   ├── calculations.py          # Totals, currency, discounts
│   ├── project_manager.py       # File operations & exports
│   ├── project_data.py          # Structured project snapshot
│   ├── project_io.py            # JSON persistence helpers
│   ├── rates_importer.py        # Excel rate import
│   ├── outlook_import/          # Outlook `.msg` parsing
│   ├── sc_xml_parser.py         # Smartcat report parser
│   ├── trados_xml_parser.py     # Trados report parser
│   ├── ms_graph_client.py       # Microsoft Graph client
│   ├── online_rates.py          # Currency conversion utilities
│   ├── activity_logger.py       # Markdown activity log writer
│   ├── env_loader.py            # `.env` discovery helpers
│   ├── service_config.py        # Environment configuration
│   ├── logging_utils.py         # Logging configuration
│   ├── user_config.py           # User preference storage
│   └── ...
├── services/                    # Integration adapters
│   ├── excel_export.py          # Qt table exports (LogTab, MemoQ, legacy)
│   └── ms_graph.py              # Wrapper around the Graph client
├── rates1/                      # Embedded legacy rates tabs
├── templates/                   # Excel/PDF templates and assets
├── utils/                       # UI helpers (history, theme)
├── tests/                       # Pytest suite
├── requirements.txt             # Dependency list
├── main.spec                    # PyInstaller configuration
└── resource_utils.py            # Resource lookup for dev & bundles
```

A detailed architecture description is available in [ARCHITECTURE_EN.md](ARCHITECTURE_EN.md).

---

## 🚀 Getting started

### Prerequisites
- Python 3.11

### Installation

1. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python main.py
   ```

### Additional requirements

- **Windows + Excel** — needed for PDF export and COM interactions (`pywin32`).
- **Outlook `.msg`** — handled through the bundled `extract_msg` dependency.
- **Internet access** — required for Microsoft 365 rate downloads and manual currency updates.

---

## 🧪 Testing

The project uses `pytest`:

```bash
pytest
```

Tests cover calculations, XML/Outlook parsing, configuration helpers, and import/export utilities.

---

## 📦 Technology stack

- **Python 3.11** — primary runtime
- **PySide6** — Qt-based desktop UI
- **pandas** & **openpyxl** — tabular calculations and Excel I/O
- **Pillow** — image/icon handling
- **langcodes** & **Babel** — localisation and language code normalisation
- **requests** — Microsoft Graph networking
- **msal** — Microsoft 365 authentication
- **psutil** — process diagnostics and Excel shutdown helpers
- **extract_msg** — Outlook `.msg` parsing
- **python-dotenv** — environment configuration loading
- **pywin32** *(Windows only)* — COM interop for PDF export

---

## 💾 User data storage

### Configuration files

**`languages.json`** — custom language directory, stored in:
- Windows: `%APPDATA%/ProjectCalculator`
- macOS: `~/Library/Application Support/ProjectCalculator`
- Linux: `~/.config/ProjectCalculator`

**`pm_history.json`** — project manager history, stored alongside `languages.json`.

**`logic/legal_entities.json`** — bundled legal entity reference shipped with the app.

---

## 🔨 Building a standalone bundle

Use the provided PyInstaller spec:

```bash
pyinstaller main.spec
```

The spec file packages templates, localisation data (Babel/langcodes), and the application icon. Adjust resource lists or output names as needed.

---

## 📝 Logs & feedback

- Access the latest log via **Project → Open Log**; files rotate between `./logs`, `~/.smeta/logs`, or the system temp directory.
- The structured activity log (`activity.log`) records major actions together with optional data snapshots.

---

## 🤝 Contributing

Pull requests are welcome. Before submitting changes:

1. Run `pytest` and ensure the suite passes.
2. Update the documentation if UI behaviour or import/export flows change.

When reporting bugs, please include:
- application version;
- reproduction steps;
- relevant log excerpts (if possible).
