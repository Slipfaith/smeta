# RateApp

> Professional translation project cost calculation tool

RateApp is a Python/PySide6 desktop application that automates the preparation of commercial proposals for translation agencies. The application aggregates data from multiple sources (Trados, Smartcat, Outlook, Excel, SharePoint) and enables rapid generation of accurate cost estimates with discounts, markups, and additional services.

---

## ✨ Key Features

### 📋 Project Management
- **Project Card** — Complete client information, contacts, legal entity, and currency
- **Multilingual Interface** — Switch between Russian and English
- **Flexible VAT Configuration** — Support for various tax rates

### 🌍 Language Pairs
- **Custom Directories** — Create and manage your own language lists
- **Automatic Normalization** — Correct language code recognition via `langcodes` and `Babel`
- **New Words Mode** — Calculate costs based only on new words and repetitions
- **Cross-platform Storage** — Settings synchronization across devices

### 📥 Smart Data Import

#### CAT System Reports
- **Drag & Drop Interface** — Simply drag XML files into the application window
- **Supported Formats:**
  - Trados Studio (all versions)
  - Smartcat
- **Automatic Distribution** — Volume categories detected automatically
- **Data Validation** — Highlighting of warnings and potential errors

#### Outlook Emails
- **`.msg` File Parsing** — Extract information from client correspondence
- **Auto-fill** — Client data transferred to project card
- **HTML Table Processing** — Recognition of structured data in email body

#### Rate Cards
- **Excel Import** — Load rates from corporate spreadsheets (`R1_*`, `R2_*` sheets)
- **Multi-currency Support** — Price lists in different currencies
- **Microsoft 365 Integration:**
  - Direct access to SharePoint/OneDrive files
  - OAuth 2.0 authentication via MSAL
  - File search by path or `fileId`
  - Interactive browser-based authorization

### 💰 Financial Calculations

- **Current Exchange Rates** — Automatic updates or manual input
- **Flexible Discount System** — Individual terms for each client
- **Markups and Services** — Add additional work with arbitrary units of measurement
- **Transparent Calculations** — Detailed breakdown of all cost components
- **Smart Rounding** — Correct handling of cents and kopecks

### 📤 Professional Export

#### Output Formats
- **Excel Proposals** — Formatted commercial offers based on templates
- **PDF Documents** — Ready for client delivery (requires Excel on Windows)
- **JSON Projects** — Save and load current work state
- **Operational Tables** — Internal calculation forms for team

#### Template Capabilities
- **Corporate Branding** — Logos, colors, company style
- **Customizable Sections** — Add arbitrary sections
- **Auto-formatting** — Currencies, numbers, percentages per regional standards
- **Specialized Formats:**
  - LogTab export
  - MemoQ reports
  - Styled sheets with conditional formatting

### 🔄 Automatic Updates

- **Release Monitoring** — Check for new versions on GitHub
- **Smart Installation:**
  - Windows: automatic installer launch
  - macOS/Linux: notification with download link
- **Security** — Digital signature verification for updates

### 📊 Monitoring and Debugging

- **Detailed Logging** — Record all operations for diagnostics
- **Quick Access** — Open latest log via application menu
- **File Rotation** — Automatic cleanup of old logs
- **Flexible Location** — Choose directory for log storage

---

## 🏗️ Architecture

### Project Structure

```
RateApp/
│
├── 📱 main.py                      # Application entry point
│
├── 🎨 gui/                         # User interface
│   ├── main_window.py             # Main window
│   ├── panels/                    # Functional panels
│   │   ├── project_card.py        # Project card
│   │   ├── language_pairs.py      # Language pair management
│   │   ├── services.py            # Additional services
│   │   └── log.py                 # Event log
│   ├── dialogs/                   # Modal windows
│   └── models/                    # Data models for UI
│
├── 🧠 logic/                       # Business logic
│   ├── outlook_import/            # Outlook email parsing
│   ├── calculations.py            # Financial calculations
│   ├── excel_exporter.py          # Excel proposal generation
│   ├── pdf_exporter.py            # PDF export
│   ├── importers.py               # XML/MSG/Excel import
│   ├── rates_importer.py          # Rate loading
│   ├── ms_graph_client.py         # Microsoft Graph API
│   ├── online_rates.py            # Currency exchange rates
│   ├── project_manager.py         # Project management
│   ├── project_io.py              # Project serialization
│   ├── translation_config.py      # Interface localization
│   ├── user_config.py             # User settings
│   ├── legal_entities.py          # Legal entity directory
│   └── language_codes.py          # Language normalization
│
├── 🔌 services/                    # Integration services
│   ├── excel_export.py            # Table export
│   └── ms_graph.py                # MS Graph wrapper
│
├── 📦 templates/                   # Export resources
│   ├── excel/                     # Excel templates
│   ├── images/                    # Logos and images
│   └── fonts/                     # PDF fonts
│
├── 🔄 updater/                     # Update system
│   ├── update_checker.py          # Version checking
│   └── release_metadata.py        # Release metadata
│
├── 🧪 tests/                       # Test coverage
│   ├── test_calculations.py
│   ├── test_parsers.py
│   └── test_importers.py
│
├── ⚙️ utils/                       # Helper utilities
│   └── resource_utils.py          # Resource management
│
├── 📄 requirements.txt             # Project dependencies
├── 🔧 main.spec                    # PyInstaller configuration
├── 📖 ARCHITECTURE.md              # Architecture documentation
└── 📝 README.md                    # This file
```

**Detailed Description:** See [ARCHITECTURE.md](ARCHITECTURE.md) for in-depth explanation of layers, modules, and data flows.

---

## 🚀 Quick Start

### System Requirements

- **Python**: 3.11 or higher
- **OS**: Windows 10/11, macOS 10.15+, Linux (Ubuntu 20.04+)
- **Memory**: minimum 4 GB RAM
- **Disk Space**: 500 MB

### Installation from Source

#### 1. Clone the Repository
```bash
git clone https://github.com/your-org/rateapp.git
cd rateapp
```

#### 2. Create Virtual Environment
```bash
python -m venv .venv

# Activation:
# Windows:
.venv\Scripts\activate

# macOS/Linux:
source .venv/bin/activate
```

#### 3. Install Dependencies
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

#### 4. Run Application
```bash
python main.py
```

### Additional Components

#### For Windows Users
```bash
# PDF export and Excel COM object interaction
pip install pywin32
```

**Requirements:**
- Microsoft Excel installed (any version)
- Active Office license

#### For Developers
```bash
# Development and testing tools
pip install pytest pytest-cov black flake8 mypy
```

---

## 🧪 Testing

### Running Tests

```bash
# All tests
pytest

# With code coverage
pytest --cov=logic --cov=services --cov-report=html

# Specific module
pytest tests/test_calculations.py

# With verbose output
pytest -v
```

### Test Coverage

Tests cover:
- ✅ Financial calculations and rounding
- ✅ XML parsers (Trados, Smartcat)
- ✅ Outlook email import
- ✅ Language code normalization
- ✅ Configuration utilities
- ✅ Data export and formatting
- ✅ User input validation

---

## 📦 Technology Stack

### Core Libraries

| Component | Technology | Purpose |
|-----------|------------|---------|
| **UI Framework** | PySide6 6.5+ | Qt-based graphical interface |
| **Data Processing** | pandas 2.0+ | Tabular calculations and transformations |
| **Excel I/O** | openpyxl 3.1+ | Read and write Excel with formatting |
| **Image Processing** | Pillow 10.0+ | Logo and image handling |
| **Localization** | langcodes, Babel | Language normalization and localization |
| **HTTP Client** | requests 2.31+ | API requests and update downloads |
| **Authentication** | msal 1.24+ | OAuth 2.0 for Microsoft 365 |
| **Email Parsing** | extract_msg 0.45+ | Parse Outlook `.msg` files |
| **System Utils** | psutil 5.9+ | Process and resource monitoring |
| **Configuration** | python-dotenv 1.0+ | Environment variable management |

### Platform Dependencies

| Platform | Library | What It's For |
|----------|---------|---------------|
| **Windows** | pywin32 | Excel COM interface, PDF export |
| **macOS** | - | Native support via Qt |
| **Linux** | - | Native support via Qt |

---

## 💾 Data Storage

### User Settings

#### Configuration File Locations

| File | Contents | Path |
|------|----------|------|
| `languages.json` | Language directory | See below |
| `pm_history.json` | Manager history | See below |
| `user_settings.json` | Interface settings | See below |

**Directories by Platform:**

- **Windows**: `%APPDATA%\ProjectCalculator`
- **macOS**: `~/Library/Application Support/ProjectCalculator`
- **Linux**: `~/.config/ProjectCalculator`

#### System Data

| File | Purpose | Location |
|------|---------|----------|
| `legal_entities.json` | Legal entities | Application directory |
| `templates/` | Export templates | Application directory |

### Log Files

Logs are saved in the first available directory:
1. `./logs` (project directory)
2. `~/.smeta/logs` (home directory)
3. System temporary directory

**Access:** Menu → Project → Open Log

---

## 🔨 Building Application

### PyInstaller Build

```bash
# Simple build
pyinstaller main.spec

# With cleanup of previous artifacts
pyinstaller --clean main.spec

# With verbose output
pyinstaller --clean --log-level=DEBUG main.spec
```

### `main.spec` Structure

The spec file automatically includes:
- ✅ Excel/PDF templates from `templates/`
- ✅ Babel and langcodes localization data
- ✅ Application icon and metadata
- ✅ Legal entity and language directories
- ✅ Configuration files

### Output Artifacts

After building, `dist/` will contain:
- `RateApp.exe` (Windows) or `RateApp` (macOS/Linux)
- Required libraries and dependencies
- Packaged resources

### Splash Screen

Use the lightweight [`splash.py`](splash.py) launcher to display a `.mov`
animation immediately while the main application starts in the background. The
window closes automatically once the main GUI is ready.

```bash
# Example: run the splash screen from sources
python splash.py -- --animation /path/to/splash.mov python main.py
```

By default the script looks for `templates/splash.mov`. When building with
PyInstaller, create two separate executables (`Splash.exe` and `RateApp.exe`)
so that the splash screen can start instantly and hand over control to the main
program.

---

## 🔄 Update System

### Automatic Checking

**Menu → Updates → Check for Updates**

1. Application queries GitHub Releases
2. Compares current version with latest available
3. Offers download if update exists

### Installing Updates

#### Windows
- Automatic `.exe` installer download
- Digital signature verification
- Launch installation with settings preservation

#### macOS/Linux
- Notification about new version availability
- Link to download archive
- Installation instructions

### Manual Update

```bash
# Download latest version
git pull origin main

# Update dependencies
pip install --upgrade -r requirements.txt

# Restart application
python main.py
```

---

## 📚 Documentation

- **[ARCHITECTURE.md](ARCHITECTURE.md)** — Detailed architecture description