# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Linx2DOI** is a PyQt6-based desktop application that extracts DOI (Digital Object Identifiers) from Word documents and generates APA-formatted bibliographies. The name is a play on words: "links" (references/URLs) + "DOI".

The application processes references with URLs, automatically searches for DOIs via multiple sources (CrossRef, PubMed, PMC), and allows manual DOI entry for records that cannot be auto-resolved.

## Project Structure

```
Linx2DOI/
├── main.py                     # Entry point
├── requirements.txt            # Dependencies
├── icon.ico                    # Application icon
├── README.md                   # User documentation
├── USER_GUIDE.md               # Detailed user guide
├── CHANGELOG.md                # Version history
├── ROADMAP.md                  # Future development plan
├── src/
│   ├── __init__.py
│   ├── core/                   # Business logic
│   │   ├── __init__.py
│   │   ├── config.py           # Configuration constants
│   │   ├── document_parser.py  # .docx parsing
│   │   ├── doi_resolver.py     # DOI resolution strategies
│   │   ├── citation_formatter.py # APA citation formatting
│   │   ├── html_generator.py   # HTML report generation
│   │   ├── ris_exporter.py     # RIS format export (NEW v1.5)
│   │   └── worker.py           # Background worker thread
│   └── gui/                    # GUI components
│       ├── __init__.py
│       ├── main_window.py      # Main application window
│       └── delegates.py        # Table cell delegates
└── main_old.py                 # Original monolithic file (backup)
```

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```

## Module Responsibilities

### Core Modules

**config.py**
- API endpoints (CrossRef, NCBI E-utilities)
- Application settings and constants
- Retry/timeout parameters

**document_parser.py** (`DocumentParser` class)
- Extracts numbered references from .docx files
- Parses both paragraphs and table cells
- Expected format: `1. Article Title..... https://url.com`
- Methods: `collect_items_from_document()`, `extract_item_from_text()`

**doi_resolver.py** (`DOIResolver` class)
- Resolves DOIs using multiple strategies (priority order):
  1. Manual DOI input (highest priority)
  2. Text extraction via regex
  3. DOI URL parsing (doi.org links)
  4. PubMed API lookup (PMID)
  5. PMC API lookup (PMCID)
  6. ResearchGate title extraction → CrossRef search
  7. Fallback: CrossRef search by article title
- Requires user email for NCBI API compliance
- Method: `resolve_doi(item, manual_doi, auto_search)`

**citation_formatter.py** (`CitationFormatter` class)
- Retrieves article metadata from CrossRef API
- Formats author lists per APA 7th edition rules
- Generates complete APA citations with HTML markup
- Methods: `get_apa_citation(doi)`, `format_authors(authors)`

**html_generator.py** (`HTMLGenerator` class)
- Generates self-contained HTML reports
- Color-coded status badges (no data: red, manual DOI: green, duplicate: orange)
- Detects duplicate DOIs and marks them with "ДУБЛЬ № X" status
- Statistics summary including duplicate count
- Clickable article titles (links to DOI)
- Methods: `generate_html_ordered(items, filename, email)`, `generate_table_data(items)`

**ris_exporter.py** (`RISExporter` class) - **NEW in v1.5**
- Exports bibliographic data to RIS format for Mendeley/Zotero/EndNote
- Parses metadata from APA citations
- Extracts authors, title, journal, year, volume, issue, pages, DOI
- Methods: `generate_ris(items, filename)`, `save_ris_file(items, filepath, source_filename)`
- Helper methods for metadata extraction: `_extract_authors()`, `_extract_journal()`, etc.

**worker.py** (`WorkerThread` class)
- Qt background thread for document processing
- Three processing modes:
  1. Full document processing
  2. Selective reprocessing (user-selected rows)
  3. Manual DOI-only processing (preserves unchanged records)
- Emits signals for progress updates, logging, and results
- Coordinates all core modules

### GUI Modules

**delegates.py**
- `HyperlinkDelegate` - Custom Qt table delegate for clickable hyperlinks, opens URLs/DOIs in default browser
- `ManualDOIDelegate` - **NEW in v1.5** - Custom delegate for manual DOI input with top alignment

**main_window.py** (`MainWindow` class)
- PyQt6 main window with tabbed interface
- Two color themes: `AppColors` (dark) and `LightColors` (light, default)
- Theme switching button in header
- Email input, file selection, results table, progress bar, log area
- Table features:
  - Editable "Manual DOI" column with top-aligned text
  - Hidden "Select" column (reserved for future features)
  - Clickable column headers (click "✓" to select/deselect all)
  - Auto-resizing columns
- Buttons:
  - Start Processing
  - Apply Manual DOI
  - Open Result (HTML)
  - Export to RIS (NEW v1.5)
  - Theme Toggle (NEW v1.5)
  - Help
- Connects worker signals to UI updates
- Methods: `toggle_theme()`, `export_to_ris()`, `apply_theme()`

## Architecture Patterns

### State Preservation
When processing manual DOIs or selected items, the application preserves previous results for unchanged records. This is critical for incremental updates:

```python
# In worker.py
self.previous_items = previous_items if previous_items else []
# Unmodified records retain their previous state
```

### Signal-Based Communication
Worker thread uses Qt signals to communicate with GUI thread (thread-safe):
- `progress` → progress bar updates
- `log` → log message appends
- `table_data` → table population
- `finished_success` / `finished_error` → completion handling

### API Rate Limiting
Both DOIResolver and CitationFormatter implement:
- Exponential backoff on HTTP 429 responses
- Configurable delays between requests (`RATE_LIMIT_DELAY`)
- Retry logic with max attempts (`MAX_RETRIES`)

## Development Workflow

### Adding a New DOI Source

1. Add extraction method to `doi_resolver.py`:
   ```python
   def extract_xxx_from_url(self, url):
       # Parse URL and extract identifier
       pass
   ```

2. Add to resolution chain in `resolve_doi()`:
   ```python
   if not dois:
       self.log(f"  🔍 Checking XXX...")
       xxx_id = self.extract_xxx_from_url(item['url'])
       if xxx_id:
           doi = self.get_doi_from_xxx_api(xxx_id)
           if doi:
               dois.add(doi)
   ```

### Modifying HTML Output

Edit `html_generator.py`:
- CSS styles in `generate_html_ordered()` method
- Item rendering logic in the loop over `items`
- Statistics calculation at the top of the method

### Changing Processing Logic

Modify `worker.py`:
- `process_item()` for single-item processing
- `process_only_manual_items()` for manual DOI workflow
- `run()` for overall orchestration

## Testing Strategy

No automated tests are present. Manual testing workflow:

1. Create test .docx with references containing:
   - Direct DOI links
   - PubMed URLs (`pubmed.ncbi.nlm.nih.gov/PMID`)
   - PMC URLs (`pmc.ncbi.nlm.nih.gov/articles/PMCID`)
   - ResearchGate URLs
   - Plain text with article titles

2. Verify:
   - All items extracted correctly (numbered list format)
   - DOI resolution works for each URL type
   - Manual DOI input overrides automatic results
   - HTML output maintains original order
   - Selective re-parsing preserves unmodified items

3. Test edge cases:
   - Missing DOIs (verify "no data" status)
   - Malformed URLs
   - Rate limiting (large documents)
   - Invalid email format

## Common Pitfalls

1. **Email Requirement**: NCBI API requires a valid email. Application validates format before processing.

2. **Rate Limiting**: CrossRef and NCBI have rate limits. Use exponential backoff and delays between requests.

3. **Document Format**: Parser expects numbered list format (`1. Title..... URL`). Free-form text won't be extracted.

4. **State Management**: When modifying worker.py, ensure `previous_items` is passed correctly to preserve unmodified records during partial updates.

5. **Thread Safety**: UI updates MUST use Qt signals. Direct UI manipulation from `WorkerThread` will cause crashes.

6. **Module Imports**: Use relative imports within src/ package:
   ```python
   from ..core.config import SETTINGS_ORG  # Correct
   from src.core.config import SETTINGS_ORG  # Incorrect for package code
   ```

## Code Style

- Comments and UI labels are in Russian
- Variable/function names are in English
- Email stored in QSettings (key: "user_email", persistent across sessions)
- Stylesheet definitions inline within widget creation
- Log messages use emoji prefixes for visual parsing (🔍, ✅, ❌, 📊)

## Migration from Monolithic Version

The original `main.py` has been preserved as `main_old.py`. The refactored version maintains identical functionality with improved maintainability:

- **Single Responsibility**: Each module handles one concern
- **Testability**: Core logic isolated from GUI
- **Reusability**: DOIResolver and CitationFormatter can be used independently
- **Readability**: Smaller, focused files instead of 1800+ line monolith
