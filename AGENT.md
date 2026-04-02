# AGENT.md

## Project purpose
Linx2DOI is a desktop PyQt6 application.

It parses references from Word (.docx) documents, resolves DOIs using external APIs (CrossRef, PubMed, PMC), and generates:
- APA formatted bibliography (HTML)
- RIS export
- APA docx publication list

The project is a GUI application with a processing engine.
It is NOT a library, web service, or CLI utility.

---

## Core behavior constraints

The agent must implement requested changes with minimal side effects.

Never:
- rewrite working modules
- refactor architecture unless explicitly requested
- introduce new frameworks
- change UI toolkit (must remain PyQt6)
- change processing order in DOI resolution
- remove incremental processing behavior

Prefer:
- local edits
- additive changes
- preserving public method signatures

---

## Architecture

src/core — pure processing logic
src/gui — UI only

Rules:
- gui may import core
- core must not import gui
- worker is boundary layer (thread orchestration)

Processing flow:

DocumentParser → DOIResolver → CitationFormatter → HTMLGenerator / RISExporter
                    ↓
            PubMedSearcher (optional fallback)

WorkerThread coordinates modules and preserves previous results.

---

## Critical invariants

1. Manual DOI has highest priority
2. Incremental processing must preserve unchanged records
3. Output order must match source document numbering
4. API rate limiting must remain functional
5. UI updates only via Qt signals
6. NCBI requests require user email

Breaking any of these is a bug.

---

## DOI resolution priority (must not change)

1 Manual DOI
2 DOI regex in text
3 DOI URL parsing
4 PubMed PMID (from URL)
5 PMC PMCID (from URL)
6 ResearchGate title extraction → CrossRef
7 CrossRef title search
8 PubMed bibliography search (parse reference → search by title+authors)

---

## Editing rules per module

document_parser.py
Responsible only for extracting structured items from .docx
Do not add API logic here

doi_resolver.py
Network logic only
Do not add formatting or UI logic
May import and use pubmed_searcher as fallback

pubmed_searcher.py
PubMed bibliography search only
- Parse reference (GOST, APA, Vancouver, Hybrid)
- Search PubMed by title + authors
- Fuzzy matching with normalization
Do not modify document or UI state

citation_formatter.py
Only metadata → APA formatting

html_generator.py  
Presentation only  
Do not perform network requests

worker.py  
Coordinates modules and preserves state  
Be extremely careful: partial processing relies on previous_items

main_window.py  
UI logic only  
No blocking operations allowed

---

## Threading rules

WorkerThread must:
- perform all heavy work
- communicate via Qt signals only
Direct UI modification from worker is forbidden

---

## Validation

Manual testing workflow:

- mixed URL types (doi, pubmed, pmc, researchgate)
- missing DOI
- large document (rate limit)
- manual DOI override
- selective processing

Expected behavior:
unchanged records retain previous results

---

## Code style constraints

Python 3.11  
Relative imports inside src package only  
English identifiers, Russian UI text  
No global state  
Logging instead of print

---

## Typical tasks

Allowed:
- add DOI source
- adjust parsing edge case
- modify HTML output
- add export field
- fix resolution bug
- add PubMed search enhancement

Not allowed:
- redesign architecture
- merge modules
- move logic into GUI
- convert to async framework

---

## When uncertain
Ask for clarification instead of refactoring.

