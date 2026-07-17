# Home library inventory template (marketing asset)

Public-facing spreadsheet template that doubles as a zero-friction Bookplate
CSV import. Built 2026-07-17; round-trip verified against `CsvImportService`
in the inventorymgmt test env (4/4 rows, all columns mapped, enrichment off).

## Files

- `home-library-inventory.xlsx` — three tabs: My library (data + dropdowns,
  frozen header, autofilter, ISBN column forced to text), Lists (dropdown
  vocabularies the user edits), Start here (four steps + one Bookplate line).
  Palette: Burgundy #561A27 headers for book columns, Coral #D75340 for the
  your-copy columns (Location, Notes), per the Bookplate style guide.
- `home-library-inventory.csv` — flat version, same headers and starter rows.
- `landing-copy.md` — draft page copy plus internal caveats.

## Design constraints (deliberate)

- Headers are generic on purpose and all map through the importer's aliases.
  Three color groups mirror the app's structure: the book (burgundy: ISBN,
  Title, Author, Genre, Format, Publisher, Publication year, Pages), your
  copy (coral: Location, Acquired → `books.acquired_on`), your reading
  (green: Status, Rating, Date read, My notes → the importing user's
  journal, cloning the Goodreads-import semantics; loose status words like
  "finished" normalize, a date with no status implies Read). Do not rename
  headers. Dates accept year-only ("2015") or full ("2019-12-25"). No
  copy-level Notes column (ruling 2026-07-17: no destination exists; My
  notes is reading notes). Tags deliberately absent, same reasoning as the
  Goodreads/LibraryThing importers: shelf tags don't map to Bookplate's
  defined-list tags. Extra user-added columns are harmless; unknown headers
  are ignored on import.
- Format vocabulary is exactly Hardcover / Softcover because
  `CatalogBook::BINDINGS` allows only those two values, and any other value
  crashes the whole import (RecordInvalid, not a row error). See pre-publish
  blockers below.
- Physical-only on purpose: e-book and audiobook rows are skipped on import
  for free organizations, so the public template avoids the media format
  column entirely. The upgrade story belongs in landing copy, not the sheet.
- No condition/quantity/price columns (business import shape stays private),
  no header-alias hints, nothing that implies fetch/correction behavior.

## Pre-publish status (updated 2026-07-17, same day)

1. FIXED: importer crash on unknown binding values ("Paperback" killed the
   whole file). `CatalogBook.normalize_binding` now collapses loose values.
2. FIXED: Bookplate had no generic library import at all (only the business
   side did). New /library_import page (Add books card) accepts CSV and
   .xlsx directly, so both template files import as-is once deployed.
3. RESOLVED 2026-07-17: Notes column dropped from the template (Serene's
   call). Copy-level notes don't exist as a product concept (`books` has no
   notes column; reading notes live in the Journal). If copy-level notes
   ever ship, re-add the column here and map it on import.

## Regenerating

`make_template.py` (session scratchpad; needs a venv with openpyxl).
Copy it next to this README if this template needs future edits.
