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

- Headers are generic on purpose and all map through the importer's aliases:
  ISBN, Title, Author, Genre, Format, Publisher, Publication year, Pages,
  Location, Notes. Do not rename them.
- Format vocabulary is exactly Hardcover / Softcover because
  `CatalogBook::BINDINGS` allows only those two values, and any other value
  crashes the whole import (RecordInvalid, not a row error). See pre-publish
  blockers below.
- Physical-only on purpose: e-book and audiobook rows are skipped on import
  for free organizations, so the public template avoids the media format
  column entirely. The upgrade story belongs in landing copy, not the sheet.
- No condition/quantity/price columns (business import shape stays private),
  no header-alias hints, nothing that implies fetch/correction behavior.

## Pre-publish blockers (app-side, need go-ahead)

1. Importer crash on unknown binding values ("Paperback" kills the file,
   including real Goodreads exports, which use that value). Normalize or
   row-error it before pointing the public at CSV import.
2. Notes column is accepted by the sheet but silently dropped by the
   importer. Either add a notes mapping to `CsvImportService` or cut the
   promise carefully (current landing copy stays silent on notes).

## Regenerating

`make_template.py` (session scratchpad; needs a venv with openpyxl).
Copy it next to this README if this template needs future edits.
