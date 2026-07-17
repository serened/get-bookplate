"""Generate the Bookplate home-library-inventory XLSX template.

Palette (Bookplate style guide, exactly four + ink):
burgundy 561A27, coral D75340, green A4A056, cream F2EEE8, ink 1A120A.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

BURGUNDY = "561A27"
CORAL = "D75340"
GREEN = "A4A056"
CREAM = "F2EEE8"
INK = "1A120A"

OUT_DIR = os.path.join(os.path.dirname(__file__), "library-template")
os.makedirs(OUT_DIR, exist_ok=True)

HEADERS = [
    "ISBN", "Title", "Author", "Genre", "Format",
    "Publisher", "Publication year", "Pages", "Location", "Notes",
]
# Columns A-H describe the book; I-J describe your copy.
BOOK_COLS = 8
WIDTHS = [17, 32, 22, 16, 14, 22, 16, 8, 16, 36]

STARTERS = [
    ["9780141439518", "Pride and Prejudice", "Jane Austen", "Fiction",
     "Softcover", "Penguin Classics", 2002, 480, "Living room", ""],
    ["9780547928227", "The Hobbit", "J. R. R. Tolkien", "Fantasy",
     "Softcover", "Mariner Books", 2012, 300, "Living room", "Gift from Dad"],
    ["9780451524935", "1984", "George Orwell", "Fiction",
     "Softcover", "Signet Classics", 1961, 328, "Office", ""],
    ["9781400033416", "Beloved", "Toni Morrison", "Fiction",
     "Hardcover", "Vintage", 2004, 324, "Bedroom", "Book club pick, May 2024"],
]

GENRES = ["Fiction", "Nonfiction", "Mystery", "Science fiction", "Fantasy",
          "Romance", "Biography", "History", "Poetry", "Children's"]
LOCATIONS = ["Living room", "Bedroom", "Office", "Hallway", "Storage"]
FORMATS = ["Hardcover", "Softcover"]

LAST_ROW = 500  # dropdowns and ISBN text format reach this far

wb = Workbook()

header_font = Font(bold=True, color=CREAM, size=11)
book_fill = PatternFill("solid", fgColor=BURGUNDY)
copy_fill = PatternFill("solid", fgColor=CORAL)
thin_cream = Border(bottom=Side(style="thin", color=CREAM))

# ---- My library ----
ws = wb.active
ws.title = "My library"

for i, h in enumerate(HEADERS, start=1):
    c = ws.cell(row=1, column=i, value=h)
    c.font = header_font
    c.fill = book_fill if i <= BOOK_COLS else copy_fill
    c.alignment = Alignment(vertical="center")
    c.border = thin_cream
    ws.column_dimensions[chr(64 + i)].width = WIDTHS[i - 1]
ws.row_dimensions[1].height = 22

for r, row in enumerate(STARTERS, start=2):
    for i, v in enumerate(row, start=1):
        ws.cell(row=r, column=i, value=v)

# ISBNs must stay text so ISBN-10 leading zeros survive
for r in range(2, LAST_ROW + 1):
    ws.cell(row=r, column=1).number_format = "@"

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:J{LAST_ROW}"

def dropdown(target_range, list_range):
    dv = DataValidation(type="list", formula1=f"'Lists'!{list_range}",
                        allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(target_range)

dropdown(f"D2:D{LAST_ROW}", "$A$2:$A$41")   # Genre
dropdown(f"E2:E{LAST_ROW}", "$C$2:$C$41")   # Format
dropdown(f"I2:I{LAST_ROW}", "$B$2:$B$41")   # Location

# ---- Lists ----
ls = wb.create_sheet("Lists")
for col, (title, values) in enumerate(
        [("Genres", GENRES), ("Locations", LOCATIONS), ("Formats", FORMATS)],
        start=1):
    c = ls.cell(row=1, column=col, value=title)
    c.font = header_font
    c.fill = book_fill
    c.border = thin_cream
    for r, v in enumerate(values, start=2):
        ls.cell(row=r, column=col, value=v)
    ls.column_dimensions[chr(64 + col)].width = 20
ls.row_dimensions[1].height = 22
note = ls.cell(row=1, column=5,
               value="These lists feed the dropdowns on the My library tab. "
                     "Edit them to match your shelves.")
note.font = Font(italic=True, color=INK)
ls.column_dimensions["E"].width = 60

# ---- Start here ----
sh = wb.create_sheet("Start here")
sh.sheet_view.showGridLines = False
sh.column_dimensions["B"].width = 90
title = sh.cell(row=2, column=2, value="Home library inventory")
title.font = Font(name="Georgia", size=18, bold=True, color=BURGUNDY)
steps = [
    "1. Add one row per book on the My library tab. The ISBN is on the back cover or the copyright page.",
    "2. Genre, Format, and Location are dropdowns. Edit the choices on the Lists tab to match your shelves.",
    "3. Keep the header row as it is. Sort and filter from row 2 down.",
    "4. The example rows show the idea. Replace them with your own books.",
]
for i, s in enumerate(steps, start=4):
    c = sh.cell(row=i, column=2, value=s)
    c.font = Font(size=12, color=INK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    sh.row_dimensions[i].height = 20
credit = sh.cell(row=10, column=2,
                 value="Made by Bookplate (bookplate.app). If the spreadsheet ever starts to feel "
                       "like homework, import this exact file and your covers and book details "
                       "fill in for you.")
credit.font = Font(size=11, italic=True, color=CORAL)
credit.alignment = Alignment(wrap_text=True, vertical="top")
sh.row_dimensions[10].height = 34

path = os.path.join(OUT_DIR, "home-library-inventory.xlsx")
wb.save(path)
print("wrote", path)
