# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))

from cellartracker import cellartracker
from ctdataframe import CTDataFrame
from itertools import cycle
#from matplotlib import pyplot as plt
#from matplotlib import style
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter


from ctconfig import CELLAR_USERNAME, CELLAR_PASSWORD, COLOR_HEADER, COLOR_ALT1, COLOR_ALT2, FONT_NAME, FONT_SIZE, IMAGE_NAME, MAX_BOTTLE_CHECKBOXES, SPACE_PADDING, TEXT_PADDING, COLUMN_NAMES, CURRENT_YEAR

# Get the data from CellarTracker and convert it to a DataFrame
CT_DATA : list = cellartracker.CellarTracker(CELLAR_USERNAME, CELLAR_PASSWORD).get_list()
ct_data_frame = CTDataFrame(CT_DATA)

# TODO: Move the following to ctworksheet.py or similar, to separate concerns and improve maintainability.
#
#  Set up the Excel workbook and worksheet
output_filename = "CellarTracker.xlsx"
sheet_name = "Cellar Tracker"
# with Workbook(output_filename) as workbook:
workbook = Workbook()
worksheet = workbook.active
# Set up workbook and defaults
assert worksheet is not None
worksheet.title = sheet_name
worksheet.page_setup.draft = False
worksheet.page_setup.blackAndWhite = False
worksheet.page_setup.fitToHeight = 0
worksheet.page_setup.fitToWidth = 1
worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
worksheet.page_setup.paperSize = worksheet.PAPERSIZE_TABLOID
worksheet.page_margins.header = 0.3
worksheet.page_margins.footer = 0.3
worksheet.page_margins.left = 0.5
worksheet.page_margins.right = 0.5
worksheet.page_margins.top = 0.5
worksheet.page_margins.bottom = 0.5
worksheet.sheet_properties.pageSetUpPr.fitToPage = True

#  Set up the header row with bold font and background color
worksheet.append(COLUMN_NAMES)

colors = cycle([COLOR_ALT1, COLOR_ALT2])
WHITE = "FFFFFFFF"
BLACK = "FF000000"
RED = "FFFF0000"
YELLOW = "FFFFFF00"
GREEN = "FF00FF00"
MAUVE = "FFCC66CC"
FILL_TYPE = "solid"
thickness = Side(border_style="thick", color=BLACK)
cell_border = Border(
    left=thickness, right=thickness, top=thickness, bottom=thickness
)
bold_font = Font(name=FONT_NAME, bold=True, color=BLACK, size=FONT_SIZE)
for cell in worksheet[1]: # Assuming the first row is row 1
    cell.font = bold_font
    #cell.fill = PatternFill(start_color=BLACK, end_color=COLOR_HEADER, fill_type=FILL_TYPE)
    cell.fill = PatternFill(start_color=GREEN, fill_type=FILL_TYPE)
    cell.border = cell_border


# worksheet.set_row(0, cell_format=data_header)
COL_WIDTHS = [0] * len(COLUMN_NAMES)
for col, header in enumerate(COLUMN_NAMES):
    worksheet.cell(row=1, column=col + 1, value=header)
    COL_WIDTHS[col] = len(header)

for col in worksheet.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)  # Get the column letter (e.g., 'A')
    
    # 3. Determine the maximum length of values in the current column
    for cell in col:
        if cell.value is not None:
            max_len = max(max_len, len(str(cell.value)))
            
    # 4. Set the column width with some extra padding
    # Adding a safety margin (e.g., +3) keeps text readable and uncompressed
    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)

last_row = 0
alignments = {
    0: "right",
    1: "center",
    2: "center",
    3: "left",
    4: "left",
    5: "left",
    6: "right",
    7: "right",
    8: "right",
    9: "left",
}
BOTTLE_COUNT = 0
MONEY_COLUMNS = [6, 7, 8]
STRING_COLUMNS = [3, 4, 5]
BEST_BY = [2]
CELL_STYLE = NamedStyle(name="cell_style")
CELL_STYLE.font = Font(size=FONT_SIZE)
for row, CT_DATA in enumerate(ct_data_frame.itertuples()):
    row += 1
    last_row = row
    BOTTLE_COUNT = CT_DATA[1]
    color = next(colors)
    CELL_STYLE.alignment = Alignment(vertical='top')
    CELL_STYLE.fill = PatternFill(start_color=color, fill_type=FILL_TYPE)
    worksheet.row_dimensions[row + 1].height = None
    # print(f"ROW: {row}")
    for col, value in enumerate(CT_DATA[1:]):
        CELL_STYLE.alignment = Alignment(horizontal=alignments[col], wrap_text=True)
        CELL_STYLE.border = Border(left=None, right=None, top=None, bottom=thickness)
        CELL_STYLE.number_format = 'General'
        # print(f"ROW: {row}, COL: {col}, VALUE: {value}")
        target_cell = Cell(worksheet=worksheet, row=row + 1, column=col + 1)
        target_cell.style = CELL_STYLE
        if value == 0:
            value = ""
        if col in MONEY_COLUMNS:
           CELL_STYLE.number_format = "$#,##0 "
        if col in BEST_BY:
            if isinstance(value, int) and int(value) < CURRENT_YEAR:
                color = RED
            if isinstance(value, int) and int(value) == CURRENT_YEAR:
                color = YELLOW
            CELL_STYLE.fill = PatternFill(start_color=color, fill_type=FILL_TYPE)
        if isinstance(value, str):
            COL_WIDTHS[col] = max(len(value), COL_WIDTHS[col])
        worksheet.cell(row=row + 1, column=col + 1).style = CELL_STYLE
        worksheet.cell(row=row + 1, column=col + 1, value=value)

    square_cells_start = len(CT_DATA) - 1
    square_cells_end = len(CT_DATA) - 1 + min(CT_DATA[1], MAX_BOTTLE_CHECKBOXES)
    CELL_STYLE.fill = PatternFill(start_color=color, fill_type=FILL_TYPE)
    # print(f"BOTTLE_COUNT: {BOTTLE_COUNT}, square_cells_start: {square_cells_start}, square_cells_end: {square_cells_end}")
    for col in range(
        square_cells_start, square_cells_end, 1
    ):
        thickness = Side(border_style="double", color=BLACK)
        CELL_STYLE.border = Border(left=thickness, right=thickness, top=thickness, bottom=thickness)
        CELL_STYLE.fill = PatternFill(start_color=MAUVE, fill_type=FILL_TYPE)
        CELL_STYLE.number_format = 'General'
        worksheet.cell(row=row + 1, column=col + 1).style = CELL_STYLE
        worksheet.cell(row=row + 1, column=col + 1, value=" ")
        # print(f"ROW: {row + 1}, COL: {col + 1}, LETTER: {get_column_letter(col + 1)}")
        worksheet.column_dimensions[get_column_letter(col + 1)].width = SPACE_PADDING
    #if row >= 10:
    #    break
    
for col in range(len(COLUMN_NAMES)):
    print(f"COL: {col}, COL_WIDTHS[col]: {COL_WIDTHS[col]}")
    padding = SPACE_PADDING
    if col in STRING_COLUMNS:
        padding = TEXT_PADDING
    worksheet.column_dimensions[get_column_letter(col + 1)].width = COL_WIDTHS[col] + padding

# worksheet.autofilter(0, 0, 0, len(COLUMN_NAMES) - 1)
workbook.save(output_filename)

exit(0)

#    for col in range(square_cells_start, square_cells_end):
#        worksheet.insert_image(
#            last_row,
#            col,
#            SQUARE_IMAGE_NAME,
#            {"x_offset": 0, "y_offset": 0},
#        )


# Print out Totals and Chart
last_row += 2
print_cell(
    workbook,
    worksheet,
    last_row,
    1,
    "Totals",
    COLOR_HEADER,
    alignments[0],
    1,
)
print_cell(
    workbook,
    worksheet,
    last_row,
    2,
    "",
    COLOR_HEADER,
    alignments[0],
    1,
)
worksheet.insert_image(last_row, 3, IMAGE_NAME, {"x_offset": 0, "y_offset": 0})

# Print summaries
MONEY_FIELDS = ["Least Expensive", "Most Expensive", "Value"]
STRING_FIELDS = ["Producer", "Wine", "Type"]
for value in SUMMARY:
    print(f"SUMMARY Value: {value} = Summary[value]: {SUMMARY[value]}")
    last_row += 1
    color = next(colors)
    print_cell_money(
        workbook,
        worksheet,
        last_row,
        1,
        value,
        color,
        alignments[0],
        1,
    )
    if value in MONEY_FIELDS:
        print_cell_money(
            workbook,
            worksheet,
            last_row,
            2,
            SUMMARY[value],
            color,
            alignments[0],
            1,
        )
    else:
        print_cell(
            workbook,
            worksheet,
            last_row,
            2,
            SUMMARY[value],
            color,
            alignments[0],
            1,
        )
    if isinstance(value, str):
        COL_WIDTHS[1] = max(len(value), COL_WIDTHS[1])

# Print Wine Categories
for key, value in sorted(WINE_TYPES.items(), key=lambda x: x[0]):
    print(f"Wine Types KEY: {key} : VALUE: {value}")
    last_row += 1
    color = next(colors)
    print_cell(
        workbook,
        worksheet,
        last_row,
        1,
        key,
        color,
        alignments[0],
        1,
    )
    print_cell(
        workbook,
        worksheet,
        last_row,
        2,
        value,
        color,
        alignments[0],
        1,
    )
    if isinstance(value, str):
        COL_WIDTHS[1] = max(len(value), COL_WIDTHS[1])

print(f"COL WIDTHS: {COL_WIDTHS}")
for col, header in enumerate(COLUMN_NAMES):
    worksheet.set_column(col, col, min(COL_WIDTHS[col], 50))
for col in range(len(COLUMN_NAMES), MAX_COUNT + 1):
    worksheet.set_column(col, col, 2)


# excel = win32.gencache.EnsureDispatch("Excel.Application")
# wb = excel.Workbooks.Open(Path(output_filename).absolute())
# ws = wb.Worksheets(sheet_name)
# ws.Columns.AutoFit()
# wb.Save()
# excel.Application.Quit()

wb = Workbook(Path(output_filename).absolute())
ws = wb.Worksheets(sheet_name)
ws.Columns.AutoFit()
wb.Save()
