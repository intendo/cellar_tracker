from datetime import date
from itertools import cycle
#from matplotlib import pyplot as plt
#from matplotlib import style
from numpy import array as narray
from openpyxl.cell import Cell
from pandas import DataFrame
from unidecode import unidecode
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle

from cellartracker import cellartracker
from decouple import config
from configparser import ConfigParser
from pathlib import Path
from re import sub as subst

# GLOBALS
config = ConfigParser()
config.read('settings.ini')
#config['AUTH']['CELLAR_USERNAME']
# https://cellartracker.readthedocs.io/en/latest/readme.html
CT_DATA = cellartracker.CellarTracker(
    username=config.get("AUTH", "CELLAR_USERNAME"),
    password=config.get("AUTH", "CELLAR_PASSWORD")
).get_list()

CURRENT_YEAR = date.today().year

COLOR_HEADER = "70ad47"
COLOR_ALT1 = "FFFFFF"
COLOR_ALT2 = "E2EFDA"
#    "Inventory",
COLUMN_NAMES = [
    "Count",
    "Vintage",
    "Best By",
    "Producer",
    "Wine",
    "Type",
    "Price",
    "CTValue",
    "Value",
]
ENCODE = "utf-8-sig"
FONT_SIZE = 16
IMAGE_NAME = "CellarTrackerHistogram.png"
MAX_BOTTLE_CHECKBOXES = 12
# SQUARE_IMAGE_NAME = "Square.png"
SPACE_PADDING = 3
TEXT_PADDING = 10

# Build a dataframe from the data
DFRAME = DataFrame.from_records(CT_DATA)
# with open("rawdata.txt", "w") as outfile:
#     print(DFRAME.to_csv(index=False), file=outfile)
# Drop the columns that we don't care about
DFRAME = DFRAME.drop(
    columns=[
        "iWine",
        "WineBarcode",
        "Pending",
        "Location",
        "Bin",
        "Size",
        "Valuation",
        "MyValue",
        "WBValue",
        "MenuPrice",
        "Currency",
        "Locale",
        "SortProducer",
        "Color",
        "Category",
        "MasterVarietal",
        "Designation",
        "Vineyard",
        "WA",
        "WS",
        "IWC",
        "BH",
        "AG",
        "WE",
        "JR",
        "RH",
        "JG",
        "GV",
        "JK",
        "LD",
        "CW",
        "WFW",
        "PR",
        "SJ",
        "WD",
        "RR",
        "JH",
        "MFW",
        "WWR",
        "IWR",
        "CHG",
        "TT",
        "TWF",
        "DR",
        "FP",
        "JM",
        "PG",
        "WAL",
        "JS",
        "CT",
        "MY",
        "BeginConsume",
        "UPC",
    ],
)
# Remove the empty value (9999) and convert value to integer
DFRAME["EndConsume"] = DFRAME["EndConsume"].apply(lambda x: "" if (x == "9999") else x)
DFRAME["EndConsume"] = DFRAME["EndConsume"].apply(lambda x: int(x) if (x != "") else x)
# Convert Quantity/Count to integer
DFRAME["Quantity"] = DFRAME["Quantity"].apply(lambda x: int(x))
# Convert float Value of wine to integer
DFRAME["CTValue"] = DFRAME["CTValue"].apply(lambda x: round(float(x)))
# Remove zero values
# Convert float Value of wine to integer
DFRAME["Price"] = DFRAME["Price"].apply(lambda x: round(float(x)))
# Remove zero values
DFRAME["Vintage"] = DFRAME["Vintage"].apply(lambda x: "NV" if (x == "1001") else int(x))
DFRAME["Wine"] = DFRAME["Wine"].apply(lambda x: unidecode(str(x)))
DFRAME["Type"] = DFRAME["Type"].apply(lambda x: unidecode(str(x)))

# Limit the field lengths characters
DFRAME["Producer"] = DFRAME["Producer"].apply(lambda x: x[:30])
DFRAME["Type"] = DFRAME["Type"].apply(lambda x: x[:10])
# Sort the dataframe by wine type, year, varietal, and winery
DFRAME.sort_values(
    ["Type", "Vintage", "Wine", "Producer", "EndConsume"],
    ascending=[False, True, True, True, True],
    inplace=True,
)
# Rename columns
DFRAME = DFRAME.rename(
    columns={
        "Quantity": "Count",
        "EndConsume": "Best By",
    }
)

# Convert Value from price per bottle to Value of all bottles
DFRAME["Value"] = DFRAME.apply(
    lambda row: (int(row["Count"]) * int(max(row["CTValue"], row["Price"]))),
    axis=1,
)
DFRAME["Wine"] = DFRAME.apply(
    lambda row: subst(r"\s*" + row["Producer"] + r"\s*", "", row["Wine"])
    + ", "
    + row["Appellation"],
    axis=1,
)

# dataframe_to_pdf(df, "test_1.pdf")
# print(DFRAME.to_csv(index=False))

SUMMARY = {"Bottles": DFRAME["Count"].sum()}
SUMMARY["Value"] = DFRAME["Value"].sum(skipna=True)
SUMMARY["Producers"] = len(DFRAME.Producer.unique())
SUMMARY["Countries"] = len(DFRAME.Country.unique())
SUMMARY["Regions"] = len(DFRAME.Region.unique())
SUMMARY["Appellations"] = len(DFRAME.Appellation.unique())
SUMMARY["Vintages"] = len(DFRAME.Vintage.unique())
max_ctvalue = max(DFRAME.CTValue.unique())
max_price = max(DFRAME.Price.unique())
MAX_COUNT = max(DFRAME.Count.unique())
SUMMARY["Most Expensive"] = max(max_ctvalue, max_price)
min_ctvalue = min(i for i in DFRAME.CTValue.unique() if i > 0)
min_price = min(i for i in DFRAME.Price.unique() if i > 0)
SUMMARY["Least Expensive"] = min(min_ctvalue, min_price)
VINTAGES = sorted(filter(lambda x: x not in ["NV"], DFRAME["Vintage"]))
NV_VINTAGES = sorted(filter(lambda x: x in ["NV"], DFRAME["Vintage"]))

# for i in SUMMARY:
#     print(f"SUMMARY: {i} = {SUMMARY[i]}")
# print(f"MAX_COUNT: {MAX_COUNT}")

year_min: int = min(VINTAGES)
year_max = max(VINTAGES)
# print(f"Year MIN: {year_min}")
# print(f"Year MAX: {year_max}")
# Generate a list of years (including "NV") as strings
year_list = ["NV", *[f"{x}" for x in range(year_min - 1, CURRENT_YEAR + 1)]]
# print(f"Number of Years: {len(year_list)}")
# print(f"Year List: {year_list}")
year_list_str = list(map(lambda x: str(x).replace("20", "", 1), year_list))
# print(f"Year List String: {year_list_str}")

for i in NV_VINTAGES:
    VINTAGES.insert(0, i)
# print(f"Vintages: {VINTAGES}")

vintage_data = {}
for vintage in VINTAGES:
    # print(f"vintage: #{vintage}#")
    for year in year_list:
        # print(f"year: #{year}#")
        if year not in vintage_data:
            vintage_data[year] = 0
        if year == vintage:
            # print(f"year/vintage MATCH: #{year}#")
            vintage_data[year] = vintage_data[year] + 1

data_values = list(vintage_data.values())
# print(f"Vintage Data: {vintage_data}")
# print(f"Len of Data Values: {len(data_values)}")
# print(f"Data Values: {data_values}")


WINE_TYPES = {
    "Red": 0,
    "Rose": 0,
    "White": 0,
    "Rose - Spa": 0,
    "White - Sp": 0,
}
for i in range(len(DFRAME)):
    key: str = str(DFRAME.at[i, "Type"])
    vintage: str = str(DFRAME.at[i, "Vintage"])
    count: int = int(float(str(DFRAME.at[i, "Count"])))
    # count: int = int(float(DFRAME.at[i, "Count"]))
    # print(f"COUNT: {count} = DFRAME.at[{i}, 'Count'] = {DFRAME.at[i, 'Count']}")
    WINE_TYPES[key] = count if key not in WINE_TYPES else WINE_TYPES[key] + count
    # print(f"WINE_TYPE KEY: {key} = VALUE: {WINE_TYPES[key]}")

# Initialize a dictionary to hold the wine stack data for each wine type and year
WINE_STACK_DATA: dict[str, dict[str, int]] = {}
for key, value in sorted(WINE_TYPES.items(), key=lambda x: x[0]):
    WINE_STACK_DATA[key] = {"NV": 0}
    for year in year_list:
        WINE_STACK_DATA[key][year] = 0
        #print(f"WINE TYPE KEY: {key} : {WINE_TYPES[key]}, YEAR: {year} : {WINE_STACK_DATA[key][year]}")

# print(f"WINE STACK DATA: {WINE_STACK_DATA}")

for i in range(len(DFRAME)):
    key: str = str(DFRAME.at[i, "Type"])
    vintage: str = str(DFRAME.at[i, "Vintage"])
    if key not in WINE_STACK_DATA:
        WINE_STACK_DATA[key] = {"NV": 0}
        for year in year_list:
            WINE_STACK_DATA[key][year] = 0
    WINE_STACK_DATA[key][vintage] += 1
# print(f"WINE STACK DATA: {WINE_STACK_DATA}")
# print(f"SUMMARY: {SUMMARY}")

WINE_COLORS = {
    "Red": "red",
    "Rose": "pink",
    "White": "yellow",
    "Rose - Spa": "blue",
    "White - Sp": "green",
}
x1 = year_list_str
# Ensure numeric numpy arrays (convert any non-int types to int)
y1 = narray([int(v) for v in WINE_STACK_DATA["Red"].values()])
y2 = narray([int(v) for v in WINE_STACK_DATA["Rose"].values()])
y3 = narray([int(v) for v in WINE_STACK_DATA["White"].values()])
y4 = narray([int(v) for v in WINE_STACK_DATA["Rose - Spa"].values()])
y5 = narray([int(v) for v in WINE_STACK_DATA["White - Sp"].values()])

# Reduce the data to just the Columns we care about
DFRAME = DFRAME[COLUMN_NAMES]
# print("DFRAME Columns: ", DFRAME.columns)
# print(DFRAME.to_csv(index=False))


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
bold_font = Font(bold=True, color=BLACK, size=FONT_SIZE)
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
for row, CT_DATA in enumerate(DFRAME.itertuples()):
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
