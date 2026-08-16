# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))

from cellartracker import cellartracker
from ctconfig import CELLAR_USERNAME, CELLAR_PASSWORD
from ctdataframe import CTDataFrame
from ctworkbook import CTWorkbook
#from matplotlib import pyplot as plt
#from matplotlib import style
import sys


def main(args):
    # Get the data from CellarTracker and convert it to a DataFrame
    ct_data : list = cellartracker.CellarTracker(CELLAR_USERNAME, CELLAR_PASSWORD).get_list()
    ct_data_frame = CTDataFrame(ct_data)
    ct_workbook = CTWorkbook(ct_data_frame)
    return 0  # Standard success exit code

if __name__ == "__main__":
    # sys.argv reads inputs directly from the terminal
    sys.exit(main(sys.argv))
    
exit(0)


"""
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
"""