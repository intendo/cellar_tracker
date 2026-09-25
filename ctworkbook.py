"""This file contains utilities to populate Workbook from CellarTracker API."""

# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))
            
from ctconfig import COLUMN_NAMES, CURRENT_YEAR
from ctconfig import COLOR_ALT1, COLOR_ALT2, CELL_BORDER_COLOR, FONT_COLOR
from ctconfig import DRINK_PAST_DUE_COLOR, DRINK_NOW_COLOR, HEADER_BG_COLOR, CHECKBOX_COLOR
from ctconfig import CURRENCY_STYLE, FONT_NAME, FONT_SIZE, MAX_BOTTLE_CHECKBOXES
from ctconfig import SPACE_PADDING, TEXT_PADDING, WORKBOOK_FILENAME
from ctdataframe import CTDataFrame
from ctutils import _LOGGER
from itertools import cycle
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter
from pandas import DataFrame
from typing import List


class CTWorkbookError(Exception):
    def __init__(self, message, bad_value=None):
        super().__init__(message)
        self.bad_value = bad_value


class CTWorkbook:
    """Helper class for organizing and manipulating CellarTracker data as a Workbook."""

    def __init__(self, ct_data_frame: DataFrame, ct_summary_data: dict[str, int]):
        """Raises CTWorkbookError if there is any problems."""
        self.ct_data_frame = ct_data_frame
        self.ct_summary_data = ct_summary_data
        #self.create_workbook()

    def create_workbook(self):
        """Creates a workbook from the CellarTracker data."""

        output_filename = WORKBOOK_FILENAME
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
        assert worksheet.sheet_properties.pageSetUpPr is not None
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        
        # TODO: should this be combined with HEADERS
        # Alignment for each data column in the worksheet
        alignments = {
            COLUMN_NAMES.index("Count"): "right",
            COLUMN_NAMES.index("Bin"): "center",
            COLUMN_NAMES.index("Vintage"): "center",
            COLUMN_NAMES.index("Best By"): "left",
            COLUMN_NAMES.index("Producer"): "left",
            COLUMN_NAMES.index("Wine"): "left",
            COLUMN_NAMES.index("Type"): "right",
            COLUMN_NAMES.index("Price"): "right",
            COLUMN_NAMES.index("CTValue"): "right",
            COLUMN_NAMES.index("Value"): "left",
        }

        best_by_columns = [COLUMN_NAMES.index("Best By")]
        money_columns = [COLUMN_NAMES.index("Price"), COLUMN_NAMES.index("CTValue"), COLUMN_NAMES.index("Value")]
        string_columns = [COLUMN_NAMES.index("Producer"), COLUMN_NAMES.index("Wine"), COLUMN_NAMES.index("Type")]

        bold_font = Font(name=FONT_NAME, bold=True, color=FONT_COLOR, size=FONT_SIZE)
        bottle_count = 0
        cell_style = NamedStyle(name="cell_style")
        cell_style.font = Font(size=FONT_SIZE)
        colors = cycle([COLOR_ALT1, COLOR_ALT2])
        fill_type = "solid"
        last_row = 0
        thickness = Side(border_style="thick", color=CELL_BORDER_COLOR)
        cell_border = Border(
            left=thickness, right=thickness, top=thickness, bottom=thickness
        )
   
        #  TODO:
        #    - This section seems awkward because it only work on the header
        #      and not the worksheet as a whole
        #    - This is partly because the font is bold on the header row
        #      so it has to be treated specially
        #    - The COLUMN_NAMES have to be added to the worksheet before
        #      you can set the widths
        #  Set up the header row with bold font and background color
        worksheet.append(COLUMN_NAMES)

        # Set up the header row with bold font and background color
        for cell in worksheet[1]: # Assuming the first row is row 1
            cell.font = bold_font
            cell.fill = PatternFill(start_color=HEADER_BG_COLOR, fill_type=fill_type)
            cell.border = cell_border

        col_widths = [0] * len(COLUMN_NAMES)
        for col, header in enumerate(COLUMN_NAMES):
            worksheet.cell(row=1, column=col + 1, value=header)
            col_widths[col] = len(header)

        for col in worksheet.columns:
            max_len = 0
            # Get the column letter (e.g., 'A')
            col_letter = get_column_letter(col[0].column)  # type: ignore
            
            # Determine the maximum length of values in the current column
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            
            # Set the column width with some extra padding
            # Adding a safety margin (e.g., +3) keeps text readable and uncompressed
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
            # print(f"Column {col_letter} width set to {worksheet.column_dimensions[col_letter].width}")
        # exit(0)

        # Process all the data from Cellar Tracker into the worksheet
        for row, ct_data in enumerate(self.ct_data_frame.itertuples()):
            row += 1
            last_row = row
            bottle_count = ct_data[1]
            color = next(colors)
            cell_style.alignment = Alignment(vertical='top')
            cell_style.fill = PatternFill(start_color=color, fill_type=fill_type)
            worksheet.row_dimensions[row + 1].height = None
            print(f"ROW: {row}")
            for col, value in enumerate(ct_data[1:]):
                # print(f"COL: {col}, ROW: {row}, VALUE: {value}")
                # if col >= 5:
                #    exit(0)
                cell_style.alignment = Alignment(horizontal=alignments[col], wrap_text=True)
                cell_style.border = Border(left=None, right=None, top=None, bottom=thickness)
                cell_style.number_format = 'General'
                # _LOGGER.debug(f"ROW: {row}, COL: {col}, VALUE: {value}")
                target_cell = Cell(worksheet=worksheet, row=row + 1, column=col + 1)
                target_cell.style = cell_style
                if value == 0:
                    value = ""
                if col in money_columns:
                   cell_style.number_format = CURRENCY_STYLE
                if col in best_by_columns:
                    if isinstance(value, int) and int(value) < CURRENT_YEAR:
                        color = DRINK_PAST_DUE_COLOR
                    if isinstance(value, int) and int(value) == CURRENT_YEAR:
                        color = DRINK_NOW_COLOR
                    cell_style.fill = PatternFill(start_color=color, fill_type=fill_type)
                if isinstance(value, str):
                    col_widths[col] = max(len(value), col_widths[col])
                worksheet.cell(row=row + 1, column=col + 1).style = cell_style
                worksheet.cell(row=row + 1, column=col + 1, value=value)

            square_cells_start = len(ct_data) - 1
            square_cells_end = len(ct_data) - 1 + min(ct_data[1], MAX_BOTTLE_CHECKBOXES)
            cell_style.fill = PatternFill(start_color=color, fill_type=fill_type)
            # print(f"BOTTLE_COUNT: {BOTTLE_COUNT}, square_cells_start: {square_cells_start}, square_cells_end: {square_cells_end}")
            for col in range(
                square_cells_start, square_cells_end, 1
            ):
                thickness = Side(border_style="double", color=CELL_BORDER_COLOR)
                cell_style.border = Border(left=thickness, right=thickness, top=thickness, bottom=thickness)
                cell_style.fill = PatternFill(start_color=CHECKBOX_COLOR, fill_type=fill_type)
                cell_style.number_format = 'General'
                worksheet.cell(row=row + 1, column=col + 1).style = cell_style
                worksheet.cell(row=row + 1, column=col + 1, value=" ")
                # print(f"ROW: {row + 1}, COL: {col + 1}, LETTER: {get_column_letter(col + 1)}")
                worksheet.column_dimensions[get_column_letter(col + 1)].width = SPACE_PADDING
            #if row >= 10:
            #    break
    
        for col in range(len(COLUMN_NAMES)):
            print(f"COL: {col}, COL_WIDTHS[col]: {col_widths[col]}")
            padding = SPACE_PADDING
            if col in string_columns:
                padding = TEXT_PADDING
            worksheet.column_dimensions[get_column_letter(col + 1)].width = col_widths[col] + padding

        # Print out Totals and Chart
        last_row += 3
        thickness = Side(border_style="double", color=CELL_BORDER_COLOR)
        cell_style.alignment = Alignment(horizontal="left", wrap_text=False)
        cell_style.border = Border(left=thickness, right=thickness, top=thickness, bottom=thickness)
        cell_style.fill = PatternFill(start_color=HEADER_BG_COLOR, fill_type=fill_type)
        worksheet.cell(row=last_row, column=1).style = cell_style
        worksheet.cell(row=last_row, column=1, value="Totals")

        # TODO Create image here and insert it into the worksheet
        # worksheet.insert_image(last_row, 3, IMAGE_NAME, {"x_offset": 0, "y_offset": 0})
        
        # Print summaries
        MONEY_FIELDS = ["Least Expensive", "Most Expensive", "CTValue", "Value"]
        STRING_FIELDS = ["Producer", "Wine", "Type"]
        for value in self.ct_summary_data:
            print(f"SUMMARY Value: {value}, Data: {self.ct_summary_data[value]}")
            last_row += 1
            cell_style.number_format = 'General'
            thickness = Side(border_style="double", color=CELL_BORDER_COLOR)
            cell_style.alignment = Alignment(horizontal="left", wrap_text=False)
            cell_style.border = Border(left=thickness, right=thickness, top=thickness, bottom=thickness)
            cell_style.fill = PatternFill(start_color=HEADER_BG_COLOR, fill_type=fill_type)
            worksheet.cell(row=last_row, column=1).style = cell_style
            worksheet.cell(row=last_row, column=1, value=f"{value}")

            color = next(colors)
            if f"{value}" in MONEY_FIELDS:
                print(f"WTF: {value}")
                cell_style.number_format = CURRENCY_STYLE
            cell_style.alignment = Alignment(horizontal="right", wrap_text=False)
            cell_style.border = Border(left=None, right=None, top=None, bottom=thickness)
            cell_style.fill = PatternFill(start_color=color, fill_type=fill_type)
            worksheet.cell(row=last_row, column=2).style = cell_style
            worksheet.cell(row=last_row, column=2).value = self.ct_summary_data[value]
            worksheet.cell(row=last_row, column=2).data_type = 'n'

#        
#        # Print Wine Categories
#        for key, value in sorted(WINE_TYPES.items(), key=lambda x: x[0]):
#            print(f"Wine Types KEY: {key} : VALUE: {value}")
#            last_row += 1
#            color = next(colors)
#            print_cell(
#                workbook,
#                worksheet,
#                last_row,
#                1,
#                key,
#                color,
#                alignments[0],
#                1,
#            )
#            print_cell(
#                workbook,
#                worksheet,
#                last_row,
#                2,
#                value,
#                color,
#                alignments[0],
#                1,
#            )
#            if isinstance(value, str):
#                COL_WIDTHS[1] = max(len(value), COL_WIDTHS[1])
#
#        print(f"COL WIDTHS: {COL_WIDTHS}")
#        for col, header in enumerate(COLUMN_NAMES):
#            worksheet.set_column(col, col, min(COL_WIDTHS[col], 50))
#        for col in range(len(COLUMN_NAMES), MAX_COUNT + 1):
#            worksheet.set_column(col, col, 2)
                                 
        # worksheet.autofilter(0, 0, 0, len(COLUMN_NAMES) - 1)
        workbook.save(output_filename)
