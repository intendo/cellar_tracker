"""This file contains utilities to populate Workbook from CellarTracker API."""

# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))
            
from ctconfig import COLUMN_NAMES, CURRENT_YEAR, COLOR_ALT1, COLOR_ALT2, FONT_NAME, FONT_SIZE, MAX_BOTTLE_CHECKBOXES, SPACE_PADDING, TEXT_PADDING, WORKBOOK_FILENAME
from ctconfig import CELL_BORDER_COLOR, FONT_COLOR, DRINK_PAST_DUE_COLOR, DRINK_NOW_COLOR, HEADER_BG_COLOR, CHECKBOX_COLOR
from ctdataframe import CTDataFrame
from ctutils import _LOGGER
from itertools import cycle
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font, NamedStyle
from openpyxl.utils import get_column_letter
from typing import List


class CTWorkbookError(Exception):
    def __init__(self, message, bad_value=None):
        super().__init__(message)
        self.bad_value = bad_value


class CTWorkbook:
    """Helper class for organizing and manipulating CellarTracker data as a Workbook."""

    def __init__(self, ct_data_frame: CTDataFrame):
        """Raises CTWorkbookError if there is any problems."""
        self.ct_data_frame = ct_data_frame
        self.dframes = self.create_workbook()

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
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        
        #  Set up the header row with bold font and background color
        worksheet.append(COLUMN_NAMES)

        colors = cycle([COLOR_ALT1, COLOR_ALT2])
        fill_type = "solid"
        thickness = Side(border_style="thick", color=CELL_BORDER_COLOR)
        cell_border = Border(
            left=thickness, right=thickness, top=thickness, bottom=thickness
        )
   
        # Set up the header row with bold font and background color
        bold_font = Font(name=FONT_NAME, bold=True, color=FONT_COLOR, size=FONT_SIZE)
        for cell in worksheet[1]: # Assuming the first row is row 1
            cell.font = bold_font
            cell.fill = PatternFill(start_color=HEADER_BG_COLOR, fill_type=fill_type)
            cell.border = cell_border

        # worksheet.set_row(0, cell_format=data_header)
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

        # Alignment for each column 
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

        bottle_count = 0
        money_columns = [6, 7, 8]
        string_columns = [3, 4, 5]
        best_by = [2]
        cell_style = NamedStyle(name="cell_style")
        cell_style.font = Font(size=FONT_SIZE)
        last_row = 0

        for row, ct_data in enumerate(self.ct_data_frame.itertuples()):
            row += 1
            last_row = row
            bottle_count = ct_data[1]
            color = next(colors)
            cell_style.alignment = Alignment(vertical='top')
            cell_style.fill = PatternFill(start_color=color, fill_type=fill_type)
            worksheet.row_dimensions[row + 1].height = None
            # print(f"ROW: {row}")
            for col, value in enumerate(ct_data[1:]):
                cell_style.alignment = Alignment(horizontal=alignments[col], wrap_text=True)
                cell_style.border = Border(left=None, right=None, top=None, bottom=thickness)
                cell_style.number_format = 'General'
                # print(f"ROW: {row}, COL: {col}, VALUE: {value}")
                target_cell = Cell(worksheet=worksheet, row=row + 1, column=col + 1)
                target_cell.style = cell_style
                if value == 0:
                    value = ""
                if col in money_columns:
                   cell_style.number_format = "$#,##0 "
                if col in best_by:
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
#        last_row += 2
#        worksheet.cell(row=last_row, column=1).style = cell_style
#            "Totals",
#            HEADER_BG_COLOR,
#            alignments[0],
#
#        # TODO Create image here and insert it into the worksheet
#        # worksheet.insert_image(last_row, 3, IMAGE_NAME, {"x_offset": 0, "y_offset": 0})
        
#        # Print summaries
#        MONEY_FIELDS = ["Least Expensive", "Most Expensive", "Value"]
#        STRING_FIELDS = ["Producer", "Wine", "Type"]
#        for value in SUMMARY:
#            print(f"SUMMARY Value: {value} = Summary[value]: {SUMMARY[value]}")
#            last_row += 1
#            color = next(colors)
#            print_cell_money(
#                workbook,
#                worksheet,
#                last_row,
#                1,
#                value,
#                color,
#                alignments[0],
#                1,
#            )
#            if value in MONEY_FIELDS:
#                print_cell_money(
#                    workbook,
#                    worksheet,
#                    last_row,
#                    2,
#                    SUMMARY[value],
#                    color,
#                    alignments[0],
#                    1,
#                )
#            else:
#                print_cell(
#                    workbook,
#                    worksheet,
#                    last_row,
#                    2,
#                    SUMMARY[value],
#                    color,
#                    alignments[0],
#                    1,
#                )
#            if isinstance(value, str):
#                COL_WIDTHS[1] = max(len(value), COL_WIDTHS[1])
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
