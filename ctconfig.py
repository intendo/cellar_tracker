from configparser import ConfigParser
from datetime import date
from os import path as os_path

# Locate the file dynamically relative to this script
BASE_DIR = os_path.dirname(os_path.abspath(__file__))
INI_PATH = os_path.join(BASE_DIR, "settings.ini")

# Instantiate and read
_parser = ConfigParser()
_parser.read(INI_PATH)

# Expose settings as strongly-typed dictionary structures or objects
CELLAR_USERNAME = _parser.get("AUTH", "CELLAR_USERNAME", fallback="")
CELLAR_PASSWORD = _parser.get("AUTH", "CELLAR_PASSWORD", fallback="")

CELL_BORDER_COLOR = _parser.get("COLORS", "CELL_BORDER_COLOR", fallback="FF000000")
CHECKBOX_COLOR = _parser.get("COLORS", "CHECKBOX_COLOR", fallback="FF6DEDED")
COLOR_ALT1 = _parser.get("COLORS", "ALT1", fallback="FFFFFF")
COLOR_ALT2 = _parser.get("COLORS", "ALT2", fallback="E2EFDA")
DRINK_PAST_DUE_COLOR = _parser.get("COLORS", "DRINK_PAST_DUE_COLOR", fallback="FFFF0000")
DRINK_NOW_COLOR = _parser.get("COLORS", "DRINK_NOW_COLOR", fallback="FFFFFF00")
FONT_COLOR = _parser.get("COLORS", "FONT_COLOR", fallback="FF000000")
HEADER_BG_COLOR = _parser.get("COLORS", "HEADER_BG_COLOR", fallback="FF44CC44")

CURRENCY_STYLE = _parser.get("STYLES", "CURRENCY_STYLE", fallback="$#,##0")
FONT_NAME = _parser.get("SETTINGS", "FONT_NAME", fallback="Ariel Narrow")
FONT_SIZE = _parser.getint("SETTINGS", "FONT_SIZE", fallback=16)
IMAGE_NAME = _parser.get("SETTINGS", "IMAGE_NAME", fallback="CellarTrackerHistogram.png")
MAX_BOTTLE_CHECKBOXES = _parser.getint("SETTINGS", "MAX_BOTTLE_CHECKBOXES", fallback=12)
SPACE_PADDING = _parser.getint("SETTINGS", "SPACE_PADDING", fallback=3)
TEXT_PADDING = _parser.getint("SETTINGS", "TEXT_PADDING", fallback=16)
WORKBOOK_FILENAME = _parser.get("SETTINGS", "WORKBOOK_FILENAME", fallback="CellarTracker.xlsx")


CURRENT_YEAR = date.today().year

COLUMN_NAMES = [
    "Count",
    "Bin",
    "Vintage",
    "Best By",
    "Producer",
    "Wine",
    "Type",
    "Price",
    "CTValue",
    "Value",
]

WINE_COLORS = {
    "Red": "red",
    "Rose": "pink",
    "White": "yellow",
    "Rose - Spa": "blue",
    "White - Sp": "green",
}
        