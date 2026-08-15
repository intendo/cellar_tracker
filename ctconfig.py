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

COLOR_HEADER = _parser.get("COLORS", "HEADER", fallback="70AD47")
COLOR_ALT1 = _parser.get("COLORS", "ALT1", fallback="FFFFFF")
COLOR_ALT2 = _parser.get("COLORS", "ALT2", fallback="E2EFDA")

ENCODE = _parser.get("SETTINGS", "ENCODE", fallback="utf-8-sig")
FONT_NAME = _parser.get("SETTINGS", "FONT_NAME", fallback="Ariel Narrow")
FONT_SIZE = _parser.getint("SETTINGS", "FONT_SIZE", fallback=16)
IMAGE_NAME = _parser.get("SETTINGS", "IMAGE_NAME", fallback="CellarTrackerHistogram.png")
MAX_BOTTLE_CHECKBOXES = _parser.getint("SETTINGS", "MAX_BOTTLE_CHECKBOXES", fallback=12)
# SQUARE_IMAGE_NAME = "Square.png"
SPACE_PADDING = _parser.getint("SETTINGS", "SPACE_PADDING", fallback=3)
TEXT_PADDING = _parser.getint("SETTINGS", "TEXT_PADDING", fallback=16)

CURRENT_YEAR = date.today().year

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
