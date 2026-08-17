# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))

from cellartracker import cellartracker
from ctconfig import CELLAR_USERNAME, CELLAR_PASSWORD
from ctdataframe import CTDataFrame
from ctworkbook import CTWorkbook
from ctutils import _LOGGER
from pandas import DataFrame
import sys
from typing import List


def main(args):
    # Get the data from CellarTracker and convert it to a DataFrame
    ct_data : list = cellartracker.CellarTracker(CELLAR_USERNAME, CELLAR_PASSWORD).get_list()
    _LOGGER.debug(f"Retrieved {len(ct_data)} records from CellarTracker.")
    ct_metadata = CTDataFrame(ct_data)
    cd_data_frame: DataFrame = ct_metadata.get_dataframe()
    ct_summary: dict[str, int] = ct_metadata.get_summary_data()
    ct_workbook = CTWorkbook(cd_data_frame, ct_summary)
    ct_workbook.create_workbook()
    return 0  # Standard success exit code

if __name__ == "__main__":
    # sys.argv reads inputs directly from the terminal
    _LOGGER.setLevel("DEBUG")
    sys.exit(main(sys.argv))
    
exit(0)