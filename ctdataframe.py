"""This file contains utilities to populate DataFrame from CellarTracker API."""

# Adds the parent directory of your package to the search path
from sys import path
from pathlib import Path
path.append(str(Path(__file__).resolve().parent.parent))
            
from ctconfig import COLUMN_NAMES, CURRENT_YEAR, WINE_COLORS
from ctutils import _LOGGER
from datetime import date
from numpy import array as narray
from pandas import DataFrame
from re import sub as subst
from typing import List
from unidecode import unidecode


class CTDataFrameError(Exception):
    def __init__(self, message, bad_value=None):
        super().__init__(message)
        self.bad_value = bad_value


class CTDataFrame:
    """Helper class for organizing and manipulating CellarTracker data as a DataFrame."""

    def __init__(self, ct_data: List[dict]):
        """Raises CTDataFrameError if there is any problems."""
        _LOGGER.debug(f"Initializing CTDataFrame with {len(ct_data)} records.")
        self.ct_data = ct_data
        self.dframe = DataFrame()
        self.summary: dict[str, int] = {}

    def __contains__(self, item):
        """Helper for checking for the presence of an option"""
        return item in self.ct_data

    def has_key(self, item):
        """Helper for checking for the presence of an option"""
        return item in self.ct_data

    def __iter__(self):
        yield from self.ct_data

    def __getitem__(self, key):
        return (
            self.ct_data[key]
            if key in self.ct_data
            else CTDataFrameError(f"Key '{key}' not found in CTDataFrame.", bad_value=key)
        )

    def itertuples(self):
        yield from self.dframe.itertuples()

    def get_dataframe(self) -> DataFrame:
        self.create_dataframe()
        return self.dframe

    def get_summary_data(self) -> dict[str, int]:
        return self.summary

    def create_dataframe(self) -> None:
        """ Build a dataframe from the data"""
        self.dframe = DataFrame.from_records(self.ct_data)
        _LOGGER.debug(f"Entering")
        # with open("rawdata.txt", "w") as outfile:
        #     print(self.dframe.to_csv(index=False), file=outfile)
        # exit(0)
        # Drop the columns that we don't care about
        self.dframe = self.dframe.drop(
            columns=[
                "iWine",
                "WineBarcode",
                "Pending",
                "Location",
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
        self.dframe["EndConsume"] = self.dframe["EndConsume"].apply(lambda x: "" if (x == "9999") else x)
        self.dframe["EndConsume"] = self.dframe["EndConsume"].apply(lambda x: int(x) if (x != "") else x)
        # Convert Quantity/Count to integer
        self.dframe["Quantity"] = self.dframe["Quantity"].apply(lambda x: int(x))
        # Convert float Value of wine to integer
        self.dframe["CTValue"] = self.dframe["CTValue"].apply(lambda x: round(float(x)))
        # Remove zero values
        # Convert float Value of wine to integer
        self.dframe["Price"] = self.dframe["Price"].apply(lambda x: round(float(x)))
        # Remove zero values
        self.dframe["Vintage"] = self.dframe["Vintage"].apply(lambda x: "NV" if (x == "1001") else int(x))
        self.dframe["Wine"] = self.dframe["Wine"].apply(lambda x: unidecode(str(x)))
        self.dframe["Type"] = self.dframe["Type"].apply(lambda x: unidecode(str(x)))
        self.dframe["Bin"] = self.dframe["Bin"].apply(lambda x: unidecode(str(x)))

        # print(f"DFRAME Columns: {self.dframe["Bin"]}")
        # exit(0)

        # Limit the field lengths characters
        self.dframe["Producer"] = self.dframe["Producer"].apply(lambda x: x[:30])
        self.dframe["Type"] = self.dframe["Type"].apply(lambda x: x[:10])
        # Sort the dataframe by wine type, year, varietal, and winery
        self.dframe.sort_values(
            ["Type", "Vintage", "Wine", "Producer", "EndConsume"],
            ascending=[False, True, True, True, True],
            inplace=True,
        )
        # Rename columns
        self.dframe = self.dframe.rename(
            columns={
                "Quantity": "Count",
                "EndConsume": "Best By",
            }
        )

        # Convert Value from price per bottle to Value of all bottles
        self.dframe["Value"] = self.dframe.apply(
            lambda row: (int(row["Count"]) * int(max(row["CTValue"], row["Price"]))),
            axis=1,
        )
        self.dframe["Wine"] = self.dframe.apply(
            lambda row: subst(r"\s*" + row["Producer"] + r"\s*", "", row["Wine"])
            + ", "
            + row["Appellation"],
            axis=1,
        )
        #  TODO: Figure out how to combine row["Bin"] 
        #self.dframe["Bin"] = self.dframe.apply(
        #    lambda x: subst(r"\s*" + x + r"\s*", "", x["Bin"]),
        #    axis=1,
        #)

        # dataframe_to_pdf(df, "test_1.pdf")
        #print(self.dframe.to_csv(index=False))
        #print(f"Bin: {self.dframe["Bin"]}")
        #exit(0)

        self.summary = {"Bottles": self.dframe["Count"].sum()}
        self.summary["Value"] = self.dframe["Value"].sum(skipna=True)
        self.summary["CTValue"] = self.dframe["CTValue"].sum(skipna=True)
        self.summary["Producers"] = len(self.dframe.Producer.unique())
        self.summary["Countries"] = len(self.dframe.Country.unique())
        self.summary["Regions"] = len(self.dframe.Region.unique())
        self.summary["Appellations"] = len(self.dframe.Appellation.unique())
        self.summary["Vintages"] = len(self.dframe.Vintage.unique())
        max_ctvalue = max(self.dframe.CTValue.unique())
        max_price = max(self.dframe.Price.unique())
        MAX_COUNT = max(self.dframe.Count.unique())
        self.summary["Most Expensive"] = max(max_ctvalue, max_price)
        min_ctvalue = min(i for i in self.dframe.CTValue.unique() if i > 0)
        min_price = min(i for i in self.dframe.Price.unique() if i > 0)
        self.summary["Least Expensive"] = min(min_ctvalue, min_price)
        VINTAGES = sorted(filter(lambda x: x not in ["NV"], self.dframe["Vintage"]))
        NV_VINTAGES = sorted(filter(lambda x: x in ["NV"], self.dframe["Vintage"]))

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
        for i in range(len(self.dframe)):
            key: str = str(self.dframe.at[i, "Type"])
            vintage: str = str(self.dframe.at[i, "Vintage"])
            count: int = int(float(str(self.dframe.at[i, "Count"])))
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

        for i in range(len(self.dframe)):
            key: str = str(self.dframe.at[i, "Type"])
            vintage: str = str(self.dframe.at[i, "Vintage"])
            if key not in WINE_STACK_DATA:
                WINE_STACK_DATA[key] = {"NV": 0}
                for year in year_list:
                    WINE_STACK_DATA[key][year] = 0
            WINE_STACK_DATA[key][vintage] += 1
        # print(f"WINE STACK DATA: {WINE_STACK_DATA}")
        # print(f"SUMMARY: {SUMMARY}")

        x1 = year_list_str
        # Ensure numeric numpy arrays (convert any non-int types to int)
        y1 = narray([int(v) for v in WINE_STACK_DATA["Red"].values()])
        y2 = narray([int(v) for v in WINE_STACK_DATA["Rose"].values()])
        y3 = narray([int(v) for v in WINE_STACK_DATA["White"].values()])
        y4 = narray([int(v) for v in WINE_STACK_DATA["Rose - Spa"].values()])
        y5 = narray([int(v) for v in WINE_STACK_DATA["White - Sp"].values()])

        # Reduce the data to just the Columns we care about
        self.dframe = self.dframe[COLUMN_NAMES]
        # print("DFRAME Columns: ", DFRAME.columns)
        # print(DFRAME.to_csv(index=False)

        #return self.dframe
