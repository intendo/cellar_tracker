# import codecs
import pandas as pd
from datetime import date
from xlsxwriter import Workbook
from itertools import cycle

from cellartracker import cellartracker

data = [
    {
        "iWine": "3528378",
        "WineBarcode": "W3528378_750ml",
        "Quantity": "3",
        "Pending": "0",
        "Location": "Cellar",
        "Bin": "",
        "Size": "750ml",
        "Price": "0",
        "Valuation": "",
        "MyValue": "0",
        "WBValue": "",
        "CTValue": "0",
        "MenuPrice": "0",
        "Currency": "USD",
        "Vintage": "2015",
        "Wine": "Hogue Cellars Cabernet Sauvignon Reserve",
        "Locale": "USA, Washington, Columbia Valley, Wahluke Slope",
        "Country": "USA",
        "Region": "Washington",
        "SubRegion": "Columbia Valley",
        "Appellation": "Wahluke Slope",
        "Producer": "Hogue Cellars",
        "SortProducer": "Hogue Cellars",
        "Type": "Red",
        "Color": "Red",
        "Category": "Dry",
        "Varietal": "Cabernet Sauvignon",
        "MasterVarietal": "Cabernet Sauvignon",
        "Designation": "Reserve",
        "Vineyard": "Unknown",
        "WA": "",
        "WS": "",
        "IWC": "",
        "BH": "",
        "AG": "",
        "WE": "",
        "JR": "",
        "RH": "",
        "JG": "",
        "GV": "",
        "JK": "",
        "LD": "",
        "CW": "",
        "WFW": "",
        "PR": "",
        "SJ": "",
        "WD": "",
        "RR": "",
        "JH": "",
        "MFW": "",
        "WWR": "",
        "IWR": "",
        "CHG": "",
        "TT": "",
        "TWF": "",
        "DR": "",
        "FP": "",
        "JM": "",
        "PG": "",
        "WAL": "",
        "JS": "",
        "CT": "89",
        "MY": "",
        "BeginConsume": "",
        "EndConsume": "",
        "UPC": "087754010950",
    },
    {
        "iWine": "4406096",
        "WineBarcode": "W4406096_750ml",
        "Quantity": "5",
        "Pending": "0",
        "Location": "Cellar",
        "Bin": "",
        "Size": "750ml",
        "Price": "0",
        "Valuation": "",
        "MyValue": "0",
        "WBValue": "",
        "CTValue": "0",
        "MenuPrice": "0",
        "Currency": "USD",
        "Vintage": "2017",
        "Wine": "Hogue Cellars Cabernet Sauvignon Reserve",
        "Locale": "USA, Washington, Columbia Valley, Wahluke Slope",
        "Country": "USA",
        "Region": "Washington",
        "SubRegion": "Columbia Valley",
        "Appellation": "Wahluke Slope",
        "Producer": "Hogue Cellars",
        "SortProducer": "Hogue Cellars",
        "Type": "Red",
        "Color": "Red",
        "Category": "Dry",
        "Varietal": "Cabernet Sauvignon",
        "MasterVarietal": "Cabernet Sauvignon",
        "Designation": "Reserve",
        "Vineyard": "Unknown",
        "WA": "",
        "WS": "",
        "IWC": "",
        "BH": "",
        "AG": "",
        "WE": "",
        "JR": "",
        "RH": "",
        "JG": "",
        "GV": "",
        "JK": "",
        "LD": "",
        "CW": "",
        "WFW": "",
        "PR": "",
        "SJ": "",
        "WD": "",
        "RR": "",
        "JH": "",
        "MFW": "",
        "WWR": "",
        "IWR": "",
        "CHG": "",
        "TT": "",
        "TWF": "",
        "DR": "",
        "FP": "",
        "JM": "",
        "PG": "",
        "WAL": "",
        "JS": "",
        "CT": "",
        "MY": "",
        "BeginConsume": "",
        "EndConsume": "",
        "UPC": "087754010950",
    },
    {
        "iWine": "2926809",
        "WineBarcode": "W2926809_750ml",
        "Quantity": "25",
        "Pending": "0",
        "Location": "Cellar",
        "Bin": "",
        "Size": "750ml",
        "Price": "0",
        "Valuation": "24",
        "MyValue": "0",
        "WBValue": "",
        "CTValue": "24",
        "MenuPrice": "0",
        "Currency": "USD",
        "Vintage": "2015",
        "Wine": "Hogue Cellars Cabernet Sauvignon Terroir",
        "Locale": "USA, Washington, Columbia Valley, Red Mountain",
        "Country": "USA",
        "Region": "Washington",
        "SubRegion": "Columbia Valley",
        "Appellation": "Red Mountain",
        "Producer": "Hogue Cellars",
        "SortProducer": "Hogue Cellars",
        "Type": "Red",
        "Color": "Red",
        "Category": "Dry",
        "Varietal": "Cabernet Sauvignon",
        "MasterVarietal": "Cabernet Sauvignon",
        "Designation": "Terroir",
        "Vineyard": "Unknown",
        "WA": "",
        "WS": "",
        "IWC": "",
        "BH": "",
        "AG": "",
        "WE": "",
        "JR": "",
        "RH": "",
        "JG": "",
        "GV": "",
        "JK": "",
        "LD": "",
        "CW": "",
        "WFW": "",
        "PR": "",
        "SJ": "",
        "WD": "",
        "RR": "",
        "JH": "",
        "MFW": "",
        "WWR": "",
        "IWR": "",
        "CHG": "",
        "TT": "",
        "TWF": "",
        "DR": "",
        "FP": "",
        "JM": "",
        "PG": "",
        "WAL": "",
        "JS": "",
        "CT": "",
        "MY": "",
        "BeginConsume": "",
        "EndConsume": "",
        "UPC": "",
    },
]
# client = cellartracker.CellarTracker(
#    username="lcurtis62", password="Way2Go90-"
# )
# client.get_list()  # Return List
# client.get_inventory()  # Return Inventory
# client.get_notes()  # Return Notes
# client.get_private_notes()  # Return PrivateNotes
# client.get_purchase()  # Return Purchase
# client.get_pending()  # Return Pending
# client.get_consumed()  # Return Consumed
# client.get_availability()  # Return Availability
# client.get_tag()  # Return Tag
# client.get_pro_review()  # Return ProReview
# client.get_bottles()  # Return Bottles
# client.get_food_tag()  # Return FoodTag


current_year = date.today().year

ENCODE = "utf-8-sig"
HEADERS = [
    "Count",
    "Year",
    "Best By",
    "Winery - Varietal",
    "Type",
    "Price",
    "Value",
    "Inventory",
]

# df = pd.read_csv("./CellarData.csv", encoding="utf-8-sig")
# df = pd.read_csv("./CellarData.csv", encoding=ENCODE)
df = pd.DataFrame.from_records(data)
# df = df.drop(
#    columns=[
#        "iWine",
#        "iWineBottlesEquiv",
#        "iWineAverageVintage",
#        "EndConsumeBottles",
#    ],
# )
df.sort_values(
    ["TypeiWine", "VintageiWine", "WineiWine"],
    ascending=[False, True, True],
    inplace=True,
)
df["EndConsume"] = df["EndConsume"].apply(lambda x: "" if (x == 9999) else x)
df["EndConsume"] = df["EndConsume"].apply(
    lambda x: "*" * (current_year - x) + str(x) + "*" * (current_year - x)
    if (x != "" and x < current_year)
    else x
)
df["iWineValue"] = df["iWineValue"].apply(lambda x: round(x))
df["iWineValue"] = df["iWineValue"].apply(lambda x: "" if (x == 0) else x)
df["iWinePrice"] = df["iWinePrice"].apply(lambda x: "" if (x == 0) else x)
df["VintageiWine"] = df["VintageiWine"].apply(
    lambda x: "NV" if (x == 1001) else x
)
df["Inventory"] = df["iWineBottles"].apply(lambda x: "|_| " * min(x, 12))
df["WineiWine"] = df["WineiWine"].apply(lambda x: x.replace("\ufffd", "e"))
df["WineiWine"] = df["WineiWine"].apply(lambda x: x[:50])
df["TypeiWine"] = df["TypeiWine"].apply(lambda x: x.replace("\ufffd", "e"))
df["TypeiWine"] = df["TypeiWine"].apply(lambda x: x[:10])

df = df.rename(
    columns={
        "iWineBottles": "Count",
        "iWinePrice": "Price",
        "iWineValue": "Value",
        "EndConsume": "Best By",
        "VintageiWine": "Year",
        "WineiWine": "Winery - Varietal",
        "TypeiWine": "Type",
    }
)

df = df[
    [
        "Count",
        "Year",
        "Best By",
        "Winery - Varietal",
        "Type",
        "Price",
        "Value",
        "Inventory",
    ]
]


with Workbook("hello.xlsx") as workbook:
    worksheet = workbook.add_worksheet("Cellar Tracker")
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.center_horizontally()
    worksheet.center_vertically()
    worksheet.set_margins(0.5, 0.5, 0.5, 0.5)

    data_header = workbook.add_format({"bg_color": "#70ad47"})
    data_format1 = workbook.add_format({"bg_color": "#FFFFFF"})
    data_format1.set_shrink()
    data_format2 = workbook.add_format({"bg_color": "#E2EFDA"})
    data_format2.set_shrink()
    formats = cycle([data_format1, data_format2])

    worksheet.set_row(0, cell_format=data_header)
    col_widths = [0] * len(HEADERS)
    for col, header in enumerate(HEADERS):
        print(f"column {col} {header}")
        worksheet.write(0, col, header)
        col_widths[col] = len(header)

    worksheet.autofilter(0, 0, 0, len(HEADERS) - 1)

    for row, data in enumerate(df.itertuples()):
        row = row + 1
        data_format = next(formats)
        worksheet.set_row(row, cell_format=data_format)
        for col, value in enumerate(data[1:]):
            worksheet.write(row, col, value)
            if isinstance(value, str):
                col_widths[col] = max(len(value), col_widths[col])

    for col, header in enumerate(HEADERS):
        print(f"column {col} {col_widths[col]}")
        worksheet.set_column(col, col, col_widths[col])
