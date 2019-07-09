# read the TSLA stock data from yahoo,
# generate a table with the data from the
# last 30 trading days, and present it in html

import datetime as dt
import pandas as pd
import pandas_datareader.data as pdr
import openpyxl
from openpyxl import Workbook


def finance_script():
    symbol = ''
    wb = Workbook("Finance_Data.xlsx")
    wb.save("Finance_Data.xlsx")
    wb = openpyxl.load_workbook("Finance_Data.xlsx")
    ws = wb.active

    start = dt.datetime.now() - dt.timedelta(days=30)
    end = dt.datetime.now()

    df = pdr.DataReader('TSLA', 'yahoo', str(start), str(end))

    with pd.ExcelWriter("Finance_Data.xlsx") as writer:
        df.to_excel(writer)

    wb = writer

    wb.save()


finance_script()
