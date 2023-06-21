import pandas as pd

from webScraping import *

def updateLadder(wb=None):
    if not wb: wb = xw.Book.caller()
    sheet = wb.sheets['Home']

    ladder_df = getLadder()

    data = ladder_df.reset_index(drop=True).to_numpy()
    sheet.range("P5").value = data


if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    updateLadder(wb)