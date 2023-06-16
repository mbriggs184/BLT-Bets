import xlwings as xw
import os

from webScraping import *

def main():
    pass

@xw.func
def test(wb=None):
    if not wb: wb = xw.Book.caller()
    wb.sheets['Home'].range('A1').value = "Hello World"

@xw.func
def importFixture(wb=None):
    if not wb: wb = xw.Book.caller()

    sheet = wb.sheets['Fixture']

    lastRow = sheet.cells.last_cell.row
    if lastRow > 7:
        sheet.range(f"A8:E{lastRow}").clear_contents()

    year = int(sheet.range("C3").value)
    seasonID = int(sheet.range("C4").value)
    numRounds = int(sheet.range("C5").value)

    fixture = getSeasonFixture(year, seasonID, numRounds)
    fixture.addToSpreadsheet(sheet=wb.sheets['Fixture'])

if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    importFixture(wb)
    