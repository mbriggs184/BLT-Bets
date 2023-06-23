import xlwings as xw
import os
import pandas as pd

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading

from webScraping import *
from classes import *
from progressBar import *

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

    # Clear sheet
    lastRow = sheet.cells.last_cell.row
    if lastRow > 7:
        sheet.range(f"A8:F{lastRow}").clear_contents()

    year = int(sheet.range("C3").value)
    seasonID = int(sheet.range("C4").value)
    numRounds = int(sheet.range("C5").value)

    # Get the fixture and add it to the spreadsheet
    fixture = getSeasonFixture(year, seasonID, numRounds)
    fixture.addToSpreadsheet(sheet=wb.sheets['Fixture'])

@xw.func
def importPlayers(wb=None):
    if not wb: wb = xw.Book.caller()

    sheet = wb.sheets['Players']

    # Clear sheet
    pasteRange = sheet.range("A8").expand()
    pasteRange.clear_contents()

    players = runFunctionWithStatusBar(ws.getPlayersInfo, *[])

    # Add all the information to a dataframe
    df = pd.DataFrame(columns = ['First Name','Last Name','Team','Number','Position','Weight','Height','DoB','Photo Link','Profile Link'])
    for player in players:
        df = player.addToDataframe(df)

    # Paste the dataframe to the spreadsheet
    sheet["A8"].options(pd.DataFrame, header=False, index=False, expand='table').value = df


def runFunction(function, progress_bar):
    global result
    result = function(progress_bar)

if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    importPlayers(wb)
    