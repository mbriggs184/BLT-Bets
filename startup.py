import pandas as pd
from datetime import datetime

from webScraping import *

def updateLadder(wb=None):
    if not wb: wb = xw.Book.caller()
    sheet = wb.sheets['Home']

    ladder_df = getLadder()

    data = ladder_df.reset_index(drop=True).to_numpy()
    sheet.range("P5").value = data

def updateFixture(wb=None):
    if not wb: wb = xw.Book.caller()
    sheet = wb.sheets['Home']

    # Check what round we are on
    pass

def updatePastGamesData(wb=None):
    if not wb: wb = xw.Book.caller()
    
    Game.getPlayerTeams(wb)

    fixture = Fixture.createFromSpreadsheet(wb)
    fixture.loadFromSpreadsheet(wb)
    fixture.loadGamesData(wb)
    # TODO: Update the PastGamesData sheet 
    # TODO: Update the fixture sheet when the game is loaded
    # TODO: Update each teamsheet, maybe in another function?

if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    updatePastGamesData(wb)