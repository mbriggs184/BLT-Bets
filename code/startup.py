# Import libararies
import pandas as pd
import xlwings as xw
from datetime import datetime

# Import other functions and classes
from classes.fixture import *
from classes.game import *
from classes.team import *
from webScraping import *

def updateHomeSheet(wb=None):
    if not wb: wb = xw.Book.caller()

    updateHomeLadder(wb)
    updateHomeFixture(wb)



def main(wb):
    if not wb: wb = xw.Book.caller()

    # Spreadsheet first loads

    # Load the player teams dict into the Game class
    Game.getPlayerTeams(wb)

    # Get the season fixture
    fixture = Fixture.createFromSpreadsheet(wb)
    fixture.loadFromSpreadsheet()

    # Populate the fixture with past games and determine if there are any games to load
    fixture.loadGamesData()

    # Update team sheets
    teams = loadTeamSheets(wb)

    # Update Home screen fixture

    # Update Home screen Ladder

    pass


def loadTeamSheets(wb=None):
    if not wb: wb = xw.Book.caller()

    # Get a list of teams
    configSheet = wb.sheets['Config']
    teamsRange = configSheet.range("A4:B21")
    teams_df = configSheet.range(teamsRange).options(pd.DataFrame, index=False, header=False, type=int).value
    teams_list = teams_df.iloc[:, 0].unique().tolist()
    teamsAbr_list = teams_df.iloc[:, 1].unique().tolist()

    # Create a team object for each team
    teams = {}
    for i, team in enumerate(teams_list):
        teams[team] = Team(wb, team, teamsAbr_list[i])

    return teams


def updateHomeLadder(wb=None):
    if not wb: wb = xw.Book.caller()
    sheet = wb.sheets['Home']

    ladder_df = getLadder()

    data = ladder_df.reset_index(drop=True).to_numpy()
    sheet.range("P5").value = data

def updateHomeFixture(wb=None):
    if not wb: wb = xw.Book.caller()
    sheet = wb.sheets['Home']

    fixture = Fixture.createFromSpreadsheet(wb)
    fixture.loadFromSpreadsheet()
    thisWeeksGames_df = fixture.getThisWeeksGames()

    data = thisWeeksGames_df.reset_index(drop=True).to_numpy()
    sheet.range("I5").value = data


if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    updateHomeFixture(wb)