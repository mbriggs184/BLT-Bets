import xlwings as xw
import pandas as pd

from classes.player import *
from classes.fixture import *

class Team:
    def __init__(self, wb, name, abbreviation):
        self.wb = wb
        self.name = name
        self.abbreviation = abbreviation
        self.players = self.__loadPlayers()
        self.fixture = {}
        self.games = {}


    def loadFixture(self, fixture):

        for key, value in fixture.games.items():
            print(key)
            print(value)



        pass

    #region Private Methods
    def __loadPlayers(self):
        playersSheet = self.wb.sheets["Players"]

        # Get a dataframe of all players in the team
        columnRange = playersSheet.range("A7:J7").expand('down')
        allplayers_df = playersSheet.range(columnRange).options(pd.DataFrame, index=False, header=True).value
        players_df = allplayers_df[allplayers_df["Team"] == self.name]
        
        # create a Player object for each player in the team and add to the players dictionary
        players = {}
        for index, row in players_df.iterrows():
            firstName = row['First Name']
            lastName = row['Last Name']
            team = row['Team']
            number = int(row['Number'])
            position = row['Position']
            weight = row['Weight']
            height = row['Height']
            DoB = row['DoB']
            photoLink = row['Photo Link']
            profileLink = row['Profile Link']

            players[f'{firstName} {lastName}'] = Player(firstName, lastName, team, number, position, weight, height, DoB, photoLink, profileLink)

        return players
    


    def __loadExistingGames(self):
        games = {}
        return games
    #endregion