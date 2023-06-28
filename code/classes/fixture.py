import datetime
import xlwings as xw
import pandas as pd
import numpy as np
import threading
import sys

# from webScraping import *
sys.path.append('../')
from webScraping import *
from progressBar import *

from classes.game import *

class Fixture:

    # TODO: Make some of the methods private

    def __init__(self, wb, year, numRounds):
        self.wb = wb
        self.year = year
        self.numRounds = numRounds
        self.games = {}
        self.df = pd.DataFrame(columns = ['Round', 'Date', 'Home Team', 'Away Team', 'GameID', 'GameLoaded'])
    
    def getThisWeeksGames(self):
        
        df = pd.DataFrame(columns = ['Day', 'Date', 'Home Team', 'Away Team', 'Generate Report'])

        rounds = {}

        currentRound = 0
        lastGameOfRound = "1/1/2000"

        for game in self.games.values():
            if game.roundNumber != currentRound:
                lastGameOfRound = game.date
                currentRound = game.roundNumber
            lastGameOfRound = game.date if game.date > lastGameOfRound else lastGameOfRound
            rounds[currentRound] = lastGameOfRound

        # Get the current round
        for round, lastGame in rounds.items():
            roundFinished = lastGame + datetime.timedelta(days=1) < datetime.datetime.now()
            if not roundFinished:
                currentRound = round
                break
        
        # Add this week's games to the dataframe
        for game in self.games.values():
            if game.roundNumber == currentRound:
                day = game.date.strftime('%A')
                date = game.date
                homeTeam = game.homeTeam
                awayTeam = game.awayTeam

                row = [day, date, homeTeam, awayTeam, "Generate Report"]
                df.loc[len(df)] = row

        return df


    def addGame(self, game):
        self.games[game.gameID] = game

    def addGames(self, games):
        for game in games:
            self.addGame(game)

    def convertToDf(self):
        for game in self.games.values():
            row = [game.roundNumber, game.date, game.homeTeam, game.awayTeam, game.gameID, game.gameLoaded]
            self.df.loc[len(self.df)] = row
        self.df = self.df.sort_values(by = ['Round', 'Date', 'GameID'], ascending=True)

    def addToSpreadsheet(self, sheet):
        if self.df.empty: 
            self.convertToDf()

        data = self.df.reset_index(drop=True).to_numpy()
        sheet.range("A8").value = data

    def loadFromSpreadsheet(self):
        fixtureSheet = self.wb.sheets['Fixture']

        # Get the data into a dataframe
        lastRow = fixtureSheet.used_range.last_cell.row

        fixtureColumns = fixtureSheet.range(f"A7:F7").value
        fixtureValues = fixtureSheet.range(f"A8:F{lastRow}").value

        seasonFixture_df = pd.DataFrame(fixtureValues, columns = fixtureColumns)

        # Get the data into the fixture object
        spreadsheetRow = 8
        for i in range(1, self.numRounds+1):
            round_df = seasonFixture_df[seasonFixture_df['Round'] == i]
            for _, row in round_df.iterrows():
                roundNumber = int(row['Round'])
                date = row['Date']
                homeTeam = row['Home Team']
                awayTeam = row['Away Team']
                gameID = int(row['GameID'])
                gameLoaded = bool(row['Game Loaded'])
                game = Game(gameID, roundNumber, date, homeTeam, awayTeam, gameLoaded)
                
                # Keep track of the row in the spreadsheet
                game.addSpreadsheetRow(spreadsheetRow)
                spreadsheetRow += 1

                self.addGame(game)

    def loadGamesData(self):
        # Create a list of completed games that need to be loaded
        gamesToLoad = []
        for game in self.games.values():
            gameID = game.gameID
            gameLoaded = game.gameLoaded
            gamePlayed = game.date + datetime.timedelta(hours=3) < datetime.datetime.now()
            if gameLoaded:
                # Find the game in PastGamesData
                found = self.__findGameInPastGamesData(gameID)
                if not found:
                    game.markGameLoaded(self.wb, False)
                    gamesToLoad.append(game)
            elif gamePlayed:
                # Add it to a list of games that need to be loaded
                gamesToLoad.append(game)
            else:
                # Game hasn't been played yet
                pass
                
        # runFunctionWithProgressBar(ws.getMatchData, *[self.wb, gamesToLoad])
        getMatchData(self.wb, gamesToLoad)
        # games = ws.getMatchData(gamesToLoad)
        # for game in games:
        #     self.addGame(game)

    def __findGameInPastGamesData(self, gameID):
        if not hasattr(self, 'loadedGames'):
            sheet = self.wb.sheets['PastGamesData']
            columnRange = sheet.range("A8").expand('down')
            df = sheet.range(columnRange).options(pd.DataFrame, index=False, header=False, type=int).value
            self.loadedGames = df.iloc[:, 0].unique().tolist()
            self.loadedGames = list(map(int, self.loadedGames))
        if gameID in self.loadedGames:
            return True
        return False


    def loadPlayerStatsFromPastGamesData(self, gameID):
        # TODO: Go to the PastGamesData sheet and get the data
        # Each row should have the gameID and then the player stats
        # Load the whole sheet into a dataframe and filter the df by the gameID
        # Add the filtered dataframe to the game object with that gameID
        sheet = self.wb.sheets['PastGamesData']
        playerStats_df = pd.DataFrame()
        self.games[gameID].addPlayerStats(playerStats_df)

        found = False

        return found

    @classmethod
    def createFromSpreadsheet(cls, wb):
        sheet = wb.sheets['Fixture']

        year = sheet.range("C3").value
        numRounds = int(sheet.range("C5").value)
        
        return cls(wb, year, numRounds)