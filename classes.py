import datetime
import xlwings as xw
import pandas as pd
import numpy as np

# from webScraping import *
import webScraping as ws

class Fixture:
    def __init__(self, year, numRounds):
        self.year = year
        self.numRounds = numRounds
        self.teams = []
        self.games = {}
        self.df = pd.DataFrame(columns = ['Round', 'Date', 'Home Team', 'Away Team', 'GameID', 'GameLoaded'])
    
    def addGame(self, game):
        self.games[game.gameID] = game

    def addGames(self, games):
        for game in games:
            self.addGame(self, game)

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

    def loadFromSpreadsheet(self, wb):
        fixtureSheet = wb.sheets['Fixture']

        # Get the data into a dataframe
        lastRow = fixtureSheet.used_range.last_cell.row

        fixtureColumns = fixtureSheet.range(f"A7:F7").value
        fixtureValues = fixtureSheet.range(f"A8:F{lastRow}").value

        seasonFixture_df = pd.DataFrame(fixtureValues, columns = fixtureColumns)

        # Get the data into the fixture object
        for i in range(1, self.numRounds+1):
            round_df = seasonFixture_df[seasonFixture_df['Round'] == i]
            round = []
            for _, row in round_df.iterrows():
                roundNumber = int(row['Round'])
                date = row['Date']
                homeTeam = row['Home Team']
                awayTeam = row['Away Team']
                gameID = int(row['GameID'])
                gameLoaded = bool(row['Game Loaded'])
                game = Game(gameID, roundNumber, date, homeTeam, awayTeam, gameLoaded)

                self.addGame(game)

    def loadGamesData(self, wb):
        # Create a list of completed games that need to be loaded
        gamesToLoad = []
        for game in self.games.values():
            gameID = game.gameID
            gameLoaded = game.gameLoaded
            gamePlayed = game.date + datetime.timedelta(hours=3) < datetime.datetime.now()
            if gameLoaded:
                # Get the game data from the PastGamesData sheet
                self.loadPlayerStatsFromPastGamesData(self, wb, gameID)
                pass
            elif gamePlayed:
                # Add it to a list of games that need to be loaded
                gamesToLoad.append(game)
            else:
                # Game hasn't been played yet
                pass
                
        games = ws.getMatchData(gamesToLoad)
        for game in games:
            self.addGame(game)

    def loadPlayerStatsFromPastGamesData(self, wb, gameID):
        # TODO: Go to the PastGamesData sheet and get the data
        # Each row should have the gameID and then the player stats
        # Load the whole sheet into a dataframe and filter the df by the gameID
        # Add the filtered dataframe to the game object with that gameID
        sheet = wb.sheets['PastGamesData']
        playerStats_df = pd.DataFrame()
        self.games[gameID].addPlayerStats(playerStats_df)

    @classmethod
    def createFromSpreadsheet(cls, wb):
        sheet = wb.sheets['Fixture']

        year = sheet.range("C3").value
        numRounds = int(sheet.range("C5").value)
        
        return cls(year, numRounds)


class Game:
    def __init__(self, gameID, roundNumber, date, homeTeam, awayTeam, gameLoaded):
        self.gameID = gameID
        self.roundNumber = roundNumber
        self.date = date
        self.homeTeam = homeTeam
        self.awayTeam = awayTeam

        self.gameLoaded = gameLoaded

        self.location = "Unknown"

    def addLocation(self, location):
        self.location = location

    def addPlayerStats(self, playerStats_df):
        self.playerStats = playerStats_df

class Player:
    def __init__(self, firstName, lastName, birthDate, height, weight, position, team):
        self.firstName = firstName
        self.lastName = lastName
        self.birthDate = birthDate
        # self.age = time.year - birthDate.year
        self.height = height
        self.weight = weight
        self.position = position
        self.team = team


class Team:
    def __init__(self, name, abbreviation):
        self.name = name
        self.abbreviation = abbreviation
        self.players = []

