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

    def loadGamesData(self, wb):
        # Create a list of completed games that need to be loaded
        gamesToLoad = []
        for game in self.games.values():
            gameID = game.gameID
            gameLoaded = game.gameLoaded
            gamePlayed = game.date + datetime.timedelta(hours=3) < datetime.datetime.now()
            if gameLoaded:
                # Get the game data from the PastGamesData sheet
                found = self.loadPlayerStatsFromPastGamesData(self, wb, gameID)
                if not found:
                    game.gameLoaded = False
                    gamesToLoad.append(game)
            elif gamePlayed:
                # Add it to a list of games that need to be loaded
                gamesToLoad.append(game)
            else:
                # Game hasn't been played yet
                pass
                
        ws.getMatchData(wb, gamesToLoad)
        # games = ws.getMatchData(gamesToLoad)
        # for game in games:
        #     self.addGame(game)

    def saveGames(self, wb):
        sheet = wb.sheets["PastGamesData"]
        
        # Load all the games into a dataframe
        for game in self.games:
            pass

    def loadPlayerStatsFromPastGamesData(self, wb, gameID):
        # TODO: Go to the PastGamesData sheet and get the data
        # Each row should have the gameID and then the player stats
        # Load the whole sheet into a dataframe and filter the df by the gameID
        # Add the filtered dataframe to the game object with that gameID
        sheet = wb.sheets['PastGamesData']
        playerStats_df = pd.DataFrame()
        self.games[gameID].addPlayerStats(playerStats_df)

        found = False

        return found

    @classmethod
    def createFromSpreadsheet(cls, wb):
        sheet = wb.sheets['Fixture']

        year = sheet.range("C3").value
        numRounds = int(sheet.range("C5").value)
        
        return cls(year, numRounds)


class Game:

    playerTeams = {}

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

    def addSpreadsheetRow(self, row):
        self.spreadsheetRow = row

    def addPlayerStats(self, playerStats_df):
        # Add a GameID column
        newColumn = pd.Series([self.gameID]*len(playerStats_df), name="GameID")
        playerStats_df.insert(loc=0, column="GameID", value=newColumn)

        # Add a team column
        newColumn = pd.Series(["Unknown"]*len(playerStats_df), name="Team")
        playerStats_df.insert(loc=2, column="Team", value=newColumn)

        # Add in the team name
        for index, row in playerStats_df.iterrows():
            playerName =  row['Player']
            try: 
                playerStats_df.at[index, "Team"] = Game.playerTeams[playerName]
            except KeyError:
                msg = f"Couldn't find {playerName} in the Players tab.\n"
                msg += "Try to find what his name is in the players tab and add the difference to the Config tab."
                print(msg)
        self.playerStats = playerStats_df

    def addToSpreadsheet(self, wb):
        sheet = wb.sheets['PastGamesData']

        # Convert to a list
        playerStats = self.playerStats.values.tolist()

        # Determine the start cell
        row = max(sheet.range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row, 8)
        startCell = sheet.range(f"A{row}")

        # Determine the end cell
        endCell = startCell.offset(len(playerStats) - 1, len(playerStats[0]) - 1)
        
        # Paste in the PastGamesData sheet
        sheet.range(startCell, endCell).value = playerStats

    def markGameLoaded(self, wb, loaded):
        self.gameLoaded = loaded
        sheet = wb.sheets['Fixture']
        if not self.spreadsheetRow is None:
            sheet.range(f"F{self.spreadsheetRow}").value = loaded

    @classmethod
    def getPlayerTeams(cls, wb):
        sheet = wb.sheets['Players']
        playersRange = sheet.range("A8").expand()
        players = playersRange.value

        for player in players:
            firstName = player[0]
            lastName = player[1]
            team = player[2]
            cls.playerTeams[f"{firstName} {lastName}"] = team

        # Get name differences for afl.com
        sheet = wb.sheets['Config']
        differencesRange = sheet.range("A27").expand()
        differences = differencesRange.value

        for difference in differences:
            team = cls.playerTeams[difference[0]]
            cls.playerTeams[difference[1]] = team

class Player:
    def __init__(self, firstName, lastName, team, number, position, weight, height, birthDate, photoLink, profileLink):
        self.firstName = firstName
        self.lastName = lastName
        self.team = team
        self.number = number
        self.position = position
        self.weight = weight
        self.height = height
        self.birthDate = birthDate
        self.photoLink = photoLink
        self.profileLink = profileLink

    def addToDataframe(self, df):
        row = [self.firstName, self.lastName, self.team, self.number, self.position, self.weight, self.height, self.birthDate, self.photoLink, self.profileLink]
        df.loc[len(df)] = row
        return df

class Team:
    def __init__(self, name, abbreviation):
        self.name = name
        self.abbreviation = abbreviation
        self.players = []

