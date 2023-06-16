import datetime
import xlwings as xw
import pandas as pd
import numpy as np

class Fixture:    
    def __init__(self, year, numRounds):
        self.year = year
        self.numRounds = numRounds
        self.teams = []
        self.rounds = {}
        self.df = pd.DataFrame(columns = ['Round', 'Date', 'Home Team', 'Away Team', 'GameID'])

    def addRound(self, roundNumber, games):
        self.rounds[roundNumber] = games
    
    def convertToDf(self):
        for roundNumber, games in self.rounds.items():
            for game in games:
                row = [roundNumber, game.date, game.homeTeam, game.awayTeam, game.gameID]
                self.df.loc[len(self.df)] = row
        self.df = self.df.sort_values(by = ['Round', 'Date', 'GameID'], ascending=True)

    def addToSpreadsheet(self, sheet):
        if self.df.empty: 
            self.convertToDf()

        data = self.df.reset_index(drop=True).to_numpy()
        sheet.range("A8").value = data


class Game:
    def __init__(self, gameID, date, homeTeam, awayTeam):
        self.gameID = gameID
        self.date = date
        self.homeTeam = homeTeam
        self.awayTeam = awayTeam

        self.location = "Unknown"

    def addLocation(self, location):
        self.location = location

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

