import time

class Fixture:    
    def __init__(self, year):
        self.year = year
        self.numRounds = 0
        self.teams = []
        self.games = []


class Game:
    def __init__(self, gameID, location, homeTeam, awayTeam):
        self.gameID = gameID
        self.location = location
        self.homeTeam = homeTeam
        self.awayTeam = awayTeam


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