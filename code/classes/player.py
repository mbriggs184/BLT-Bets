import pandas as pd
import xlwings as xw



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

        self.roundStats = {}

    def addToDataframe(self, df):
        row = [self.firstName, self.lastName, self.team, self.number, self.position, self.weight, self.height, self.birthDate, self.photoLink, self.profileLink]
        df.loc[len(df)] = row
        return df

    def addRoundStats(self, round, stats):
        self.roundStats[round] = stats
        
    #region Private methods

    #endregion




def fixPlayersSheetTeams(wb):
    # TODO: Integrate this into the import players function

    sheet = wb.sheets['Players']

    # Fix the team names
    teamFix = {}
    teamFix['Adelaide'] = 'Adelaide Crows'
    teamFix['Brisbane'] = 'Brisbane Lions'
    teamFix['Geelong'] = 'Geelong Cats'
    teamFix['Gold Coast'] = 'Gold Coast Suns'
    teamFix['GWS'] = 'GWS Giants'
    teamFix['Sydney'] = 'Sydney Swans'
    teamFix['West Coast'] = 'West Coast Eagles'
    teamFix['Bulldogs'] = 'Western Bulldogs'

    teamsRange = sheet.range("C8").expand('down')
    for cell in teamsRange:
        if cell.value in teamFix:
            cell.value = teamFix[cell.value]

def fixPlayersSheetNames(wb):
    # TODO: Integrate this into the import players function

    sheet = wb.sheets['Players']

    # Fix the player names and replace them
    nameFix = {}
    nameFix['Adelaide'] = 'Adelaide Crows'

    nameFix['Tom J. Lynch'] = ['Tom', 'Lynch']
    nameFix['Bradley Close'] = ['Brad', 'Close']
    nameFix['Thomas Barrass'] = ['Tom', 'Barrass']
    nameFix['Thomas Cole'] = ['Tom', 'Cole']
    nameFix['Ashley Johnson'] = ['Ash', 'Johnson']

    namesRange = sheet.range("A8").expand('down')
    for cell in namesRange:
        firstNameCell = cell
        lastNameCell = cell.offset(0, 1)

        firstName = firstNameCell.value
        lastName = lastNameCell.value
        name = f"{firstName} {lastName}"
        if name in nameFix:
            firstNameCell.value = nameFix[name][0]
            lastNameCell.value = nameFix[name][1]


if __name__ == '__main__':
    pass