import pandas as pd
import xlwings as xw


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

        # Add a round column
        newColumn = pd.Series([self.roundNumber]*len(playerStats_df), name="Round")
        playerStats_df.insert(loc=1, column="Round", value=newColumn)

        # Add a team column
        newColumn = pd.Series(["Unknown"]*len(playerStats_df), name="Team")
        playerStats_df.insert(loc=3, column="Team", value=newColumn)

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