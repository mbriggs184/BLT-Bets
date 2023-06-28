import pandas as pd
import xlwings as xw

from classes.fixture import *
from classes.game import *
from classes.team import *

def formatTeamSheet(wb):
    if not wb: wb = xw.Book.caller()

    # Load the fixture
    fixture = Fixture.createFromSpreadsheet(wb)
    fixture.loadFromSpreadsheet()

    # Get a list of teams
    configSheet = wb.sheets['Config']
    teamsRange = configSheet.range("A4:A21")

    for team in teamsRange.value:
        sheet = wb.sheets[team]
        sheet.clear_contents()
        sheet.used_range.clear()

        games = {}
        for game in fixture.games.values():
            if team == game.homeTeam:
                games[game.roundNumber] = game.awayTeam
            if team == game.awayTeam:
                games[game.roundNumber] = game.homeTeam
                
        sheet.range("A1").api.Font.Size = 36
        sheet.range("A2").api.Font.Size = 36
        sheet.range("A2").value = team

        sheet.range("A7").value = "Name"
        sheet.range("B7").value = "#"
        sheet.range("C7").value = "Position"
        sheet.range("D7").value = "Avg Disposals"
        sheet.range("E7").value = "Avg Goals"

        sheet.range("A5:A7").merge()
        sheet.range("B5:B7").merge()
        sheet.range("C5:C7").merge()
        sheet.range("D5:D7").merge()
        sheet.range("E5:E7").merge()

        sheet.range("A5:A7").wrap_text = True
        sheet.range("B5:B7").wrap_text = True
        sheet.range("C5:C7").wrap_text = True
        sheet.range("D5:D7").wrap_text = True
        sheet.range("E5:E7").wrap_text = True

        sheet.range("A7").column_width = 35
        sheet.range("B7").column_width = 3
        sheet.range("C7").column_width = 18.29
        sheet.range("D7").column_width = 9
        sheet.range("E7").column_width = 5

        sheet.range("A1:BD7").api.Font.Bold = True
        sheet.range("A1:BD4").color = (255, 255, 255)
        sheet.range("A5:BD7").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter        
        
        # Freeze panes
        sheet.activate()
        active_window = wb.app.api.ActiveWindow
        active_window.FreezePanes = False
        active_window.SplitColumn = 5
        active_window.SplitRow = 7
        active_window.FreezePanes = True

        # Borders
        rng = sheet.range('A5:E7')
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlInsideVertical).Weight = 3
        
        rng = sheet.range('A8:E100')
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3
        rng.api.Borders(xw.constants.BordersIndex.xlInsideVertical).Weight = 2

        column = 6
        for round, opposition in games.items():        
            sheet.range(5, column).value = f"Round {round}"
            sheet.range(6, column).value = f"VS {opposition}"
            sheet.range(7, column).value = "Disposals"
            sheet.range(7, column + 1).value = "Goals"

            sheet.range(sheet.range(5, column), sheet.range(5, column+1)).merge()
            sheet.range(sheet.range(6, column), sheet.range(6, column+1)).merge()

            sheet.range(1, column).column_width = 10
            sheet.range(1, column + 1).column_width = 10

            # Borders
            rng = sheet.range(sheet.range(5, column), sheet.range(6, column+1))
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3

            rng = sheet.range(sheet.range(7, column), sheet.range(7, column+1))
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlInsideVertical).Weight = 2

            rng = sheet.range(sheet.range(8, column), sheet.range(100, column+1))
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeBottom).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeTop).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeLeft).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlEdgeRight).Weight = 3
            rng.api.Borders(xw.constants.BordersIndex.xlInsideVertical).Weight = 2



            column += 2

        pass

class BordersIndex:
    xlDiagonalDown = 5
    xlDiagonalUp = 6 
    xlEdgeBottom = 9 
    xlEdgeLeft = 7 
    xlEdgeRight = 10
    xlEdgeTop = 8 
    xlInsideHorizontal = 12
    xlInsideVertical = 11 


if __name__ == "__main__":
    wb = xw.Book('BLT Bets.xlsm')
    formatTeamSheet(wb)