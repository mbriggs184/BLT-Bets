import xlwings as xw
import os

def main():
    pass

@xw.func
def test(wb=None):
    if not wb: wb = xw.Book.caller()
    wb.sheets['Home'].range('A1').value = "Hello World"

@xw.func
def importFixture(wb=None, year=None):
    if not wb: wb = xw.Book.caller()
    wb.sheets['Home'].range('A1').value = "Hello World"


if __name__ == '__main__':
    wb = xw.Book('BLT Bets.xlsm')
    test(wb)
    