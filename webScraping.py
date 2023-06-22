import time
import datetime
import pandas as pd
import re
import xlwings as xw
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from classes import *

def getMatchData(games):
    """Use webscraping to get the match data

    Args:
        games(list(Object)): A list of matchIds to scrape

    Returns:
        list(Object): contains the games with the match data
    """
    # 

    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
    
    matchData = {}
    for game in games:

        gameId = game.gameId

        # Open the webpage of the match
        url = f"https://www.afl.com.au/afl/matches/{gameId}#player-stats"
        driver.get(url)

        # Wait for the webpage to load
        time.sleep(3)

        # Scrape the player stats
        table = driver.find_element(By.TAG_NAME, "table")
        playerStats_df = pd.DataFrame() 
        playerStats_df = pd.read_html(table.get_attribute("outerHTML"))

        game.gameLoaded = True
        game.addPlayerStats(playerStats_df)
   
    # Close the chrome webdriver
    driver.close()

    return games

def getLadder():
    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
    
    # Open the afl ladder webpage
    url = f"https://www.espn.com.au/afl/standings"
    driver.get(url)

    # Wait for the webpage to load
    time.sleep(3)

    # Scrape the player stats
    table = driver.find_element(By.TAG_NAME, "table")
    ladder_df = pd.DataFrame() 
    ladder_df = pd.read_html(table.get_attribute("outerHTML"))
    ladder_df = ladder_df[0]

    # Close the chrome webdriver
    driver.close()

    # Add position column
    newColumn = pd.Series([0]*len(ladder_df), name="Position")
    ladder_df.insert(loc=0, column="Position", value=newColumn)

    # Rename Team column
    ladder_df = ladder_df.rename(columns={"Unnamed: 0": "Team"})

    # Parse team name and position
    for index, row in ladder_df.iterrows():
        team = row["Team"]
        
        if team[1].isdigit():
            position = team[:2]
            team = team[2:]
        else:
            position = team[0]
            team = team[1:]

        ladder_df.at[index, "Position"] = int(position)

        count = 0
        while team[-1].isupper():
            team = team[:-1]
            if count > 4:
                raise Exception(f"Could not parse team {row['Team']}")

        ladder_df.at[index, "Team"] = team

    return ladder_df

def getSeasonFixture(year, seasonID, numRounds):
    """Use webscraping to get the game fixture for the given year

    Args:
        year (int): The year of the fixture
        seasonID (int): The code of the year for the URL
        numRounds (int): The number of rounds in the season

    Raises:
        Exception: Couldn't find the game id for a game

    Returns:
        object: The fixture object 
    """  

    # Create fixture object
    fixture = Fixture(year, numRounds)

    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe')

    
    for i in range(1, numRounds + 1):
        print(f"Getting fixture for round {i}")

        # Open the webpage of the week
        url = f"https://www.afl.com.au/fixture?Competition=1&CompSeason={seasonID}&MatchTimezone=MY_TIME&Regions=2&ShowBettingOdds=1&GameWeeks={i}&Teams=2&Venues=3#byround"
        driver.get(url)
        time.sleep(2)

        # Get data from the webpage and process it
        soup = BeautifulSoup(driver.page_source, "html.parser")
        divs = soup.find_all("section", {"class": "match-list"})
        games = processFixture(divs[0], year, i)
        
        # Add the week's games to the fixture object
        fixture.addGames(i, games)

    # Close the chrome webdriver
    driver.close()
    
    return fixture

def processFixture(html_string, year, roundNumber):
    """Processes the information from the HTML of the game fixture for a week

    Args:
        html_string (string): Raw HTML from the AFL website
        year (int): The year of the fixture

    Raises:
        Exception: Couldn't find the game id for a game

    Returns:
        list: A list of game objects for the week
    """
    
    games = []

    daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    pattern = r'\b(?:{})\b'.format('|'.join(daysOfWeek))

    # Iterate over all the HTML tags
    day = datetime.datetime(2000, 1, 1)
    for tag in html_string:
        rowText = tag.text
        if re.search(pattern, rowText, re.IGNORECASE):
            # New day
            day = rowText.split(" ")
            day = datetime.datetime(year, datetime.datetime.strptime(day[3], "%B").month, int(day[4]))
        else:
            # Get the game ID
            index = str(tag).find("data-match-id")
            if index!= -1:
                startIndex = index + len("data-match-id") + 2
                endIndex = startIndex + 4
                gameID = int(str(tag)[startIndex:endIndex])
            else:
                print(f"Could not find the game ID for: {day}")
                game = Game(0000, 0, day, "homeTeam", "awayTeam", False)
                games.append(game)
                break


            # Get the home and away teams
            homeTeam, awayTeam = findHomeAndAwayTeams(rowText)

            game = Game(gameID, roundNumber, day, homeTeam, awayTeam, False)
            games.append(game)

    return games

def findHomeAndAwayTeams(rowText):
    """Find the home and away teams from a list

    Args:
        rowText (list): list to find the teams in

    Returns:
        string: home and away teams
    """
    gameInfo = rowText.split(" ")
    gameInfo = [entry for entry in gameInfo if entry]

    # Remove excess
    for index, word in enumerate(gameInfo):
        if word == "v":
            separatorIndex = index
        elif word in ("Match", "Where"):
            gameInfo = gameInfo[:index]
            break
    
    # Get home team
    homeTeam = " ".join(gameInfo[:separatorIndex])

    # Get away team
    awayTeam = " ".join(gameInfo[separatorIndex + 1:])

    return homeTeam, awayTeam


if __name__ == "__main__":

    # matchIds = [4890, 4891, 4892]
    # matchData = getMatchData(matchIds)

    # for matchId, playerStats in matchData.items():
    #     print(f"\n\n{matchId}\n")
    #     print(playerStats)
    getSeasonFixture()
