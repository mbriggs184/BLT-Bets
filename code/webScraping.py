import sys
import time
import datetime
import pandas as pd
import re
import requests
import xlwings as xw
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

from tkinter import *
from tkinter.ttk import Progressbar

from classes.fixture import *
from classes.game import *

def getMatchData(wb, games, progressBar=None):
    """Use webscraping to get the match data

    Args:
        games(list(Object)): A list of matchIds to scrape

    Returns:
        list(Object): contains the games with the match data
    """

    # Set up Chrome options for running in headless mode
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    
    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe', options=chrome_options)
    
    # Loop through the games and get the match data
    count = 0
    for game in games:

        gameID = game.gameID

        # Open the webpage of the match
        url = f"https://www.afl.com.au/afl/matches/{gameID}#player-stats"
        driver.get(url)

        # Wait for the webpage to load
        time.sleep(2)

        # Scrape the player stats
        table = driver.find_element(By.TAG_NAME, "table")
        playerStats_df = pd.DataFrame() 
        playerStats_df = pd.read_html(table.get_attribute("outerHTML"))[0]

        game.addPlayerStats(playerStats_df)
        game.addToSpreadsheet(wb)
        game.markGameLoaded(wb, True)

        # Update the progress bar
        count += 1
        if progressBar is not None:
            percentComplete = count / len(games) * 100
            progressBar['value'] = percentComplete
            progressBar.update()
   
    # Close the chrome webdriver
    driver.close()

    return games

def getLadder():
    # TODO: try to do this without webdriver

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

def getSeasonFixture(wb, year, seasonID, numRounds, progressBar=None):
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

    # TODO: try to do this without webdriver

    # Create fixture object
    fixture = Fixture(wb, year, numRounds)

    # Set up Chrome options for running in headless mode
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    
    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe', options=chrome_options)

    for i in range(1, numRounds + 1):
        print(f"Getting fixture for round {i}")

        # Open the webpage of the week
        url = f"https://www.afl.com.au/fixture?Competition=1&CompSeason={seasonID}&MatchTimezone=MY_TIME&Regions=2&ShowBettingOdds=1&GameWeeks={i}&Teams=2&Venues=3#byround"
        driver.get(url)
        time.sleep(3)

        # Get data from the webpage and process it
        soup = BeautifulSoup(driver.page_source, "html.parser")
        divs = soup.find_all("section", {"class": "match-list"})
        games = processFixture(divs[0], year, i)
        
        # Add the week's games to the fixture object
        fixture.addGames(games)

        # Update the progress bar
        if progressBar is not None:
            percentComplete = i / numRounds * 100
            progressBar['value'] = percentComplete
            progressBar.update()

    # Close the chrome webdriver
    driver.close()
    
    return fixture

def getPlayersInfo(progressBar=None):
    print("Getting players info")

    # Open the afl players webpage
    url = f"https://www.zerohanger.com/afl/players/"
    response = requests.get(url)
    
    # Get all the links to the players
    soup = BeautifulSoup(response.text, "html.parser")
    links = soup.find_all("a")
    playerLinks = []
    regex = r"[^/?]+/\?players"
    for link in links:
        href = link.get("href")
        if href is not None:
            if re.match(regex, href, re.IGNORECASE):
                playerLinks.append(href)
                
    # Loop through each player and get his stats
    players = []
    for link in playerLinks:
        # Open the webpage of the player
        url = f"https://www.zerohanger.com/afl/players/{link}"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        # Get name
        div_tag = soup.find('div', class_='player-profile-logo-right')
        h1_tag = div_tag.find('h1')
        span_tag = h1_tag.find("span")
        firstName = span_tag.get_text(strip=True) or "Unknown"
        lastName = h1_tag.get_text(strip=True)[len(firstName):] or "Unknown"

        # Get team
        element = soup.find('h2', class_='player-profile-club hide-mobile')
        team = element.text or "Unknown"

        # Get Number, Position, Weight, Height, DoB
        div_tag = soup.find('div', class_='player-profile-details')
        tr_tags = div_tag.find_all('tr')

        number = "Unknown"
        position = "Unknown"
        weight = "Unknown"
        height = "Unknown"
        birthDate = "Unknown"

        for tr_tag in tr_tags:
            label = tr_tag.find('td', class_='label-title').text
            value = tr_tag.find('td').find_next_sibling().text

            if label == 'Number':
                number = value
            elif label == 'Position':
                position = value
            elif label == 'Weight':
                weight = value
            elif label == 'Height':
                height = value
            elif label == "Date of Birth":
                birthDate = value

        # Get Photo Link
        div_tag = soup.find('div', class_='player-profile-image')
        img_tag = div_tag.find('img')
        photoLink = img_tag['src'] or "Unknown"

        players.append(Player(firstName, lastName, team, number, position, weight, height, birthDate, photoLink, url))

        # Percentage complete
        if progressBar is not None:
            percentComplete = len(players) / len(playerLinks) * 100
            progressBar['value'] = percentComplete
            progressBar.update()

    return players

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
    pass