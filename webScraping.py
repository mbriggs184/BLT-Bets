import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


def getMatchData(matchIds):
    """Use webscraping to get the match data

    Args:
        matchIds (list(int)): A list of matchIds to scrape

    Returns:
        dict: contains the match data for each match
    """
    # 

    # Open the chrome webdriver
    driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\chromedriver.exe')
    
    matchData = {}
    for matchId in matchIds:
        # Open the webpage of the match
        url = f"https://www.afl.com.au/afl/matches/{matchId}#player-stats"
        driver.get(url)

        # Wait for the webpage to load
        time.sleep(3)

        # Scrape the player stats
        table = driver.find_element(By.TAG_NAME, "table")
        playerStats_df = pd.DataFrame() 
        playerStats_df = pd.read_html(table.get_attribute("outerHTML"))

        matchData[matchId] = playerStats_df
   
    # Close the chrome webdriver
    driver.close()

    return matchData

def getGameFixture(year):
    # Use webscraping to get the game fixture for the given year

    pass


if __name__ == "__main__":

    matchIds = [4890, 4891, 4892]
    matchData = getMatchData(matchIds)

    for matchId, playerStats in matchData.items():
        print(f"\n\n{matchId}\n")
        print(playerStats)
    